"""Telegram document-to-Markdown bot. Run locally (polling) or on Render (webhook)."""
from __future__ import annotations

import asyncio
import hashlib
import logging
import os
import re
import tempfile
from pathlib import Path

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import Application, CallbackQueryHandler, CommandHandler, ContextTypes, MessageHandler, filters
from telegram.error import NetworkError, TimedOut

from converter import ConversionError, SUPPORTED, convert_file, estimate_tokens

logger = logging.getLogger(__name__)
MAX_UPLOAD_BYTES = min(int(os.getenv('MAX_UPLOAD_MB', '15')), 20) * 1024 * 1024
PROCESS_LOCK = asyncio.Semaphore(2)

START_TEXT = (
    '📄 بوت تحويل الملفات إلى Markdown\n\n'
    'أرسل ملف PDF أو Word (DOCX) أو PowerPoint (PPTX) أو Excel (XLSX) '
    'أو TXT أو CSV أو HTML أو صورة، وأعيده إليك بصيغة .md.\n\n'
    '✅ التحويل دون استخدام API ذكاء اصطناعي مدفوع\n'
    '✅ OCR عربي/إنجليزي للصور وPDF المصوّر\n'
    '✅ يحتفظ بالنصوص والجداول قدر الإمكان\n\n'
    'الخيارات: /mode لاختيار التنسيق، /ocr on لإعادة قراءة PDF بالـOCR، /help للمساعدة.\n\n'
    'مهم: Markdown لا يعني دائمًا توكنز أقل، والرياضيات والصور تحتاج مراجعة.'
)


def _keyboard(mode: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[
        InlineKeyboardButton(('✅ ' if mode == 'faithful' else '') + 'دقيق', callback_data='mode:faithful'),
        InlineKeyboardButton(('✅ ' if mode == 'compact' else '') + 'موفر', callback_data='mode:compact'),
    ]])


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    mode = context.user_data.get('mode', 'faithful')
    await update.effective_message.reply_text(START_TEXT, reply_markup=_keyboard(mode))


async def mode_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if context.args and context.args[0].lower() in ('compact', 'faithful'):
        context.user_data['mode'] = context.args[0].lower()
    mode = context.user_data.get('mode', 'faithful')
    await update.effective_message.reply_text(
        'اختر «دقيق» للحفاظ على النص المستخرج، أو «موفر» لإزالة الزوائد الشكلية فقط.\n'
        'الوضع الحالي: ' + ('دقيق' if mode == 'faithful' else 'موفر'),
        reply_markup=_keyboard(mode),
    )


async def set_mode(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    mode = (query.data or '').split(':')[-1]
    if mode not in ('faithful', 'compact'):
        return
    context.user_data['mode'] = mode
    await query.edit_message_reply_markup(reply_markup=_keyboard(mode))
    await query.message.reply_text('تم تفعيل وضع ' + ('دقيق ✅' if mode == 'faithful' else 'موفر ✅'))


async def ocr_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if context.args:
        requested = context.args[0].lower()
        if requested in ('on', 'off'):
            context.user_data['force_ocr'] = requested == 'on'
    force = context.user_data.get('force_ocr', False)
    await update.effective_message.reply_text(
        ('OCR إجباري ✅ سيعاد تحليل صفحات PDF حتى لو تحتوي نصًا رقميًا.\n' if force else 'OCR تلقائي ✅ للملفات المصورة.\n')
        + 'التغيير: /ocr on أو /ocr off'
    )


def output_name(source: str, mode: str) -> str:
    stem = Path(source).stem
    stem = re.sub(r'[\\/<>:"|?*\x00-\x1f]', '_', stem).strip(' .')[:62] or 'file'
    return f'{stem}{"_compact" if mode == "compact" else ""}.md'


async def download_with_retries(attachment, destination: Path) -> None:
    """Retry transient Telegram network problems while fetching file metadata/data."""
    for attempt in range(4):
        try:
            remote = await attachment.get_file(
                connect_timeout=25, read_timeout=45, write_timeout=45, pool_timeout=25
            )
            await remote.download_to_drive(
                custom_path=destination,
                connect_timeout=25, read_timeout=90, write_timeout=45, pool_timeout=25
            )
            return
        except (TimedOut, NetworkError) as exc:
            if attempt == 3:
                raise
            logger.warning(
                "Temporary Telegram file-transfer error (%s); retry %s/3",
                type(exc).__name__, attempt + 1
            )
            await asyncio.sleep((2, 4, 8)[attempt])


async def convert_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if not msg or not update.effective_chat or update.effective_chat.type != 'private':
        return
    doc = msg.document
    photo = msg.photo[-1] if msg.photo else None
    if not doc and not photo:
        return
    filename = doc.file_name or 'document' if doc else 'image.jpg'
    suffix = Path(filename).suffix.lower()
    if suffix not in SUPPORTED:
        await msg.reply_text('هذا النوع غير مدعوم حاليًا. الأنواع: PDF DOCX PPTX XLSX TXT MD CSV HTML JPG PNG WEBP.')
        return
    file_size = (doc.file_size if doc else photo.file_size) or 0
    if file_size > MAX_UPLOAD_BYTES:
        await msg.reply_text(f'حجم الملف أكبر من الحد ({MAX_UPLOAD_BYTES // 1024 // 1024} ميجابايت). قسّمه إلى ملفات أصغر.')
        return
    mode = context.user_data.get('mode', 'faithful')
    force_ocr = context.user_data.get('force_ocr', False)
    status = await msg.reply_text('⏳ جاري استخراج المحتوى من الملف...')
    try:
        async with PROCESS_LOCK:
            with tempfile.TemporaryDirectory(prefix='tgmd-') as tmp:
                incoming = Path(tmp) / ('input' + suffix)
                await download_with_retries(doc or photo, incoming)
                # CPU/OCR conversion is synchronous; run off the Telegram event loop.
                result = await asyncio.to_thread(convert_file, incoming, mode, force_ocr)
                outfile = Path(tmp) / output_name(filename, mode)
                outfile.write_text(result.markdown, encoding='utf-8')
                before, method = estimate_tokens(result.raw_markdown)
                after, _ = estimate_tokens(result.markdown)
                delta = before - after
                report = (
                    f'✅ تم التحويل: {filename}\n'
                    f'📄 عدد أحرف الناتج: {len(result.markdown):,}\n'
                    f'🔢 توكنز النص المستخرج ≈ {before:,}\n'
                    f'🔢 توكنز الناتج ≈ {after:,}\n'
                    f'{"⬇️ التوفير التقريبي" if delta >= 0 else "⬆️ الزيادة التقريبية"}: {abs(delta):,}\n'
                    f'ℹ️ التقدير باستعمال {method} وليس فاتورة نموذج محدد.'
                )
                if result.warnings:
                    report += '\n\n' + '\n'.join('⚠️ ' + w for w in result.warnings[:2])
                with outfile.open('rb') as fileobj:
                    await msg.reply_document(document=fileobj, filename=outfile.name, caption=report[:1000])
    except ConversionError as ex:
        await msg.reply_text('❌ ' + str(ex))
    except (TimedOut, NetworkError):
        logger.warning('Telegram transfer failed after four attempts.')
        await msg.reply_text('❌ لم ينجح تنزيل الملف من تيليجرام بسبب تأخر الاتصال، حتى بعد عدة محاولات. جرّب مرة أخرى لاحقًا.')
    except Exception:
        logger.exception('Failed to handle conversion request')
        await msg.reply_text('❌ حدث خطأ غير متوقع أثناء التحويل. جرّب ملفًا أصغر أو غيّر وضع OCR.')
    finally:
        try:
            await status.delete()
        except Exception:
            pass


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logger.error('Telegram update error: %s', type(context.error).__name__)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
    # Avoid logging Telegram API URLs: they contain the BOT_TOKEN.
    logging.getLogger('httpx').setLevel(logging.WARNING)
    logging.getLogger('httpcore').setLevel(logging.WARNING)
    token = os.getenv('BOT_TOKEN', '').strip()
    if not token:
        raise SystemExit('Missing BOT_TOKEN in environment.')
    app = (Application.builder().token(token)
           .connect_timeout(25).read_timeout(45).write_timeout(45).pool_timeout(25)
           .build())
    app.add_handler(CommandHandler(['start', 'help'], start))
    app.add_handler(CommandHandler('mode', mode_command))
    app.add_handler(CommandHandler('ocr', ocr_command))
    app.add_handler(CallbackQueryHandler(set_mode, pattern=r'^mode:(faithful|compact)$'))
    app.add_handler(MessageHandler((filters.Document.ALL | filters.PHOTO) & filters.ChatType.PRIVATE, convert_message))
    app.add_error_handler(error_handler)
    webhook_root = os.getenv('WEBHOOK_BASE_URL', '').strip().rstrip('/')
    if not webhook_root and os.getenv('RENDER_EXTERNAL_HOSTNAME'):
        webhook_root = 'https://' + os.environ['RENDER_EXTERNAL_HOSTNAME']
    if webhook_root:
        # Consistent ASCII URL-safe secret, never printed; configurable salt is optional.
        salt = os.getenv('WEBHOOK_SECRET', '')
        secret = hashlib.sha256((token + ':telegram-webhook:' + salt).encode()).hexdigest()
        port = int(os.getenv('PORT', '10000'))
        app.run_webhook(
            listen='0.0.0.0', port=port, url_path='telegram',
            webhook_url=f'{webhook_root}/telegram', secret_token=secret,
            drop_pending_updates=False,
            bootstrap_retries=5,
        )
    else:
        logger.info('Starting polling mode (local computer)')
        app.run_polling(drop_pending_updates=False)


if __name__ == '__main__':
    main()
