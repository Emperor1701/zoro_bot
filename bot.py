import os
from datetime import datetime, timezone, timedelta
from io import BytesIO

from dateutil.relativedelta import relativedelta
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputFile,
)
from telegram.constants import ParseMode
from telegram.helpers import mention_html
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ChatMemberHandler,
    MessageHandler,
    filters,
)

from psycopg_pool import ConnectionPool


# =========================
# إعدادات أساسية
# =========================
BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
OWNER_ID = int(os.getenv("OWNER_ID", "0").strip())
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()

DAILY_CHECK_HOUR_UTC = int(os.getenv("DAILY_CHECK_HOUR_UTC", "9"))
WARN_STAGES_OWNER = [7, 3, 1]
GROUP_MENTION_CHUNK_SIZE = 30

KING_USERNAME = "@Al_K_i_n_g"

if not BOT_TOKEN or OWNER_ID == 0 or not DATABASE_URL:
    raise SystemExit("Missing env vars. Please set BOT_TOKEN, OWNER_ID, DATABASE_URL.")


# =========================
# اتصال PostgreSQL (Pool)
# =========================
pool = ConnectionPool(conninfo=DATABASE_URL, min_size=1, max_size=5, open=True)


def init_db():
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            CREATE TABLE IF NOT EXISTS members (
                user_id BIGINT PRIMARY KEY,
                username TEXT,
                full_name TEXT,
                joined_at TIMESTAMPTZ NOT NULL,
                expires_at TIMESTAMPTZ NOT NULL,
                warn_stage_sent INTEGER
            );
            """)
            cur.execute("CREATE INDEX IF NOT EXISTS idx_members_expires_at ON members (expires_at);")

            cur.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            );
            """)

            # افتراضيات
            cur.execute("""
            INSERT INTO settings(key, value) VALUES
                ('default_duration_months', '3'),
                ('notify_group_chat_id', ''),
                ('group_notify_enabled', '0'),
                ('group_notify_stages', '7')
            ON CONFLICT (key) DO NOTHING;
            """)
        conn.commit()


def get_setting(key: str, default: str = "") -> str:
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT value FROM settings WHERE key=%s", (key,))
            row = cur.fetchone()
            return row[0] if row and row[0] is not None else default


def set_setting(key: str, value: str):
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            INSERT INTO settings(key, value) VALUES(%s, %s)
            ON CONFLICT (key) DO UPDATE SET value=EXCLUDED.value
            """, (key, value))
        conn.commit()


def get_default_months() -> int:
    try:
        return int(get_setting("default_duration_months", "3"))
    except Exception:
        return 3


def compute_expiry(joined_at: datetime) -> datetime:
    return joined_at + relativedelta(months=+get_default_months())


def upsert_member(user_id: int, username: str, full_name: str, joined_at: datetime):
    expires = compute_expiry(joined_at)
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            INSERT INTO members(user_id, username, full_name, joined_at, expires_at, warn_stage_sent)
            VALUES (%s, %s, %s, %s, %s, NULL)
            ON CONFLICT (user_id) DO UPDATE SET
                username=EXCLUDED.username,
                full_name=EXCLUDED.full_name,
                joined_at=EXCLUDED.joined_at,
                expires_at=EXCLUDED.expires_at,
                warn_stage_sent=NULL
            """, (user_id, username or "", full_name or "", joined_at, expires))
        conn.commit()


def get_counts() -> int:
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) FROM members")
            return int(cur.fetchone()[0])


def fetch_all():
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            SELECT user_id, username, full_name, joined_at, expires_at, warn_stage_sent
            FROM members
            ORDER BY joined_at DESC
            """)
            return cur.fetchall()


def fetch_expired(now: datetime):
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            SELECT user_id, username, full_name, joined_at, expires_at, warn_stage_sent
            FROM members
            WHERE expires_at < %s
            ORDER BY expires_at ASC
            """, (now,))
            return cur.fetchall()


def fetch_expiring_within(now: datetime, days: int):
    end = now + timedelta(days=days)
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("""
            SELECT user_id, username, full_name, joined_at, expires_at, warn_stage_sent
            FROM members
            WHERE expires_at >= %s AND expires_at <= %s
            ORDER BY expires_at ASC
            """, (now, end))
            return cur.fetchall()


def find_member(query: str):
    with pool.connection() as conn:
        with conn.cursor() as cur:
            if query.isdigit():
                cur.execute("""
                SELECT user_id, username, full_name, joined_at, expires_at, warn_stage_sent
                FROM members WHERE user_id=%s
                """, (int(query),))
                return cur.fetchone()

            q = query.lstrip("@")
            cur.execute("""
            SELECT user_id, username, full_name, joined_at, expires_at, warn_stage_sent
            FROM members WHERE lower(username)=lower(%s)
            """, (q,))
            return cur.fetchone()


def extend_member_days(user_id: int, days: int) -> bool:
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT expires_at FROM members WHERE user_id=%s", (user_id,))
            row = cur.fetchone()
            if not row:
                return False
            old_exp = row[0]
            new_exp = old_exp + timedelta(days=days)
            cur.execute("""
            UPDATE members SET expires_at=%s, warn_stage_sent=NULL WHERE user_id=%s
            """, (new_exp, user_id))
        conn.commit()
    return True


def set_warn_stage(user_id: int, stage: int | None):
    with pool.connection() as conn:
        with conn.cursor() as cur:
            cur.execute("UPDATE members SET warn_stage_sent=%s WHERE user_id=%s", (stage, user_id))
        conn.commit()


# =========================
# صلاحيات المالك
# =========================
def is_owner(update: Update) -> bool:
    u = update.effective_user
    return bool(u and u.id == OWNER_ID)


async def owner_only(update: Update) -> bool:
    if is_owner(update):
        return True
    try:
        if update.message:
            await update.message.reply_text("هذا البوت مخصص للمالك فقط.")
        elif update.callback_query:
            await update.callback_query.answer("صلاحيات غير كافية.", show_alert=True)
    except Exception:
        pass
    return False


# =========================
# واجهة الأزرار
# =========================
WAITING_SEARCH = "waiting_search"
WAITING_EXTEND_ID = "waiting_extend_id"

def group_notify_status_text() -> str:
    enabled = get_setting("group_notify_enabled", "0") == "1"
    gid = get_setting("notify_group_chat_id", "").strip()
    stages = get_setting("group_notify_stages", "7").strip() or "7"
    return (
        f"تنبيه الكروب: {'مفعّل' if enabled else 'موقّف'}\n"
        f"الكروب المحفوظ: {gid if gid else '(غير محدد)'}\n"
        f"مراحل التنبيه بالكروب: {stages}"
    )


def menu_text() -> str:
    months = get_setting("default_duration_months", "3")
    return (
        f"لوحة تحكم المالك ✅\n"
        f"مدة الاشتراك الافتراضية: {months} شهر\n"
        f"{group_notify_status_text()}\n\n"
        f"اختر إجراء:"
    )


def main_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📊 عدد المشتركين", callback_data="count")],
        [InlineKeyboardButton("🧾 تصدير Excel (مع تصفية)", callback_data="export_xlsx")],
        [InlineKeyboardButton("🔎 بحث عن مشترك", callback_data="ask_search")],
        [InlineKeyboardButton("➕ تمديد اشتراك (+90 يوم)", callback_data="ask_extend_90")],
        [InlineKeyboardButton("🔔 فحص التنبيه الآن", callback_data="run_warn_now")],
        [InlineKeyboardButton("⏳ تغيير مدة الاشتراك الافتراضية", callback_data="duration_menu")],
        [InlineKeyboardButton("📣 إعدادات تنبيه الكروب", callback_data="group_menu")],
    ])


def duration_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("1 شهر", callback_data="set_duration_1")],
        [InlineKeyboardButton("3 شهور", callback_data="set_duration_3")],
        [InlineKeyboardButton("12 شهر (سنة)", callback_data="set_duration_12")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="back_main")],
    ])


def group_menu():
    enabled = get_setting("group_notify_enabled", "0") == "1"
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ تفعيل" if not enabled else "⛔ إيقاف", callback_data="toggle_group_notify")],
        [InlineKeyboardButton("مراحل: 7 فقط", callback_data="set_group_stages_7")],
        [InlineKeyboardButton("مراحل: 7+3+1", callback_data="set_group_stages_7_3_1")],
        [InlineKeyboardButton("📌 شرح /setgroup", callback_data="how_setgroup")],
        [InlineKeyboardButton("⬅️ رجوع", callback_data="back_main")],
    ])


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await owner_only(update):
        return
    await update.message.reply_text(menu_text(), reply_markup=main_menu())


# =========================
# /setgroup داخل الكروب
# =========================
async def setgroup(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await owner_only(update):
        return

    chat = update.effective_chat
    if chat.type not in ("group", "supergroup"):
        await update.message.reply_text("استخدم الأمر داخل الكروب فقط.")
        return

    set_setting("notify_group_chat_id", str(chat.id))
    await update.message.reply_text("✅ تم حفظ هذا الكروب لإرسال تنبيهات الاشتراك.")

    # تفعيل تلقائي
    if get_setting("group_notify_enabled", "0") != "1":
        set_setting("group_notify_enabled", "1")
        await update.message.reply_text("✅ تم تفعيل تنبيه الكروب تلقائيًا.")


# =========================
# تتبع دخول الأعضاء
# =========================
async def on_chat_member(update: Update, context: ContextTypes.DEFAULT_TYPE):
    result = update.chat_member
    if not result:
        return

    new_status = result.new_chat_member.status
    old_status = result.old_chat_member.status

    joined = (old_status in ("left", "kicked")) and (new_status in ("member", "restricted", "administrator"))
    if not joined:
        return

    user = result.from_user
    joined_at = datetime.now(timezone.utc)

    full_name = " ".join([p for p in [user.first_name, user.last_name] if p]) or ""
    username = user.username or ""
    upsert_member(user.id, username, full_name, joined_at)


# =========================
# Excel Export
# =========================
def autosize(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 45)


def add_sheet(wb: Workbook, title: str, rows):
    ws = wb.create_sheet(title=title)
    headers = ["user_id", "username", "full_name", "joined_at_utc", "expires_at_utc", "warn_stage_sent"]
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    autosize(ws)


def build_xlsx_bytes(all_rows, expiring_rows, expired_rows):
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "All"
    headers = ["user_id", "username", "full_name", "joined_at_utc", "expires_at_utc", "warn_stage_sent"]
    ws0.append(headers)
    for r in all_rows:
        ws0.append(list(r))
    autosize(ws0)

    add_sheet(wb, "Expiring_7_Days", expiring_rows)
    add_sheet(wb, "Expired", expired_rows)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# =========================
# التنبيهات + منشن بالكروب
# =========================
def stage_for_remaining_days(remaining_days: int) -> int | None:
    if remaining_days <= 1:
        return 1
    if remaining_days <= 3:
        return 3
    if remaining_days <= 7:
        return 7
    return None


def parse_group_stages() -> set[int]:
    raw = (get_setting("group_notify_stages", "7") or "7").strip()
    parts = [p.strip() for p in raw.split(",") if p.strip()]
    out = set()
    for p in parts:
        if p.isdigit():
            out.add(int(p))
    return (out & {7, 3, 1}) if out else {7}


async def notify_group_mentions(context: ContextTypes.DEFAULT_TYPE, stage: int, people):
    enabled = get_setting("group_notify_enabled", "0") == "1"
    gid = get_setting("notify_group_chat_id", "").strip()
    if (not enabled) or (not gid):
        return

    allowed_stages = parse_group_stages()
    if stage not in allowed_stages:
        return

    group_id = int(gid)

    mentions = []
    for user_id, full_name in people:
        display = full_name if full_name else str(user_id)
        mentions.append(mention_html(user_id, display))

    title = f"🔔 تنبيه: باقي {stage} أيام على انتهاء الاشتراك"
    subtitle = (
        "إذا كان لديك أي استفسار أو كنت تريد تجديد اشتراكك\n"
        f"يرجى مراسلة الكينغ لتجديد الاشتراك: {KING_USERNAME}\n"
        "وشكرًا."
    )

    for i in range(0, len(mentions), GROUP_MENTION_CHUNK_SIZE):
        chunk = mentions[i:i + GROUP_MENTION_CHUNK_SIZE]
        text = title + "\n\n" + subtitle + "\n\n" + "\n".join(f"• {m}" for m in chunk)
        await context.bot.send_message(
            chat_id=group_id,
            text=text,
            parse_mode=ParseMode.HTML,
            disable_web_page_preview=True,
        )


async def run_warning_check(context: ContextTypes.DEFAULT_TYPE, manual: bool = False):
    now = datetime.now(timezone.utc)
    candidates = fetch_expiring_within(now, 7)

    to_notify_owner = {7: [], 3: [], 1: []}
    to_notify_group = {7: [], 3: [], 1: []}  # stage -> list[(user_id, full_name)]

    for user_id, username, full_name, _joined_at, expires_at, warn_stage_sent in candidates:
        remaining_seconds = (expires_at - now).total_seconds()
        remaining_days = int((remaining_seconds + 86399) // 86400)  # ceil

        stage = stage_for_remaining_days(remaining_days)
        if stage is None:
            continue

        # منع التكرار (إذا أرسلنا مرحلة أقرب أو نفسها)
        if warn_stage_sent is not None and int(warn_stage_sent) <= stage:
            continue

        u = f"@{username}" if username else "(بدون يوزرنيم)"
        exp_iso = expires_at.isoformat()

        to_notify_owner[stage].append((user_id, full_name, u, exp_iso, remaining_days))
        to_notify_group[stage].append((user_id, full_name))

    sent_any = False

    for stage in WARN_STAGES_OWNER:
        items = to_notify_owner[stage]
        if not items:
            continue

        sent_any = True
        lines = [f"🔔 تنبيه قبل {stage} أيام (UTC)", ""]
        for user_id, full_name, u, exp_iso, remaining_days in items[:60]:
            lines.append(f"- {full_name} | {u} | ID: {user_id} | ينتهي: {exp_iso} | باقي: {remaining_days} يوم")
        if len(items) > 60:
            lines.append("")
            lines.append(f"… ويوجد {len(items) - 60} آخرين (صدّر Excel لرؤية الجميع).")

        await context.bot.send_message(chat_id=OWNER_ID, text="\n".join(lines))

        # تنبيه بالكروب حسب الإعدادات
        await notify_group_mentions(context, stage, to_notify_group[stage])

        # تعليم المرحلة كمرسلة
        for user_id, *_ in items:
            set_warn_stage(user_id, stage)

    if manual and not sent_any:
        await context.bot.send_message(chat_id=OWNER_ID, text="✅ لا يوجد تنبيهات الآن.")


async def job_daily_warning(context: ContextTypes.DEFAULT_TYPE):
    await run_warning_check(context, manual=False)


# =========================
# الأزرار والرسائل
# =========================
async def on_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()

    if not await owner_only(update):
        return

    data = q.data

    if data == "back_main":
        await q.message.reply_text(menu_text(), reply_markup=main_menu())
        return

    if data == "count":
        await q.message.reply_text(f"عدد المشتركين المسجّلين: {get_counts()}")
        return

    if data == "export_xlsx":
        now = datetime.now(timezone.utc)
        all_rows = fetch_all()
        expiring = fetch_expiring_within(now, 7)
        expired = fetch_expired(now)
        bio = build_xlsx_bytes(all_rows, expiring, expired)
        bio.name = "subscribers.xlsx"
        await q.message.reply_document(document=InputFile(bio), caption="ملف Excel ✅ (All / Expiring_7_Days / Expired)")
        return

    if data == "ask_search":
        context.user_data[WAITING_SEARCH] = True
        await q.message.reply_text("أرسل رقم الـID أو اليوزرنيم (مثال: 123456 أو @username).")
        return

    if data == "ask_extend_90":
        context.user_data[WAITING_EXTEND_ID] = True
        await q.message.reply_text("أرسل رقم ID للمشترك لتمديده +90 يوم (3 شهور).")
        return

    if data == "run_warn_now":
        await q.message.reply_text("✅ سأفحص التنبيهات الآن…")
        await run_warning_check(context, manual=True)
        return

    if data == "duration_menu":
        await q.message.reply_text("اختر مدة الاشتراك الافتراضية للأعضاء الجدد:", reply_markup=duration_menu())
        return

    if data.startswith("set_duration_"):
        months = data.split("_")[-1]
        if months not in ("1", "3", "12"):
            await q.message.reply_text("قيمة غير صالحة.")
            return
        set_setting("default_duration_months", months)
        await q.message.reply_text(f"✅ تم تعيين المدة الافتراضية إلى: {months} شهر")
        await q.message.reply_text(menu_text(), reply_markup=main_menu())
        return

    if data == "group_menu":
        await q.message.reply_text(group_notify_status_text(), reply_markup=group_menu())
        return

    if data == "toggle_group_notify":
        cur = get_setting("group_notify_enabled", "0")
        newv = "0" if cur == "1" else "1"
        set_setting("group_notify_enabled", newv)
        await q.message.reply_text("✅ تم تحديث حالة تنبيه الكروب.")
        await q.message.reply_text(group_notify_status_text(), reply_markup=group_menu())
        return

    if data == "set_group_stages_7":
        set_setting("group_notify_stages", "7")
        await q.message.reply_text("✅ تم ضبط مراحل تنبيه الكروب إلى: 7 فقط")
        await q.message.reply_text(group_notify_status_text(), reply_markup=group_menu())
        return

    if data == "set_group_stages_7_3_1":
        set_setting("group_notify_stages", "7,3,1")
        await q.message.reply_text("✅ تم ضبط مراحل تنبيه الكروب إلى: 7 + 3 + 1")
        await q.message.reply_text(group_notify_status_text(), reply_markup=group_menu())
        return

    if data == "how_setgroup":
        await q.message.reply_text(
            "لتحديد الكروب الذي تُرسل إليه التنبيهات:\n"
            "1) أضف البوت كـ Admin بالكروب\n"
            "2) داخل الكروب اكتب: /setgroup\n"
            "بعدها فعّل تنبيه الكروب من الأزرار إن لزم."
        )
        return


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not await owner_only(update):
        return

    text = (update.message.text or "").strip()

    if context.user_data.get(WAITING_SEARCH):
        context.user_data[WAITING_SEARCH] = False
        row = find_member(text)
        if not row:
            await update.message.reply_text("لم يتم العثور على هذا المشترك.")
            return

        user_id, username, full_name, joined_at, expires_at, warn_stage = row
        u = f"@{username}" if username else "(لا يوجد)"
        warn_txt = str(warn_stage) if warn_stage is not None else "None"
        await update.message.reply_text(
            "✅ بيانات المشترك:\n"
            f"- ID: {user_id}\n"
            f"- Username: {u}\n"
            f"- Name: {full_name}\n"
            f"- Joined (UTC): {joined_at.isoformat()}\n"
            f"- Expires (UTC): {expires_at.isoformat()}\n"
            f"- Warn stage sent: {warn_txt}"
        )
        return

    if context.user_data.get(WAITING_EXTEND_ID):
        context.user_data[WAITING_EXTEND_ID] = False
        if not text.isdigit():
            await update.message.reply_text("أرسل رقم ID فقط.")
            return
        ok = extend_member_days(int(text), 90)
        if not ok:
            await update.message.reply_text("هذا الـID غير موجود في قاعدة البيانات.")
            return
        await update.message.reply_text("✅ تم التمديد +90 يوم (3 شهور).")
        return

    await update.message.reply_text("اكتب /start لفتح لوحة التحكم.", reply_markup=main_menu())


# =========================
# تشغيل
# =========================
def main():
    init_db()

    app = Application.builder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start, filters=filters.ChatType.PRIVATE))
    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.TEXT & filters.ChatType.PRIVATE, on_text))

    app.add_handler(CommandHandler("setgroup", setgroup))
    app.add_handler(ChatMemberHandler(on_chat_member, ChatMemberHandler.CHAT_MEMBER))

    app.job_queue.run_daily(
        job_daily_warning,
        time=datetime.now(timezone.utc).replace(hour=DAILY_CHECK_HOUR_UTC, minute=0, second=0, microsecond=0).time(),
        name="daily_warning_check",
    )

    app.run_polling(allowed_updates=["message", "callback_query", "chat_member"])


if __name__ == "__main__":
    main()
