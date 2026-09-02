# HR Vakansiya Boti

Telegram orqali bo'sh ish o'rinlariga anketa qabul qiluvchi bot.

## Nima qiladi

Foydalanuvchi `/start` bosganda bot ketma-ket so'raydi:
1. Ism
2. Familiya
3. Yosh
4. Telefon raqam (yozib yoki kontakt tugmasi orqali)
5. Lavozim (qaysi vakansiyaga)
6. Ish tajribasi
7. Rezyume (PDF/Word fayl, yoki `/skip`)

So'ng ma'lumotlar ko'rsatiladi, foydalanuvchi `/confirm` bilan tasdiqlaydi.

Tasdiqlangan anketa:
- **SQLite bazaga** (`applications.db`) saqlanadi
- **Admin/HR Telegram chatiga** xabar sifatida yuboriladi (rezyume fayli bilan birga, agar biriktirilgan bo'lsa)

## O'rnatish

1. Kerakli kutubxonalarni o'rnating:
```bash
pip install -r requirements.txt
```

2. BotFather orqali bot yarating va tokenni oling (@BotFather ga `/newbot` yuboring).

3. Admin chat ID sini aniqlang:
   - Agar shaxsiy xabar sifatida qabul qilmoqchi bo'lsangiz — botga o'zingiz `/start` yozing, so'ng @userinfobot orqali o'z ID'ingizni bilib oling.
   - Agar guruhga yubormoqchi bo'lsangiz — botni guruhga qo'shing, guruh ID'sini oling (odatda manfiy son, masalan `-1001234567890`).

4. Environment o'zgaruvchilarini o'rnating:
```bash
export BOT_TOKEN="123456:ABC-DEF..."
export ADMIN_CHAT_ID="123456789"
```

Windows (PowerShell) uchun:
```powershell
$env:BOT_TOKEN="123456:ABC-DEF..."
$env:ADMIN_CHAT_ID="123456789"
```

5. Botni ishga tushiring:
```bash
python bot.py
```

## Muhim eslatmalar

- Agar guruhga xabar yuborilishini istasangiz, botni o'sha guruhga qo'shib, admin qiling va guruh xabarlarini o'chirib qo'ymang (Privacy Mode kerak emas, chunki bot o'zi yozadi).
- Bazadagi ma'lumotlarni ko'rish uchun istalgan SQLite klient (masalan, DB Browser for SQLite) dan `applications.db` faylini oching, yoki:
```bash
sqlite3 applications.db "SELECT * FROM applications;"
```
- Botni doimiy ishlab turishi uchun serverda `systemd`, `pm2`, yoki `screen`/`tmux` yordamida fon rejimida ishga tushiring, yoki Docker konteynerga joylashtiring.
- Ishlab chiqarish (production) muhitida `run_polling()` o'rniga webhook ishlatish tavsiya etiladi (yuqori yuklama bo'lsa).

## Anketa ustunlarini o'zgartirish

`bot.py` faylida:
- Yangi savol qo'shish uchun: yangi `state` qo'shing (masalan `MANZIL`), tegishli `get_manzil` funksiyasini yozing, va `ConversationHandler`dagi `states` lug'atiga qo'shing.
- Bazaga yangi ustun qo'shish uchun `init_db()` va `save_application()` funksiyalarini yangilang.
