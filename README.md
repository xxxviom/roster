# Розподіл по секторах ACC

Застосунок розподіляє людей по секторах і сесіях з графіка Excel (.xlsx). Працює офлайн, дані нікуди не надсилаються.

- `docs/` — веб-застосунок (PWA). Публікується через GitHub Pages: **Settings → Pages → Branch `main`, папка `/docs`**.
  Адреса: https://xxxviom.github.io/roster/
- `apk/` — оболонка Capacitor для Android. APK збирає workflow `.github/workflows/apk.yml`
  після змін у `docs/` або `apk/` і викладає файл `roster-acc.apk` у **Releases → «Сектори ACC (APK)»**.
  Підпис debug-ключем: перед установкою треба дозволити встановлення з невідомих джерел.
