# AbsenceTool - Payroll Process Platform

Statyczna platforma webowa do obsługi procesów payroll.

## Tool #1: Holiday Balance Tool (web)

Wersja webowa narzędzia z procesami:

- **Process 1a** - Build "Flexi-absence input" from Current Holiday Report
- **Process 1b** - Compare Holiday Balance tab in Current Report vs master file
- **Process 2** - Compare master file totals vs payslip PDFs

### Co ważne

- Działa **bez instalacji Pythona**.
- Działa **bez pip / virtualenv / pakietów lokalnych**.
- To aplikacja front-end (HTML/CSS/JS) uruchamiana w przeglądarce.

## Uruchomienie lokalne (zero-install)

1. Otwórz `index.html` w przeglądarce (Chrome/Edge/Firefox).
2. Załaduj pliki przez UI:
   - source report (`.xlsx/.xls`),
   - master file (`.xlsx`),
   - payslip PDFs (`.pdf`) dla Process 2.
3. Kliknij odpowiedni przycisk Run.
4. Raporty są pobierane jako pliki `.xlsx` przez przeglądarkę.

Nie trzeba instalować nic w systemie.

## GitHub Pages

Repo zawiera workflow: `.github/workflows/deploy-pages.yml`.

Publikacja uruchamia się automatycznie po pushu do gałęzi `main`.

Kroki jednorazowe w GitHub:

1. Wejdź w `Settings -> Pages`.
2. W sekcji **Build and deployment** ustaw **Source: GitHub Actions**.
3. Zapisz ustawienia.

Po kolejnym pushu do `main` strona będzie dostępna pod adresem:

`https://<twoj-login>.github.io/<nazwa-repo>/`