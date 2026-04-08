# AbsenceTool - Payroll Process Platform

Statyczna platforma webowa do obsługi procesów payroll.

## Tool #1: Holiday Balance Tool (web)

Wersja webowa narzędzia z procesami:

- **Process 1** - Stworzenie inputu do Flexi z Current Holiday Report
- **Process 2** - Update master file i utworzenie nowego pliku do pobrania
- **Process 3** - Porównanie sald między source a masterem
- **Process 4** - Porównanie mastera z payslip PDF

### Co ważne

- Działa **bez instalacji Pythona**.
- Działa **bez pip / virtualenv / pakietów lokalnych**.
- To aplikacja front-end (HTML/CSS/JS) uruchamiana w przeglądarce.

## Uruchomienie lokalne (zero-install)

1. Otwórz `index.html` w przeglądarce (Chrome/Edge/Firefox).
2. Załaduj pliki przez UI:
   - source report (`.xlsx/.xls`),
   - master file (`.xlsx`),
   - payslip PDFs (`.pdf`) dla Process 4.
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