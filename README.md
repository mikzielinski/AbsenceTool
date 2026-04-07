# AbsenceTool - Payroll Process Platform

Statyczna platforma webowa do obsługi procesów payroll, gotowa do hostowania na GitHub Pages.

## Co jest zaimplementowane

- Moduł `Holiday comparer - Denmark`
  - porównanie dat świąt z payrollu z oficjalnym kalendarzem DK,
  - zakres lat konfigurowany przez użytkownika,
  - walidacja formatu dat (`YYYY-MM-DD`),
  - raport: daty brakujące i nadmiarowe.
- Szkielet pod kolejne moduły payroll (placeholdery w UI).

## Uruchomienie lokalne

To aplikacja statyczna. Wystarczy otworzyć plik `index.html` w przeglądarce.

## GitHub Pages

Repo zawiera workflow: `.github/workflows/deploy-pages.yml`.

Publikacja uruchamia się automatycznie po pushu do gałęzi `main`.

Kroki jednorazowe w GitHub:

1. Wejdź w `Settings -> Pages`.
2. W sekcji **Build and deployment** ustaw **Source: GitHub Actions**.
3. Zapisz ustawienia.

Po kolejnym pushu do `main` strona będzie dostępna pod adresem:

`https://<twoj-login>.github.io/<nazwa-repo>/`