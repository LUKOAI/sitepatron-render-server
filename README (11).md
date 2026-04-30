# Apps Script — backup kodu Google Apps Script

Ten folder zawiera backup 1:1 kodu z projektu Apps Script **SitePatron Email Automation** (`script.google.com/u/0/home/projects/1KNmyFHDHKvqyQp2weVj6dfWzqAvTh3UE_IuOJIQcSstuUtsC2eQl3Y9m`).

Apps Script nie ma natywnego git-a — ten folder służy jako historia zmian.

## Pliki w tym folderze

Nazwy są identyczne jak w panelu Files w Apps Script editor:

| Plik | Linie | Co robi |
|---|---|---|
| `Sitepatron emailautomation.gs` | 636 | Wysyłka maili: kolejka, batch, checkbox-y, menu główne, mapowanie Country→Language |
| `pitch_addon.gs` | 717 | Generowanie PDF deck przez Render Service na Railway, tworzenie draftów Gmail z PDF jako załącznikiem |
| `buildDeckTemplate.gs` | 408 | Jednorazowa budowa template Google Slides (15 slajdów z placeholderami) |

## Stan obecny — data: 2026-04-30

**Wszystkie 3 pliki: dosłowna kopia tego co działa w Apps Script. ZERO modyfikacji.**

To jest punkt wyjścia. Jeśli kiedykolwiek coś pójdzie źle przy aktualizacji skryptów, zawsze można wrócić do tego stanu.

## Procedura aktualizacji w przyszłości

1. Każda przyszła zmiana = osobny commit z jasnym opisem (np. `Sitepatron emailautomation.gs: nowe szablony COLDPITCH/PITCH/LANDING1/2/PURCHASE`)
2. W Apps Script editor: Ctrl+A → Delete → wklej nową wersję pliku → Save
3. Test
4. Jeśli coś się popsuje: `git checkout poprzedni_commit -- "apps-script/Sitepatron emailautomation.gs"` i wracasz do poprzedniej wersji

## Zewnętrzne zależności (nie w tym repo)

- **Google Sheet** `SitePatron_Management` — `1c9nZWirDH1izHeDFASpVUARVshWAwi94PfprRH7vPtk`
- **Render Service** — Python na Railway, kod jest w głównym folderze tego samego repo `LUKOAI/sitepatron-render-server`
- **Shared Drive** dla generowanych PDF — `0AORrQ2XLr8xcUk9PVA`
- **Slides Template** — `1Z3HRS4dqCetfXY4ow7bRJHq4uWNwT7K63i6RwA6NCSo`
