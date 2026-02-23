# Test-CSV för PROMPT

Filen `spreadsheet-prompt.csv` i denna katalog innehåller rader som testar `=PROMPT(...)` på olika sätt.

## Importera i Google Sheets

1. Skapa ett nytt kalkylark eller öppna det där Mass Prompt är installerat.
2. **Fil → Importera → Ladda upp** (eller klistra in): välj `spreadsheet-prompt.csv` från mappen `test/`.
3. Vid import: välj att ersätta eller infoga på aktuell sida; avgränsare **komma**.
4. Sätt API-nyckel via **Mass Prompt → Sätt API-nyckel** om du inte redan gjort det.

## Innehåll

- **Kolumn A:** Beskrivning av testfallet.
- **B–D:** Indata 1–3 (värden som används i prompten).
- **E:** Prompt-mall (med `{0}`, `{1}` osv., eventuellt `[fält1,fält2]` för flera utdata).
- **F:** Formel som anropar PROMPT med cellerna på samma rad.

Rader som testar fel kräver att du själv triggar dem (t.ex. radera API-nyckeln för raden "Saknad API-nyckel").

## Testade fall

| Typ | Beskrivning |
|-----|-------------|
| Fel | För få argument (endast prompt) → `ERROR: Minst ett indata och en prompt krävs` |
| Enkel utdata | 1, 2 respektive 3 indata med enkel prompt |
| Flera utdata | Prompt med `[firstname,lastname]` eller `[country,city]` → spill till höger |
| Ett fält i [] | `[city]` → en extra kolumn (spill) |
| Tom indata | Prompt med tomt `{0}` |
| Fel | Saknad API-nyckel → `ERROR: Sätt API-nyckel via menyn...` (radera nyckel och ladda om) |
| Fel | Multi-output med icke-JSON-svar → `ERROR: Ogiltigt strukturerat svar` (beror på modellens svar) |
