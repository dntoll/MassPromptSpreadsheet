# Google Spreadsheet Mass Prompt

Plugin/funktion för Google Spreadsheet som kör AI-prompts mot cellvärden och skriver tillbaka resultatet i cellen.

## Teknisk grund

- **Plattform**: Google Apps Script (container-bound script eller add-on).
- **AI**: Anropar **OpenAI** (Chat Completions). Standardmodell är **gpt-4o-mini**. Du använder din egen API-nyckel från OpenAI.
- **Konfiguration**: API-nyckel sätts via menyn (se **Inmatning av API-nyckel**); lagras i Script Properties.

## Installation av skriptet

1. Öppna det Google Spreadsheet där du vill använda Mass Prompt.
2. Gå till **Tillägg → Apps Script** (svenska) eller **Extensions → Apps Script** (engelska). Det öppnar scriptredigeraren kopplad till just detta kalkylark.
3. Ta bort eventuell tom `function myFunction() {}` i standardfilen (**Kod.gs** eller **Kod.js** på svenska, **Code.gs** på engelska).
4. Kopiera in koden från detta projekt: lägg till eller ersätt **tre skriptfiler** – **Code.gs**, **Template.gs** och **ApiOpenAI.gs** – med innehållet från motsvarande filer i mappen `appsscript/`. Skapa en HTML-fil med namnet **Sidebar** (Plus → HTML-fil) och klistra in innehållet från `appsscript/Sidebar.html`. Alla fyra filerna behövs (Code.gs, Template.gs, ApiOpenAI.gs, Sidebar.html).
5. **(Valfritt, för minimala behörigheter – krävs för att sidebar ska fungera)** Lägg till manifestfilen: i Apps Script, öppna **Projektinställningar** (kugghjulet) och kryssa i **Visa "appsscript.json"-manifestfil i redigeraren**. Skapa eller ersätt filen **appsscript.json** med innehållet från `appsscript/appsscript.json` i detta projekt. Kontrollera att `oauthScopes` innehåller både `spreadsheets.currentonly` och `script.container.ui` (den andra behövs för menyn och sidopanelen). Spara. Om du fortfarande får fel om "permissions not sufficient to call Ui.showSidebar": återkalla godkännandet (se nedan) och godkänn igen.
6. Spara (Ctrl+S). Namnge projektet om du vill (t.ex. "Mass Prompt").
7. Första gången: **Kör en funktion** eller öppna kalkylarket på nytt och godkänn behörigheter om Google visar en säkerhetsdialog (åtkomst till att köra som dig, lagra Script Properties). Om Google visar **"Google har inte verifierat den här appen"**: klicka på **Avancerat** och välj **Fortsätt till [projektnamn] (osäker)**. Scriptet är ditt eget och körs bara i det här kalkylarket; det är normalt att gå vidare vid egen utveckling.
8. Efter installation: i kalkylarket ska menyn **Mass Prompt** (eller motsvarande) finnas i menyraden.

**Om du får "Specified permissions are not sufficient to call Ui.showSidebar":** Kontrollera att `appsscript.json` innehåller båda scope (steg 5). Spara manifestet. Återkalla sedan det gamla godkännandet så att Google frågar igen: gå till [Google-kontots säkerhet](https://myaccount.google.com/permissions), hitta "Kalkylark" eller ditt projekt under "Åtkomst till tredje part" och ta bort åtkomsten. Öppna kalkylarket igen och klicka på **Mass Prompt → Sätt API-nyckel** – då ska en ny behörighetsruta visas; godkänn så inkluderas `script.container.ui` och sidebar fungerar.

Detta är ett **container-bound script** (kopplat till ett specifikt kalkylark). Vill du använda Mass Prompt i ett annat blad måste scriptet kopieras till det bladets Apps Script-projekt.

## Behörigheter

Scriptet är utformat för **minimala behörigheter**. Det begär endast: (1) åtkomst till **det kalkylark du har öppnat** (inte alla dina kalkylark), (2) behörighet att visa **meny och sidopanel** (sidebar) i kalkylarket, och (3) möjlighet att skicka **externa anrop** till OpenAI API (för att köra =PROMPT). Det begär inte åtkomst till Gmail, Drive-filer eller andra Google-tjänster. Om du använder medföljande **appsscript.json** (manifest) begränsas behörigheterna uttryckligen till detta. Om du inte lägger in manifest kan Google visa bredare behörigheter baserat på automatisk detektering.

## Inmatning av API-nyckel

- **Vem delar nyckeln:** API-nyckeln lagras i **Script Properties** (projektnivå) och används av alla som kan redigera dokumentet. Endast scriptet kan läsa eller skriva nyckeln; den är inte tillgänglig för andra tjänster.
- **Var nyckeln aldrig ska vara:** Inte i en cell, inte i scriptkoden. Endast i Script Properties, satt via menyn.

**Så sätter du nyckeln:**

1. I kalkylarket: klicka på menyn **Mass Prompt** (eller det namn implementationen använder).
2. Välj **Sätt API-nyckel** (eller liknande).
3. En sidopanel (sidebar) öppnas med ett dolt fält för API-nyckel.
4. Klistra in din API-nyckel och klicka **Spara**. Nyckeln sparas i Script Properties och används vid anrop till `=PROMPT(...)`. När du sparar (eller raderar) API-nyckel uppdateras alla celler som innehåller `=PROMPT(...)` automatiskt, så du behöver inte redigera formler eller indata för att se det nya resultatet.

Om implementationen har det: använd **Radera API-nyckel** i menyn för att ta bort nyckeln från Script Properties.

**Säkerhet:** Nyckeln lagras på Googles servrar (Script Properties) och används endast av scriptet för validering och API-anrop; scriptet begär inte fler rättigheter än de som beskrivs under **Behörigheter**. Begränsa vilka som har redigeringsrätt till **scriptet** (Tillägg → Apps Script / Extensions → Apps Script) om du vill att färre ska kunna se eller ändra nyckeln. Indata från celler skickas till OpenAI som separat struktur (uppgift + data), och modellen instrueras att behandla data endast som substitutionsvärden – det minskar risken för prompt injection, men ger ingen full garanti.

## Signatur

```
=PROMPT(indata_1; [indata_2; …]; prompt_cell)
```

- **indata_1, indata_2, …**: Celler eller värden som ska användas i prompten. Minst ett indata krävs. Flera indata mappas till placeholders `{0}`, `{1}`, `{2}` osv. i ordning.
- **prompt_cell**: Cell som innehåller prompt-mallen. Sista argumentet är alltid prompt-mallen.

Argument avgränsas med semikolon (;) i svensk och många europeiska locale, med komma (,) i t.ex. engelska (US). Funktionen returnerar antingen ett enda textvärde (en cell) eller flera värden som sprids horisontellt (se **Flera utdata**).

## Placeholders

Placeholders använder **curly brackets** `{}` (vanlig standard i prompt-mallar). Index motsvarar ordningen på indata-argumenten.

- **`{0}`** — ersätts med första indata-argumentet.
- **`{1}`** — ersätts med andra indata-argumentet (om det finns).
- **`{2}`** — tredje, osv.

## Flera utdata (utdata-schema)

Om prompt-mallen innehåller **hakparenteser med kommaseparerade fältnamn**, t.ex. `[firstname,lastname]`, tolkas det som att modellen ska svara med ett strukturerat svar (ett JSON-objekt med dessa nycklar). Resultatet sprids då över flera celler (en rad horisontellt), en cell per fält.

- **Syntax:** `[fält1,fält2,...]` – första förekomsten av `[ ... ]` i prompten parsas; fältnamnen trimmas.
- **Exempel:** `Extract [firstname,lastname] from {0}` → modellen ombes att returnera JSON med nycklarna `firstname` och `lastname`; första cellen får förnamn, nästa cell efternamn.
- Utan utdata-schema i prompten gäller som tidigare ett enda textvärde per anrop.
- Vid ogiltigt eller icke-JSON-svar från modellen visas `ERROR: Ogiltigt strukturerat svar` i formelcellen.

## Användning

**Exempel med ett indata:**

- **A3** innehåller: `Hello my name is Henrik`
- **B1** innehåller: `If the {0} contains a name then output only the name as output`
- **B3** innehåller: `=PROMPT(A3;B1)` → resultat: `Henrik`

**Exempel med flera indata:**

- **A3** innehåller: `Henrik`
- **B3** innehåller: `Stockholm`
- **C1** innehåller: `The person {0} lives in {1}.`
- **C3** innehåller: `=PROMPT(A3;B3;C1)` → resultat: `The person Henrik lives in Stockholm.`

**Exempel med flera utdata (spill till höger):**

- **A3** innehåller: `John Doe`
- **B1** innehåller: `Extract [firstname,lastname] from {0}`
- **B3** innehåller: `=PROMPT(A3;B1)` → första cellen (B3) får t.ex. `John`, nästa cell (C3) får t.ex. `Doe`. Resultatet sprids horisontellt.

Se **Inmatning av API-nyckel** om du ännu inte satt nyckel.

## Fel och begränsningar

- Vid fel (nätverksfel, ogiltig API-nyckel, rate limit m.m.) returnerar funktionen ett felmeddelande i cellen, t.ex. `ERROR: <kort beskrivning>`. Implementeringen kan komplettera med loggning.
- Google Apps Script har begränsningar för anpassade funktioner (t.ex. max körtid ~30 s). Vid många rader rekommenderas att inte köra hundratals `=PROMPT`-anrop samtidigt; batch eller meny/trigger kan användas i en framtida version.
