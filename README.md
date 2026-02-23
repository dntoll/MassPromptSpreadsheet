# Google Spreadsheet Mass Prompt

Plugin/function for Google Spreadsheet that runs AI prompts against cell values and writes the result back into the cell.

**Example:** Cell A1 contains `Hello my name is Henrik`. In B1: `=PROMPT(A1; "Extract only the name from: {0}")` → result is **Henrik**. With the `[firstname,lastname]` syntax in the prompt you get multiple outputs in separate cells, e.g. `=PROMPT(A1; "Extract [firstname,lastname] from {0}")` → **John** and **Doe** in two cells.

## Technical overview

- **Platform**: Google Apps Script (container-bound script or add-on).
- **AI**: Calls **OpenAI** (Chat Completions). Default model is **gpt-4o-mini**. You use your own OpenAI API key.
- **Configuration**: API key is set via the menu (see **API key setup**); stored in Script Properties.

## Installation

1. Open the Google Spreadsheet where you want to use Mass Prompt.
2. Go to **Extensions → Apps Script**. This opens the script editor bound to that spreadsheet.
3. Remove any empty `function myFunction() {}` in the default file (**Code.gs**).
4. Copy in the code from this project: add or replace **three script files** – **Code.gs**, **Template.gs**, and **ApiOpenAI.gs** – with the contents of the corresponding files in the `appsscript/` folder. Create an HTML file named **Sidebar** (Plus → HTML file) and paste the contents of `appsscript/Sidebar.html`. All four files are required (Code.gs, Template.gs, ApiOpenAI.gs, Sidebar.html).
5. **(Optional, for minimal permissions – required for sidebar)** Add the manifest: in Apps Script, open **Project settings** (gear icon) and check **Show "appsscript.json" manifest file in editor**. Create or replace the file **appsscript.json** with the contents of `appsscript/appsscript.json` in this project. Ensure `oauthScopes` includes both `spreadsheets.currentonly` and `script.container.ui` (the latter is needed for the menu and sidebar). Save. If you still get "permissions not sufficient to call Ui.showSidebar": revoke the approval (see below) and approve again.
6. Save (Ctrl+S). Name the project if you like (e.g. "Mass Prompt").
7. First time: **Run a function** or reopen the spreadsheet and approve permissions when Google shows a security dialog (run as you, store Script Properties). If Google shows **"Google hasn't verified this app"**: click **Advanced** and choose **Go to [project name] (unsafe)**. The script is your own and runs only in this spreadsheet; it's normal to proceed for personal use.
8. After installation, the **Mass Prompt** menu (or equivalent) should appear in the spreadsheet menu bar.

**If you get "Specified permissions are not sufficient to call Ui.showSidebar":** Check that `appsscript.json` contains both scopes (step 5). Save the manifest. Then revoke the previous approval so Google asks again: go to [Google account security](https://myaccount.google.com/permissions), find "Spreadsheet" or your project under "Third-party access" and remove access. Reopen the spreadsheet and click **Mass Prompt → Set API key** – a new permission prompt should appear; approve so `script.container.ui` is included and the sidebar works.

This is a **container-bound script** (tied to one spreadsheet). To use Mass Prompt in another sheet, copy the script into that sheet’s Apps Script project.

## Permissions

The script is designed for **minimal permissions**. It requests only: (1) access to **the spreadsheet you have open** (not all your spreadsheets), (2) permission to show the **menu and sidebar** in that spreadsheet, and (3) the ability to make **external requests** to the OpenAI API (to run =PROMPT). It does not request access to Gmail, Drive files, or other Google services. With the provided **appsscript.json** (manifest), permissions are explicitly limited to this. Without the manifest, Google may show broader permissions based on auto-detection.

## API key setup

- **Who shares the key:** The API key is stored in **Script Properties** (project level) and used by everyone who can edit the document. Only the script can read or write the key; it is not available to other services.
- **Where the key must never be:** Not in a cell, not in the script code. Only in Script Properties, set via the menu.

**How to set the key:**

1. In the spreadsheet: click the **Mass Prompt** menu (or whatever name the implementation uses).
2. Choose **Set API key** (or similar).
3. A sidebar opens with a hidden field for the API key.
4. Paste your API key and click **Save**. The key is stored in Script Properties and used for `=PROMPT(...)` calls. When you save (or clear) the API key, all cells containing `=PROMPT(...)` are refreshed automatically, so you don’t need to edit formulas or inputs to see the new result.

If the implementation provides it: use **Clear API key** in the menu to remove the key from Script Properties.

**Security:** The key is stored on Google’s servers (Script Properties) and used only by the script for validation and API calls; the script does not request more rights than described under **Permissions**. Restrict who can edit the **script** (Extensions → Apps Script) if you want fewer people to see or change the key. Input from cells is sent to OpenAI as a separate structure (task + data), and the model is instructed to treat data only as substitution values – that reduces prompt injection risk but does not guarantee it.

## Signature

```
=PROMPT(input_1; [input_2; …]; prompt_cell)
```

- **input_1, input_2, …**: Cells or values to use in the prompt. At least one input is required. Multiple inputs map to placeholders `{0}`, `{1}`, `{2}`, etc. in order.
- **prompt_cell**: Cell containing the prompt template. The last argument is always the prompt template.

Arguments are separated by semicolon (;) in many European locales, and by comma (,) in e.g. English (US). The function returns either a single text value (one cell) or multiple values that spill horizontally (see **Multiple outputs**).

## Placeholders

Placeholders use **curly brackets** `{}`. The index matches the order of the input arguments.

- **`{0}`** — replaced with the first input.
- **`{1}`** — replaced with the second input (if present).
- **`{2}`** — third, etc.

## Multiple outputs (output schema)

If the prompt template contains **square brackets with comma-separated field names**, e.g. `[firstname,lastname]`, it is treated as a request for a structured response (a JSON object with those keys). The result then spills across multiple cells (one row horizontally), one cell per field.

- **Syntax:** `[field1,field2,...]` – the first occurrence of `[ ... ]` in the prompt is parsed; field names are trimmed.
- **Example:** `Extract [firstname,lastname] from {0}` → the model is asked to return JSON with keys `firstname` and `lastname`; the first cell gets the first name, the next cell the last name.
- Without an output schema in the prompt, behaviour is unchanged: one text value per call.
- If the model returns invalid or non-JSON text, `ERROR: Invalid structured response` is shown in the formula cell.

## Usage examples

**One input:**

- **A3** contains: `Hello my name is Henrik`
- **B1** contains: `If the {0} contains a name then output only the name as output`
- **B3** contains: `=PROMPT(A3;B1)` → result: `Henrik`

**Multiple inputs:**

- **A3** contains: `Henrik`
- **B3** contains: `Stockholm`
- **C1** contains: `The person {0} lives in {1}.`
- **C3** contains: `=PROMPT(A3;B3;C1)` → result: `The person Henrik lives in Stockholm.`

**Multiple outputs (spill to the right):**

- **A3** contains: `John Doe`
- **B1** contains: `Extract [firstname,lastname] from {0}`
- **B3** contains: `=PROMPT(A3;B1)` → first cell (B3) gets e.g. `John`, next cell (C3) gets e.g. `Doe`. The result spills horizontally.

See **API key setup** if you haven’t set the key yet.

## Errors and limitations

- On failure (network error, invalid API key, rate limit, etc.) the function returns an error message in the cell, e.g. `ERROR: <short description>`. The implementation may add logging.
- Google Apps Script limits custom functions (e.g. ~30 s max runtime). With many rows, avoid running hundreds of `=PROMPT` calls at once; a batch or menu/trigger can be used in a future version.
