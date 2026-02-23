/**
 * @OnlyCurrentDoc
 * Mass Prompt – Google Apps Script (container-bound).
 * Custom function PROMPT(indata_1; …; prompt_cell) anropar OpenAI Chat Completions och returnerar modellens svar.
 * API-nyckel sätts via menyn Mass Prompt → Sätt API-nyckel.
 */

var SCRIPT_PROP_API_KEY = 'API_KEY';
var SCRIPT_PROP_OPENAI_ORG = 'OPENAI_ORG';
var ERROR_MIN_ARGS = 'ERROR: Minst ett indata och en prompt krävs';
var ERROR_NO_API_KEY = 'ERROR: Sätt API-nyckel via menyn Mass Prompt → Sätt API-nyckel';
var ERROR_STRUCTURED_RESPONSE = 'ERROR: Ogiltigt strukturerat svar';

/**
 * Anpassad funktion för kalkylarket: =PROMPT(indata_1; indata_2; …; prompt_cell)
 * Sista argumentet = prompt-mall (med {0}, {1}, …), övriga = indata. Anropar OpenAI och returnerar svaret.
 * Om prompten innehåller [field1,field2,…] returneras resultatet som flera celler (spill till höger).
 */
function PROMPT() {
  var args = Array.prototype.slice.call(arguments);
  if (!args || args.length < 2) {
    return ERROR_MIN_ARGS;
  }
  var apiKey = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_API_KEY);
  if (!apiKey || (typeof apiKey === 'string' && apiKey.trim() === '')) {
    return ERROR_NO_API_KEY;
  }
  var rawInputs = args.slice(0, -1);
  var templateRaw = args[args.length - 1];
  var template = templateRaw != null ? String(normalizeCellValue(templateRaw)) : '';
  var normalizedInputs = [];
  for (var i = 0; i < rawInputs.length; i++) {
    normalizedInputs.push(normalizeCellValue(rawInputs[i]));
  }

  var outputKeys = parseOutputSchema(template);
  if (outputKeys && outputKeys.length > 0) {
    var result = sendPromptStructuredMulti(template, normalizedInputs, outputKeys, apiKey);
    if (result.indexOf('ERROR: ') === 0) {
      return result;
    }
    try {
      var raw = result.trim();
      if (raw.indexOf('```') === 0) {
        raw = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '');
      }
      var obj = JSON.parse(raw);
      if (obj === null || typeof obj !== 'object' || Array.isArray(obj)) {
        return ERROR_STRUCTURED_RESPONSE;
      }
      var row = [];
      for (var k = 0; k < outputKeys.length; k++) {
        var key = outputKeys[k];
        var val = obj[key];
        row.push(val !== undefined && val !== null ? String(val) : '');
      }
      return [row];
    } catch (e) {
      return ERROR_STRUCTURED_RESPONSE;
    }
  }

  var result = sendPromptStructured(template, normalizedInputs, apiKey);
  if (result.indexOf('ERROR: ') === 0) {
    return result;
  }
  return result;
}

/**
 * Returnerar ett cellvänligt värde när val är enkel värde eller 2D/1D-array från område.
 */
function normalizeCellValue(val) {
  if (val === null || val === undefined) {
    return '';
  }
  if (Array.isArray(val)) {
    if (val.length === 0) return '';
    if (Array.isArray(val[0])) {
      return val[0][0] !== undefined && val[0][0] !== null ? val[0][0] : '';
    }
    return val[0] !== undefined && val[0] !== null ? val[0] : '';
  }
  return val;
}

/**
 * Skapar menyn Mass Prompt vid öppning av kalkylarket.
 * Sätt en utlösare "Vid öppning av dokument" för onOpen om menyn inte dyker upp automatiskt.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mass Prompt')
    .addItem('Sätt API-nyckel', 'showApiKeySidebar')
    .addItem('Radera API-nyckel', 'clearApiKey')
    .addToUi();
}

/**
 * Returnerar sparad organisation (endast för att fylla i sidebar). Exponerar inte API-nyckel.
 */
function getOrgForSidebar() {
  return PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_OPENAI_ORG) || '';
}

/**
 * Öppnar sidopanelen för inmatning av API-nyckel och organisation.
 */
function showApiKeySidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('API-nyckel')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Tvingar omräkning av alla celler som innehåller PROMPT i formeln,
 * så att de får uppdaterat resultat efter att API-nyckel sparats eller raderats.
 */
function refreshPromptCells() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getDataRange();
    if (!range) return;
    var formulas = range.getFormulas();
    var numRows = formulas.length;
    if (numRows === 0) return;
    for (var r = 0; r < numRows; r++) {
      var row = formulas[r];
      if (!row) continue;
      for (var c = 0; c < row.length; c++) {
        var formula = row[c];
        if (formula && typeof formula === 'string' && formula.indexOf('PROMPT') !== -1) {
          range.getCell(r + 1, c + 1).setFormula(formula);
        }
      }
    }
  } catch (e) {
    Logger.log('refreshPromptCells: ' + e.toString());
  }
}

/**
 * Sparar API-nyckel och valfritt OpenAI-organisations-id i Script Properties. Anropas från sidebar.
 * Uppdaterar därefter alla celler med =PROMPT(...) så att de räknas om.
 */
function saveApiKey(apiKey, orgId) {
  if (!apiKey || (typeof apiKey === 'string' && apiKey.trim() === '')) {
    throw new Error('Ange en API-nyckel.');
  }
  PropertiesService.getScriptProperties().setProperty(SCRIPT_PROP_API_KEY, apiKey.trim());
  if (orgId != null && typeof orgId === 'string' && orgId.trim() !== '') {
    PropertiesService.getScriptProperties().setProperty(SCRIPT_PROP_OPENAI_ORG, orgId.trim());
  } else {
    PropertiesService.getScriptProperties().deleteProperty(SCRIPT_PROP_OPENAI_ORG);
  }
  refreshPromptCells();
}

/**
 * Tar bort API-nyckel och organisations-id från Script Properties och visar bekräftelse.
 * Uppdaterar därefter alla celler med =PROMPT(...) så att de visar felmeddelande.
 */
function clearApiKey() {
  PropertiesService.getScriptProperties().deleteProperty(SCRIPT_PROP_API_KEY);
  PropertiesService.getScriptProperties().deleteProperty(SCRIPT_PROP_OPENAI_ORG);
  refreshPromptCells();
  SpreadsheetApp.getUi().alert('API-nyckeln och organisationen har tagits bort.');
}
