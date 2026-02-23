/**
 * @OnlyCurrentDoc
 * Mass Prompt – Google Apps Script (container-bound).
 * Custom function PROMPT(input_1; …; prompt_cell) calls OpenAI Chat Completions and returns the model response.
 * API key is set via menu Mass Prompt → Set API key.
 */

var SCRIPT_PROP_API_KEY = 'API_KEY';
var SCRIPT_PROP_OPENAI_ORG = 'OPENAI_ORG';
var ERROR_MIN_ARGS = 'ERROR: At least one input and one prompt required';
var ERROR_NO_API_KEY = 'ERROR: Set API key via menu Mass Prompt → Set API key';
var ERROR_STRUCTURED_RESPONSE = 'ERROR: Invalid structured response';

/**
 * Custom spreadsheet function: =PROMPT(input_1; input_2; …; prompt_cell)
 * Last argument = prompt template (with {0}, {1}, …), others = inputs. Calls OpenAI and returns the response.
 * If the prompt contains [field1,field2,…] the result is returned as multiple cells (spill to the right).
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
 * Returns a cell-friendly value when val is a single value or 2D/1D array from a range.
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
 * Creates the Mass Prompt menu when the spreadsheet is opened.
 * Set an "On open" trigger for onOpen if the menu does not appear automatically.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mass Prompt')
    .addItem('Set API key', 'showApiKeySidebar')
    .addItem('Clear API key', 'clearApiKey')
    .addToUi();
}

/**
 * Returns the saved organization (only for populating the sidebar). Does not expose the API key.
 */
function getOrgForSidebar() {
  return PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_OPENAI_ORG) || '';
}

/**
 * Opens the sidebar for entering the API key and organization.
 */
function showApiKeySidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('API key')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Forces recalculation of all cells that contain PROMPT in their formula,
 * so they show updated results after the API key is saved or cleared.
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
 * Saves the API key and optional OpenAI organization id in Script Properties. Called from the sidebar.
 * Then refreshes all cells with =PROMPT(...) so they recalculate.
 */
function saveApiKey(apiKey, orgId) {
  if (!apiKey || (typeof apiKey === 'string' && apiKey.trim() === '')) {
    throw new Error('Enter an API key.');
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
 * Removes the API key and organization id from Script Properties and shows confirmation.
 * Then refreshes all cells with =PROMPT(...) so they show the error message.
 */
function clearApiKey() {
  PropertiesService.getScriptProperties().deleteProperty(SCRIPT_PROP_API_KEY);
  PropertiesService.getScriptProperties().deleteProperty(SCRIPT_PROP_OPENAI_ORG);
  refreshPromptCells();
  SpreadsheetApp.getUi().alert('API key and organization have been removed.');
}
