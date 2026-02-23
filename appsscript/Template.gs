/**
 * API-agnostic template substitution for Mass Prompt.
 * Replaces placeholders {0}, {1}, {2}, … with values from the inputs array.
 * No reference to SpreadsheetApp, UrlFetchApp, or OpenAI.
 */

/**
 * Fills a prompt template with input values.
 * @param {string} template - Template with placeholders {0}, {1}, {2}, …
 * @param {Array} inputs - Array of values (string or number); index i corresponds to {i}.
 * @returns {string} The filled prompt. Missing value for {i} is replaced with empty string.
 */
function fillTemplate(template, inputs) {
  if (template == null || typeof template !== 'string') {
    return '';
  }
  if (!inputs || !Array.isArray(inputs)) {
    return template;
  }
  var result = template;
  for (var i = 0; i < inputs.length; i++) {
    var placeholder = '{' + i + '}';
    var value = inputs[i] != null ? String(inputs[i]) : '';
    result = result.split(placeholder).join(value);
  }
  return result;
}

/**
 * Parses the output schema from the prompt template: the first [field1,field2,...] is treated as the list of field names.
 * @param {string} template - Prompt template (may contain [firstname,lastname] etc.).
 * @returns {Array.<string>|null} Array of field names (trimmed), or null if no schema found.
 */
function parseOutputSchema(template) {
  if (template == null || typeof template !== 'string') {
    return null;
  }
  var match = template.match(/\[([^\]]+)\]/);
  if (!match || !match[1]) {
    return null;
  }
  var part = match[1].trim();
  if (part.length === 0) {
    return null;
  }
  var keys = part.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });
  return keys.length > 0 ? keys : null;
}
