/**
 * API-agnostisk mall-ersättning för Mass Prompt.
 * Ersätter placeholders {0}, {1}, {2}, … med värden från inputs-arrayen.
 * Ingen referens till SpreadsheetApp, UrlFetchApp eller OpenAI.
 */

/**
 * Fyller en prompt-mall med indatavärden.
 * @param {string} template - Mall med placeholders {0}, {1}, {2}, …
 * @param {Array} inputs - Array av värden (sträng eller tal); index i motsvarar {i}.
 * @returns {string} Den fyllda prompten. Saknas värde för {i} ersätts med tom sträng.
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
 * Parsar utdata-schema från prompt-mall: första [field1,field2,...] tolkas som lista med fältnamn.
 * @param {string} template - Prompt-mall (kan innehålla [firstname,lastname] etc.).
 * @returns {Array.<string>|null} Array med fältnamn (trimmat), eller null om inget schema hittas.
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
