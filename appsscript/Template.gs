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
