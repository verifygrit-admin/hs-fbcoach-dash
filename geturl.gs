/**
 * GETURL — Extracts the hyperlink URL from a cell.
 * Handles both manually inserted links (rich text) and =HYPERLINK() formulas.
 *
 * Usage:
 *   =GETURL("B10")           — cell B10 on the active sheet
 *   =GETURL("Sheet1!B10")    — cell B10 on a specific sheet
 *
 * To use across a column, drag down:
 *   =GETURL("B"&ROW())       — auto-adjusts row as you drag
 *
 * @param {string} address - Cell address as a string, e.g. "B10" or "Sheet1!B10"
 * @return {string} The URL embedded in that cell, or "" if none found.
 * @customfunction
 */
function GETURL(address) {
  if (!address || typeof address !== 'string') return '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let range;

  try {
    if (address.includes('!')) {
      const [sheetName, cellRef] = address.split('!');
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return 'Sheet not found: ' + sheetName;
      range = sheet.getRange(cellRef);
    } else {
      range = ss.getActiveSheet().getRange(address);
    }
  } catch (e) {
    return 'Invalid address: ' + address;
  }

  // 1. Try rich text (manually inserted hyperlinks via Insert > Link)
  const richText = range.getRichTextValue();
  if (richText) {
    // Check whole-cell link
    const cellUrl = richText.getLinkUrl();
    if (cellUrl) return cellUrl;
    // Check individual runs (partial links within cell text)
    const runs = richText.getRuns();
    for (const run of runs) {
      const runUrl = run.getLinkUrl();
      if (runUrl) return runUrl;
    }
  }

  // 2. Try =HYPERLINK("url", "text") formula
  const formula = range.getFormula();
  if (formula) {
    const match = formula.match(/=HYPERLINK\(\s*"([^"]+)"/i);
    if (match) return match[1];
  }

  return '';
}
