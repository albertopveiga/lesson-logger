/**
 * Invoice Generator — Google Apps Script API
 * Version: v2026-04-18 14:40 UTC (adds authorizeAll helper for Drive/Sheets scope)
 *
 * Proxies Mollie API calls and reads the "Facturas" intake spreadsheet for
 * parent/student data. Paste into a NEW Google Apps Script project (or add
 * to existing). Deploy as Web App (Execute as: Me, Access: Anyone).
 *
 * IMPORTANT: This script now needs Drive/Sheets permissions. After pasting
 * you MUST redeploy, and the first run will prompt for reauthorization so
 * the executing account can read the spreadsheet below.
 *
 * Architecture notes:
 * - Mollie AR customer list has no public API, so we do NOT call Mollie for
 *   customer data. The "Facturas" sheet (filled in by every new parent at
 *   intake) is the source of truth for parent/student contact details.
 * - Mollie's Sales Invoices API requires the full inline recipient payload
 *   on every POST, even if the AR customer already exists. Empirically
 *   verified 2026-04-18. The UUID-as-recipientIdentifier trick does not work.
 */

// ─── Configuration ───────────────────────────────────────────────────────────

// The "Facturas" intake spreadsheet — the parent/student registration form
// writes to this sheet. Tab gid 1134169331 is the "Parents/Student" tab.
var SPREADSHEET_ID = '11WEK_G0_OhA18VxuzGziXnkvwi8FIT-JLsaxqlgNmbo';
var PARENTS_TAB_GID = 1134169331;

// Column headers we look for in the intake tab. These must exist exactly in
// the header row (header matching is case-insensitive and whitespace-tolerant).
// If any of these go missing, lookupRecipients returns an error.
var COL_PARENT_NAME  = "Parent/Guardian's Full Name";
var COL_PARENT_EMAIL = "Parent/Guardian's Email";
var COL_COUNTRY      = 'Country Of Residence';
var COL_ADDRESS      = 'Address (including city/town)';
var COL_POSTCODE     = 'Postcode';

// ─── One-time authorization helpers ──────────────────────────────────────────
//
// Apps Script only grants scopes (like Sheets read) after the script owner
// approves them in an interactive prompt. Deploying code that USES a scope
// doesn't grant the scope — you have to run any function that touches that
// scope from the editor UI first, which triggers the OAuth consent flow.
//
// To unblock the web app after pasting this file:
//   1. Open the Apps Script editor.
//   2. Pick `authorizeAll` from the function dropdown next to the Run button.
//   3. Click Run. Google will prompt: "Authorization required" → Review
//      permissions → pick your account → "Allow" on Drive/Sheets scope.
//   4. Done — the web app now has permission. No need to redeploy.

function authorizeAll() {
  // Touches both scopes the web app needs: SpreadsheetApp (Sheets) and
  // UrlFetchApp (external HTTPS). Returns a short status string so you see
  // something in the editor's execution log.
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheetByGid(ss, PARENTS_TAB_GID);
  var tabName = sheet ? sheet.getName() : '(no tab with that gid)';
  // Dry-run a UrlFetch to api.mollie.com (no auth, expect 401 — we just want
  // the scope consent triggered).
  try {
    UrlFetchApp.fetch('https://api.mollie.com/v2/sales-invoices', {
      muteHttpExceptions: true, method: 'get'
    });
  } catch (e) { /* ignore */ }
  return 'OK — sheet title: "' + ss.getName() + '", tab: "' + tabName + '"';
}

// ─── Entry point ─────────────────────────────────────────────────────────────

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ error: 'Invalid JSON' });
  }

  var action = data.action || '';
  var result;

  try {
    switch (action) {
      case 'createDraftInvoice':
        result = createDraftInvoice(data.apiKey, data.invoice);
        break;
      case 'createBatchInvoices':
        result = createBatchInvoices(data.apiKey, data.invoices);
        break;
      case 'lookupRecipients':
        result = lookupRecipients(data.parentNames || []);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return json(result);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Sheet lookup ────────────────────────────────────────────────────────────

/**
 * Looks up recipient data for each requested parent name against the intake
 * spreadsheet. Dedups rows by lowercased parent email (first row wins —
 * siblings share a parent, we only want one address per parent).
 *
 * Returns: {
 *   recipients: {
 *     "<original parent name>": <recipient-or-null>,
 *     ...
 *   },
 *   diagnostics: { totalRows, uniqueParents, matched, unmatched: [names] }
 * }
 *
 * Each recipient-or-null is either the parsed recipient or null if no row
 * was found for that parent name. Shape:
 *   { email, givenName, familyName, streetAndNumber, postalCode, city,
 *     country: 'NL', locale: 'en_GB', _source: 'sheet',
 *     _parsed: { cityFromField, postcodeCleaned } }
 */
function lookupRecipients(parentNames) {
  if (!Array.isArray(parentNames)) return { error: 'parentNames must be an array' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = findSheetByGid(ss, PARENTS_TAB_GID);
  if (!sheet) return { error: 'Tab with gid ' + PARENTS_TAB_GID + ' not found in spreadsheet' };

  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { error: 'Sheet is empty or missing header row' };

  // Find the header row. In the Facturas sheet the first couple of rows are
  // a pricing table; the real header row starts with "Date/Time". Scan until
  // we find a row that contains the parent-name column.
  var headerIdx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].some(function(c) { return normalizeHeader(c) === normalizeHeader(COL_PARENT_NAME); })) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) return { error: 'Could not find header row (no "' + COL_PARENT_NAME + '" column)' };

  var headers = rows[headerIdx].map(normalizeHeader);
  var col = {
    name:     headers.indexOf(normalizeHeader(COL_PARENT_NAME)),
    email:    headers.indexOf(normalizeHeader(COL_PARENT_EMAIL)),
    country:  headers.indexOf(normalizeHeader(COL_COUNTRY)),
    address:  headers.indexOf(normalizeHeader(COL_ADDRESS)),
    postcode: headers.indexOf(normalizeHeader(COL_POSTCODE))
  };
  var missing = [];
  Object.keys(col).forEach(function(k) { if (col[k] === -1) missing.push(k); });
  if (missing.length) return { error: 'Missing expected columns: ' + missing.join(', ') };

  // Dedup: group by lowercased parent email, first row wins.
  var byEmail = {};
  var byName = {};
  var totalRows = 0;
  for (var r = headerIdx + 1; r < rows.length; r++) {
    var row = rows[r];
    var name  = String(row[col.name]  || '').trim();
    var email = String(row[col.email] || '').trim().toLowerCase();
    if (!name && !email) continue; // skip empty rows
    totalRows++;
    var key = email || ('name:' + normalizeName(name));
    if (!byEmail[key]) {
      byEmail[key] = buildRecipient(row, col);
    }
    // Index by normalized name — first winner stays.
    var nameKey = normalizeName(name);
    if (nameKey && !byName[nameKey]) {
      byName[nameKey] = byEmail[key];
    }
  }

  // Look up each requested parent name.
  var out = {};
  var unmatched = [];
  parentNames.forEach(function(pn) {
    var nameKey = normalizeName(pn);
    var hit = byName[nameKey] || null;
    out[pn] = hit;
    if (!hit) unmatched.push(pn);
  });

  return {
    recipients: out,
    diagnostics: {
      totalRows: totalRows,
      uniqueParents: Object.keys(byEmail).length,
      matched: parentNames.length - unmatched.length,
      unmatched: unmatched
    }
  };
}

function findSheetByGid(spreadsheet, gid) {
  var sheets = spreadsheet.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return null;
}

function normalizeHeader(h) {
  return String(h || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function normalizeName(n) {
  // Same algorithm as the browser's normalize() — strip accents, lowercase,
  // collapse whitespace, drop punctuation.
  var s = String(n || '').toLowerCase();
  s = s.normalize ? s.normalize('NFD').replace(/[\u0300-\u036f]/g, '') : s;
  s = s.replace(/[^a-z0-9\s]/g, '').replace(/\s+/g, ' ').trim();
  return s;
}

function splitName(full) {
  var parts = String(full || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return { givenName: '', familyName: '' };
  if (parts.length === 1) return { givenName: parts[0], familyName: '' };
  return {
    givenName: parts.slice(0, -1).join(' '),
    familyName: parts[parts.length - 1]
  };
}

/**
 * Build a recipient object from a single row. Applies best-effort parsing to
 * the messy Address/Postcode columns (city is embedded, not separate).
 */
function buildRecipient(row, col) {
  var rawName     = String(row[col.name]     || '').trim();
  var rawEmail    = String(row[col.email]    || '').trim();
  var rawCountry  = String(row[col.country]  || '').trim();
  var rawAddress  = String(row[col.address]  || '').trim();
  var rawPostcode = String(row[col.postcode] || '').trim();

  var nameSplit = splitName(rawName);
  var parsed = parseAddressAndPostcode(rawAddress, rawPostcode);

  return {
    email:           rawEmail.toLowerCase(),
    givenName:       nameSplit.givenName,
    familyName:      nameSplit.familyName,
    streetAndNumber: parsed.street,
    postalCode:      parsed.postcode,
    city:            parsed.city,
    country:         'NL',                 // we assume NL; form can override
    locale:          'en_GB',
    _source:         'sheet',
    _raw: {
      countryFromSheet: rawCountry,
      addressFromSheet: rawAddress,
      postcodeFromSheet: rawPostcode
    }
  };
}

/**
 * Best-effort parser for the intake sheet's address and postcode fields.
 *
 * Real examples and how we handle them:
 *   address "Adriaan Pauwstraat 25", postcode "2582AN"
 *     → street="Adriaan Pauwstraat 25", postcode="2582 AN", city=""
 *
 *   address "Kolenwagenslag 55, Den Haag", postcode "2584SL"
 *     → street="Kolenwagenslag 55", postcode="2584 SL", city="Den Haag"
 *
 *   address "De Bruynestraat 9, 2597RD, The Hague", postcode "2597RD"
 *     → street="De Bruynestraat 9", postcode="2597 RD", city="The Hague"
 *
 *   address "Strijplaan 208", postcode "2285HX Rijswijk"
 *     → street="Strijplaan 208", postcode="2285 HX", city="Rijswijk"
 *
 *   address "Nederhoflaan, 48, Den Haag", postcode "2553EW"
 *     → street="Nederhoflaan 48", postcode="2553 EW", city="Den Haag"
 */
function parseAddressAndPostcode(address, postcode) {
  // --- 1. Pull a clean postcode + maybe a city out of the Postcode field.
  var pc = '';
  var cityFromPostcode = '';
  var pcMatch = String(postcode || '').match(/(\d{4})\s*([A-Za-z]{2})(?:\s+(.+))?/);
  if (pcMatch) {
    pc = (pcMatch[1] + ' ' + pcMatch[2].toUpperCase()).trim();
    if (pcMatch[3]) cityFromPostcode = pcMatch[3].trim();
  }

  // --- 2. Parse the Address field.
  // Strip any embedded postcode (like "..., 2597RD, The Hague") so we don't
  // double-count it in street/city.
  var addr = String(address || '')
    .replace(/\b\d{4}\s*[A-Za-z]{2}\b/g, '')
    .replace(/,\s*,/g, ',')
    .trim()
    .replace(/^,|,$/g, '')
    .trim();

  var street = addr;
  var cityFromAddress = '';

  // Comma-separated? Try: street is first N-1 parts joined, city is the last.
  if (addr.indexOf(',') !== -1) {
    var parts = addr.split(',').map(function(p) { return p.trim(); }).filter(Boolean);
    if (parts.length >= 2) {
      var last = parts[parts.length - 1];
      if (looksLikeCity(last)) {
        cityFromAddress = last;
        parts.pop();
      }
      // Collapse remaining parts into "Street Number": if one is pure digits
      // (like the "48" in "Nederhoflaan, 48, Den Haag") join with a space.
      street = parts.join(' ').replace(/\s+/g, ' ').trim();
    }
  } else {
    // No commas. Look for a trailing city keyword.
    var trailMatch = addr.match(/^(.+?)\s+(Den Haag|The Hague|Amsterdam|Rotterdam|Rijswijk|Utrecht|Hague|Leiden|Delft|Voorburg|Wassenaar|Scheveningen|Zoetermeer)\s*$/i);
    if (trailMatch) {
      street = trailMatch[1].trim();
      cityFromAddress = trailMatch[2].trim();
    }
  }

  // --- 3. Prefer address-derived city; fall back to postcode-derived city.
  var city = cityFromAddress || cityFromPostcode || '';

  return { street: street, postcode: pc, city: city };
}

function looksLikeCity(s) {
  // A "city" should be mostly letters (optionally with hyphens/spaces) and
  // NOT start with a digit. Reject things that look like house numbers or
  // postcodes.
  if (!s) return false;
  var t = String(s).trim();
  if (/^\d/.test(t)) return false;
  if (/^\d{4}\s*[A-Za-z]{2}$/.test(t)) return false;
  return /[A-Za-z]{2,}/.test(t);
}

// ─── Create a single draft sales invoice ─────────────────────────────────────

function createDraftInvoice(apiKey, invoice) {
  // invoice: {
  //   recipientIdentifier,            // stable slug, e.g. "parent-bruno-pallotta"
  //   customerName,                   // used only for logging/labels
  //   recipient: {                    // passed through verbatim to Mollie
  //     type, email, givenName, familyName,
  //     streetAndNumber, postalCode, city, country, locale
  //   },
  //   lines: [{ description, quantity, unitPrice }],
  //   memo
  // }

  // Only forward non-empty recipient fields — Mollie rejects empty strings in
  // some address fields but accepts their absence. Also drop any _source /
  // _raw / _parsed helpers we attached for debugging on the browser side.
  var rawRecipient = invoice.recipient || {};
  var recipient = {};
  Object.keys(rawRecipient).forEach(function(k) {
    if (k.charAt(0) === '_') return; // drop our own diagnostic fields
    var v = rawRecipient[k];
    if (v !== null && v !== undefined && String(v).trim() !== '') {
      recipient[k] = v;
    }
  });
  if (!recipient.type) recipient.type = 'consumer';
  if (!recipient.locale) recipient.locale = 'en_GB';

  var body = {
    status: 'draft',
    currency: 'EUR',
    vatScheme: 'standard',
    vatMode: 'inclusive',
    paymentTerm: '14 days',
    recipientIdentifier: invoice.recipientIdentifier || '',
    recipient: recipient,
    lines: invoice.lines.map(function(line) {
      return {
        description: line.description,
        quantity: line.quantity,
        vatRate: '0',
        unitPrice: {
          currency: 'EUR',
          value: formatMollieAmount(line.unitPrice)
        }
      };
    }),
    memo: invoice.memo || ''
  };

  var response = UrlFetchApp.fetch('https://api.mollie.com/v2/sales-invoices', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var responseText = response.getContentText();
  var responseData;
  try { responseData = JSON.parse(responseText); } catch (e) { responseData = { raw: responseText }; }

  if (code >= 200 && code < 300) {
    return {
      success: true,
      invoiceId: responseData.id,
      invoiceNumber: responseData.invoiceNumber || '',
      status: responseData.status
    };
  } else {
    // Surface Mollie's detail/title/field verbatim so the browser can display it.
    var msg = responseData.detail || responseData.title || responseText;
    if (responseData.fields && responseData.fields.length) {
      msg += ' (fields: ' + responseData.fields.map(function(f) { return f.field + ': ' + f.message; }).join('; ') + ')';
    }
    return { error: 'Mollie error (' + code + '): ' + msg };
  }
}

// ─── Create batch of draft invoices ──────────────────────────────────────────

function createBatchInvoices(apiKey, invoices) {
  var results = [];

  for (var i = 0; i < invoices.length; i++) {
    var inv = invoices[i];
    try {
      var result = createDraftInvoice(apiKey, inv);
      results.push({
        customerName: inv.customerName,
        result: result
      });
      // Small delay to avoid rate limiting
      if (i < invoices.length - 1) {
        Utilities.sleep(500);
      }
    } catch (err) {
      results.push({
        customerName: inv.customerName,
        result: { error: err.toString() }
      });
    }
  }

  return { success: true, results: results };
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function formatMollieAmount(amount) {
  // Mollie expects "10.00" format
  var num = parseFloat(amount) || 0;
  return num.toFixed(2);
}
