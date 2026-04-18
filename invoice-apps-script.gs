/**
 * Invoice Generator — Google Apps Script API
 *
 * Proxies Mollie API calls from the browser app.
 * Paste into a NEW Google Apps Script project (or add to existing).
 * Deploy as Web App (Execute as: Me, Access: Anyone).
 *
 * NOTE: We don't fetch customers from Mollie anymore. The AR/Invoicing customer
 * list has no public API, and Mollie's sales-invoices endpoint doesn't require
 * an existing customer record — `recipientIdentifier` is a free-form string
 * you assign per parent, and `recipient` carries the actual contact data.
 */

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid JSON' }))
      .setMimeType(ContentService.MimeType.JSON);
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
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
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
  // some address fields but accepts their absence.
  var rawRecipient = invoice.recipient || {};
  var recipient = {};
  Object.keys(rawRecipient).forEach(function(k) {
    var v = rawRecipient[k];
    if (v !== null && v !== undefined && String(v).trim() !== '') {
      recipient[k] = v;
    }
  });
  if (!recipient.type) recipient.type = 'consumer';
  if (!recipient.locale) recipient.locale = 'nl_NL';

  var body = {
    status: 'draft',
    currency: 'EUR',
    vatScheme: 'standard',
    vatMode: 'inclusive',
    paymentTerm: '30 days',
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
