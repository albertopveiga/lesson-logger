/**
 * Invoice Generator — Google Apps Script API
 * 
 * Proxies Mollie API calls from the browser app.
 * Paste into a NEW Google Apps Script project (or add to existing).
 * Deploy as Web App (Execute as: Me, Access: Anyone).
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
      case 'listCustomers':
        result = listMollieCustomers(data.apiKey);
        break;
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

// ─── List all Mollie customers ───────────────────────────────────────────────

function listMollieCustomers(apiKey) {
  var allCustomers = [];
  var url = 'https://api.mollie.com/v2/customers?limit=250';
  
  while (url) {
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + apiKey },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      return { error: 'Mollie API error: ' + response.getContentText() };
    }
    
    var data = JSON.parse(response.getContentText());
    if (data._embedded && data._embedded.customers) {
      data._embedded.customers.forEach(function(c) {
        allCustomers.push({
          id: c.id,
          name: c.name || '',
          email: c.email || ''
        });
      });
    }
    url = (data._links && data._links.next) ? data._links.next.href : null;
  }
  
  return { success: true, customers: allCustomers };
}

// ─── Create a single draft sales invoice ─────────────────────────────────────

function createDraftInvoice(apiKey, invoice) {
  // invoice: { customerId, customerName, customerEmail, lines: [{ description, quantity, unitPrice }] }
  
  var body = {
    status: 'draft',
    currency: 'EUR',
    vatScheme: 'standard',
    vatMode: 'inclusive',
    paymentTerm: '30days',
    recipientIdentifier: invoice.customerId,
    recipient: {
      type: 'consumer',
      email: invoice.customerEmail || '',
      locale: 'nl_NL'
    },
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
  var responseData = JSON.parse(response.getContentText());
  
  if (code >= 200 && code < 300) {
    return { 
      success: true, 
      invoiceId: responseData.id,
      invoiceNumber: responseData.invoiceNumber || '',
      status: responseData.status
    };
  } else {
    return { 
      error: 'Mollie error (' + code + '): ' + (responseData.detail || responseData.title || JSON.stringify(responseData))
    };
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
