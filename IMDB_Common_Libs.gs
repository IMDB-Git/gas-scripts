/**
 * Reads a required Script Property. Throws if missing so callers fail fast.
 * All secrets (API keys, passwords, tokens) must be stored as Script Properties
 * via Project Settings > Script Properties, or via setScriptPropertyPrompt().
 */
function getScriptProp_(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error('Script Property mancante: ' + key);
  return value;
}

// FattureInCloud — URL pubblico (non segreto)
var FICUrl = "https://api-v2.fattureincloud.it";

// SCOT srl — URL pubblici (non segreti)
var prodScotBaseURL = "https://api.portalescotsrl.it";
var testScotBaseURL = "https://testapi.portalescotsrl.it";


/**
 * Requests an authentication token from the SCOT portal.
 *
 * @param {string} username - The username credential.
 * @param {string} password - The password credential.
 * @return {string|null} - The token if the request succeeds, otherwise null.
 */
function getScotToken(test = false)
{
  var baseURL = test ? testScotBaseURL : prodScotBaseURL;
  var url = baseURL + "/api/token/";

  var payload = {
    username: getScriptProp_('SCOT_USERNAME'),
    password: getScriptProp_('SCOT_PASSWORD')
  };
  
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var jsonResponse = JSON.parse(response.getContentText());
    
    if (code !== 200) {
      Logger.log("Token request failed with error code: " + code);
      if (jsonResponse.error) {
        Logger.log("Error message: " + jsonResponse.error);
      }
      return null;
    } else {
      Logger.log("Token request successful. Token: " + jsonResponse.token);
      return jsonResponse.token;
    }
  } catch (e) {
    Logger.log("Exception during token request: " + e);
    return null;
  }
}

// ===== SCOT API =====

/**
 * Invia un ordine in uscita al portale SCOT (/api/uscite/).
 * PUBLIC: chiamata da IMDB_Ordini_Scripts.
 */
function scotOrdiniUscita(orderId, clientId, header, rows, clienteNome, campagnaNome, files = null) {
  var url = prodScotBaseURL + "/api/uscite/";
  var token = getScotToken();
  if (!token) { Logger.log("Impossibile ottenere il token"); return; }

  var payload = { order_id: orderId, client: clientId, header: header, rows: rows };
  if (files && Array.isArray(files) && files.length) { payload.files = files; }
  Logger.log(payload);

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var json = JSON.parse(response.getContentText());
    var htmlJSON = formatJsonGeneric(payload);
    if (code === 200) {
      Logger.log("Ordine Uscita: " + orderId + " inviato con successo: %s", htmlJSON);
      sendEmailViaSMTP(htmlJSON, "ordini@ilmassimodelbere.it", "SPEDIZIONE: " + orderId + " Cliente: " + clienteNome + " Campagna: " + campagnaNome, "IMDB Logistics");
      return json;
    } else {
      Logger.log("Errore invio ordine: " + orderId + " (%s): %s", code, JSON.stringify(json));
      return null;
    }
  } catch (e) {
    Logger.log("Eccezione invio ordine: " + orderId + " : %s", e);
    return null;
  }
}

function testscotOrdiniUscita_() {
  var token = getScotToken();
  if (!token) { Logger.log("Impossibile ottenere il token"); return; }
  var header = {
    business_name: "La Mia Azienda SRL",
    document_date: (new Date()).toISOString(),
    attachment: false, address: "Via Roma 1", location: "Milano",
    province: "MI", zip_code: "20100", nation: "IT", urgent: false,
    delivery_date: (new Date(new Date().getTime() + 3*24*3600*1000)).toISOString(),
    appointment: false, email: "info@azienda.it", tel_reference: "0234567890",
    carrier_note: "Consegna al piano", warehouse_note: "Att.n. imballi fragili",
    cash_on_delivery_value: 0.0, cash_on_delivery_type: ""
  };
  var rows = [{ id: 1, code: "ABABR03", quantity: 2 }, { id: 2, code: "ABABR14", quantity: 1 }];
  var orderName = "OD" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMddmmss");
  var result = scotOrdiniUscita(token, orderName, "MDB", header, rows);
  Logger.log(result ? "Succeded! Result: %s" : "ERROR! Result: %s", JSON.stringify(result));
}

function scotOrdiniUscita_Stato_(token, orderId, clientId) {
  var url = prodScotBaseURL + "/api/uscite/stato/";
  var payload = { order_id: orderId, client: clientId };
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  };
  try {
    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    var data = JSON.parse(resp.getContentText());
    if (code === 200) {
      Logger.log("Stato ordine %s: %s %s", orderId, JSON.stringify(data), resp);
      return data;
    } else {
      Logger.log("Errore Stato ordine (%s): %s", code, JSON.stringify(data));
      return null;
    }
  } catch (e) {
    Logger.log("Eccezione Stato ordine: %s", e);
    return null;
  }
}

function testscotOrdiniUscita_Stato_() {
  var token = getScotToken();
  if (!token) { Logger.log("Token non ottenuto"); return; }
  var stato = scotOrdiniUscita_Stato_(token, "5E01003342", "MDB");
  if (!stato) { Logger.log("Recupero stato fallito"); return; }
  Logger.log("Order ID: %s", stato.order_id);
  Logger.log("Status code: %s", scotGetOrderStatusDescription_(stato.status));
  if (stato.acquisition_date) Logger.log("Acquired: %s", stato.acquisition_date);
  if (stato.conclusion_date)  Logger.log("Concluded: %s", stato.conclusion_date);
  if (stato.rows) stato.rows.forEach(function(r) {
    Logger.log("Riga %s (ID %s): code=%s, req=%s, proc=%s", r.row_number, r.id, r.code, r.quantity_required, r.quantity_processed);
  });
}

function scotGetOrderStatusDescription_(code) {
  var key = typeof code === 'string' ? parseInt(code, 10) : code;
  var statusMap = {
    0: "In Elaborazione", 10: "In Acquisizione", 20: "Acquisito (non valido)",
    30: "Acquisito (valido)", 50: "Da Elaborare (Attivato)", 55: "In Elaborazione",
    60: "Evadibile (elaborato)", 62: "Elaborato - Da Preparare", 63: "In Preparazione",
    65: "Attesa esecuzione rimpiazzi", 68: "In Preparazione attività",
    69: "Non Prelevabile", 70: "Prelevabile", 80: "Prelevabile (senza impegni)",
    90: "In Prelievo", 95: "In Viaggio", 100: "Parzialmente Prelevato",
    110: "Prelevato", 180: "Pesato", 200: "Spuntato", 500: "Concluso",
    600: "Annullato", 1000: "Inevadibile", 5000: "Aggregato a Lista"
  };
  return statusMap.hasOwnProperty(key) ? statusMap[key] : "Sconosciuto (" + key + ")";
}

// ===== Functions moved from IMDB_Ordini_Scripts =====

/**
 * Finds the column number of a target header in a given header row of a sheet.
 */
function getEmailRiepilogoColumn(currentSheet, targetHeader, headerRow)
{

  // Get all header values from row 33 across all columns
  var headers = currentSheet.getRange(headerRow, 1, 1, currentSheet.getLastColumn()).getValues()[0];

  // Initialize the variable to store the column number
  var targetColumn = -1;

  // Loop through headers to find the matching column title
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim() === targetHeader) {
      targetColumn = i + 1; // Column numbers are 1-indexed in Google Sheets
      break;
    }
  }

  if (targetColumn === -1) {
    throw new Error('Column with header ' + targetHeader + ' not found in row ' + headerRow);
  }

  return targetColumn;
}

/**
 * Sends a shipping order to the SCOT portal and logs the result.
 */
function inviaSpedizioni(numeroOrdine, scotUsciteHeader, scotUsciteRows, clienteNome, campagnaNome)
{
    if (numeroOrdine === null)
      return;

    var result =  scotOrdiniUscita(numeroOrdine, "MDB", scotUsciteHeader, scotUsciteRows, clienteNome, campagnaNome);
    if (result)
    {
      Logger.log("SCOT Uscite: ordine " + numeroOrdine +" succeded! Result: %s", JSON.stringify(result));

      return true;
    }
    else
    {
      Logger.log("SCOT Uscite: ordine " + numeroOrdine +" ERROR! Result: %s", JSON.stringify(result));
      return false;
    }
}

/**
 * Submits the "Order Received" Mautic form for a customer.
 */
function submitFormOrderReceived(email, firstName, lastName, phoneNumber, actualCampaign, actualDate)
{
  // URL for the external form submission
  var url = "https://www.ilmassimodelbere.it/Mautic/form/submit?formId=38";

  // Build the payload that mimics the form fields
  var payload = {
    'mauticform[email]': email,
    'mauticform[nome]': firstName,
    'mauticform[cognome]': lastName,
    'mauticform[telefono]': phoneNumber,
    'mauticform[data_ordine_mettere_la_da]': actualDate,
    'mauticform[descrizione_campagna]': actualCampaign,
    'mauticform[formId]': '38',
    'mauticform[return]': '',
    'mauticform[formName]': 'itimdbinternalordinevinoorderreceived'
  };

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  // Perform the POST request
  var response = UrlFetchApp.fetch(url, options);

  // Log details for diagnostics
  Logger.log("Submitted form with:");
  Logger.log("  Email: " + email);
  Logger.log("  Name: " + lastName + " " + firstName);
  Logger.log("  Data: " + actualDate);
  Logger.log("  Campagna: " + actualCampaign);

  return response.getContentText();
}

/**
 * Submits the "Shipping Started" Mautic form for a customer.
 */
function submitFormShippingStarted(email, firstName, lastName, actualCampaign, actualDate)
{
  // URL for the external form submission
  var url = "https://www.ilmassimodelbere.it/Mautic/form/submit?formId=30";

  // Build the payload that mimics the form fields
  var payload = {
    'mauticform[nome]': firstName,
    'mauticform[cognome]': lastName,
    'mauticform[email]': email,
    'mauticform[data_spedizione_mettere_l]': actualDate,
    'mauticform[descrizione_campagna_iden]': actualCampaign,
    'mauticform[formId]': '30',
    'mauticform[return]': '',
    'mauticform[formName]': 'itimdbinternalshippingstarted'
  };

  var options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  // Perform the POST request
  var response = UrlFetchApp.fetch(url, options);

  // Log details for diagnostics
  Logger.log("Submitted form with:");
  Logger.log("  Email: " + email);
  Logger.log("  Name: " + lastName + " " + firstName);
  Logger.log("  Data: " + actualDate);
  Logger.log("  Campagna: " + actualCampaign);
  Logger.log("Response: " + response.getContentText());

  return response.getContentText();
}

/**
 * Converts a column letter (e.g. "A", "AB") to its 1-based column number.
 */
function columnLetterToNumber(columnLetter) {
  var columnNumber = 0;
  var length = columnLetter.length;
  for (var i = 0; i < length; i++) {
    columnNumber *= 26;
    columnNumber += (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return columnNumber;
}

/**
 * Returns the 1-based column index for the first matching header name, or 0 if not found.
 */
function getColumnByHeaderName(headerRow, headerName)
{
  for (var i = 0; i < headerRow.length; i++)
  {
    if (String(headerRow[i] || '').trim() === headerName)
    {
      return i + 1;
    }
  }

  return 0;
}

/**
 * Returns the 1-based column index for the first matching header among possibleNames, or 0 if none found.
 */
function getColumnByHeaderNameMultiple(headerRow, possibleNames)
{
  for (var p = 0; p < possibleNames.length; p++)
  {
    var col = getColumnByHeaderName(headerRow, possibleNames[p]);

    if (col)
    {
      return col;
    }
  }

  return 0;
}

/**
 * Returns true if value is a valid Date object or a parseable date string.
 */
function isValidDateValue(value)
{
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime()))
  {
    return true;
  }

  var s = String(value || '').trim();

  if (!s)
  {
    return false;
  }

  var d = new Date(s);

  return !isNaN(d.getTime());
}

// ===== End of functions moved from IMDB_Ordini_Scripts =====

// This function makes the actual API call to create the invoice
function createFICOrderInvoice(invoiceType, clientRagioneSociale, clientEntityType, clientType, clientNome, clientCognome, clientCodice, clientAddress, clientCity, clientCAP, clientProvincia, clientEmail, clientPhone, clientVatNumber, clientPEC, clientSDI, clientCodiceFiscale, clientNoteSpedizione, clientNoteCliente, invoiceDate, invoiceSubject, invoiceVisibleSubject, invoiceAmount, invoiceItems, showPaymentMethod, paymentID, paymentMethod, paymentEMethod, paymentNotes, paymentsItems, useGrossPrice = true)
{
    var endpoint = "/c/" + getScriptProp_('FIC_COMPANY_ID') + "/issued_documents";
    var headers = {
      "Content-Type": "application/json",
      Authorization: "Bearer " + getScriptProp_('FIC_BEARER'),
    };

    var itemsList = mapInvoiceItems(invoiceItems);
    var paymentsList = mapPaymentsList(paymentsItems);

    var body = 
    {
      data: 
      {
        type: invoiceType,
        e_invoice: true,
        ei_data:
        {
          payment_method: paymentEMethod
        },
        amount_gross: invoiceAmount,
        "amount_net": invoiceAmount / 1.22,
        "amount_vat": invoiceAmount - (invoiceAmount / 1.22),
        visible_subject: invoiceVisibleSubject,
        subject: invoiceSubject,
        use_gross_prices: useGrossPrice,
        entity: 
        {
            code: clientCodice,
            name: clientRagioneSociale,
            type: clientType,
            entity_type: clientEntityType,
            first_name: clientNome,
            last_name: clientCognome,
            vat_number: clientVatNumber,
            tax_code: clientCodiceFiscale,
            address_street: clientAddress,
            address_city: clientCity,
            address_postal_code: clientCAP,
            address_province: clientProvincia,
            country: "Italia",
            email: clientEmail,
            certified_email: clientPEC,
            phone: clientPhone,
        },
        date: invoiceDate,
        currency: 
        {
          id: "EUR",
          exchange_rate: "1.00000",
          symbol: "€"
        },
        language: 
        {
          code: "it",
          name: "Italiano"
        }, 
        payment_method:
        {
          id: paymentID,
          name: paymentMethod
        },
        show_payment_method: showPaymentMethod,
        notes: paymentNotes,
        payments_list: paymentsList,
        /*payments_list: 
        [
          {
            amount: invoiceAmount,
            due_date: clientDataSaldo,
            paid_date: clientDataSaldo,
            id: null,// 305871974,
            forfettari_revenue: invoiceAmount,
            payment_terms: null,
            status: "paid",
            payment_account: {
              id: 1232535,
              name: "Credit Agricole",
              virtual: false
            },
            ei_raw: null
          },
        ],*/
        items_list: itemsList
      }
    };

    Logger.log(body);

    var options = {
        "method": "post",
        "headers": headers,
        "payload": JSON.stringify(body),
        "muteHttpExceptions": false

    };

    try {
        var response = UrlFetchApp.fetch(FICUrl + endpoint, options);
        data = JSON.parse(response.getContentText());
        Logger.log("Invoice created succesfully with id: " + data.data.number);
        return data;
    } catch (e) {
        //SpreadsheetApp.getUi().alert(e.message)
        Logger.log(e.message);
        return e.message;
    }
}


// This function makes the actual API call to create the invoice
function createFICInvoice(invoiceType, clientRagioneSociale, clientEntityType, clientType, clientNome, clientCognome, clientCodice, clientAddress, clientCity, clientCAP, clientProvincia, clientEmail, clientPhone, clientVatNumber, clientPEC, clientSDI, clientCodiceFiscale, clientNoteSpedizione, clientNoteCliente, invoiceDate, invoiceSubject, invoiceVisibleSubject, invoiceAmount, clientNoteSaldo, clientDataSaldo, clientPaidAmount, invoiceDescription, paymentID, paymentMethod, paymentEMethod)
{
    var endpoint = "/c/" + getScriptProp_('FIC_COMPANY_ID') + "/issued_documents";
    var headers = {
      "Content-Type": "application/json",
      Authorization: "Bearer " + getScriptProp_('FIC_BEARER'),
    };

    var body = 
    {
      data: 
      {
        type: invoiceType,
        e_invoice: true,
        ei_data:
        {
          payment_method: paymentEMethod
        },
        amount_gross: invoiceAmount,
        "amount_net": invoiceAmount / 1.22,
        "amount_vat": invoiceAmount - (invoiceAmount / 1.22),
        visible_subject: invoiceVisibleSubject,
        subject: invoiceSubject,
        use_gross_prices: true,
        entity: 
        {
            code: clientCodice,
            name: clientRagioneSociale,
            type: clientType,
            entity_type: clientEntityType,
            first_name: clientNome,
            last_name: clientCognome,
            vat_number: clientVatNumber,
            tax_code: clientCodiceFiscale,
            address_street: clientAddress,
            address_city: clientCity,
            address_postal_code: clientCAP,
            address_province: clientProvincia,
            country: "Italia",
            email: clientEmail,
            certified_email: clientPEC,
            phone: clientPhone,
        },
        date: invoiceDate,
        currency: 
        {
          id: "EUR",
          exchange_rate: "1.00000",
          symbol: "€"
        },
        language: 
        {
          code: "it",
          name: "Italiano"
        }, 
        payment_method:
        {
          id: paymentID,
          name: paymentMethod
        },
        show_payment_method: true,
        notes: clientNoteSaldo + clientDataSaldo,
        payments_list: 
        [
          {
            amount: invoiceAmount,
            due_date: clientDataSaldo,
            paid_date: clientDataSaldo,
            id: null,// 305871974,
            forfettari_revenue: invoiceAmount,
            payment_terms: null,
            status: "paid",
            payment_account: {
              id: 1232535,
              name: "Credit Agricole",
              virtual: false
            },
            ei_raw: null
          },
        ],
        items_list: 
        [
          {
            "product_id": null,
            "code": "",
            "name": "",
            "measure": "",
            "category": "",
            "id": null, //1123147833,
            "apply_withholding_taxes": false,
            "discount": 0,
            "discount_highlight": false,
            "in_dn": false,
            "qty": 1,
            "gross_price": invoiceAmount,
            "vat": {
              "id": 0,
              "value": 22,
              "description": ""
            },
            "stock": null,
            "description": invoiceDescription,
            "gross_price": invoiceAmount,
            "not_taxable": false,
            "ei_raw": null
          }
        ],
      }
    }

    Logger.log(body);

    var options = {
        "method": "post",
        "headers": headers,
        "payload": JSON.stringify(body)
    };

    try {
        var response = UrlFetchApp.fetch(FICUrl + endpoint, options);
        data = JSON.parse(response.getContentText());
        Logger.log("Invoice created succesfully with id: " + data.data.number);
        return data;
    } catch (e) {
        //SpreadsheetApp.getUi().alert(e.message)
        Logger.log(e.message);
        return e.message;
    }
}

function convertDateString(inputData) 
{
  try
  {
    const parts = inputData.split('/');
    const giorno = parts[0].padStart(2, '0');
    const mese   = parts[1].padStart(2, '0');
    const anno   = parts[2];
    return `${anno}-${mese}-${giorno}`;
  }
  catch (e)
  {
    return null;
  }
}

/**
 * Restituisce l’array items_list mappato dai tuoi invoiceItems,
 * gestendo anche il caso in cui ogni item sia un array contenente
 * l’oggetto reale.
 */
function mapInvoiceItems(invoiceItems) {
  return invoiceItems.map((raw, idx) => {
    // Se raw è un array prendi il primo elemento, altrimenti usa raw stesso
    const item = Array.isArray(raw) ? raw[0] : raw;

    const vat = item.vat || {};

    return {
      product_id:              item.product_id   ?? null,
      code:                    item.code         ?? '',
      name:                    item.name         ?? '',
      measure:                 item.measure      ?? '',
      category:                item.category     ?? '',
      id:                      item.id           ?? null,
      apply_withholding_taxes: Boolean(item.apply_withholding_taxes),
      discount:                Number(item.discount    ?? 0),
      discount_highlight:      Boolean(item.discount_highlight),
      in_dn:                   Boolean(item.in_dn),
      qty:                     Number(item.qty         ?? 1),
      gross_price:             Number(item.gross_price ?? 0),
      net_price:               Number(item.net_price ?? 0),
      vat: {
        id:        Number(vat.id    ?? 0),
        value:     Number(vat.value ?? 0),
        description: vat.description ?? ''
      },
      stock:                 item.stock    ?? null,
      description:           item.description ?? '',
      not_taxable:           Boolean(item.not_taxable),
      ei_raw:                item.ei_raw   ?? null
    };
  });
}

/**
 * Mappa un array di paymentItems nella struttura richiesta da payments_list.
 *
 * @param {Array<Object|Array>} paymentItems — array di oggetti (o di array contenenti un solo oggetto)
 * @return {Array<Object>} — array di pagamenti formattati
 */
function mapPaymentsList(paymentItems) {
  return paymentItems.map((raw, idx) => {
    // Se raw è un array, prendi il primo elemento; altrimenti usa raw direttamente
    const item = Array.isArray(raw) ? raw[0] : raw;

    // Debug: verifica l’oggetto
    Logger.log(`PaymentItem[${idx}]: ${JSON.stringify(item)}`);

    return {
      amount:             Number(item.amount    ?? 0),
      due_date:           item.due_date         ?? null,
      paid_date:          item.paid_date        ?? null,
      id:                 item.id               ?? null,
      forfettari_revenue: Number(item.forfettari_revenue ?? item.amount ?? 0),
      payment_terms:      item.payment_terms    ?? null,
      status:             item.status           ?? 'pending',
      payment_account: {
        id:      item.payment_account?.id    ?? null,
        name:    item.payment_account?.name  ?? '',
        virtual: Boolean(item.payment_account?.virtual)
      },
      ei_raw:             item.ei_raw           ?? null
    };
  });
}


// Utility function to format price values as "1.234,00€"
function formatPrice(value) {
  return value.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,').replace('.', ',') + '€';
}

function getCustomDateCode() {
  const now = new Date();

  // Anno → '5' per 2025, poi '6', ..., '9', 'A', 'B', ...
  const year = now.getFullYear();
  let yearCode = '';
  if (year < 2030) {
    yearCode = String(year).slice(-1); // Ultima cifra
  } else {
    yearCode = String.fromCharCode(65 + (year - 2030)); // 'A' = 2030, 'B' = 2031, ...
  }

  // Mese → A=Gennaio, B=Febbraio, ... L=Dicembre
  const monthCode = String.fromCharCode(65 + now.getMonth()); // getMonth() = 0–11

  // Giorno → due cifre
  const dayCode = ("0" + now.getDate()).slice(-2);

  return yearCode + monthCode + dayCode;
}




function sendEmailViaSMTP(htmlContent, recipientEmail, subject, senderName, ccEmail = '', bccEmail = '', fileBlob = null, provider = 'Aruba') {
  
  var smtpServer;
  var port;
  var username;
  var password;
  
  if (provider === "Aruba")
  {  
    Logger.log("Sending email with Aruba provider");
    smtpServer = "smtps.aruba.it";
    port = 465;
    username = "ordini@ilmassimodelbere.it";
    password = getScriptProp_('SMTP_ARUBA_PASSWORD');
  }
  else   if (provider === "TurboSMTP")
  {
    Logger.log("Sending email with TurboSMTP provider");
    smtpServer = "pro.turbo-smtp.com";
    port = 465;
    username = "admin@ilmassimodelbere.it";
    password = getScriptProp_('SMTP_TURBOSMTP_PASSWORD');
  }
  else
  {
    Logger.log("No SMTP provider specified!");
    return "No SMTP provider specified";
  }

  const emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log(`Remaining email quota: ${emailQuotaRemaining}`);
  if (emailQuotaRemaining <= 0)
  {
     return ("Errore: nessuna email residua da inviare!");
  }

  var props = {
    "mail.smtp.host": smtpServer,
    "mail.smtp.port": port,
    "mail.smtp.auth": "true",
    "mail.smtp.socketFactory.port": port,
    "mail.smtp.socketFactory.class": "javax.net.ssl.SSLSocketFactory",
    "mail.smtp.socketFactory.fallback": "false"
  };

  try {
    if (fileBlob)
    {
      MailApp.sendEmail({
        to: recipientEmail,
        cc: ccEmail,
        bcc: bccEmail,
        subject: subject,
        htmlBody: htmlContent,
        attachments: [fileBlob],
        name: senderName
      });
    }
    else
    {
      MailApp.sendEmail({
        to: recipientEmail,
        cc: ccEmail,
        bcc: bccEmail,
        subject: subject,
        htmlBody: htmlContent,
        name: senderName
      });
    }
    Logger.log("Email sent successfully. Recipient: " + recipientEmail + " Subject: " + subject);
    return false;
  } catch (e) {
    Logger.log("Error sending email: " + e.message);
    return e.message;
  }
}

function testCreateProforma_ ()
{
createFICInvoice ("Pipppo Pluto", "client", "person", "Massimo", "Bombino", "", "Via di qui, 32", "Milano", "20123", "PV", "ordini@ilmassimodelbere.it", "+39-348-2639796", "02948720186", "pec@pec.it", "0000000", "BMBMSM70L14F205Y", "Spedire subito", "Ottimo cliente", "IMDB Acquisto", "IMDB - Acconto per merce da consegnare", "497", "Fattura saldata con pagamento Stripe del ","2025-04-26",  "497", "IMDB - Acconto per merce da consegnare (rif. IMDB VIP Club Diamond)");
//  createFICInvoice ("", "client", "person", "", "", "2443", "", "", "", "", "", "", "", "", "", "", "", "", "IMDB Acquisto", "IMDB - Acconto per merce da consegnare", "497", "Fattura saldata con pagamento Stripe del ","2025-04-26",  "497", "IMDB - Acconto per merce da consegnare (rif. IMDB VIP Club Diamond)");
}

/**
 * Test del webhook.php su Aruba
 */
function testWebhookWhatsappBridge() {
  const url = 'https://www.ilmassimodelbere.it/php/3141592654_1123581321_UebUcch.php';
  
  // Build the payload.
  var payload = 
  {
    phone: '393482639796',
    template: '20250505_hello_word_position',
    params: ['Max']

  };

  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // per vedere comunque corpo e codice di risposta
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    Logger.log('HTTP Code: ' + response.getResponseCode());
    Logger.log('Response Body: ' + response.getContentText());
  } catch (e) {
    Logger.log('Errore nella fetch: ' + e);
  }
}


function jsonToHtml(obj, indent = 0)
{
  let html = '';
  const margin = indent * 20; // px

  if (Array.isArray(obj))
  {
    obj.forEach((item, i) =>
    {
      html += `<div style="margin-left:${margin}px"><b>[${i}]</b></div>`;
      html += jsonToHtml(item, indent + 1);
    });
  }
  else if (typeof obj === 'object' && obj !== null)
  {
    Object.keys(obj).forEach(key =>
    {
      const value = obj[key];
      if (typeof value === 'object' && value !== null)
      {
        html += `<div style="margin-left:${margin}px"><b>${key}:</b></div>`;
        html += jsonToHtml(value, indent + 1);
      }
      else
      {
        html += `<div style="margin-left:${margin}px"><b>${key}:</b> ${value}</div>`;
      }
    });
  }
  else
  {
    html += `<div style="margin-left:${margin}px">${obj}</div>`;
  }

  return html;
}

function formatJsonGeneric(json)
{
  return `<div style="font-family:Tahoma,Arial,sans-serif; font-size:14px;">${jsonToHtml(json)}</div>`;
}


/**
 * Request "Giacenza" information from the SCOT portal for a given client.
 *
 * Endpoint: /api/giacenza/
 * Request body JSON structure:
 *   { "client": "string len(5)" }
 *
 * Expected successful response (HTTP 200) is a JSON object with a property "InStockList",
 * which is an array of objects having the properties: code (string), name (string),
 * quantity (integer) and quantity_in_orders (integer).
 *
 * @param {string} token - A valid authentication token (obtained from getScotToken()).
 * @param {string} clientId - The client code (string of length 5) for which to request "Giacenza".
 * @return {Object|null} - The parsed JSON response object if successful; otherwise, null.
 */
function getGiacenze(token, clientId) {
  var endpoint = "/api/giacenza/";
  var url = scotBaseURL + endpoint;
  
  // Build the payload according to the API specification.
  var payload = {
    client: clientId
  };
  
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: {
      "Authorization": "Bearer " + token
    },
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    if (responseCode === 200) {
      Logger.log("Giacenza request successful. Response: " + response.getContentText());
      // Parse and return the JSON response.
      return JSON.parse(response.getContentText());
    } else {
      Logger.log("Giacenza request failed. Code: " + responseCode + ", Response: " + response.getContentText());
      return null;
    }
  } catch (e) {
    Logger.log("Error while fetching giacenza: " + e);
    return null;
  }
}



// =======================
// CONFIGURAZIONE MAUTIC
// =======================
const MAUTIC_BASE_URL     = 'https://www.ilmassimodelbere.it/Mautic'; // senza /api
const MAUTIC_API_BASE_URL = MAUTIC_BASE_URL + '/api';

// =======================
// SERVICE OAUTH2 MAUTIC
// =======================
function getMauticService() {
  return OAuth2.createService('Mautic')
    .setAuthorizationBaseUrl(MAUTIC_BASE_URL + '/oauth/v2/authorize')
    .setTokenUrl(MAUTIC_BASE_URL + '/oauth/v2/token')
    .setClientId(getScriptProp_('MAUTIC_CLIENT_ID'))
    .setClientSecret(getScriptProp_('MAUTIC_CLIENT_SECRET'))
    .setCallbackFunction('mauticAuthCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('')
    .setParam('response_type', 'code');
}

// =======================
// CALLBACK DI AUTORIZZAZIONE
// =======================
function mauticAuthCallback(request) {
  const service = getMauticService();
  const authorized = service.handleCallback(request);

  if (authorized) {
    return HtmlService.createHtmlOutput(
      'Autorizzazione Mautic riuscita. Puoi chiudere questa scheda e tornare al foglio.'
    );
  } else {
    return HtmlService.createHtmlOutput(
      'Autorizzazione Mautic negata. Riprova dal foglio Google.'
    );
  }
}

// =======================
// (OPZIONALE) LOG REDIRECT URI
// =======================
function logMauticRedirectUri() {
  const service = getMauticService();
  Logger.log(service.getRedirectUri());
}

function testGetMauticCustomerData ()
{
  Logger.log(JSON.stringify(getMauticCustomerData('Farina')));
}

/**
 * LIBRERIA - Nessuna UI.
 *
 * Restituisce i dati del contatto Mautic in forma JSON.
 *
 * Il parametro può essere:
 * - ID numerico (max 6 cifre)  -> ritorna un singolo JSON (/contacts/{id})
 * - Email                     -> ricerca; se 1 match ritorna JSON, se >1 ritorna array di JSON
 * - Cognome (lastname)        -> ricerca; se 1 match ritorna JSON, se >1 ritorna array di JSON
 *
 * NOTE:
 * - Gestisce cognomi composti (es. "Pinco Pallino") facendo query con virgolette: lastname:"..."
 * - In caso di errori, logga su Logger e lancia eccezione (nessuna UI).
 */
function getMauticCustomerData(key) {
  if (!key) {
    const msg = 'Parametro mancante per la ricerca contatto Mautic.';
    Logger.log(msg);
    throw new Error(msg);
  }

  key = String(key).trim();

  const service = getMauticService();
  if (!service.hasAccess()) {
    const authUrl = service.getAuthorizationUrl();
    const msg =
      'Mautic non autorizzato. Apri questo URL in un browser, autorizza l’app e riprova:\n\n' +
      authUrl;
    Logger.log(msg);
    throw new Error(msg);
  }

  const accessToken = service.getAccessToken();

  function fetchJson_(url) {
    const resp = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    const status = resp.getResponseCode();
    const body   = resp.getContentText();

    if (status !== 200) {
      const msg = 'Errore Mautic HTTP ' + status + ' su URL: ' + url + '\n' + body;
      Logger.log(msg);
      throw new Error(msg);
    }

    return JSON.parse(body);
  }

  // ==========================
  // 1) ID numerico (1–6 cifre): fetch diretto
  // ==========================
  if (/^\d{1,6}$/.test(key)) {
    const customerUrl = MAUTIC_API_BASE_URL + '/contacts/' + key;
    return fetchJson_(customerUrl);
  }

  // ==========================
  // 2) Ricerca: email o lastname
  //    - se lastname contiene spazi -> lastname:"..."
  // ==========================
  const isEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(key);

  let searchQuery;
  if (isEmail) {
    searchQuery = 'email:' + key;
  } else {
    // Evita rotture se key contiene doppi apici
    const safeKey = key.replace(/"/g, '\\"');

    // Cognomi doppi: cerca frase esatta
    if (/\s/.test(safeKey)) {
      searchQuery = 'lastname:"' + safeKey + '"';
    } else {
      searchQuery = 'lastname:' + safeKey;
    }
  }

  const searchUrl  = MAUTIC_API_BASE_URL + '/contacts?search=' + encodeURIComponent(searchQuery);
  const searchData = fetchJson_(searchUrl);

  // Dipende dalla versione: a volte {contacts:{...}}, a volte direttamente l'oggetto
  const contactsObj = searchData.contacts || searchData || {};

  const ids = [];
  for (var id in contactsObj) {
    if (contactsObj.hasOwnProperty(id)) ids.push(String(id));
  }

  if (ids.length === 0) {
    const msg = 'Nessun contatto trovato per ricerca: ' + searchQuery;
    Logger.log(msg);
    throw new Error(msg);
  }

  // ==========================
  // 3) Recupero dettagli /contacts/{id} per ogni match
  // ==========================
  const MAX_FETCH = 20; // guard-rail per cognomi molto comuni
  if (ids.length > MAX_FETCH) {
    Logger.log('Trovati ' + ids.length + ' contatti per "' + key + '". Limito a ' + MAX_FETCH + '.');
  }

  const results = [];
  const take = Math.min(ids.length, MAX_FETCH);

  for (var i = 0; i < take; i++) {
    const contactId = ids[i];
    const customerUrl = MAUTIC_API_BASE_URL + '/contacts/' + contactId;
    results.push(fetchJson_(customerUrl));
  }

  // Se uno solo, ritorna oggetto singolo
  return (results.length === 1) ? results[0] : results;
}


/**
 * Meta Conversions API (CAPI) - Google Apps Script
 * - Funnel stages / Conversion Leads (CRM) friendly: event_name può essere una fase (custom event)
 * - Supporta lead_id da Meta Lead Ads (es. "l:1234567890123456") -> user_data.lead_id (NO hashing)
 *
 * Endpoint: https://graph.facebook.com/{API_VERSION}/{PIXEL_ID}/events?access_token=...  :contentReference[oaicite:8]{index=8}
 */

const META_CAPI_DEFAULTS = {
  apiVersion: "v24.0",                 // v24.0 è indicata come latest Graph API attualmente :contentReference[oaicite:9]{index=9}
  currency: "EUR",
  actionSource: "system_generated",  // valido per WhatsApp/Messenger/IG :contentReference[oaicite:10]{index=10}
  // Stadi consigliati (puoi personalizzarli liberamente, conta il significato) :contentReference[oaicite:11]{index=11}
  stageMap: {
    LEAD_GENERATED: "Intake",
    MARKETING_QUALIFIED: "Qualified",
    SALES_OPPORTUNITY: "Opportunity",
    CONVERTED: "Converted",
    DISQUALIFIED: "Disqualified",
  }
};

/**
 * Invia un evento CAPI.
 *
 * @param {string} status  Es: "CONVERTED", "QUALIFIED", "LEAD_GENERATED", oppure una fase custom ("VIP Qualified")
 * @param {Object} customer Dati cliente (raw). Esempio in metaCapiExample_().
 * @param {number|null} value Valore monetario opzionale (es. 129.90). Se null/undefined, non inviato.
 * @param {Object=} options Opzioni
 * @returns {Object} risposta JSON Meta
 */

// evita perdita cifre oltre MAX_SAFE_INTEGER
function safeLeadId_(digitsStr) {
  const maxSafe = BigInt(Number.MAX_SAFE_INTEGER);
  const bi = BigInt(digitsStr);
  return (bi <= maxSafe) ? Number(digitsStr) : digitsStr; // stringa se troppo grande
}

function metaCapiSend(status, customer, value, options) {
  // === LOG INIZIALE (safe) ===
  try {
    const c = customer || {};
    const o = options || {};

    const safeCustomer = {
      leadId: c.leadId || c.lead_id || "",
      externalId: c.externalId || "",
      hasEmail: !!c.email,
      hasPhone: !!c.phone,
      firstName: c.firstName ? String(c.firstName).substring(0, 4) + "***" : "",
      lastName:  c.lastName  ? String(c.lastName).substring(0, 4) + "***" : "",
      city: c.city || "",
      state: c.state || "",
      zip: c.zip || "",
      country: c.country || "",
      birthDate: c.birthDate || c.birth_date || "", // YYYYMMDD (non super sensibile, ma se vuoi lo mascheriamo)
      gender: c.gender || ""
    };

    const safeOptions = {
      actionSource: o.actionSource,
      currency: o.currency,
      eventTimeSec: o.eventTimeSec,
      eventId: o.eventId,
      testEventCode: o.testEventCode,
      messagingChannel: o.messagingChannel,
      leadEventSource: o.leadEventSource,
      eventSource: o.eventSource
    };

    Logger.log(
      "metaCapiSend_ input => " +
      JSON.stringify({
        status: status,
        value: value,
        customer: safeCustomer,
        options: safeOptions
      })
    );
  } catch (e) {
    Logger.log("metaCapiSend_ input log error: " + e);
  }

  // ... resto della tua funzione ...

  const props = PropertiesService.getScriptProperties();
  const pixelId = props.getProperty("META_PIXEL_ID");
  const accessToken = props.getProperty("META_ACCESS_TOKEN");
  const apiVersion = (props.getProperty("META_API_VERSION") || META_CAPI_DEFAULTS.apiVersion).trim();

  if (!pixelId || !accessToken) throw new Error("META_PIXEL_ID o META_ACCESS_TOKEN mancanti. Esegui metaCapiSetup_().");

  options = options || {};
  const actionSource = options.actionSource || META_CAPI_DEFAULTS.actionSource;
  const currency = options.currency || META_CAPI_DEFAULTS.currency;

  // status -> event_name (custom stage) oppure standard/custom passato dall'utente
  const eventName = mapStatusToEventName_(status);

  // event_time: unix seconds (GMT). Max 7 giorni nel passato :contentReference[oaicite:12]{index=12}
  const eventTime = options.eventTimeSec || Math.floor(Date.now() / 1000);

  const userData = buildUserData_(customer);

  // user_data deve contenere almeno un parametro valido :contentReference[oaicite:13]{index=13}
  if (Object.keys(userData).length === 0) {
    throw new Error("user_data vuoto: serve almeno un identificativo (email, telefono, ecc.)");
  }

  // event_id consigliato per deduplica :contentReference[oaicite:14]{index=14}
  const eventId = options.eventId || buildDeterministicEventId_(eventName, customer, eventTime);

  const event = {
    event_name: eventName,
    event_time: eventTime,
    action_source: actionSource,
    user_data: userData,
    custom_data: {
      event_source: "crm",
      lead_event_source: options.leadEventSource || "chatwoot"
    },
    event_id: eventId
  };

  // ✅ Fix: required when action_source is business_messaging
  if (actionSource === "business_messaging") {
    const ch = String(options.messagingChannel || "").trim().toLowerCase();
    if (!ch) throw new Error("action_source=business_messaging richiede options.messagingChannel = messenger|whatsapp|instagram");
    if (!["messenger", "whatsapp", "instagram"].includes(ch)) {
      throw new Error("messagingChannel non valido. Usa: messenger | whatsapp | instagram");
    }
    event.messaging_channel = ch;
  }


  // Se action_source=website, Meta richiede event_source_url (+ spesso user agent) :contentReference[oaicite:15]{index=15}
  if (actionSource === "website") {
    if (!options.eventSourceUrl) throw new Error("action_source=website richiede options.eventSourceUrl");
    event.event_source_url = options.eventSourceUrl;
    if (options.clientUserAgent) event.user_data.client_user_agent = options.clientUserAgent;
    if (options.clientIpAddress) event.user_data.client_ip_address = options.clientIpAddress;
  }
  // Custom data: base CRM (non sovrascrivere)
  event.custom_data = event.custom_data || {};
  event.custom_data.event_source = event.custom_data.event_source || "crm";
  event.custom_data.lead_event_source = event.custom_data.lead_event_source || (options.leadEventSource || "chatwoot");

  // ✅ Se evento Purchase (o se value presente) aggiungi currency/value SENZA perdere i campi CRM
  const isPurchase = String(eventName).toLowerCase() === "purchase";

  if (isPurchase) 
  {
    // Meta richiede currency per Purchase
    event.custom_data.currency = options.currency || META_CAPI_DEFAULTS.currency;
    // value: se non lo hai, manda 0 oppure imposta un default sensato
    if (typeof value === "number" && isFinite(value)) {
      event.custom_data.value = round2(value);
    } else {
      event.custom_data.value = 0;
    }
  } else if (typeof value === "number" && isFinite(value)) {
    // Per altri eventi, se passi value lo mettiamo con currency
    event.custom_data.currency = options.currency || META_CAPI_DEFAULTS.currency;
    event.custom_data.value = round2(value);
  }


  const payload = {
    data: [event]
  };

  Logger.log ("Meta CAPI Payload: " + JSON.stringify(payload));

  // Per test in Events Manager -> Test Events, puoi passare options.testEventCode
  if (options.testEventCode) payload.test_event_code = options.testEventCode;

  return postToMetaCapi_(apiVersion, pixelId, accessToken, payload);
}

/**
 * Esempio d'uso
 */
function metaCapiExample_() {
  const customer = {
    leadId: "l:1234567890123456", // da Meta Lead Ads, se disponibile
    email: "Mario.Rossi@example.com",
    phone: "+39 333 123 4567",
    firstName: "Mario",
    lastName: "Ròssi",
    city: "Roma",
    state: "RM",
    zip: "00100",
    country: "IT",
    externalId: "crm-contact-98765" // tuo ID interno (consigliato)
  };

  const res = metaCapiSend(
    "Purchase",
    customer,
    129.90,
    {
      actionSource: "system_generated", // WhatsApp/Chatwoot
      messagingChannel: "whatsapp",
      testEventCode: "TEST19291" // opzionale
    }
  );

  Logger.log(JSON.stringify(res, null, 2));
}

/* ----------------------- Helpers ----------------------- */

function mapStatusToEventName_(status) {
  const s = String(status || "").trim();
  if (!s) throw new Error("status/event stage vuoto");

  // normalizza sinonimi rapidi
  const up = s.toUpperCase();
  const map = META_CAPI_DEFAULTS.stageMap;

  if (map[up]) return map[up];

  // accetta anche QUALIFIED -> MARKETING_QUALIFIED
  if (up === "QUALIFIED") return map.MARKETING_QUALIFIED;

  // altrimenti usa lo status come custom event name (max 50 caratteri consigliato da integrazioni comuni)
  // (Meta accetta eventi custom; attenzione a non farli troppo lunghi) :contentReference[oaicite:16]{index=16}
  return s.length > 50 ? s.substring(0, 50) : s;
}

function buildUserData_(customer) {
  customer = customer || {};
  const ud = {};

  const leadIdDigits = extractLeadIdDigits(customer.leadId || customer.lead_id);
  if (leadIdDigits) {
    const maxSafe = BigInt(Number.MAX_SAFE_INTEGER);
    try {
      const bi = BigInt(leadIdDigits);
      ud.lead_id = (bi <= maxSafe) ? Number(leadIdDigits) : leadIdDigits;
    } catch (e) {
      ud.lead_id = leadIdDigits;
    }
  }

  if (customer.email) ud.em = [sha256Hex(normalizeEmail(customer.email))];
  if (customer.phone) ud.ph = [sha256Hex(normalizePhone(customer.phone))];

  if (customer.firstName) ud.fn = [sha256Hex(normalizeName(customer.firstName))];
  if (customer.lastName) ud.ln = [sha256Hex(normalizeName(customer.lastName))];

  // ✅ DOB + gender da CF (o altri canali)
  const birth = String(customer.birthDate || customer.birth_date || customer.db || '').trim();
  if (birth && /^\d{8}$/.test(birth)) {
    ud.db = [sha256Hex(birth)];
  }

  const g = String(customer.gender || customer.ge || '').trim().toLowerCase();
  if (g === 'm' || g === 'f') {
    ud.ge = [sha256Hex(g)];
  }

  if (customer.city) ud.ct = [sha256Hex(normalizeCity(customer.city))];
  if (customer.state) ud.st = [sha256Hex(normalizeState_(customer.state))];
  if (customer.zip) ud.zp = [sha256Hex(normalizeZip(customer.zip))];
  if (customer.country) ud.country = [sha256Hex(normalizeCountry(customer.country))];

  if (customer.externalId) ud.external_id = [sha256Hex(String(customer.externalId).trim())];

  if (customer.fbp) ud.fbp = String(customer.fbp).trim();
  if (customer.fbc) ud.fbc = String(customer.fbc).trim();

  return ud;
}


function postToMetaCapi_(apiVersion, pixelId, accessToken, payload) {
  const url = `https://graph.facebook.com/${encodeURIComponent(apiVersion)}/${encodeURIComponent(pixelId)}/events?access_token=${encodeURIComponent(accessToken)}`;
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  let json;
  try { json = JSON.parse(text); } catch (e) { json = { raw: text }; }

  if (code < 200 || code >= 300) {
    Logger.log(`Meta CAPI HTTP ${code}: ${text}`);
    throw new Error(`Meta CAPI HTTP ${code}: ${text}`);
  }
  return json;
}

function buildDeterministicEventId_(eventName, customer, eventTime) {
  // Se hai un ID evento CRM vero, passa options.eventId per deduplica migliore.
  const parts = [
    "crm",
    eventName,
    extractLeadIdDigits(customer && (customer.leadId || customer.lead_id)) || "",
    customer && customer.externalId ? String(customer.externalId) : "",
    customer && customer.email ? normalizeEmail(customer.email) : "",
    String(eventTime)
  ].join("|");
  return sha256Hex(parts).substring(0, 32);
}

/* --- Normalizzazione + hashing (SHA-256) --- */

function sha256Hex(input) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    input,
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")).join("");
}

function normalizeEmail(email) {
  return String(email).trim().toLowerCase();
}

function normalizePhone(phone) {
  // Meta: rimuovi simboli/lettere e zeri iniziali; includi prefisso internazionale :contentReference[oaicite:24]{index=24}
  let digits = String(phone).replace(/[^\d]/g, "");
  digits = digits.replace(/^0+/, "");
  return digits;
}

function normalizeName(name) {
  return stripDiacritics_(String(name).trim().toLowerCase())
    .replace(/[^a-z0-9]/g, ""); // semplificazione robusta
}

function normalizeCity(city) {
  return stripDiacritics_(String(city).trim().toLowerCase())
    .replace(/[^a-z0-9]/g, "");
}

function normalizeState_(state) {
  return stripDiacritics_(String(state).trim().toLowerCase())
    .replace(/[^a-z0-9]/g, "");
}

function normalizeZip(zip) {
  return String(zip).trim().toLowerCase().replace(/[\s-]/g, "");
}

function normalizeCountry(country) {
  // ISO-3166-1 alpha-2 lowercase :contentReference[oaicite:25]{index=25}
  return String(country).trim().toLowerCase();
}

function stripDiacritics_(s) {
  // rimuove accenti (es. Ròssi -> rossi)
  return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function extractLeadIdDigits(leadIdRaw) {
  if (!leadIdRaw) return "";
  const s = String(leadIdRaw).trim();
  // accetta "l:123..." oppure "123..."
  const m = s.match(/^(?:l:)?(\d{10,20})$/i);
  return m ? m[1] : "";
}

function round2(n) {
  return Math.round(n * 100) / 100;
}


// Lookup function to get province acronym
function getProvinceAcronym(provinceName) {
  var provinceMap = {
    "AGRIGENTO": "AG",
    "ALESSANDRIA": "AL",
    "ANCONA": "AN",
    "AOSTA": "AO",
    "AREZZO": "AR",
    "ASCOLI PICENO": "AP",
    "ASTI": "AT",
    "AVELLINO": "AV",
    "BARI": "BA",
    "BARLETTA-ANDRIA-TRANI": "BT",
    "BELLUNO": "BL",
    "BENEVENTO": "BN",
    "BERGAMO": "BG",
    "BIELLA": "BI",
    "BOLOGNA": "BO",
    "BOLZANO": "BZ",
    "BRESCIA": "BS",
    "BRINDISI": "BR",
    "CAGLIARI": "CA",
    "CALTANISSETTA": "CL",
    "CAMPOBASSO": "CB",
    "CASERTA": "CE",
    "CATANIA": "CT",
    "CATANZARO": "CZ",
    "CHIETI": "CH",
    "COMO": "CO",
    "COSENZA": "CS",
    "CREMONA": "CR",
    "CROTONE": "KR",
    "CUNEO": "CN",
    "ENNA": "EN",
    "FERMO": "FM",
    "FERRARA": "FE",
    "FIRENZE": "FI",
    "FOGGIA": "FG",
    "FORLI-CESENA": "FC",
    "FROSINONE": "FR",
    "GENOVA": "GE",
    "GORIZIA": "GO",
    "GROSETTO": "GR",
    "IMPERIA": "IM",
    "ISERNIA": "IS",
    "L'AQUILA": "AQ",
    "LA SPEZIA": "SP",
    "LATINA": "LT",
    "LECCE": "LE",
    "LECCO": "LC",
    "LIVORNO": "LI",
    "LODI": "LO",
    "LUCCA": "LU",
    "MACERATA": "MC",
    "MANTOVA": "MN",
    "MASSA-CARRARA": "MS",
    "MATERA": "MT",
    "MESSINA": "ME",
    "MILANO": "MI",
    "MODENA": "MO",
    "MONZA E BRIANZA": "MB",
    "MONZA BRIANZA": "MB",
    "NAPOLI": "NA",
    "NOVARA": "NO",
    "NUORO": "NU",
    "ORISTANO": "OR",
    "PADOVA": "PD",
    "PALERMO": "PA",
    "PARMA": "PR",
    "PAVIA": "PV",
    "PERUGIA": "PG",
    "PESARO E URBINO": "PU",
    "PESARO-URBINO": "PU",
    "PESCARA": "PE",
    "PIACENZA": "PC",
    "PISA": "PI",
    "PISTOIA": "PT",
    "PORDENONE": "PN",
    "POTENZA": "PZ",
    "PRATO": "PO",
    "RAGUSA": "RG",
    "RAVENNA": "RA",
    "REGGIO CALABRIA": "RC",
    "REGGIO EMILIA": "RE",
    "RIETI": "RI",
    "RIMINI": "RN",
    "ROMA": "RM",
    "ROVIGO": "RO",
    "SALERNO": "SA",
    "SASSARI": "SS",
    "SAVONA": "SV",
    "SIENA": "SI",
    "SIRACUSA": "SR",
    "SONDRIO": "SO",
    "SUD SARDEGNA": "SU",
    "TARANTO": "TA",
    "TERAMO": "TE",
    "TERNI": "TR",
    "TORINO": "TO",
    "TRAPANI": "TP",
    "TRENTO": "TN",
    "TREVISO": "TV",
    "TRIESTE": "TS",
    "UDINE": "UD",
    "VARESE": "VA",
    "VENEZIA": "VE",
    "VERBANO-CUSIO-OSSOLA": "VB",
    "VERCELLI": "VC",
    "VERONA": "VR",
    "VIBO VALENTIA": "VV",
    "VICENZA": "VI",
    "VITERBO": "VT"
  };

  // Convert the provided province name to uppercase and trim whitespace.
  var normalizedName = provinceName.trim().toUpperCase();
  return provinceMap[normalizedName] || normalizedName;
}

/**
 * Validazione headless dati spedizione/anagrafica.
 *
 * Comportamento:
 * - nessuna UI
 * - in caso di errore: Logger.log(...) + return "messaggio errore"
 * - in caso OK: normalizza i campi dentro data + return ""
 *
 * Campi attesi in data:
 * - surname
 * - firstName
 * - provinciaDestinatario
 * - provinciaRaw
 * - indirizzo
 * - localita
 * - zipCode
 * - contact
 *
 * Campi opzionali/funzioni esterne usate:
 * - getMauticFieldNormalized(contact, fieldName)
 * - formatPhoneNumber_(raw)
 * - IMDBCommonLibs.getProvinceAcronym(provincia)
 *
 * Campi valorizzati/normalizzati in data:
 * - provinciaDestinatario
 * - phoneRaw
 * - telefonoFormatted
 * - nazione
 * - noteRaw
 * - noteSpedizione
 * - zipCode
 * - ragioneSociale
 */
function validateShippingDataHeadless(data) {
  data = data || {};

  var surname = String(data.surname || "").trim();
  var firstName = String(data.firstName || "").trim();
  var provinciaDestinatario = String(data.provinciaDestinatario || "").trim();
  var provinciaRaw = String(data.provinciaRaw || "").trim();
  var indirizzo = String(data.indirizzo || "").trim();
  var localita = String(data.localita || "").trim();
  var zipCode = String(data.zipCode || "").trim();
  var contact = data.contact || null;

  function fail_(msg) {
    Logger.log(msg);
    return msg;
  }

  // Provincia
  if (provinciaDestinatario.length > 2) {
    provinciaDestinatario = getProvinceAcronym(provinciaDestinatario);
  }

  if (provinciaDestinatario.length != 2) {
    return fail_(surname + ": Provincia errata: " + provinciaDestinatario + " (PROVINCIA=" + provinciaRaw + ")");
  }

  // Telefono (preferisci telefono, poi mobile, poi phone)
  var phoneRaw = "";
  if (contact) {
    phoneRaw =
      getMauticFieldNormalized(contact, "telefono") ||
      getMauticFieldNormalized(contact, "mobile") ||
      getMauticFieldNormalized(contact, "phone") ||
      "";
  }

  var telefonoFormatted = "";
  if (String(phoneRaw).trim().length) {
    telefonoFormatted = formatPhoneNumber(phoneRaw);
  }

  // Nazione
  var nazione = "IT";
  // Se vorrai riattivarla:
  // var nazione = (getMauticFieldNormalized(contact, 'nazione') || getMauticFieldNormalized(contact, 'country') || 'IT').toUpperCase();

  // Note spedizione
  var noteRaw = "";
  if (contact) {
    noteRaw =
      getMauticFieldNormalized(contact, "note_spedizione") ||
      getMauticFieldNormalized(contact, "address2") ||
      "";
  }

  // Check lunghezze
  var ragioneSociale = (surname + " " + firstName).trim();
  if ((ragioneSociale.length >= 30) || (ragioneSociale.length <= 6)) {
    return fail_(surname + ": Ragione sociale troppo lunga/corta: " + ragioneSociale);
  }

  if ((indirizzo.length >= 50) || (indirizzo.length <= 5)) {
    return fail_(surname + ": Indirizzo troppo lungo/corto: " + indirizzo);
  }

  if ((localita.length >= 30) || (localita.length <= 2)) {
    return fail_(surname + ": Località troppo lunga/corta: " + localita);
  }

  if (data.codiceFiscale.length != 16)
  {
    return fail_(surname + ": Codice fiscale errato: " + data.codiceFiscale);
  }

  // CAP
  if (/^\d+$/.test(zipCode) && zipCode.length < 5) {
    zipCode = zipCode.padStart(5, "0");
  }

  if (!/^\d{5}$/.test(zipCode) || zipCode === "00000") {
    return fail_(surname + ": CAP errato: " + zipCode + " " + zipCode.length);
  }

  // Note spedizione lunghezza
  if (String(noteRaw).length > 70) {
    return fail_(surname + ": Note di spedizione troppo lunghe: " + noteRaw);
  }

  var noteSpedizione = "";
  if (String(noteRaw).length >= 56) {
    noteSpedizione = noteRaw;
  } else {
    noteSpedizione = noteRaw;
    if (telefonoFormatted) {
      noteSpedizione = (noteSpedizione ? (noteSpedizione + " ") : "") + "Tel:" + telefonoFormatted;
    }
  }

  // Scrive i valori normalizzati nel payload in uscita
  data.provinciaDestinatario = provinciaDestinatario;
  data.phoneRaw = phoneRaw;
  data.telefonoFormatted = telefonoFormatted;
  data.nazione = nazione;
  data.noteRaw = noteRaw;
  data.noteSpedizione = noteSpedizione;
  data.zipCode = zipCode;
  data.ragioneSociale = ragioneSociale;
  data.noteSpedizione = noteSpedizione;

  return "";
}

// =======================
// HELPER: normalizedValue (es. SILVER/GOLD/DIAMOND/BLACK/N.A. oppure text/email)
// =======================
function getMauticFieldNormalized(contact, alias) {
  if (!contact || !contact.fields) return null;

  const fields = contact.fields;

  for (const groupName in fields) {
    if (!fields.hasOwnProperty(groupName)) continue;
    const group = fields[groupName];
    const field = group && group[alias];

    if (field) {
      if (typeof field.normalizedValue !== 'undefined' && field.normalizedValue !== null) {
        return field.normalizedValue;
      }
      if (typeof field.value !== 'undefined' && field.value !== null) {
        return field.value;
      }
    }
  }

  // fallback: se c'è in fields.all come valore semplice
  const all = fields.all || {};
  if (typeof all[alias] !== 'undefined' && all[alias] !== null) {
    return all[alias];
  }

  return null;
}

/* =========================================================
 * Helpers: Mautic field read (robusto)
 * ========================================================= */

function getMauticFieldNormalizedSafe(contact, fieldNames) {
  if (!contact) return '';
  const names = (fieldNames || []).map(n => String(n).trim()).filter(Boolean);

  // 1) proprietà dirette
  for (let i = 0; i < names.length; i++) {
    const key = names[i];
    if (contact[key] != null && contact[key] !== '') return String(contact[key]).trim();

    const low = key.toLowerCase();
    if (contact[low] != null && contact[low] !== '') return String(contact[low]).trim();
  }

  // 2) container fields (Mautic spesso: contact.fields.all/core/custom...)
  const fields = contact.fields || {};
  const containers = [fields.all, fields.core, fields.social, fields.personal, fields.custom]
    .filter(Boolean);

  for (let c = 0; c < containers.length; c++) {
    const obj = containers[c];
    for (let i = 0; i < names.length; i++) {
      const key = names[i];
      const low = key.toLowerCase();

      const v = (obj[key] !== undefined) ? obj[key] : obj[low];
      if (v === undefined || v === null || v === '') continue;

      if (typeof v === 'object' && v !== null && v.value !== undefined) {
        const vv = v.value;
        if (vv !== null && vv !== undefined && String(vv).trim() !== '') return String(vv).trim();
      }

      return String(v).trim();
    }
  }

  return '';
}

// Function to format the phone number: remove spaces, hyphens, and "+39" if present
function formatPhoneNumber(phone) {
  if (!phone) return ""; // Return empty if phone number is not provided
  phone = phone.toString().replace(/\s+/g, "").replace(/-/g, ""); // Remove spaces and hyphens
  if (phone.startsWith("+39")) {
    phone = phone.substring(3); // Remove "+39"
  }
  return phone;
}

// Utility function to find a matching row by a specific column (e.g., Codice Fiscale or Partita IVA)
function findMatchingRow(data, value, columnIndex) {
  for (var i = 1; i < data.length; i++) { // Start from 1 to skip the header row
    if (data[i][columnIndex] == value) {  // Compare based on the specified column index
      return data[i];
    }
  }
  return [];
}

/**
 * Funzione di supporto per mappare l'acronimo alla Regione (per logistica)
 */
function getRegioneDaAcronimo(sigla) {
  if (!sigla) return "";
  const map = {
    "AL":"PIEMONTE","AT":"PIEMONTE","BI":"PIEMONTE","CU":"PIEMONTE","NO":"PIEMONTE","TO":"PIEMONTE","VB":"PIEMONTE","VC":"PIEMONTE", "CN":"PIEMONTE",
    "AO":"VALLE D'AOSTA",
    "BG":"LOMBARDIA","BS":"LOMBARDIA","CO":"LOMBARDIA","CR":"LOMBARDIA","LC":"LOMBARDIA","LO":"LOMBARDIA","MN":"LOMBARDIA","MI":"LOMBARDIA","MB":"LOMBARDIA","PV":"LOMBARDIA","SO":"LOMBARDIA","VA":"LOMBARDIA",
    "BZ":"TRENTINO","TN":"TRENTINO",
    "BL":"VENETO","PD":"VENETO","RO":"VENETO","TV":"VENETO","VE":"VENETO","VR":"VENETO","VI":"VENETO",
    "GO":"FRIULI","PN":"FRIULI","TS":"FRIULI","UD":"FRIULI",
    "GE":"LIGURIA","IM":"LIGURIA","SP":"LIGURIA","SV":"LIGURIA",
    "BO":"EMILIA ROMAGNA","FE":"EMILIA ROMAGNA","FC":"EMILIA ROMAGNA","MO":"EMILIA ROMAGNA","PR":"EMILIA ROMAGNA","PC":"EMILIA ROMAGNA","RA":"EMILIA ROMAGNA","RE":"EMILIA ROMAGNA","RN":"EMILIA ROMAGNA",
    "AR":"TOSCANA","FI":"TOSCANA","GR":"TOSCANA","LI":"TOSCANA","LU":"TOSCANA","MS":"TOSCANA","PI":"TOSCANA","PT":"TOSCANA","PO":"TOSCANA","SI":"TOSCANA",
    "PG":"UMBRIA","TR":"UMBRIA",
    "AN":"MARCHE","AP":"MARCHE","FM":"MARCHE","MC":"MARCHE","PU":"MARCHE",
    "FR":"LAZIO","LT":"LAZIO","RI":"LAZIO","RM":"LAZIO","VT":"LAZIO",
    "AQ":"ABRUZZO","CH":"ABRUZZO","PE":"ABRUZZO","TE":"ABRUZZO",
    "CB":"MOLISE","IS":"MOLISE",
    "AV":"CAMPANIA","BN":"CAMPANIA","CE":"CAMPANIA","NA":"CAMPANIA","SA":"CAMPANIA",
    "BA":"PUGLIA","BT":"PUGLIA","BR":"PUGLIA","FG":"PUGLIA","LE":"PUGLIA","TA":"PUGLIA",
    "MT":"BASILICATA","PZ":"BASILICATA",
    "CZ":"CALABRIA","CS":"CALABRIA","KR":"CALABRIA","RC":"CALABRIA","VV":"CALABRIA",
    "AG":"SICILIA","CL":"SICILIA","CT":"SICILIA","EN":"SICILIA","ME":"SICILIA","PA":"SICILIA","RG":"SICILIA","SR":"SICILIA","TP":"SICILIA",
    "CA":"SARDEGNA","NU":"SARDEGNA","OR":"SARDEGNA","SS":"SARDEGNA","SU":"SARDEGNA"
  };
  return map[sigla.toUpperCase()] || "";
}

function firstNonEmptyValue() {
  for (var i = 0; i < arguments.length; i++) {
    var v = arguments[i];
    if (v !== null && v !== undefined && String(v).trim() !== "") return String(v).trim();
  }
  return "";
}

function pickPathValue(obj, path) {
  if (!obj) return "";
  var cur = obj;
  for (var i = 0; i < path.length; i++) {
    if (!cur || typeof cur !== "object" || !(path[i] in cur)) return "";
    cur = cur[path[i]];
  }
  if (cur === null || cur === undefined) return "";
  if (typeof cur === "object") return "";
  return String(cur).trim();
}

function isEmptyPlainObject(obj) {
  if (!obj || typeof obj !== "object") return true;
  for (var k in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, k)) {
      var v = obj[k];
      if (v !== "" && v !== null && v !== undefined) return false;
    }
  }
  return true;
}

function flattenMauticFieldsAll(src) {
  var out = {};
  if (!src || typeof src !== "object") return out;

  for (var k in src) {
    if (!Object.prototype.hasOwnProperty.call(src, k)) continue;
    var v = src[k];

    if (v && typeof v === "object" && "value" in v) {
      out[k] = v.value == null ? "" : String(v.value).trim();
    } else if (typeof v !== "object") {
      out[k] = v == null ? "" : String(v).trim();
    }
  }

  return out;
}

function normText(s) {
  return String(s || "").trim().toLowerCase();
}

function normalizeString(value) {
  return String(value || "").trim();
}

function normPhone(s) {
  var x = String(s || "").trim();
  if (!x) return "";
  x = x.replace(/[^\d+]/g, "");
  x = x.replace(/^(\+|00)?39/, "");
  x = x.replace(/[^\d]/g, "");
  return x;
}

function normalizeMauticCustomerData(raw) {
  var candidate = raw;

  if (Array.isArray(candidate)) {
    candidate = candidate.length ? candidate[0] : null;
  }

  if (!candidate || typeof candidate !== "object") {
    return { raw: String(candidate || "") };
  }

  if (candidate.contact && typeof candidate.contact === "object") {
    candidate = candidate.contact;
  } else if (candidate.contacts && Array.isArray(candidate.contacts) && candidate.contacts.length) {
    candidate = candidate.contacts[0];
  } else if (candidate.data && typeof candidate.data === "object") {
    candidate = candidate.data;
    if (Array.isArray(candidate)) candidate = candidate.length ? candidate[0] : null;
    if (candidate && candidate.contact && typeof candidate.contact === "object") candidate = candidate.contact;
  }

  if (!candidate || typeof candidate !== "object") return {};

  var fieldsAll = {};
  if (candidate.fields && candidate.fields.all && typeof candidate.fields.all === "object") {
    fieldsAll = flattenMauticFieldsAll(candidate.fields.all);
  }

  var out = {
    id: firstNonEmptyValue(
      candidate.id,
      fieldsAll.id
    ),
    email: firstNonEmptyValue(
      candidate.email,
      fieldsAll.email,
      pickPathValue(candidate, ["fields", "core", "email", "value"]),
      pickPathValue(candidate, ["fields", "all", "email", "value"])
    ),
    firstname: firstNonEmptyValue(
      candidate.firstname,
      candidate.firstName,
      fieldsAll.firstname,
      fieldsAll.first_name,
      pickPathValue(candidate, ["fields", "core", "firstname", "value"]),
      pickPathValue(candidate, ["fields", "all", "firstname", "value"])
    ),
    lastname: firstNonEmptyValue(
      candidate.lastname,
      candidate.lastName,
      fieldsAll.lastname,
      fieldsAll.last_name,
      fieldsAll.cognome,
      pickPathValue(candidate, ["fields", "core", "lastname", "value"]),
      pickPathValue(candidate, ["fields", "all", "lastname", "value"])
    ),
    fullname: firstNonEmptyValue(
      candidate.fullname,
      candidate.name,
      fieldsAll.fullname,
      fieldsAll.full_name
    ),
    phone: firstNonEmptyValue(
      candidate.phone,
      candidate.mobile,
      fieldsAll.phone,
      fieldsAll.mobile,
      fieldsAll.telefono,
      pickPathValue(candidate, ["fields", "all", "phone", "value"]),
      pickPathValue(candidate, ["fields", "all", "mobile", "value"])
    ),
    codice_fiscale: firstNonEmptyValue(
      candidate.codice_fiscale,
      candidate.codicefiscale,
      fieldsAll.codice_fiscale,
      fieldsAll.codicefiscale,
      fieldsAll.cf
    ),
    company: firstNonEmptyValue(
      candidate.company,
      fieldsAll.company
    )
  };

  if (!out.fullname) {
    out.fullname = String((out.firstname + " " + out.lastname).trim());
  }

  if (isEmptyPlainObject(out)) {
    var fallback = {};
    for (var k in candidate) {
      if (!Object.prototype.hasOwnProperty.call(candidate, k)) continue;
      var v = candidate[k];
      if (v !== null && v !== undefined && typeof v !== "object") {
        fallback[k] = String(v).trim();
      }
    }
    if (!isEmptyPlainObject(fallback)) return fallback;
  }

  return out;
}

function mauticCustomerDataToBaseFields(data) {
  var d = data && typeof data === "object" ? data : {};
  var out = {};

  if (d.id) out["ID Cliente"] = d.id;
  if (d.lastname) out["Cognome"] = d.lastname;
  if (d.firstname) out["Nome"] = d.firstname;
  if (d.email) out["Email"] = d.email;
  if (d.phone) out["Telefono"] = d.phone;
  if (d.codice_fiscale) out["Codice Fiscale"] = d.codice_fiscale;

  return out;
}

function extractMauticCandidates(raw) {
  if (!raw) return [];

  if (Array.isArray(raw)) return raw;

  if (raw.contacts && Array.isArray(raw.contacts)) return raw.contacts;

  if (raw.data && Array.isArray(raw.data)) return raw.data;

  if (raw.data && raw.data.contacts && Array.isArray(raw.data.contacts)) return raw.data.contacts;

  if (raw.contact && typeof raw.contact === "object") return [raw.contact];

  if (raw.data && raw.data.contact && typeof raw.data.contact === "object") return [raw.data.contact];

  if (typeof raw === "object") return [raw];

  return [];
}


function testsendWhatsAppCloudTemplateMessage_ ()
{
  sendWhatsAppCloudTemplateMessage("393482639796", "2025_imdb_anteprima_borgogna_2026_v_1", "it", ["Max"]);
}

/**
 * @param {string} to - The recipient's phone number in international format (e.g., "393482639796" without the plus sign).
 * @param {string} templateName - The name of the pre-approved template (e.g., "hello_world").
 * @param {string} languageCode - The language code (e.g., "en_US").
 * @param {Array} templateParameters - Optional array of text parameters for the template body.
 * @return {string} The API response.
 */

function sendWhatsAppCloudTemplateMessage(to, templateName, languageCode, templateParameters) {

  var accessToken = getScriptProp_('WHATSAPP_ACCESS_TOKEN');
  var phoneNumberId = getScriptProp_('WHATSAPP_PHONE_NUMBER_ID');

  // Construct the API endpoint URL.
  var url = 'https://graph.facebook.com/v22.0/' + phoneNumberId + '/messages';

  // Convert the array of "parameterName parameterValue" strings into the proper format.

  if (templateParameters)
  {
    var paramsArray = templateParameters.map(function(paramStr) {
      // Split the string on whitespace.
      var parts = paramStr.split(' ');
      var paramName = parts.shift(); // Take the first part as the parameter name.
      var paramValue = parts.join(' '); // The rest is the parameter value.
      return {
        type: "text",
        parameter_name: paramName,
        text: paramValue
      };
    });

    // Build the payload.
    var payload = {
      messaging_product: "whatsapp",
      to: to,
      type: "template",
      template: {
        name: templateName,
        language: {
          code: languageCode
        },
        components: [{
          type: "body",
          parameters: paramsArray
        }]
      }
    };
  }
  else
  {

    // Build the payload.
    var payload = {
      messaging_product: "whatsapp",
      to: to,
      type: "template",
      template: {
        name: templateName,
        language: {
          code: languageCode
        }
      }
    };
  }

  Logger.log("Payload: " + JSON.stringify(payload));

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: {
      "Authorization": "Bearer " + accessToken
    },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("WhatsApp API response: " + response.getContentText());
    return response;
  } catch (e) {
    Logger.log("Error sending WhatsApp message: " + e);
    //throw e;
    return response;
  }
}