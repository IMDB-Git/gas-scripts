// SCOT srl


var scotBaseURL = "https://api.portalescotsrl.it";
var scotUsername = "imdb";
var scotPassword = "Tq75mF6pbGaYSITI";
const inviaSpedizioni_DaPortale = false;

/**
 * Requests an authentication token from the SCOT portal.
 *
 * @param {string} username - The username credential.
 * @param {string} password - The password credential.
 * @return {string|null} - The token if the request succeeds, otherwise null.
 */
function getScotToken_(username, password) {
  // Replace with the correct base URL of the SCOT portal.
  var endpoint = "/api/token/"; 
  var url = scotBaseURL + endpoint;
  
  // Build the payload with credentials.
  var payload = {
    username: username,
    password: password
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


/**
 * Invia un ordine in uscita al portale SCOT (/api/uscite/).
 *
 * @param {string} token      Bearer token per l'autenticazione.
 * @param {string} orderId    order_id (maxLength:10).
 * @param {string} clientId   client ID (maxLength:5).
 * @param {Object} header     Oggetto header con almeno:
 *    business_name (string, max 30),
 *    document_date (ISO date-time),
 *    attachment (boolean)
 *    e le altre proprietà facoltative (address, location, province, ...).
 * @param {Array}  rows       Array di righe (prodotti), ciascuna con:
 *    id (integer),
 *    code (string, max 20),
 *    quantity (integer),
 *    udc (string, max 20, opzionale),
 *    batch (string, max 20, opzionale).
 * @param {Array}  [files]    (Opzionale) Array di file { file_name (max10), payload (base64) }.
 * @return {Object|null}      L’oggetto JSON di risposta se HTTP 200, altrimenti null.
 */
function scotOrdiniUscita_(orderId, clientId, header, rows, clienteNome, campagnaNome, files = null) {
  var url = scotBaseURL + "/api/uscite/";
  
  // 1) recupera il token
  var token = getScotToken_(scotUsername, scotPassword);

  if (!token) { 
    Logger.log("Impossibile ottenere il token"); 
    return; 
  }

  // Costruisci il payload completo
  var payload = {
    order_id: orderId,
    client: clientId,
    header: header,
    rows: rows
  };
  if (files && Array.isArray(files) && files.length) {
    payload.files = files;
  }

  Logger.log(payload);
  
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
    var code = response.getResponseCode();
    var json = JSON.parse(response.getContentText());

    var htmlJSON = IMDBCommonLibs.formatJsonGeneric(payload);
    
    if (code === 200) 
    {
      Logger.log("Ordine Uscita: " + orderId + " inviato con successo: %s", htmlJSON);

      IMDBCommonLibs.sendEmailViaSMTP(htmlJSON, "ordini@ilmassimodelbere.it", "SPEDIZIONE: " + orderId + " Cliente: " + clienteNome + " Campagna: " + campagnaNome, "IMDB Logistics");
      return json;
    } 
    else 
    {
      Logger.log("Errore invio ordine: " + orderId + " (%s): %s", code, JSON.stringify(json));
      return null;
    }
  } catch (e) {
    Logger.log("Eccezione invio ordine: " + orderId + " : %s", e);
    return null;
  }
}

/* ========== Esempio di utilizzo ========== */
function testscotOrdiniUscita_() {
  // 1) recupera il token
  var token = getScotToken_(scotUsername, scotPassword);

  if (!token) { 
    Logger.log("Impossibile ottenere il token"); 
    return; 
  }
  
  // 2) prepara i dati di header
  var header = {
    business_name: "La Mia Azienda SRL",
    document_date: (new Date()).toISOString(),
    attachment: false,
    address: "Via Roma 1",
    location: "Milano",
    province: "MI",
    zip_code: "20100",
    nation: "IT",
    urgent: false,
    delivery_date: (new Date(new Date().getTime() + 3*24*3600*1000)).toISOString(),
    appointment: false,
    email: "info@azienda.it",
    tel_reference: "0234567890",
    carrier_note: "Consegna al piano",
    warehouse_note: "Att.n. imballi fragili",
    cash_on_delivery_value: 0.0,
    cash_on_delivery_type: ""
  };
  
  // 3) prepara le righe (prodotti)
  var rows = [
    { id: 1, code: "ABABR03", quantity: 2 },
    { id: 2, code: "ABABR14", quantity: 1 }
  ];
  
  // 4) (opzionale) prepara eventuali file in base64
  //var files = [
  //  { file_name: "doc1.pdf", payload: "JVBERi0xLjQKJ..." }
  //];
  
  // 5) invoca la funzione
  var orderName = "OD" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMddmmss");
  var result = scotOrdiniUscita_(token, orderName, "MDB", header, rows);
  if (result) 
  {
    Logger.log("Succeded! Result: %s", JSON.stringify(result));
  }
  else
  {
    Logger.log("ERROR! Result: %s", JSON.stringify(result));
  }
}

/**
 * Recupera lo stato di un ordine uscita dal portale SCOT.
 *
 * Endpoint: /api/uscite/stato/
 * Richiede un JSON:
 *   { order_id: string, client: string }
 * Risponde con:
 *   {
 *     order_id: string,
 *     status: integer,
 *     acquisition_date?: string,
 *     conclusion_date?: string,
 *     rows: [
 *       { id, code, row_number, quantity_required, quantity_processed }, ...
 *     ]
 *   }
 *
 * @param {string} token    Bearer token ottenuto da getScotToken_().
 * @param {string} orderId  order_id (max 10 caratteri).
 * @param {string} clientId client ID (max 5 caratteri).
 * @return {Object|null}    L’oggetto JSON di risposta se HTTP 200, altrimenti null.
 */
function scotOrdiniUscita_Stato_(token, orderId, clientId) {
  var url = scotBaseURL + "/api/uscite/stato/";
  
  // Costruisci il body della richiesta
  var payload = {
    order_id: orderId,
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

/* ========== Esempio di utilizzo ========== */
function testscotOrdiniUscita_Stato_() {
  // 1) recupera il token
  var token = getScotToken_(scotUsername, scotPassword);
  if (!token) {
    Logger.log("Token non ottenuto");
    return;
  }
  
  // 2) chiama lo stato
  var stato = scotOrdiniUscita_Stato_(token, "5E01003342", "MDB");
  if (!stato) {
    Logger.log("Recupero stato fallito");
    return;
  }
  
  // 3) log dei dettagli
  Logger.log("Order ID: %s", stato.order_id);
  Logger.log("Status code: %s", scotGetOrderStatusDescription_(stato.status));
  if (stato.acquisition_date)    Logger.log("Acquired: %s", stato.acquisition_date);
  if (stato.conclusion_date)     Logger.log("Concluded: %s", stato.conclusion_date);
  
  if (stato.rows)
  {  
      stato.rows.forEach(function(r) {
      Logger.log(
        "Riga %s (ID %s): code=%s, req=%s, proc=%s",
        r.row_number, r.id, r.code, r.quantity_required, r.quantity_processed
      );
    });
  }
}

/**
 * Restituisce la descrizione testuale per un codice stato ordine SCOT.
 *
 * @param {number|string} code - Il codice numerico dello stato.
 * @return {string} La descrizione corrispondente, o "Sconosciuto" se il codice non è nella mappa.
 */
function scotGetOrderStatusDescription_(code) {
  // Assicuriamoci di trattare il codice come numero
  var key = typeof code === 'string' ? parseInt(code, 10) : code;
  
  var statusMap = {
    0:    "In Elaborazione",
    10:   "In Acquisizione",
    20:   "Acquisito (non valido)",
    30:   "Acquisito (valido)",
    50:   "Da Elaborare (Attivato)",
    55:   "In Elaborazione",
    60:   "Evadibile (elaborato)",
    62:   "Elaborato - Da Preparare",
    63:   "In Preparazione",
    65:   "Attesa esecuzione rimpiazzi",
    68:   "In Preparazione attività",
    69:   "Non Prelevabile",
    70:   "Prelevabile",
    80:   "Prelevabile (senza impegni)",
    90:   "In Prelievo",
    95:   "In Viaggio",
    100:  "Parzialmente Prelevato",
    110:  "Prelevato",
    180:  "Pesato",
    200:  "Spuntato",
    500:  "Concluso",
    600:  "Annullato",
    1000: "Inevadibile",
    5000: "Aggregato a Lista"
  };
  
  return statusMap.hasOwnProperty(key) ? statusMap[key] : "Sconosciuto (" + key + ")";
}

function findAndInsertID_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet(); // Get the active sheet


  // Get the currently selected range of cells on the active sheet
  var activeRange = currentSheet.getActiveRange();

  // Retrieve all values in the selected range as a 2D array
  var selectedValues = activeRange.getValues();

  // Get the starting row of the active range
  var startingRow = activeRange.getRow();
  
  var activeRow = 1;

  // Iterate over each row in the selected range
  for (var j = 0; j < selectedValues.length; j++) 
  {
    activeRow = startingRow + j;

    // Get the value from Column A of the current row (assuming Column A is the first column)
    var surname = currentSheet.getRange(activeRow, 2).getValue();

    if (!surname) {
      SpreadsheetApp.getUi().alert("Errore: Nessun cognome trovato nella cella selezionata.");
      Logger.log("Errore: Nessun cognome trovato nella cella selezionata.");
      continue;
    }

    var indirizziSheet = ss.getSheetByName("Indirizzi spedizione"); // Get the "Indirizzi spedizione" sheet
    if (!indirizziSheet) {
      SpreadsheetApp.getUi().alert("Errore: Il foglio 'Indirizzi spedizione' non esiste.");
      Logger.Log("Errore: Il foglio 'Indirizzi spedizione' non esiste.");
      return;
    }

    var data = indirizziSheet.getDataRange().getValues(); // Get all data from the sheet
    var matchingRows = [];

    for (var i = 1; i < data.length; i++) { // Start from row 2 (skip headers)
      if (data[i][5] === surname) { // Check if surname matches column F (index 5)
        matchingRows.push({ id: data[i][0], row: i + 1 }); // Store ID and row number
      }
    }

    if (matchingRows.length === 0) {
      SpreadsheetApp.getUi().alert("Errore: Nessuna corrispondenza trovata per il cognome '" + surname + "'.");
      Logger.log("Errore: Nessuna corrispondenza trovata per il cognome '" + surname + "'.");
      continue;
    } 
    else if (matchingRows.length === 1) 
    {
      // Only one match, insert the ID directly
      currentSheet.getRange(activeRow, 2).setValue(matchingRows[0].id);
      Logger.log("ID inserito con successo: " + matchingRows[0].id);
    } 
    else 
    {
      // Multiple matches, ask the user to select
      var ui = SpreadsheetApp.getUi();
      var choices = matchingRows.map(function(row, index) {
        return (index + 1) + ") ID: " + row.id + " (Riga " + row.row + ")";
      }).join("\n");

      var response = ui.prompt("Più corrispondenze trovate per '" + surname + "'. Seleziona l'ID corretto:", choices, ui.ButtonSet.OK_CANCEL);
      
      if (response.getSelectedButton() === ui.Button.OK) 
      {
        var selectedIndex = parseInt(response.getResponseText(), 10) - 1;
        if (selectedIndex >= 0 && selectedIndex < matchingRows.length) {
          currentSheet.getRange(activeRow, 2).setValue(matchingRows[selectedIndex].id);
          Logger.log("ID inserito con successo: " + matchingRows[selectedIndex].id);
        } else {
          SpreadsheetApp.getUi().alert("Errore: Selezione non valida.");
        }
      }
    }
  }
}

function getEmailRiepilogoColumn_(currentSheet, targetHeader, headerRow)
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

function inviaOrderReceived_() 
{
  var ui = SpreadsheetApp.getUi();
  
  // Get the current active Google Spreadsheet (IMDB_Ordini)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();  // "Current" tab
  // Get the current tab name
  var currentTabName = currentSheet.getName();
  
  var orderReceivedCol = getEmailRiepilogoColumn_(currentSheet, 'Mail "Order Received"', 33);

  // Get the currently selected range of cells on the active sheet
  var activeRange = currentSheet.getActiveRange();

  // Retrieve all values in the selected range as a 2D array
  var selectedValues = activeRange.getValues();

  // Get the starting row of the active range
  var startingRow = activeRange.getRow();
  
  var activeRowNumber = 1;

  // Iterate over each row in the selected range
  for (var i = 0; i < selectedValues.length; i++) 
  {
    activeRow = startingRow + i;

    // Get the value from Column A of the current row (assuming Column A is the first column)
    var colAValue = currentSheet.getRange(activeRow, 1).getValue();
    
    // Check if Column A starts with "Cliente"
    if (!colAValue.startsWith("Cliente")) {
      // If it doesn't start with "Cliente", abort and show an error message
      SpreadsheetApp.getUi().alert("Errore: Controlla di essere in un foglio di ordini nella riga relativa ad un cliente.");
      return; // Abort the script
    }

    var customerName = currentSheet.getRange(activeRow, 5).getValue();
    var customerSurname = currentSheet.getRange(activeRow, 4).getValue();
    var customerEmail = currentSheet.getRange(activeRow, 8).getValue();
    var customerPhone = currentSheet.getRange(activeRow, 9).getValue();

    if (customerPhone = "#N/A")
      customerPhone = "";

    if (currentSheet.getRange(activeRow, orderReceivedCol).getValue())
    {
      var responseDo = ui.alert("Hai già inviato Order Received a " + customerSurname + ", vuoi mandarlo di nuovo?", ui.ButtonSet.YES_NO);
      if (responseDo === ui.Button.NO) 
      {
        return;
      }
    }

    // Send Order Received form
    submitFormOrderReceived_(customerEmail, customerName, customerSurname, customerPhone, currentTabName, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"))

    // Update row
    currentSheet.getRange(activeRow, orderReceivedCol).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));


  }
}

function creaRiepilogo() {
  var ui = SpreadsheetApp.getUi();

  // Get the current active Google Spreadsheet (IMDB_Ordini)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();  // "Current" tab
  // Get the current tab name
  var currentTabName = currentSheet.getName();

  var riepilogoSentCol = getEmailRiepilogoColumn_(currentSheet, 'Mail "Riepilogo Ordine"', 33);
  var orderReceivedCol = getEmailRiepilogoColumn_(currentSheet, 'Mail "Order Received"', 33);
  var notificationOrdineCol = getEmailRiepilogoColumn_(currentSheet, 'Notifica Ordine', 33);
  var idAziendaCol = getEmailRiepilogoColumn_(currentSheet, 'ID Azienda', 33);

  // Get the currently selected range of cells on the active sheet
  var activeRange = currentSheet.getActiveRange();

  // Retrieve all values in the selected range as a 2D array
  var selectedValues = activeRange.getValues();

  // Get the starting row of the active range
  var startingRow = activeRange.getRow();

  var activeRowNumber = 1;

  // Iterate over each row in the selected range
  for (var i = 0; i < selectedValues.length; i++) 
  {
    activeRow = startingRow + i;

    // Get the value from Column A of the current row (assuming Column A is the first column)
    var colAValue = currentSheet.getRange(activeRow, 1).getValue();

    // Check if Column A starts with "Cliente"
    if (!colAValue.startsWith("Cliente")) {
      SpreadsheetApp.getUi().alert("Errore: Controlla di essere in un foglio di ordini nella riga relativa ad un cliente.");
      return;
    }

    // Aggiorna calcolo margini
    //calcolaLogisticaDaSelezione();

    var customerName = currentSheet.getRange(activeRow, 5).getValue();
    var customerSurname = currentSheet.getRange(activeRow, 4).getValue();
    var customerEmail = currentSheet.getRange(activeRow, 8).getValue();

    if (customerEmail === "#N/A" || customerEmail === "")
    {
      ui.alert("Email non valida: " + customerEmail);
      continue;
    }

    // Cliente azienda / HORECA?
    var idAziendaValue = idAziendaCol ? currentSheet.getRange(activeRow, idAziendaCol).getValue() : "";
    var isAzienda = false;

    if (idAziendaValue !== "" && idAziendaValue !== null) {
      var idAziendaNum = Number(String(idAziendaValue).replace(",", ".").trim());
      isAzienda = !isNaN(idAziendaNum);
    }

    var labelListino = isAzienda ? "Listino HORECA" : "Listino IMDB";
    var labelTotaleListino = isAzienda ? "Totale HORECA" : "Totale IMDB";
    var productPriceRow = isAzienda ? 9 : 14;

    // Update the relevant portion where we assign the phone number:
    var telefonoFormatted = "";
    if (String(currentSheet.getRange(activeRow, 9).getValue() || "").length)
      telefonoFormatted = IMDBCommonLibs.formatPhoneNumber(currentSheet.getRange(activeRow, 9).getValue());
    else
      telefonoFormatted = IMDBCommonLibs.formatPhoneNumber(currentSheet.getRange(activeRow, 10).getValue());

    if (telefonoFormatted === "#N/A")
      telefonoFormatted = "";

    // Send OrderReceived email
    var cellVal;
    var sendOrderReceived = true;

    cellVal = String(currentSheet.getRange(activeRow, orderReceivedCol).getValue() || '').trim();
    if (cellVal.toUpperCase() === 'SKIP') {
      sendOrderReceived = false;
    } else if (cellVal) {
      sendOrderReceived = (ui.alert("Hai già inviato Order Received a " + customerSurname + ", vuoi mandarlo di nuovo?",
        ui.ButtonSet.YES_NO
      ) === ui.Button.YES);
    }

    if (sendOrderReceived)
    {
      submitFormOrderReceived_(customerEmail, customerName, customerSurname, telefonoFormatted, currentTabName, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
      currentSheet.getRange(activeRow, orderReceivedCol).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
    }

    // Sent Whatsapp notification
    var sendNotificaOrdine = true;

    cellVal = String(currentSheet.getRange(activeRow, notificationOrdineCol).getValue() || '').trim();
    if (cellVal.toUpperCase() === 'SKIP') {
      sendNotificaOrdine = false;
    } else if (cellVal) {
      sendNotificaOrdine = (ui.alert(
        "Hai già inviato la notifica WhatsApp dell'ordine a " + customerSurname + ", vuoi mandarlo di nuovo?",
        ui.ButtonSet.YES_NO
      ) === ui.Button.YES);
    }

    if (sendNotificaOrdine)
    {
      var templateOptions = [
        "nome " + customerName,
        "campagna " + currentSheet.getName()
      ];

      var waResponse = sendWhatsAppCloudTemplateMessage_("39" + telefonoFormatted, "20250411_imdb_order_received_v1_0", "it", templateOptions);
      var waResponseObj = JSON.parse(waResponse);

      if (waResponseObj.error) 
      {
        SpreadsheetApp.getUi().alert("ERRORE: messaggio NON inviato a " + customerSurname + " " + customerName + "!");
      }
      else
      {
        Logger.log("Messaggio inviato a " + customerSurname + " " + customerName + "!");
        currentSheet.getRange(activeRow, notificationOrdineCol).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
      }
    }

    cellVal = String(currentSheet.getRange(activeRow, riepilogoSentCol).getValue() || '').trim();

    if (cellVal.toUpperCase() === 'SKIP') {
      continue;
    }

    if (cellVal) {
      var responseDo = ui.alert(
        "Hai già inviato il riepilogo a " + customerSurname + ", vuoi mandarlo di nuovo?",
        ui.ButtonSet.YES_NO
      );
      if (responseDo === ui.Button.NO) {
        continue;
      }
    }

    // Scan Row 1 from column AI onwards to get the product details
    var startColumn = getEmailRiepilogoColumn_(currentSheet, 'Codice bottiglia', 1) + 2;
    var rangeProdotti = [];
    var tableRows = [
      ["Codice", "Produttore", "Prodotto", "Quantità", labelListino, labelTotaleListino]
    ];
    var totaleBottiglie = 0;
    var totaleListino = 0;

    for (var col = startColumn; (currentSheet.getRange(1, col).getValue() != ""); col++) 
    {
      var productCode = currentSheet.getRange(1, col).getValue();
      var productMaker = currentSheet.getRange(2, col).getValue();
      var productName = currentSheet.getRange(3, col).getValue();
      var quantity = currentSheet.getRange(activeRow, col).getValue();
      var unitPrice = currentSheet.getRange(productPriceRow, col).getValue();

      if (quantity && quantity > 0) {
        rangeProdotti.push(col);
        var totalRow = quantity * unitPrice;
        tableRows.push([productCode, productMaker, productName, quantity, IMDBCommonLibs.formatPrice(unitPrice), IMDBCommonLibs.formatPrice(totalRow)]);
        totaleBottiglie += quantity;
        totaleListino += totalRow;
      }
    }

    // Add Totale Offerta / Totale da pagare
    var totaleOffertaValue = currentSheet.getRange(activeRow, 16).getValue(); // Column P
    var totaleOfferta = IMDBCommonLibs.formatPrice(totaleOffertaValue);
    var totaleDaPagareValue = currentSheet.getRange(activeRow, 19).getValue(); // Column S
    var totaleDaPagare = IMDBCommonLibs.formatPrice(totaleDaPagareValue);

    // Add Totale row
    tableRows.push(["Totale", "", "", totaleBottiglie, "", IMDBCommonLibs.formatPrice(totaleListino)]);
    tableRows[tableRows.length - 1] = tableRows[tableRows.length - 1].map(function(cell) { return "<b>" + cell + "</b>"; });

    // Prepare the user prompt
    var userPrompt = "<div style='font-family: Tahoma;'>";

    userPrompt += "<p style='font-size: 24px;'>Ciao " + customerName + ",<br><br>riepiloghiamo di seguito il tuo ordine e ti confermiamo la disponibilità delle bottiglie di seguito indicate:</p><br>";
    userPrompt += "<p style='font-size: 24px;'><i>Ti preghiamo di verificare con cura che <b>l'ordine sia corretto</b>, sia da un punto di vista delle bottiglie proposte, che dei conteggi, in modo da <b>evitare problematiche</b> in fase di consegna e di fatturazione</i></p><br>";

    userPrompt += "<span style='font-size: 24px; font-weight: bold;'><h2>Campagna: " + currentTabName + "</h2></span>";

    userPrompt += "<table border='1' cellpadding='5' cellspacing='0' style='font-family: Tahoma;'>";
    tableRows.forEach(function(row, index) {
      userPrompt += "<tr" + (index === 0 ? " style='font-weight: bold;'" : "") + ">";
      row.forEach(function(cell) {
        userPrompt += "<td>" + cell + "</td>";
      });
      userPrompt += "</tr>";
    });
    userPrompt += "</table><br><br>";


    // Voucher
    var voucherValue = currentSheet.getRange(activeRow, 18).getValue(); // Column R

    if (isAzienda) {
      userPrompt += "<span style='font-size: 24px;'>" + labelTotaleListino + ": " + IMDBCommonLibs.formatPrice(totaleListino) + " + IVA</span><br><br>";
      if (voucherValue) 
      {
        userPrompt += "<span style='font-size: 24px; font-weight: bold; color: red;'>Sconto: -" + IMDBCommonLibs.formatPrice(voucherValue / 1.22) + " + IVA</span><br><br>";
      }

      var spedizioneAziendaValue = currentSheet.getRange(activeRow, 17).getValue(); // Column Q
      if (spedizioneAziendaValue) {
        userPrompt += "<span style='font-size: 24px; '><b>Spedizione: " + IMDBCommonLibs.formatPrice(spedizioneAziendaValue / 1.22) + " + IVA</b></span><br><br>";
      }


      userPrompt += "<span style='font-size: 24px; font-weight: bold; '>Totale offerta: " + IMDBCommonLibs.formatPrice(totaleDaPagareValue / 1.22) + " + IVA</span><br><br>";

      userPrompt += "<span style='font-size: 30px; font-weight: bold; color: green; text-transform: uppercase;'>Totale ordine: " + totaleDaPagare + " <span style='font-size: 22px; text-transform: none;'>(IVA inclusa)</span></span><br><br>";
    } 
    else 
    {
      userPrompt += "<span style='font-size: 24px;'>Totale IMDB: <s>" + IMDBCommonLibs.formatPrice(totaleListino) + "</s></span><br><br>";
      var risparmio = totaleListino - totaleOffertaValue;
      userPrompt += "<span style='font-size: 24px; font-weight: bold; color: red;'>Risparmio: -" + IMDBCommonLibs.formatPrice(risparmio) + "</span><br><br>";
      userPrompt += "<span style='font-size: 24px; font-weight: bold; '>Totale offerta: " + totaleOfferta + "</span><br><br>";

      var spedizione = currentSheet.getRange(activeRow, 17).getValue(); // Column Q
      if (spedizione) {
        userPrompt += "<span style='font-size: 24px; '><b>Spedizione: " + IMDBCommonLibs.formatPrice(spedizione) + "</b></span><br><br>";
      }

      if (voucherValue) 
      {
        userPrompt += "<span style='font-size: 24px; font-weight: bold; color: red;'>Voucher: -" + IMDBCommonLibs.formatPrice(voucherValue) + "</span><br><br>";
      }
      userPrompt += "<span style='font-size: 30px; font-weight: bold; color: green; text-transform: uppercase;'>Totale ordine: " + totaleDaPagare + "</span><br><br>";
    }

    // Se VIP Club o PAGATO=OK non mandare pagamento
    if (currentSheet.getRange(activeRow, 14).getValue() === "")
    {
      if (currentSheet.getRange(activeRow, 12).getValue() === "SILVER" || 
          currentSheet.getRange(activeRow, 12).getValue() === "GOLD" ||
          currentSheet.getRange(activeRow, 12).getValue() === "DIAMOND" ||
          currentSheet.getRange(activeRow, 12).getValue() === "BLACK") 
      {
        userPrompt += "<span style='font-size: 30px; font-weight: bold; color: green; text-transform: uppercase;'>Totale da scalare dal credito: " + totaleDaPagare + "</span><br>";
        userPrompt += "<span style='font-size: 20px; color: green;'><i>(ti arriverà un'email separata con i conteggi del credito residuo ed eventuali integrazioni)</i></span><br><br>";
      }
      else
      {
        userPrompt += "<span style='font-size: 24px; font-weight: bold;'><H2>PAGAMENTO</H2></span>";
        userPrompt += "<p style='font-size: 24px;'>La velocità dei pagamenti è fondamentale per il buon funzionamento delle nostre campagne di acquisti!<br>Infatti, le bottiglie partiranno dal produttore dopo che avremo raccolto tutti i soldi.<br><br>Per cui, per velocizzare l'arrivo delle bottiglie, ti chiediamo di procedere al più presto al pagamento della cifra indicata.</p><br>";

        var campagnaName = currentTabName;
        var Cognome = currentSheet.getRange(activeRow, 4).getValue();

        userPrompt += "<p style='font-size: 24px; font-weight: bold; font-style: italic;'>Intestatario: IMDB s.r.l.s.<br>Banca: Credit Agricole<br>Sportello: Cassolnovo<br>IBAN: IT35O0623055720000030842123<br>Causale: " + Cognome + " - \"" + campagnaName + "\"<br><br>Totale da pagare: " + totaleDaPagare + "</p><br>";
        userPrompt += "<p style='font-size: 24px;'>e di mandarci cortesemente la ricevuta per accelerare le procedure.</p>";
      }
    }

    userPrompt += "<span style='font-size: 24px; font-weight: bold;'><H2>DATI DI SPEDIZIONE E DI FATTURAZIONE</H2></span>";
    userPrompt += "<p style='font-size: 24px;'><b>Se è il tuo primo ordine, o se è cambiato qualcosa nella consegna,</b> ti preghiamo di compilare il seguente modulo per indicare i dati di spedizione (e di eventuale fatturazione).</p><br>";
    userPrompt += "<p style='font-size: 24px;'><a href='https://www.ilmassimodelbere.it/Mautic/imdb-dati-di-acquisto'>Modulo dati di spedizione</a></p><br>";
    userPrompt += "<p style='font-size: 24px;'><i>PS: Il <b>Codice Fiscale è OBBLIGATORIO</b> anche per i privati!</i></p><br>";

    userPrompt += "<p style='font-size: 24px;'>Un caro saluto</p>";
    userPrompt += "<p style='font-size: 24px;'>Massimo</p><br><br>";

    userPrompt += "</div>";

    // Define Subject
    var subject = "";
    if (currentSheet.getRange(activeRow, 12).getValue() || currentSheet.getRange(activeRow, 14).getValue() != "")
      subject = "Conferma d'ordine - " + customerSurname + " " + customerName + " (Campagna " + currentTabName + ")";
    else
      subject = "Conferma d'ordine e richiesta di pagamento - " + customerSurname + " " + customerName + " (Campagna " + currentTabName + ")";

    // Show the user prompt
    var htmlContent = HtmlService.createHtmlOutput("<p style='font-size: 30px;'>Subject: " + subject + "</p>" + userPrompt).setWidth(600).setHeight(400);

    // Prepare email content
    var recipientEmail = customerEmail; 
    var bccEmail = "ordini@ilmassimodelbere.it";
    var senderName = "Il Massimo del Bere";

    ui.showModelessDialog(htmlContent, "Anteprima email");

    var response = ui.alert("Vuoi procedere con l'invio dell'email a: " + customerEmail + "?", ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) 
    {
      var emailSent = IMDBCommonLibs.sendEmailViaSMTP(userPrompt, recipientEmail, subject, senderName, '', bccEmail);

      if (!emailSent)
      {
        Logger.log("Email sent to " + customerSurname + " " + customerName);
        currentSheet.getRange(activeRow, riepilogoSentCol).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
      }
      else
      {
        ui.alert("Problema (" + emailSent + "); NON ho inviato l'email a " + customerSurname + " " + customerName);
      }
    } 
    else if (response === ui.Button.NO) 
    {
      Logger.log("Email sending canceled by the user.");
    }

    if (currentSheet.getRange(activeRow, 12).getValue() === "SILVER" || 
        currentSheet.getRange(activeRow, 12).getValue() === "GOLD" ||
        currentSheet.getRange(activeRow, 12).getValue() === "DIAMOND" ||
        currentSheet.getRange(activeRow, 12).getValue() === "BLACK") 
    {
      var DEST_SPREADSHEET_ID = '19SOEhBqA43lWavEnEERukTBrMdyFTJd6XJLcqvNQkhA';
      var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);

      var sheetName = (customerSurname + ' ' + customerName).replace(/\s+/g, ' ').trim();
      var destSheet = destSS.getSheetByName(sheetName);
      if (!destSheet) {
        Logger.log('Sheet "' + sheetName + '" non trovato nello Spreadsheet di destinazione — riga ' + activeRow);
        continue;
      }

      aggiornaSaldoVIPSelezionati(currentSheet.getRange(activeRow, 1, 1, currentSheet.getLastColumn()));
    }
  }

  return;
}

/*
function creaSpedizioniOLD() 
{
  var ui = SpreadsheetApp.getUi();
    
  // Get the current active Google Spreadsheet (IMDB_Ordini)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();  // "Current" tab

  var shippingStartedCol = getEmailRiepilogoColumn_(currentSheet, 'Shipping Started', 33);
  var noteSpedizioneCol = getEmailRiepilogoColumn_(currentSheet, 'Note', 33);
  var mktGuidesCol = getEmailRiepilogoColumn_(currentSheet, 'Guide Spedite', 33);
  var imdbVIPCol = getEmailRiepilogoColumn_(currentSheet, 'VIP', 33);
  var noteMagazzinoCol = getEmailRiepilogoColumn_(currentSheet, 'Note Magazzino', 33);
  var appTelefonicoCol = getEmailRiepilogoColumn_(currentSheet, 'Appuntamento Telefonico?', 33);
  var priorityCol = getEmailRiepilogoColumn_(currentSheet, 'Priority?', 33);
  var shippingPreparedCol = getEmailRiepilogoColumn_(currentSheet, 'Shipping Prepared', 33);
  var startColumn = getEmailRiepilogoColumn_(currentSheet, 'Codice bottiglia', 1) + 2;

  // Get the currently selected range of cells on the active sheet
  var activeRange = currentSheet.getActiveRange();

  // Retrieve all values in the selected range as a 2D array
  var selectedValues = activeRange.getValues();

  // Get the starting row of the active range
  var startingRow = activeRange.getRow();
  
  var activeRowNumber = 1;

  var rangeLen = selectedValues.length;

  // Iterate over each row in the selected range
  for (var i = 0; i < rangeLen; i++) 
  {
    activeRow = startingRow + i;

    if (currentSheet.getRange(activeRow, noteSpedizioneCol).getValue())
    {
      var responseDo = ui.alert("Sicuro di voler spedire? Hai segnato queste Note: " +   currentSheet.getRange(activeRow, noteSpedizioneCol).getValue() + ", vuoi continuare?", ui.ButtonSet.YES_NO);
      if (responseDo === ui.Button.NO) 
      {
        return;
      }
    }
    

    if (currentSheet.getRange(activeRow, shippingStartedCol).getValue())
    {
      var responseDo = ui.alert("Hai già richiesto la spedizione di: " +   currentSheet.getRange(activeRow, 4).getValue() + ", vuoi mandarlo di nuovo?", ui.ButtonSet.YES_NO);
      if (responseDo === ui.Button.NO) 
      {
        continue;
      }
    }
    
  
    // Get the value from Column A of the current row (assuming Column A is the first column)
    var colAValue = currentSheet.getRange(activeRow, 1).getValue();
    
    // Check if Column A starts with "Cliente "
    if (!colAValue.startsWith("Cliente ")) {
      // If it doesn't start with "Cliente ", abort and show an error message
      SpreadsheetApp.getUi().alert("Errore: Controlla di essere in un foglio di ordini nella riga relativa ad un cliente.");
      return; // Abort the script
    }
    
    // Get the "Indirizzi Spedizione" and "Indirizzi Fatturazione" sheets
    var indirizziSpedizioneSheet = ss.getSheetByName('Indirizzi Spedizione');
    var indirizziFatturazioneSheet = ss.getSheetByName('Indirizzi Fatturazione');

    // Get the current date in the format YYYY_MM_DD
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy_MM_d6");
    var currentDateForOrder = IMDBCommonLibs.getCustomDateCode();
    var currentDateFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "6d/MM/yyyy");

    // Create or open the "Spedizioni_YYYY_MM_DD" sheet
    var spedizioniRigheName = currentDate + "_Uscite_Righe";
    var spedizioniTestateName = currentDate + "_Uscite_Testate";
    var spedizioniRigheSheet = null;
    var spedizioniTestateSheet = null;
    
    // Check if the spreadsheet already exists and open it, otherwise create it
    var files = DriveApp.getFilesByName(spedizioniRigheName);
    if (files.hasNext()) {
      var file = files.next();
      spedizioniRigheSheet = SpreadsheetApp.open(file);
    } else {
      spedizioniRigheSheet = SpreadsheetApp.create(spedizioniRigheName);
    }

    // Check if the spreadsheet already exists and open it, otherwise create it
    var files = DriveApp.getFilesByName(spedizioniTestateName);
    if (files.hasNext()) {
      var file = files.next();
      spedizioniTestateSheet = SpreadsheetApp.open(file);
    } else {
      spedizioniTestateSheet = SpreadsheetApp.create(spedizioniTestateName);
    }

    // Create or get the tab "Spedizioni" or rename "Sheet1" if it exists
    var sheetRigheSpedizioni = spedizioniRigheSheet.getSheetByName('Sheet1');
    if (sheetRigheSpedizioni) {
      sheetRigheSpedizioni.setName('preimpostato_righe_uscite');
    } else {
      sheetRigheSpedizioni = spedizioniRigheSheet.getSheetByName('preimpostato_righe_uscite');
      if (!sheetRigheSpedizioni) {
        sheetRigheSpedizioni = spedizioniRigheSheet.insertSheet('preimpostato_righe_uscite');
      }
    }

    // Create or get the tab "Spedizioni" or rename "Sheet1" if it exists
    var sheetTestateSpedizioni = spedizioniTestateSheet.getSheetByName('Sheet1');
    if (sheetTestateSpedizioni) {
      sheetTestateSpedizioni.setName('preimpostato_testate_uscite');
    } else {
      sheetTestateSpedizioni = spedizioniTestateSheet.getSheetByName('preimpostato_testate_uscite');
      if (!sheetTestateSpedizioni) {
        sheetTestateSpedizioni = spedizioniTestateSheet.insertSheet('preimpostato_testate_uscite');
      }
    }
    
    // Set the column headers in bold in the first row
    var headers = ["N_Documento (30 Obbligatorio)", "Codice_Articolo (20 Obbligatorio)", "Quantita (Numero Obbligatorio)", "udc (20)", "lotto (20)"];

    var firstRow = sheetRigheSpedizioni.getRange(1, 1, 1, headers.length).getValues()[0];
    if (firstRow.toString() !== headers.toString()) {
      sheetRigheSpedizioni.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    }

    // Set the column headers in bold in the first row
    var headers = ["Documento (10 Campo Obbligatorio)", "Ragione Sociale (30)", "Indirizzo (50)", "Localita (30)", "Provincia (2)", "CAP (5)", "Nazione (2)", "Urgente (Si o lasciare vuoto)", "Appuntamento (Si o lasciare vuoto)", "Mail (100)", "Telefono (14)", "Note (70)", "Importo_Contrassegno (Numero con virgola)", "Tipo_Contrassegno", "Note Magazzino]"];

    var firstRow = sheetTestateSpedizioni.getRange(1, 1, 1, headers.length).getValues()[0];
    if (firstRow.toString() !== headers.toString()) {
      sheetTestateSpedizioni.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    }
    
    var rangeProdotti = [];
    var lastColumn = currentSheet.getLastColumn();
    
    for (var col = startColumn; col <= lastColumn; col++) {
      var value = currentSheet.getRange(1, col).getValue(); // Row 1
      if (value) {
        rangeProdotti.push(col); // Save the column index for products
      } else {
        break; // Stop if we find an empty cell
      }
    }
    
    // Extract the necessary information from the current row in the "Current" sheet
    var codiceCliente = currentSheet.getRange(activeRow, 2).getValue(); // Column B
    var codiceAzienda = currentSheet.getRange(activeRow, 3).getValue(); // Column B
    var surname = currentSheet.getRange(activeRow, 4).getValue();       // Column C (Surname)
    var firstName = currentSheet.getRange(activeRow, 5).getValue();     // Column D (First name)
    var codiceFiscale = currentSheet.getRange(activeRow, 6).getValue(); // Column E
    var partitaIva = currentSheet.getRange(activeRow, 7).getValue();    // Column F (Partita IVA)
    var emailCliente = currentSheet.getRange(activeRow, 8).getValue();    // Column F (Partita IVA)
    var appTelefonico = currentSheet.getRange(activeRow, appTelefonicoCol).getValue();    // Appuntamento Telefonico
    var noteMagazzino = currentSheet.getRange(activeRow, noteMagazzinoCol).getValue();    // Note magazzino
    var priorityDelivery = currentSheet.getRange(activeRow, priorityCol).getValue();    // Priority
    
    // Create Numero Ordine as YYYYMMDD_ClienteID
    var numeroOrdine = currentDateForOrder + codiceCliente.toString().padStart(6, '0');
    
    // Lookup in "Indirizzi Spedizione" tab
    var indirizziSpedizioneData = indirizziSpedizioneSheet.getDataRange().getValues();
    var destinatario = IMDBCommonLibs.findMatchingRow(indirizziSpedizioneData, codiceCliente, 0);  // Lookup Codice Fiscale in column 25 (index 24)
    
    var cliente;

    if (!codiceAzienda) {
      // If Partita IVA is empty, use "destinatario" fields for both destinatario and cliente
      cliente = destinatario;
    } else {
      // Otherwise, lookup "Indirizzi Fatturazione" by Partita IVA
      var indirizziFatturazioneData = indirizziFatturazioneSheet.getDataRange().getValues();
      cliente = IMDBCommonLibs.findMatchingRow(indirizziFatturazioneData, codiceAzienda, 0);  // Lookup Partita IVA in column 6 (index 5)
    }

    // --- Modify the script where the Provincia is assigned ---

    // Get the province values from destinatario and cliente
    var provinciaDestinatario = destinatario[49].toUpperCase();

    if (provinciaDestinatario.length > 2) 
      provinciaDestinatario = IMDBCommonLibs.getProvinceAcronym(provinciaDestinatario);
    if (provinciaDestinatario.length != 2)
    {
      Logger.log(surname + ": Provincia errata: " + provinciaDestinatario);
      ui.alert(surname + ": Provincia errata: " + provinciaDestinatario);
      continue;
    }

    // Update the relevant portion where we assign the phone number:
    if (destinatario[9].length)
      var telefonoFormatted = IMDBCommonLibs.formatPhoneNumber(destinatario[9]); // Format phone number
    else
      var telefonoFormatted = IMDBCommonLibs.formatPhoneNumber(destinatario[10]); // Format phone number

    // Check lengths

    var ragioneSociale = destinatario[4] + " " + destinatario[5];
    if ((ragioneSociale.length >= 30) || (ragioneSociale.length <= 6))
    {
      Logger.log(surname + ": Ragione sociale troppo lunga/corta: " + ragioneSociale);
      ui.alert(surname + ": Ragione sociale troppo lunga/corta: " + ragioneSociale);
      continue;
    }

    var indirizzo = destinatario[11];
    if ((indirizzo.length >= 50) || (indirizzo.length <= 5))
    {
      Logger.log(surname + ": Indirizzo troppo lungo/corto: " + indirizzo);
      ui.alert(surname + ": Indirizzo troppo lungo/corto: " + indirizzo);
      continue;
    }

    var localita = destinatario[13];
    if ((localita.length >= 30) || (localita.length <= 2))
    {
      Logger.log(surname + ": Località troppo lunga/corta: " + localita);
      ui.alert(surname + ": Località troppo lunga/corta: " + localita);
      continue;
    }    

    var zipCode = destinatario[15];
    if ((zipCode.toString().length != 5) || (zipCode === "00000"))
    {
      Logger.log(surname + ": CAP errato: " + zipCode.toString() + " " + zipCode.toString().length);
      ui.alert(surname + ": CAP errato: " + zipCode.toString() + " " + zipCode.toString().length);
      continue;
    }    

    // Update the relevant portion where we assign the phone number:
    if (destinatario[9].length)
      var telefonoFormatted = IMDBCommonLibs.formatPhoneNumber(destinatario[9]); // Format phone number
    else
      var telefonoFormatted = IMDBCommonLibs.formatPhoneNumber(destinatario[10]); // Format phone number


    if (destinatario[12].length > 70)
    {
      Logger.log(surname + ": Note di spedizione troppo lunghe: " + destinatario[12]);
      ui.alert(surname + ": Note di spedizione troppo lunghe: " + destinatario[12]);
      continue;
    }    
    else if (destinatario[12].length >= 56)
        var noteSpedizione = destinatario[12];
    else
        var noteSpedizione = destinatario[12] + " Tel:" + telefonoFormatted;

    // Check if numeroOrdine is existing

    var existingOrder = false;

    var k = 1;
    
    while (sheetTestateSpedizioni.getRange(k,1).getValue())
    {
      if (sheetTestateSpedizioni.getRange(k,1).getValue() === numeroOrdine)
      {
        existingOrder = true;
        break;
      }
      k++;
    }

    if (!existingOrder)
    {
      // Prepare the new row to append to "Testate"
      var newRow = [
        "'" + numeroOrdine,                     // Numero Ordine
        "'" + ragioneSociale, // Ragione Sociale Destinatario (Lastname + Firstname)
        indirizzo,                  // Indirizzo Destinatario (Address1)
        localita,                  // Località Destinatario (City)
        provinciaDestinatario,             // Provincia Destinatario (Provincia)
        "'" + zipCode,                  // CAP Destinatario (Zipcode)
        "IT",
        priorityDelivery, // Urgente
        appTelefonico, // Appuntamento
        "'" + emailCliente,                  // Email
        "'" + telefonoFormatted,                    // Telefono
        "'" + noteSpedizione, // Note Consegna 1 (address2)
        '',                               // Importo contrassegno
        '',                                // Tipo contrassegno
        noteMagazzino
      ];
      sheetTestateSpedizioni.appendRow(newRow);
    }

    // Send also to SCOT Portal

    // 2) prepara i dati di header
    var scotUsciteHeader = {
      business_name: ragioneSociale,
      document_date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "YYYYMMdd"),
      attachment: false,
      address: indirizzo,
      location: localita,
      province: provinciaDestinatario,
      zip_code: zipCode,
      nation: "IT",
      urgent: priorityDelivery,
      delivery_date: "", // new Date().getTime().toISOString(),
      appointment: appTelefonico,
      email: emailCliente,
      tel_reference: telefonoFormatted,
      carrier_note: noteSpedizione,
      warehouse_note: currentSheet.getRange(activeRow, noteMagazzinoCol).getValue(), // Aggiungere nel foglio,
      cash_on_delivery_value: 0.0,
      cash_on_delivery_type: ""
    };

    var scotUsciteRows = [];

/*
    // Check if Marketing material has to be sent
    if (currentSheet.getRange(activeRow, mktGuidesCol).getValue() != "OK")
    {
      // Prepare the Guida Champagne row to append to "Spedizioni"
      var newRow = [
        numeroOrdine,   // Numero Ordine
        "MKTBK01",            // Codice Articolo (from Row 1)
        "1"                  // Quantità da Spedire (current row & column)
      ];
      
      var newScotRow = {
        id : numeroOrdine,   // Numero Ordine
        code : "MKTBK01",            // Codice Articolo (from Row 1)
        quantity : "1"                  // Quantità da Spedire (current row & column)
      };
      
      // Append the new row to "Spedizioni"
      scotUsciteRows.push(newScotRow);
      sheetRigheSpedizioni.appendRow(newRow);

      // Prepare the Guida Excellence row to append to "Spedizioni"
      var newRow = [
        numeroOrdine,   // Numero Ordine
        "MKTBK02",            // Codice Articolo (from Row 1)
        "1"                  // Quantità da Spedire (current row & column)
      ];
      
      var newScotRow = {
        id : numeroOrdine,   // Numero Ordine
        code : "MKTBK02",            // Codice Articolo (from Row 1)
        quantity : "1"                  // Quantità da Spedire (current row & column)
      };
      

      // Append the new row to "Spedizioni"
      scotUsciteRows.push(newScotRow);
      sheetRigheSpedizioni.appendRow(newRow);


      // Don't send Flyer to IMDB VIP Club
      if (!currentSheet.getRange(activeRow, imdbVIPCol).getValue())
      {
        // Prepare the Guida Excellence row to append to "Spedizioni"
        var newRow = [
          numeroOrdine,   // Numero Ordine
          "MKTFL01",            // Codice Articolo (from Row 1)
          "1"                  // Quantità da Spedire (current row & column)
        ];

        // Prepare the new row to append to "Spedizioni"
        var newScotRow = {
          id: numeroOrdine,    // Numero Ordine
          code : "MKTFL01",             // Codice Articolo (from Row 1)
          quantity : "1"                   // Quantità da Spedire (current row & column)
        };
        
        // Append the new row to "Spedizioni"
        scotUsciteRows.push(newScotRow);
        sheetRigheSpedizioni.appendRow(newRow);
      }

      // Set GuideSpedite as OK
      currentSheet.getRange(activeRow, mktGuidesCol).setValue("OK");
    }
/* FINE
    // Loop through the Range_Prodotti and add new rows for each valid product entry in the current row
    for (var j = 0; j < rangeProdotti.length; j++) 
    {
      var col = rangeProdotti[j];
      
      // Get the quantity to be shipped (actual row and current column in Range_Prodotti)
      var quantitaDaSpedire = currentSheet.getRange(activeRow, col).getValue();
      
      // If the quantity is greater than 0, process the product
      if (quantitaDaSpedire > 0) 
      {
        // Get the article code from Row 1 in the current column
        var codiceArticolo = currentSheet.getRange(1, col).getValue();
        
        // Prepare the new row to append to "Spedizioni"
        var newScotRow = {
          id : numeroOrdine,                     // Numero Ordine
          code : codiceArticolo,                   // Codice Articolo (from Row 1)
          quantity : quantitaDaSpedire                // Quantità da Spedire (current row & column)
        };
        
        // Prepare the new row to append to "Spedizioni"
        var newRow = [
          numeroOrdine,                     // Numero Ordine
          codiceArticolo,                   // Codice Articolo (from Row 1)
          quantitaDaSpedire                // Quantità da Spedire (current row & column)
        ];
        
        // Append the new row to "Spedizioni"
        scotUsciteRows.push(newScotRow);
        sheetRigheSpedizioni.appendRow(newRow);
        //sheetRigheSpedizioni.appendRow(newRow.map(String).map(function(val) { return val.toUpperCase(); }));  // Make it uppercase
      }
    }
    currentSheet.getRange(activeRow, 15).setValue("Preparato: " + numeroOrdine);

    if (inviaSpedizioni_DaPortale)
    {
      var responseDo = ui.alert("Sicuro di voler inviare a SCOT la spedizione di: " +   currentSheet.getRange(activeRow, 4).getValue() + ", vuoi mandarlo di nuovo?", ui.ButtonSet.YES_NO);
      if (responseDo === ui.Button.YES) 
      {
        scotOrdiniUscita_(numeroOrdine,"MDB", scotUsciteHeader, scotUsciteRows, ragioneSociale, currentSheet.getName());
      }
    }

    submitFormShippingStarted_(emailCliente,firstName, surname, currentSheet.getName(),Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/YYYY"));
    currentSheet.getRange(activeRow, shippingPreparedCol).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));

  } // Next client
}*/

function creaSpedizioniFromMautic()
{
  var ui = SpreadsheetApp.getUi();

  // Get the current active Google Spreadsheet (IMDB_Ordini)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();  // "Current" tab

  var shippingStartedCol   = getEmailRiepilogoColumn_(currentSheet, 'Shipping Started', 33);
  var noteSpedizioneCol    = getEmailRiepilogoColumn_(currentSheet, 'Note', 33);
  var mktGuidesCol         = getEmailRiepilogoColumn_(currentSheet, 'Guide Spedite', 33);
  var imdbVIPCol           = getEmailRiepilogoColumn_(currentSheet, 'VIP', 33);
  var noteMagazzinoCol     = getEmailRiepilogoColumn_(currentSheet, 'Note Magazzino', 33);
  var appTelefonicoCol     = getEmailRiepilogoColumn_(currentSheet, 'Appuntamento Telefonico?', 33);
  var priorityCol          = getEmailRiepilogoColumn_(currentSheet, 'Priority?', 33);
  var shippingPreparedCol  = getEmailRiepilogoColumn_(currentSheet, 'Shipping Prepared', 33);
  var startColumn          = getEmailRiepilogoColumn_(currentSheet, 'Codice bottiglia', 1) + 2;


  var activeRange = currentSheet.getActiveRange();
  var selectedValues = activeRange.getValues();
  var startingRow = activeRange.getRow();
  var rangeLen = selectedValues.length;

  for (var i = 0; i < rangeLen; i++)
  {
    var activeRow = startingRow + i;

    // Skip rows hidden by filter
    if (currentSheet.isRowHiddenByFilter(activeRow))
    {
      continue;
    }

    // Skip rows hidden manually or collapsed/group-hidden
    if (currentSheet.isRowHiddenByUser(activeRow))
    {
      continue;
    }

    if (currentSheet.getRange(activeRow, noteSpedizioneCol).getValue())
    {
      var responseDo = ui.alert(
        "Sicuro di voler spedire? Hai segnato queste Note: " +
        currentSheet.getRange(activeRow, noteSpedizioneCol).getValue() +
        ", vuoi continuare?",
        ui.ButtonSet.YES_NO
      );
      if (responseDo === ui.Button.NO) {
        return;
      }
    }

    if (currentSheet.getRange(activeRow, shippingStartedCol).getValue())
    {
      var responseDo = ui.alert(
        "Hai già richiesto la spedizione di: " +
        currentSheet.getRange(activeRow, 4).getValue() +
        ", vuoi mandarlo di nuovo?",
        ui.ButtonSet.YES_NO
      );
      if (responseDo === ui.Button.NO) {
        continue;
      }
    }

    // Check Column A starts with "Cliente "
    var colAValue = currentSheet.getRange(activeRow, 1).getValue();
    if (!String(colAValue).startsWith("Cliente ")) {
      ui.alert("Errore: Controlla di essere in un foglio di ordini nella riga relativa ad un cliente.");
      return;
    }

    // Dates
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy_MM_dd");
    var currentDateForOrder = IMDBCommonLibs.getCustomDateCode();
    var currentDateFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // Create or open shipping spreadsheets
    var spedizioniRigheName = currentDate + "_Uscite_Righe";
    var spedizioniTestateName = currentDate + "_Uscite_Testate";
    var spedizioniRigheSheet = null;
    var spedizioniTestateSheet = null;

    var files = DriveApp.getFilesByName(spedizioniRigheName);
    if (files.hasNext()) spedizioniRigheSheet = SpreadsheetApp.open(files.next());
    else spedizioniRigheSheet = SpreadsheetApp.create(spedizioniRigheName);

    files = DriveApp.getFilesByName(spedizioniTestateName);
    if (files.hasNext()) spedizioniTestateSheet = SpreadsheetApp.open(files.next());
    else spedizioniTestateSheet = SpreadsheetApp.create(spedizioniTestateName);

    // Tabs righe
    var sheetRigheSpedizioni = spedizioniRigheSheet.getSheetByName('Sheet1');
    if (sheetRigheSpedizioni) {
      sheetRigheSpedizioni.setName('preimpostato_righe_uscite');
    } else {
      sheetRigheSpedizioni = spedizioniRigheSheet.getSheetByName('preimpostato_righe_uscite');
      if (!sheetRigheSpedizioni) {
        sheetRigheSpedizioni = spedizioniRigheSheet.insertSheet('preimpostato_righe_uscite');
      }
    }

    // Tabs testate
    var sheetTestateSpedizioni = spedizioniTestateSheet.getSheetByName('Sheet1');
    if (sheetTestateSpedizioni) {
      sheetTestateSpedizioni.setName('preimpostato_testate_uscite');
    } else {
      sheetTestateSpedizioni = spedizioniTestateSheet.getSheetByName('preimpostato_testate_uscite');
      if (!sheetTestateSpedizioni) {
        sheetTestateSpedizioni = spedizioniTestateSheet.insertSheet('preimpostato_testate_uscite');
      }
    }

    // Headers righe
    var headers = ["N_Documento (30 Obbligatorio)", "Codice_Articolo (20 Obbligatorio)", "Quantita (Numero Obbligatorio)", "udc (20)", "lotto (20)"];
    var firstRow = sheetRigheSpedizioni.getRange(1, 1, 1, headers.length).getValues()[0];
    if (firstRow.toString() !== headers.toString()) {
      sheetRigheSpedizioni.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    }

    // Headers testate
    headers = ["Documento (10 Campo Obbligatorio)", "Ragione Sociale (30)", "Indirizzo (50)", "Localita (30)", "Provincia (2)", "CAP (5)", "Nazione (2)", "Urgente (Si o lasciare vuoto)", "Appuntamento (Si o lasciare vuoto)", "Mail (100)", "Telefono (14)", "Note (70)", "Importo_Contrassegno (Numero con virgola)", "Tipo_Contrassegno", "Note Magazzino]"];
    firstRow = sheetTestateSpedizioni.getRange(1, 1, 1, headers.length).getValues()[0];
    if (firstRow.toString() !== headers.toString()) {
      sheetTestateSpedizioni.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    }

    // Scan Row 1 from column AH onwards to get the Range_Prodotti
    var rangeProdotti = [];
    var lastColumn = currentSheet.getLastColumn();

    for (var col = startColumn; col <= lastColumn; col++) {
      var value = currentSheet.getRange(1, col).getValue(); // Row 1
      if (value) rangeProdotti.push(col);
      else break;
    }

    // ==========================
    // DATI DAL FOGLIO (solo quelli richiesti)
    // ==========================
    var codiceCliente     = currentSheet.getRange(activeRow, 2).getValue(); // Colonna 2 => SOLO questa per Mautic
    var appTelefonico     = currentSheet.getRange(activeRow, appTelefonicoCol).getValue();
    var noteMagazzino     = currentSheet.getRange(activeRow, noteMagazzinoCol).getValue();
    var priorityDelivery  = currentSheet.getRange(activeRow, priorityCol).getValue();

    // Numero ordine (come prima)
    var numeroOrdine = currentDateForOrder + String(codiceCliente).toString().padStart(6, '0');

    var payload = retrieveMautiCustomerData_(codiceCliente);

    if (!payload) continue; // Salta in caso di errori

    var validationError = IMDBCommonLibs.validateShippingDataHeadless(payload);
    if (validationError) {
      ui.alert(validationError);
      continue;
    }

    // ==========================
    // Check if numeroOrdine is existing (identico)
    // ==========================
    var existingOrder = false;
    var k = 1;
    while (sheetTestateSpedizioni.getRange(k,1).getValue())
    {
      if (sheetTestateSpedizioni.getRange(k,1).getValue() === numeroOrdine)
      {
        existingOrder = true;
        break;
      }
      k++;
    }

    if (!existingOrder)
    {
      var newRow = [
        "'" + numeroOrdine,
        "'" + payload.ragioneSociale,
        payload.indirizzo,
        payload.localita,
        payload.provinciaDestinatario,
        "'" + payload.zipCode,
        payload.nazione,
        priorityDelivery,
        appTelefonico,
        "'" + payload.email,
        "'" + payload.telefonoFormatted,
        "'" + payload.noteSpedizione,
        '',
        '',
        noteMagazzino
      ];
      sheetTestateSpedizioni.appendRow(newRow);
    }

    // ==========================
    // SCOT Portal header (identico, ma con dati Mautic)
    // ==========================
    var scotUsciteHeader = {
      business_name: payload.ragioneSociale,
      document_date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "YYYYMMdd"),
      attachment: false,
      address: payload.indirizzo,
      location: payload.localita,
      province: payload.provinciaDestinatario,
      zip_code: payload.zipCode,
      nation: payload.nazione,
      urgent: priorityDelivery,
      delivery_date: "",
      appointment: appTelefonico,
      email: payload.email,
      tel_reference: payload.telefonoFormatted,
      carrier_note: payload.noteSpedizione,
      warehouse_note: currentSheet.getRange(activeRow, noteMagazzinoCol).getValue(),
      cash_on_delivery_value: 0.0,
      cash_on_delivery_type: ""
    };

    var scotUsciteRows = [];

    // ==========================
    // Loop prodotti (identico)
    // ==========================
    for (var j = 0; j < rangeProdotti.length; j++)
    {
      var colp = rangeProdotti[j];
      var quantitaDaSpedire = currentSheet.getRange(activeRow, colp).getValue();

      if (quantitaDaSpedire > 0)
      {
        var codiceArticolo = currentSheet.getRange(1, colp).getValue();

        var newScotRow = {
          id: numeroOrdine,
          code: codiceArticolo,
          quantity: quantitaDaSpedire
        };

        var newRowRighe = [
          numeroOrdine,
          codiceArticolo,
          quantitaDaSpedire
        ];

        scotUsciteRows.push(newScotRow);
        sheetRigheSpedizioni.appendRow(newRowRighe);
      }
    }

    currentSheet.getRange(activeRow, 15).setValue("Preparato: " + numeroOrdine);

    if (inviaSpedizioni_DaPortale)
    {
      var responseDo = ui.alert(
        "Sicuro di voler inviare a SCOT la spedizione di: " +
        currentSheet.getRange(activeRow, 4).getValue() +
        ", vuoi mandarlo di nuovo?",
        ui.ButtonSet.YES_NO
      );
      if (responseDo === ui.Button.YES)
      {
        scotOrdiniUscita_(numeroOrdine,"MDB", scotUsciteHeader, scotUsciteRows, ragioneSociale, currentSheet.getName());
      }
    }

    submitFormShippingStarted_(
      payload.email,
      payload.firstName,
      payload.surname,
      currentSheet.getName(),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/YYYY")
    );

    currentSheet.getRange(activeRow, shippingPreparedCol)
      .setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));

  } // Next client
}


function retrieveMautiCustomerData_(codiceCliente)
{
  // ==========================
  // ✅ NUOVO: Leggi TUTTO da Mautic usando SOLO codiceCliente (colonna 2)
  // ==========================
  var mauticData;
  try {
    mauticData = IMDBCommonLibs.getMauticCustomerData(String(codiceCliente).trim());
  } catch (e) {
    Logger.log("Errore getMauticCustomerData(" + codiceCliente + "): " + e);
    ui.alert("Errore Mautic su riga " + activeRow + ": " + (e && e.message ? e.message : e));
    return null;
  }

  // Per sicurezza: se dovesse arrivare array (non dovrebbe con ID), chiedi scelta
  if (Array.isArray(mauticData)) {
    var picked = pickMauticContactFromArray_(ui, mauticData, String(codiceCliente).trim());
    if (!picked) return null;
    mauticData = picked;
  }

  var contact = (mauticData && mauticData.contact) ? mauticData.contact : (mauticData || {});
  if (!contact) {
    ui.alert("Dati Mautic non validi per codiceCliente=" + codiceCliente + " (riga " + activeRow + ")");
    return null;
  }

  // ==========================
  // DATI CHE PRIMA LEGGI DAL FOGLIO -> ORA DA MAUTIC
  // ==========================
  var surname        = IMDBCommonLibs.getMauticFieldNormalized(contact, 'lastname') || '';
  var firstName      = IMDBCommonLibs.getMauticFieldNormalized(contact, 'firstname') || '';
  var email          = IMDBCommonLibs.getMauticFieldNormalized(contact, 'email') || '';
  var codiceFiscale  = IMDBCommonLibs.getMauticFieldNormalized(contact, 'codice_fiscale') || '';
  var partitaIva     = IMDBCommonLibs.getMauticFieldNormalized(contact, 'partita_iva') || IMDBCommonLibs.getMauticFieldNormalized(contact, 'vat') || '';
  var emailCliente   = IMDBCommonLibs.getMauticFieldNormalized(contact, 'email') || '';

  var telefonoFormatted = IMDBCommonLibs.formatPhoneNumber(IMDBCommonLibs.getMauticFieldNormalized(contact, 'phone') ||
                  IMDBCommonLibs.getMauticFieldNormalized(contact, 'mobile')); // Format phone number


  // ==========================
  // DATI SPEDIZIONE DA MAUTIC (con fallback robusti)
  // ==========================
  var indirizzo = IMDBCommonLibs.getMauticFieldNormalized(contact, 'indirizzo') ||
                  IMDBCommonLibs.getMauticFieldNormalized(contact, 'address1')  || '';

  var localita  = IMDBCommonLibs.getMauticFieldNormalized(contact, 'localita') ||
                  IMDBCommonLibs.getMauticFieldNormalized(contact, 'city')      || '';

  var zipCode   = IMDBCommonLibs.getMauticFieldNormalized(contact, 'cap') ||
                  IMDBCommonLibs.getMauticFieldNormalized(contact, 'zipcode') || '';

  // Provincia: usa SOLO PROVINCIA (ignora state)
  var provinciaRaw = IMDBCommonLibs.getMauticFieldNormalized(contact, 'PROVINCIA') ||
                      IMDBCommonLibs.getMauticFieldNormalized(contact, 'provincia') || '';
  var provinciaDestinatario = String(provinciaRaw).toUpperCase().trim();

  var codiceFiscale = IMDBCommonLibs.getMauticFieldNormalized(contact, 'codice_fiscale') || '';


  var payload = {
    surname: surname,
    firstName: firstName,
    email: email,
    telefonoFormatted: telefonoFormatted,
    codiceFiscale: codiceFiscale,
    provinciaDestinatario: provinciaDestinatario,
    provinciaRaw: provinciaRaw,
    indirizzo: indirizzo,
    localita: localita,
    zipCode: zipCode,
    contact: contact
  };

  return payload;
}

/**
 * Helper UI: se per qualche motivo getMauticCustomerData() tornasse un array,
 * consente all'utente di scegliere quale contatto usare.
 */
function pickMauticContactFromArray_(ui, customerDataArray, keyLabel)
{
  var n = customerDataArray.length;

  var msg = 'Sono stati trovati ' + n + ' contatti per "' + keyLabel + '":\n\n';

  for (var i = 0; i < n; i++) {
    var cd = customerDataArray[i];
    var contact = (cd && cd.contact) ? cd.contact : (cd || {});

    var firstName = IMDBCommonLibs.getMauticFieldNormalized(contact, 'firstname') || '';
    var lastName  = IMDBCommonLibs.getMauticFieldNormalized(contact, 'lastname')  || '';
    var email     = IMDBCommonLibs.getMauticFieldNormalized(contact, 'email')     || '';

    var fields = contact.fields || {};
    var all    = fields.all || {};
    var id     = (typeof contact.id !== 'undefined' && contact.id !== null)
      ? contact.id
      : (typeof all.id !== 'undefined' ? all.id : '');

    msg += (i + 1) + ') ID ' + id + ' - ' + (String(firstName) + ' ' + String(lastName)).trim() +
           (email ? ' <' + email + '>' : '') + '\n';
  }

  msg += '\nInserisci il numero del contatto da usare (1-' + n + '), oppure 0 per annullare:';

  var promptResult = ui.prompt('Seleziona contatto Mautic', msg, ui.ButtonSet.OK_CANCEL);
  var button = promptResult.getSelectedButton();
  if (button !== ui.Button.OK) return null;

  var choiceStr = String(promptResult.getResponseText()).trim();
  var choice = parseInt(choiceStr, 10);

  if (isNaN(choice) || choice < 0 || choice > n) {
    throw new Error('Scelta non valida: "' + choiceStr + '". Operazione annullata.');
  }
  if (choice === 0) return null;

  return customerDataArray[choice - 1];
}

/**
 * Raggruppa e invia le spedizioni:
 * - Legge due sheet: YYYY_MM_DD_Uscite_Testate e YYYY_MM_DD_Uscite_Righe
 * - Raggruppa eventuali duplicati in uscite_testate per Documento, utilizzando la prima occorrenza
 * - Raggruppa le righe di dettaglio da uscite_righe per N_Documento
 * - Costruisce scotUsciteHeader da uscite_testate e scotUsciteRows da uscite_righe
 * - Chiama inviaSpedizioni_(ui, numeroOrdine, scotUsciteHeader, scotUsciteRows)
 *
 * MODIFICHE:
 * - Colonna "Data Spedizione": cercata per header; se manca viene aggiunta in coda
 * - Se "Data Spedizione" contiene una data valida dd/MM/yyyy → log breve e skip ordine
 * - Le righe skip NON contribuiscono alle somme per codice/articolo
 * - Controllo coerenza aggiuntivo: totale bottiglie per ordine (prima e dopo il grouping per codice) deve combaciare
 */
function processaFileSpedizioni() 
{
  //if (!checkGiacenzeSpedizioni())
  //  return false;
  
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy_MM_dd');
  const todayHuman = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
  var ui = SpreadsheetApp.getUi();

  // Trova i file per la data corrente
  const filesTestate = DriveApp.getFilesByName(today + '_Uscite_Testate');
  const filesRighe   = DriveApp.getFilesByName(today + '_Uscite_Righe');
  if (!filesTestate.hasNext() || !filesRighe.hasNext()) 
  {
    throw new Error('File Uscite_Testate o Uscite_Righe non trovati per: ' + today);
  }

  const ssTestate = SpreadsheetApp.open(filesTestate.next());
  const ssRighe   = SpreadsheetApp.open(filesRighe.next());
  const sheetTestate = ssTestate.getSheetByName('preimpostato_testate_uscite');
  const sheetRighe   = ssRighe.getSheetByName('preimpostato_righe_uscite');
  if (!sheetTestate || !sheetRighe) 
  {
    throw new Error('Foglio preimpostato_testate_uscite o preimpostato_righe_uscite non trovato');
  }

  // =========================
  // 1) Colonna "Data Spedizione" (trova o crea in coda)
  // =========================
  const headerRow = sheetTestate.getRange(1, 1, 1, sheetTestate.getLastColumn()).getDisplayValues()[0];
  let shipDateCol = headerRow.findIndex(h => String(h || '').trim().toLowerCase() === 'data spedizione') + 1;

  if (!shipDateCol) {
    shipDateCol = sheetTestate.getLastColumn() + 1;
    sheetTestate.insertColumnAfter(sheetTestate.getLastColumn());
    sheetTestate.getRange(1, shipDateCol).setValue("Data Spedizione").setFontWeight("bold");
  } else {
    // assicura intestazione corretta (senza spaccare se già esiste)
    sheetTestate.getRange(1, shipDateCol).setValue("Data Spedizione").setFontWeight("bold");
  }

  // Leggi header e dati da testate
  const dataTestate = sheetTestate.getDataRange().getValues();
  const headersT = dataTestate.shift();
  const idxT = h => headersT.indexOf(h);

  // Helper: data valida dd/MM/yyyy
  const isValidItDate_ = (v) => {
    const s = String(v || '').trim();
    if (!s) return false;
    const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (!m) return false;
    const dd = Number(m[1]), mm = Number(m[2]), yyyy = Number(m[3]);
    if (mm < 1 || mm > 12) return false;
    if (dd < 1 || dd > 31) return false;
    const d = new Date(yyyy, mm - 1, dd);
    return d.getFullYear() === yyyy && (d.getMonth() + 1) === mm && d.getDate() === dd;
  };

  // =========================
  // 2) Raggruppa header per Documento (prima occorrenza), ma SKIP se già spedito
  // =========================
  const orderHeaders = {};
  const orderRows = {}; // riga in sheetTestate
  const skippedOrders = new Set(); // ordini già spediti (da saltare a cascata)

  dataTestate.forEach((r, i) => {
    const num = r[idxT('Documento (10 Campo Obbligatorio)')];
    if (!num) return;

    // se esiste già una prima occorrenza, ignoriamo duplicati (come prima)
    if (orderHeaders.hasOwnProperty(num)) return;

    // controllo Data Spedizione: se già valorizzata con data valida → skip ordine
    const shipCell = r[shipDateCol - 1];
    if (isValidItDate_(shipCell)) {
      Logger.log(`SKIP ordine ${num}: già spedito in data ${shipCell}`);
      skippedOrders.add(String(num));
      return;
    }

    orderHeaders[num] = r;
    orderRows[num] = i + 2; // perché dataTestate parte dalla riga 2
  });

  // Leggi header e dati da righe
  const dataRighe = sheetRighe.getDataRange().getValues();
  const headersR = dataRighe.shift();

  // Pre-calcolo gli indici per non rifare indexOf ad ogni riga
  const idxNum  = headersR.indexOf('N_Documento (30 Obbligatorio)');
  const idxCode = headersR.indexOf('Codice_Articolo (20 Obbligatorio)');
  const idxQty  = headersR.indexOf('Quantita (Numero Obbligatorio)');

  if (idxNum < 0 || idxCode < 0 || idxQty < 0) {
    throw new Error('Headers mancanti in righe: controlla N_Documento / Codice_Articolo / Quantita');
  }

  // =========================
  // 3) Somme righe: solo ordini NON skip (e presenti in orderHeaders)
  //    + controllo coerenza: totale pre/post grouping
  // =========================
  const orderDetailsRaw = {};          // per codice
  const orderTotalBefore = {};         // somma qty per ordine (prima del grouping)
  const orderTotalAfter = {};          // somma qty per ordine (dopo il grouping)

  dataRighe.forEach(r => {
    const num  = r[idxNum];
    const code = r[idxCode];
    const qty  = Number(r[idxQty]) || 0;

    if (!num || !code || qty === 0) return;

    const key = String(num);

    // Se ordine già spedito o non esiste in testate (prima occorrenza non skip) → non considerare
    if (skippedOrders.has(key)) return;
    if (!orderHeaders.hasOwnProperty(key)) return;

    // totale prima del grouping
    if (!orderTotalBefore[key]) orderTotalBefore[key] = 0;
    orderTotalBefore[key] += qty;

    // grouping per codice
    if (!orderDetailsRaw[key]) orderDetailsRaw[key] = {};
    if (!orderDetailsRaw[key][code]) orderDetailsRaw[key][code] = 0;
    orderDetailsRaw[key][code] += qty;
  });

  // Normalizza in array come prima: [{ codiceArticolo, quantita }, ...]
  const orderDetails = {};
  Object.keys(orderDetailsRaw).forEach(num => {
    orderDetails[num] = Object.keys(orderDetailsRaw[num]).map(code => ({
      codiceArticolo: code,
      quantita: orderDetailsRaw[num][code]
    }));

    // totale dopo il grouping (somma dei totali per codice)
    orderTotalAfter[num] = orderDetails[num].reduce((acc, d) => acc + (Number(d.quantita) || 0), 0);
  });

  // =========================
  // 4) Controllo coerenza esistente + controllo totale bottiglie pre/post grouping
  // =========================
  if (!checkDocumentConsistency_(orderHeaders, orderDetailsRaw))
  {
    ui.alert("Incongruenze tra gli ordini!");
    return;
  }

  // controllo aggiuntivo: totali bottiglie uguali pre/post
  const incoerenti = [];
  Object.keys(orderHeaders).forEach(num => {
    const key = String(num);
    const before = Number(orderTotalBefore[key] || 0);
    const after  = Number(orderTotalAfter[key] || 0);

    // Se non ha dettagli, verrà saltato dopo; qui non serve bloccare
    if (!orderDetails[key] || orderDetails[key].length === 0) return;

    if (before !== after) {
      incoerenti.push(`${key} (before=${before}, after=${after})`);
    }
  });

  if (incoerenti.length) {
    Logger.log("ERRORE coerenza totali bottiglie pre/post grouping: " + incoerenti.join("; "));
    ui.alert("Errore: incongruenza totali bottiglie pre/post raggruppamento:\n" + incoerenti.join("\n"));
    return;
  }

  // =========================
  // 5) Elabora ciascun ordine (solo non-skip, quindi orderHeaders già filtrati)
  // =========================
  Object.keys(orderHeaders).forEach(num => {
    const hdr = orderHeaders[num];
    const rowIndex = orderRows[num];
    const details = orderDetails[num] || [];
    if (details.length === 0) {
      // Nessuna riga di dettaglio
      return;
    }

    // Estrai campi dal header ordine
    const ragioneSociale = hdr[idxT('Ragione Sociale (30)')];
    const indirizzo      = hdr[idxT('Indirizzo (50)')];
    const localita       = hdr[idxT('Localita (30)')];
    const provincia      = hdr[idxT('Provincia (2)')];
    const cap            = hdr[idxT('CAP (5)')];
    const nazione        = hdr[idxT('Nazione (2)')];
    const urgente        = hdr[idxT('Urgente (Si o lasciare vuoto)')];
    const appuntamento   = hdr[idxT('Appuntamento (Si o lasciare vuoto)')];
    const email          = hdr[idxT('Mail (100)')];
    const telRaw         = hdr[idxT('Telefono (14)')];
    const noteSped       = hdr[idxT('Note (70)')];
    const cashVal        = parseFloat(hdr[idxT('Importo_Contrassegno (Numero con virgola)')] || 0);
    const cashType       = hdr[idxT('Tipo_Contrassegno')];
    const noteMagazzino  = hdr[idxT('Note Magazzino')];

    // Normalizza telefono
    const tel = String(telRaw).replace(/[\s\-]/g, '');

    // Costruisci header per API
    const scotUsciteHeader = {
      business_name:          ragioneSociale,
      document_date:          Utilities.formatDate(new Date(), tz, 'yyyyMMdd'),
      attachment:             false,
      address:                indirizzo,
      location:               localita,
      province:               provincia,
      zip_code:               cap,
      nation:                 nazione,
      urgent:                 urgente,
      delivery_date:          '',
      appointment:            appuntamento,
      email:                  email,
      tel_reference:          tel,
      carrier_note:           noteSped,
      warehouse_note:         noteMagazzino,
      cash_on_delivery_value: cashVal,
      cash_on_delivery_type:  cashType
    };

    const scotUsciteRows = details.map(d => ({
      id:       num,
      code:     d.codiceArticolo,
      quantity: d.quantita
    }));

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var currentSheet = ss.getActiveSheet();  // "Current" tab
    var campagnaNome = currentSheet.getName();

    if (inviaSpedizioni_(num, scotUsciteHeader, scotUsciteRows, ragioneSociale, campagnaNome))
    {
      sheetTestate.getRange(rowIndex, shipDateCol).setValue(todayHuman);
    }
    else
    {
      sheetTestate.getRange(rowIndex, shipDateCol).setValue("ERRORE!");
      ui.alert("Ordine non spedito: " + ragioneSociale);
    }
  });
}

/**
 * Raggruppa e invia le spedizioni:
 * - Legge due sheet: YYYY_MM_DD_Uscite_Testate e YYYY_MM_DD_Uscite_Righe
 * - Raggruppa eventuali duplicati in uscite_testate per Documento, utilizzando la prima occorrenza
 * - Raggruppa le righe di dettaglio da uscite_righe per N_Documento
 * - Costruisce scotUsciteHeader da uscite_testate e scotUsciteRows da uscite_righe
 * - Chiama inviaSpedizioni_(ui, numeroOrdine, scotUsciteHeader, scotUsciteRows)
 */
function processaFileSpedizioniOLD_() 
{
    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy_MM_dd');
    const todayHuman = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
    var ui = SpreadsheetApp.getUi();

    // Trova i file per la data corrente
    const filesTestate = DriveApp.getFilesByName(today + '_Uscite_Testate');
    const filesRighe   = DriveApp.getFilesByName(today + '_Uscite_Righe');
    if (!filesTestate.hasNext() || !filesRighe.hasNext()) 
    {
        throw new Error('File Uscite_Testate o Uscite_Righe non trovati per: ' + today);
    }

    const ssTestate = SpreadsheetApp.open(filesTestate.next());
    const ssRighe   = SpreadsheetApp.open(filesRighe.next());
    const sheetTestate = ssTestate.getSheetByName('preimpostato_testate_uscite');
    const sheetRighe   = ssRighe.getSheetByName('preimpostato_righe_uscite');
    if (!sheetTestate || !sheetRighe) 
    {
        throw new Error('Foglio preimpostato_testate_uscite o preimpostato_righe_uscite non trovato');
    }

    // --- GARANTIRE che la colonna P esista e abbia il titolo ---
    const lastCol = sheetTestate.getLastColumn();
    if (lastCol < 16)  // colonna P = 16
    {
        sheetTestate.insertColumnAfter(lastCol);
    }

    // Scrivi intestazione colonna P
    sheetTestate.getRange(1, 16).setValue("Data Spedizione");
    sheetTestate.getRange(1, 16).setFontWeight("bold");

    // Leggi header e dati da testate
    const dataTestate = sheetTestate.getDataRange().getValues();
    const headersT = dataTestate.shift();
    const idxT = h => headersT.indexOf(h);
    // Raggruppa un header per Documento (prima occorrenza)
    const orderHeaders = {};
    const orderRows = {}; // conserviamo anche il numero di riga nel foglio testate

    dataTestate.forEach((r,i) => {
        const num = r[idxT('Documento (10 Campo Obbligatorio)')];
        if (!orderHeaders.hasOwnProperty(num)) {
            orderHeaders[num] = r;
            orderRows[num] = i + 2; // perché dataTestate parte dalla riga 2
        }
    });

    // Leggi header e dati da righe
    const dataRighe = sheetRighe.getDataRange().getValues();
    const headersR = dataRighe.shift();

    // Pre-calcolo gli indici per non rifare indexOf ad ogni riga
    const idxNum  = headersR.indexOf('N_Documento (30 Obbligatorio)');
    const idxCode = headersR.indexOf('Codice_Articolo (20 Obbligatorio)');
    const idxQty  = headersR.indexOf('Quantita (Numero Obbligatorio)');

    // orderDetails[orderNumber] = { [codiceArticolo]: quantitaTotale }
    const orderDetailsRaw = {};

    dataRighe.forEach(r => {
      const num  = r[idxNum];
      const code = r[idxCode];
      const qty  = Number(r[idxQty]) || 0;

      // Salta righe vuote o senza codice / numero documento
      if (!num || !code || qty === 0) return;

      if (!orderDetailsRaw[num]) {
        orderDetailsRaw[num] = {};
      }
      if (!orderDetailsRaw[num][code]) {
        orderDetailsRaw[num][code] = 0;
      }

      // Somma le quantità per stesso codice/articolo
      orderDetailsRaw[num][code] += qty;
    });

    // Normalizza in array come prima: [{ codiceArticolo, quantita }, ...]
    const orderDetails = {};
    Object.keys(orderDetailsRaw).forEach(num => {
      orderDetails[num] = Object.keys(orderDetailsRaw[num]).map(code => ({
        codiceArticolo: code,
        quantita: orderDetailsRaw[num][code]
      }));
    });

    // CONTROLLA COERENZA TRA TESTATE E RIGHE
    if (!checkDocumentConsistency_(orderHeaders, orderDetailsRaw))
    {
      ui.alert("Incongruenze tra gli ordini!");
      return;
    }

    // Elabora ciascun ordine
    Object.keys(orderHeaders).forEach(num => {
        const hdr = orderHeaders[num];
        const rowIndex = orderRows[num];
        const details = orderDetails[num] || [];
        if (details.length === 0) {
            // Nessuna riga di dettaglio
            return;
        }
        // Estrai campi dal header ordine
        const ragioneSociale = hdr[idxT('Ragione Sociale (30)')];
        const indirizzo      = hdr[idxT('Indirizzo (50)')];
        const localita       = hdr[idxT('Localita (30)')];
        const provincia      = hdr[idxT('Provincia (2)')];
        const cap            = hdr[idxT('CAP (5)')];
        const nazione        = hdr[idxT('Nazione (2)')];
        const urgente        = hdr[idxT('Urgente (Si o lasciare vuoto)')];
        const appuntamento   = hdr[idxT('Appuntamento (Si o lasciare vuoto)')];
        const email          = hdr[idxT('Mail (100)')];
        const telRaw         = hdr[idxT('Telefono (14)')];
        const noteSped       = hdr[idxT('Note (70)')];
        const cashVal        = parseFloat(hdr[idxT('Importo_Contrassegno (Numero con virgola)')] || 0);
        const cashType       = hdr[idxT('Tipo_Contrassegno')];
        const noteMagazzino  = hdr[idxT('Note Magazzino')];
        const priority       = hdr[idxT('Urgente (Si o lasciare vuoto)')];

        // Normalizza telefono
        const tel = String(telRaw).replace(/[\s\-]/g, '');

        // Costruisci header per API
        const scotUsciteHeader = {
            business_name:          ragioneSociale,
            document_date:          Utilities.formatDate(new Date(), tz, 'yyyyMMdd'),
            attachment:             false,
            address:                indirizzo,
            location:               localita,
            province:               provincia,
            zip_code:               cap,
            nation:                 nazione,
            urgent:                 urgente,
            delivery_date:          '',
            appointment:            appuntamento,
            email:                  email,
            tel_reference:          tel,
            carrier_note:           noteSped,
            warehouse_note:         noteMagazzino,
            cash_on_delivery_value: cashVal,
            cash_on_delivery_type:  cashType
        };

        // Costruisci righe per API

        const scotUsciteRows = details.map(d => ({
          id:       num,
          code:     d.codiceArticolo,
          quantity: d.quantita
        }));

        var campagnaNome = "N.D.";
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var currentSheet = ss.getActiveSheet();  // "Current" tab
        var campagnaNome = currentSheet.getName();

        if (inviaSpedizioni_(num, scotUsciteHeader, scotUsciteRows, ragioneSociale, campagnaNome))
        {
          sheetTestate.getRange(rowIndex, 16).setValue(todayHuman);
        }
        else
        {
          sheetTestate.getRange(rowIndex, 16).setValue("ERRORE!");
          ui.alert("Ordine non spedito: " + ragioneSociale);
        }
    });
}

// --------------------
// Funzione di controllo
// --------------------
function checkDocumentConsistency_(orderHeaders, orderDetails) 
{
  const missingInRighe = [];
  const missingInTestate = [];

  if (orderHeaders === null || orderDetails === null)
    return;

  // Documenti presenti in Testate ma non in Righe
  Object.keys(orderHeaders)
    .filter(num => num) // evita chiavi vuote/null
    .forEach(num => {
      if (!orderDetails.hasOwnProperty(num)) {
        missingInRighe.push(num);
      }
    });

  // Documenti presenti in Righe ma non in Testate
  Object.keys(orderDetails)
    .filter(num => num)
    .forEach(num => {
      if (!orderHeaders.hasOwnProperty(num)) {
        missingInTestate.push(num);
      }
    });

  if (!missingInRighe.length && !missingInTestate.length) {
    // tutto ok
    return true;
  }

  let msg = "Incongruenze tra Testate e Righe trovate:\n\n";
  if (missingInRighe.length) {
    msg += "- Documenti presenti in TESTATE ma assenti in RIGHE: " +
           missingInRighe.join(", ") + "\n";
  }
  if (missingInTestate.length) {
    msg += "- Documenti presenti in RIGHE ma assenti in TESTATE: " +
           missingInTestate.join(", ") + "\n";
  }

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
  
  return false;

  // Se vuoi fermare il processo:
  throw new Error(msg);
  // In alternativa potresti usare:
  // SpreadsheetApp.getUi().alert("Errore coerenza dati", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}


function inviaSpedizioni_(numeroOrdine, scotUsciteHeader, scotUsciteRows, clienteNome, campagnaNome)
{
    if (numeroOrdine === null)
      return;
    
    var result = scotOrdiniUscita_(numeroOrdine, "MDB", scotUsciteHeader, scotUsciteRows, clienteNome, campagnaNome);
    if (result) 
    {
      Logger.log("SCOT Uscite: ordine " + numeroOrdine +" succeded! Result: %s", JSON.stringify(result));
//      ui.alert("SCOT Uscite: ordine " + numeroOrdine +" succeded! Result: " + JSON.stringify(result));

      // Set SPEDITO Columns
      
//      currentSheet.getRange(activeRow, shippingStartedCol).setValue(numeroOrdine);
//      currentSheet.getRange(activeRow, 15).setValue("Spedito: " + numeroOrdine);

      return true;
    }
    else
    {
      Logger.log("SCOT Uscite: ordine " + numeroOrdine +" ERROR! Result: %s", JSON.stringify(result));
      return false;
    }
}


function notificaSpedizioni() 
{
  var ui = SpreadsheetApp.getUi();
    
  // Get the current active Google Spreadsheet (IMDB_Ordini)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();  // "Current" tab

  var shippingStartedCol = getEmailRiepilogoColumn_(currentSheet, 'Shipping Started', 33);
  var notificationCol = getEmailRiepilogoColumn_(currentSheet, 'Notifica Spedizione', 33);

  // Get the currently selected range of cells on the active sheet
  var activeRange = currentSheet.getActiveRange();

  // Retrieve all values in the selected range as a 2D array
  var selectedValues = activeRange.getValues();

  // Get the starting row of the active range
  var startingRow = activeRange.getRow();

  var activeRowNumber = 1;

  // Iterate over each row in the selected range
  for (var i = 0; i < selectedValues.length; i++)
  {
    var activeRow = startingRow + i;

    // Skip rows hidden by filter
    if (currentSheet.isRowHiddenByFilter(activeRow))
    {
      continue;
    }

    // Skip rows hidden manually or collapsed/group-hidden
    if (currentSheet.isRowHiddenByUser(activeRow))
    {
      continue;
    }

    var lastName = currentSheet.getRange(activeRow, 4).getValue();       // Column C (Surname)
    var firstName = currentSheet.getRange(activeRow, 5).getValue();     // Column D (First name)

    if (currentSheet.getRange(activeRow, notificationCol).getValue())
    {
      var responseDo = ui.alert("Sicuro di voler mandare notifica a " + lastName + " " + firstName + "? Sembra che tu l'abbia già mandata...", ui.ButtonSet.YES_NO);
      if (responseDo === ui.Button.NO) 
      {
        continue; // Go next customer
      }
    }
    
    if (currentSheet.getRange(activeRow, shippingStartedCol).getValue())
    {
      var responseDo = ui.alert("Sicuro di voler mandare notifica " + lastName + " " + firstName + "? Sembra che tu  abbia già spedito...", ui.ButtonSet.YES_NO);
      if (responseDo === ui.Button.NO) 
      {
        continue; // Go next customer
      }
    }
  
    // Get the value from Column A of the current row (assuming Column A is the first column)
    var colAValue = currentSheet.getRange(activeRow, 1).getValue();
    
    // Check if Column A starts with "Cliente "
    if (!colAValue.startsWith("Cliente ")) {
      // If it doesn't start with "Cliente ", abort and show an error message
      SpreadsheetApp.getUi().alert("Errore: Controlla di essere in un foglio di ordini nella riga relativa ad un cliente.");
      continue; // Go next customer
    }
    
    // Get the "Indirizzi Spedizione" and "Indirizzi Fatturazione" sheets
    var indirizziSpedizioneSheet = ss.getSheetByName('Indirizzi Spedizione');
    var indirizziFatturazioneSheet = ss.getSheetByName('Indirizzi Fatturazione');

    // Get the current date in the format YYYY_MM_DD
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy_MM_dd");
    var currentDateForOrder = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyMM");
    var currentDateFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    // Extract the necessary information from the current row in the "Current" sheet
    var codiceCliente = currentSheet.getRange(activeRow, 2).getValue(); // Column B
    var codiceAzienda = currentSheet.getRange(activeRow, 3).getValue(); // Column B
    var codiceCliente     = currentSheet.getRange(activeRow, 2).getValue(); // Colonna 2 => SOLO questa per Mautic

    // Numero ordine (come prima)
    var numeroOrdine = currentDateForOrder + String(codiceCliente).toString().padStart(6, '0');

    var payload = retrieveMautiCustomerData_(codiceCliente);

    if (!payload) continue; // Salta in caso di errori

    var validationError = IMDBCommonLibs.validateShippingDataHeadless(payload);
    if (validationError) {
      ui.alert(validationError);
      continue;
    }

    var templateOptions = 
        ["campagna " + currentSheet.getName(),
        "nome " + payload.firstName,
        "cognome " + payload.surname,
        "codice " + payload.codiceFiscale,
        "indirizzo " + payload.indirizzo,
        "citta " + payload.localita,
        "cap " + payload.zipCode,
        "provincia " + payload.provinciaDestinatario,
        "telefono " + payload.telefonoFormatted,
        "note " + payload.noteSpedizione];//"note " + destinatario[12] + " - " + telefonoFormatted];

        var waResponse = sendWhatsAppCloudTemplateMessage_("39" + payload.telefonoFormatted, "20260206_imdb_dati_di_spedizione_v_1_2", "it", templateOptions);

        // Converte la risposta JSON in un oggetto
        var waResponseObj = JSON.parse(waResponse);

        // Controlla se l'oggetto contiene un errore
        if (waResponseObj.error) 
        {
          SpreadsheetApp.getUi().alert("ERRORE: " + waResponseObj.toString() + ". Messaggio NON inviato a " + lastName + " " + firstName + "!");
        }
        else
        {
          // Write "DONE" in the 20th column of the corresponding sheet row.
          //SpreadsheetApp.getUi().alert("Messaggio inviato a " + lastName + " " + firstName + "!");
          Logger.log("Messaggio inviato a " + lastName + " " + firstName + "!");
          currentSheet.getRange(activeRow, notificationCol).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
        }
  } // Next client
}


function submitFormOrderReceived_(email, firstName, lastName, phoneNumber, actualCampaign, actualDate) 
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
  //Logger.log("Response: " + response.getContentText());

  return response.getContentText();
}



function submitFormShippingStarted_(email, firstName, lastName, actualCampaign, actualDate) 
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

/*function creaFileOrdiniIDIKA() {
  try {
    // Determine the order row by finding the correct supplier code match
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Ask the user for the supplier code
    var ui = SpreadsheetApp.getUi();
    var supplierCodeResponse = ui.prompt("Inserire il codice del fornitore (2 lettere):");
    if (supplierCodeResponse.getSelectedButton() !== ui.Button.OK) {
      ui.alert("Operazione annullata dall'utente.");
      return;
    }
    var supplierCode = supplierCodeResponse.getResponseText();
    if (!supplierCode || supplierCode.length != 2) {
      ui.alert("Codice fornitore non valido. Riprova.");
      return;
    }

    // Determine the orderRow by searching down column 2 until the 3rd and 4th chars of the cell match the supplierCode
    var orderRow = 10;
    while (true) {
      var cellValue = sheet.getRange(orderRow, 2).getValue();
      if (cellValue && cellValue.substring(2, 4) === supplierCode) {
        break;
      }
      orderRow++;
    }
    
    // Retrieve values from the active sheet
    var orderNumber = sheet.getRange(orderRow, 1).getValue(); // Column A
    var orderCode = sheet.getRange(orderRow, 2).getValue();   // Column B
    var orderName = sheet.getRange(orderRow, 3).getValue();   // Column C
    
    // Determine the first column by finding the first occurrence of the supplier code in row 1
    var firstColumn = 30;
    while (true) {
      var cellValue = sheet.getRange(1, firstColumn).getValue();
      if (cellValue.substring(0, 2) === supplierCode) {
        break;
      }
      firstColumn++;
    }
    
    // Determine the last column by moving right until the first two letters of the next cell are different from the current cell
    var lastColumn = firstColumn;
    var currentValue = sheet.getRange(1, firstColumn).getValue();
    while (true) {
      var nextValue = sheet.getRange(1, lastColumn + 1).getValue();
      if (nextValue === "" || nextValue.substring(0, 2) !== currentValue.substring(0, 2)) {
        break;
      }
      lastColumn++;
    }
    
    // Prompt the user for the row with quantities
    var orderQuantities = 29; // parseInt(promptUser_("Inserisci la riga con i conteggi finali (inclusi OMAGGI):"));

    // Get the current date in YYYY-MM-DD format for file naming
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // Create a new Spreadsheet to store the Ordini data
    var outputSpreadsheet = SpreadsheetApp.create("Ordini Output");
    var outputSheet = outputSpreadsheet.getActiveSheet();

    // Set the headers with specified titles in bold
    var headers = [["Numero Ordine", "Codice Fornitore", "Rag.Sociale Fornitore", "Data ordine", "Codice Articolo", "Quantita da ricevere"]];
    outputSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    outputSheet.getRange(1, 1, 1, headers[0].length).setFontWeight("bold");

    // Start writing from the second row
    var outputRow = 2;
    
    // Loop through each specified column and create the output rows
    for (var currentColumn = firstColumn; currentColumn <= lastColumn; currentColumn++) {
      var orderItem = sheet.getRange(1, currentColumn).getValue();         // Row 1, current column
      var quantity = sheet.getRange(orderQuantities, currentColumn).getValue(); // orderQuantities row, current column
      
      // Skip columns where quantity is zero
      if (quantity === 0 || quantity === "") {
        continue;
      }
      
      // Write data to the output sheet
      outputSheet.getRange(outputRow, 1).setValue(orderNumber);       // Column A: Numero Ordine
      outputSheet.getRange(outputRow, 2).setValue(orderCode);         // Column B: Codice Fornitore
      outputSheet.getRange(outputRow, 3).setValue(orderName);         // Column C: Rag.Sociale Fornitore
      outputSheet.getRange(outputRow, 4).setValue(currentDate);       // Column D: Data ordine
      outputSheet.getRange(outputRow, 5).setValue(orderItem);         // Column E: Codice Articolo
      outputSheet.getRange(outputRow, 6).setValue(quantity);          // Column F: Quantita da ricevere
      
      // Move to the next row in the output file
      outputRow++;
    }

    // Wait for the data to be fully written
    SpreadsheetApp.flush();

    // Construct the output file name in YYYY-MM-DD_IMDB_orderCode format
    var outputFileName = `Ingressi_IMDBsrls_${orderCode}_${currentDate}.xlsx`;

    // Use Drive API to export Google Sheet to Excel format
    var url = `https://www.googleapis.com/drive/v3/files/${outputSpreadsheet.getId()}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    // Save the Excel file temporarily in Google Drive
    var blob = response.getBlob().setName(outputFileName);
    var folder = DriveApp.getRootFolder(); // Change to desired folder if needed
    var file = folder.createFile(blob);

    // Get the download link
    var downloadUrl = file.getDownloadUrl();

    // Prepare email content
    var htmlContent = `<p>Si trasmette il tracciato di caricamento ingressi relativo all'ordine ${orderNumber} del fornitore ${orderCode} - ${orderName}</p><p>Cordiali saluti</p><br><p>Il Massimo del Bere</p>`;
    var recipientEmail = "info@ilmassimodelbere.it"; 
    var ccEmail = "info@ilmassimodelbere.it";
    var bccEmail = "ordini@ilmassimodelbere.it";
    var subject = `Tracciati di caricamento ingressi relativi al fornitore ${orderName}`;
    var senderName = "Il Massimo del Bere";

    // Show Yes/No/Test prompt to user for sending email with HTML content preview
    var htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(600).setHeight(400);
    ui.showModalDialog(htmlOutput, "Conferma invio email");
    var response = ui.alert("Vuoi procedere con l'invio dell'email?", ui.ButtonSet.YES_NO_CANCEL);
    
    if (response === ui.Button.YES) {
      //recipientEmail = "Laura Serusi <lauraserusi@i-dika.com>, Depositi <depositi@i-dika.com>, Davide Chiappinotto <davidechiappinotto@i-dika.com>";
    } else if (response === ui.Button.NO) {
      Logger.log("Email sending canceled by the user.");
      // Clean up: delete the temporary Google Spreadsheet and Excel file from Google Drive
      DriveApp.getFileById(outputSpreadsheet.getId()).setTrashed(true);
      file.setTrashed(true);
      return;
    } else if (response === ui.Button.CANCEL) {
      Logger.log("Email test mode.");
      recipientEmail = "info@ilmassimodelbere.it";
    }

    // Send email with attachment
    sendEmailViaSMTP_(htmlContent, recipientEmail, subject, senderName, ccEmail, bccEmail, file);

    // Clean up: delete the temporary Google Spreadsheet and Excel file from Google Drive
    DriveApp.getFileById(outputSpreadsheet.getId()).setTrashed(true);
    file.setTrashed(true);

  } catch (error) {
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
  }
}*/

/*function inviaAnagraficheIDIKA() {
  try {
    // Get the active spreadsheet and the "Ricarichi MASTER (non modificare)" sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Ricarichi MASTER (non modificare)");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("Sheet 'Ricarichi MASTER (non modificare)' not found.");
      return;
    }

    var ui = SpreadsheetApp.getUi();

    // Create a new temporary Spreadsheet to hold the anagrafica data
    var outputSpreadsheet = SpreadsheetApp.create("Anagrafiche Output");
 
    // Create or get the tab "Anagrafiche" or rename "Sheet1" if it exists
    var outputSheet = outputSpreadsheet.getSheetByName('Sheet1');
    if (outputSheet) {
      outputSheet.setName('Anagrafiche');
    } else {
      outputSheet = outputSpreadsheet.getSheetByName('Anagrafiche');
      if (!outputSheet) {
        outputSheet = outputSpreadsheet.insertSheet('Anagrafiche');
      }
    }
  
    // Set the headers with specified titles in bold
    var headers = [["Codice articolo", "Descrizione articolo", "Pezzi per cartone", "Peso Lordo cartone", "Volume Cartone"]];
    outputSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    outputSheet.getRange(1, 1, 1, headers[0].length).setFontWeight("bold");

    // Start writing from the second row
    var outputRow = 2;
    var lastRow = sheet.getLastRow();

    // Loop through each row in the "Ricarichi MASTER (non modificare)" sheet and copy the values to the output sheet only if MAGAZZINO == "I-DIKA"
    for (var row = 2; row <= lastRow; row++) {
      var magazzinoValue = sheet.getRange(row, 3).getValue(); // Assuming MAGAZZINO is in column 3 (C)
      if (magazzinoValue !== "I-DIKA") {
        continue;
      }

      var codiceArticolo = sheet.getRange(row, 1).getValue();       // Column A: Codice articolo
      var descrizioneArticolo = sheet.getRange(row, 2).getValue();  // Column B: Descrizione articolo

      // Write data to the output sheet
      outputSheet.getRange(outputRow, 1).setValue(codiceArticolo);         // Column A: Codice articolo
      outputSheet.getRange(outputRow, 2).setValue(descrizioneArticolo);    // Column B: Descrizione articolo
      outputSheet.getRange(outputRow, 3).setValue(''); // Pezzi per cartone (blank for now)
      outputSheet.getRange(outputRow, 4).setValue(''); // Peso Lordo cartone (blank for now)
      outputSheet.getRange(outputRow, 5).setValue(''); // Volume Cartone (blank for now)

      // Move to the next row in the output file
      outputRow++;
    }

    // Wait for the data to be fully written
    SpreadsheetApp.flush();

    // Construct the output file name in YYYY-MM-DD format for file naming
    var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var outputFileName = `Anagrafiche_IMDBsrls_${currentDate}.xlsx`;

    // Use Drive API to export Google Sheet to Excel format
    var url = `https://www.googleapis.com/drive/v3/files/${outputSpreadsheet.getId()}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });

    // Save the Excel file temporarily in Google Drive
    var blob = response.getBlob().setName(outputFileName);
    var folder = DriveApp.getRootFolder(); // Change to desired folder if needed
    var file = folder.createFile(blob);

    // Prepare email content
    var htmlContent = `<p>Si trasmette il tracciato delle Anagrafiche aggiornato al ${currentDate}</p><p>Cordiali saluti</p><br><p>Il Massimo del Bere</p>`;
    var recipientEmail = "info@ilmassimodelbere.it";
    var ccEmail = "info@ilmassimodelbere.it";
    var bccEmail = "ordini@ilmassimodelbere.it";
    var subject = `Anagrafiche aggiornate al ${currentDate}`;
    var senderName = "Il Massimo del Bere";

    // Show Yes/No/Test prompt to user for sending email with HTML content preview
    var htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(600).setHeight(400);
    ui.showModalDialog(htmlOutput, "Conferma invio email");
    var response = ui.alert("Vuoi procedere con l'invio dell'email?", ui.ButtonSet.YES_NO_CANCEL);
    
    if (response === ui.Button.YES) {
      //recipientEmail = "Laura Serusi <lauraserusi@i-dika.com>, Depositi <depositi@i-dika.com>, Davide Chiappinotto <davidechiappinotto@i-dika.com>";
    } else if (response === ui.Button.NO) {
      Logger.log("Email sending canceled by the user.");
      // Clean up: delete the temporary Google Spreadsheet and Excel file from Google Drive
      DriveApp.getFileById(outputSpreadsheet.getId()).setTrashed(true);
      file.setTrashed(true);
      return;
    } else if (response === ui.Button.CANCEL) {
      Logger.log("Email test mode.");
      recipientEmail = "ordini@ilmassimodelbere.it";
    }

    // Send email with attachment
    sendEmailViaSMTP_(htmlContent, recipientEmail, subject, file, ccEmail, bccEmail, senderName);

    // Clean up: delete the temporary Google Spreadsheet and Excel file from Google Drive
    DriveApp.getFileById(outputSpreadsheet.getId()).setTrashed(true);
    file.setTrashed(true);

  } catch (error) {
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
  }
}*/



// Function to convert a column letter (e.g., "A", "AB") to a column number (1-based index)
function columnLetterToNumber_(columnLetter) {
  var columnNumber = 0;
  var length = columnLetter.length;
  for (var i = 0; i < length; i++) {
    columnNumber *= 26;
    columnNumber += (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return columnNumber;
}

function promptUser_(prompt) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(prompt);
  if (response.getSelectedButton() == ui.Button.OK) {
    return response.getResponseText();
  } else {
    throw new Error("Action cancelled by the user.");
  }
}

function createProformaFromSheet() {
  // Get the active sheet and the active row
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeRow = sheet.getActiveRange().getRow();

  // Retrieve customer information from columns B to G
  var customerID = sheet.getRange(activeRow, 2).getValue(); // Column B
  var surname = sheet.getRange(activeRow, 3).getValue();    // Column C
  var firstName = sheet.getRange(activeRow, 4).getValue();  // Column D
  var fiscalCode = sheet.getRange(activeRow, 5).getValue(); // Column E
  var phone = sheet.getRange(activeRow, 6).getValue();      // Column F
  var mobilePhone = sheet.getRange(activeRow, 7).getValue();// Column G

  // Prompt user to input the first and last columns to read the articles
  var ui = SpreadsheetApp.getUi();
  var firstColPrompt = ui.prompt('Enter the first column number where the articles start (e.g., 8 for column H)').getResponseText();
  var lastColPrompt = ui.prompt('Enter the last column number where the articles end').getResponseText();
  var firstCol = columnLetterToNumber_(firstColPrompt);
  var lastCol = columnLetterToNumber_(lastColPrompt);
  
  // Validate column range
  if (isNaN(firstCol) || isNaN(lastCol) || firstCol > lastCol) {
    ui.alert("Invalid column range provided.");
    return;
  }

  // Read article details for each column between firstCol and lastCol
  var items = [];
  for (var col = firstCol; col <= lastCol; col++) {
    var articleCode = sheet.getRange(1, col).getValue();    // Row 1: Article Code
    var producer = sheet.getRange(2, col).getValue();       // Row 2: Producer
    var description = sheet.getRange(3, col).getValue();    // Row 3: Description
    var quantity = sheet.getRange(activeRow, col).getValue(); // Actual row: Quantity

    // Only add the article if the quantity is greater than 0
    if (quantity > 0) {
      items.push({
        'nome': articleCode + " - " + producer + " - " + description,
        'quantita': quantity,
        'prezzo_netto': 100,  // You can adjust the price here or add it as a column in the sheet
        'cod_iva': 22  // Default to 22% VAT. You can adjust based on your need
      });
    }
  }

  // Proceed if there are any items to invoice
  if (items.length === 0) {
    ui.alert("No items to add to the proforma invoice.");
    return;
  }

  // Now prepare the proforma invoice data for the API
  var invoiceData = {
    'api_uid': 'huGs53i3E8ozJ6cgklQ5pHF8vjz0LFqc',
    'api_key': 'Z3srRK6JkGnZM35W750omo9xqJH7LFNKEogmPVTdpwJ0fl0keIxfEAVVQ6L5ev2v',
    'dati_documento': {
      'tipo': 'proforma',
      'data': new Date().toISOString().split('T')[0], // Current date in YYYY-MM-DD format
      'cliente': {
        'nome': firstName + ' ' + surname,
        'cf': fiscalCode,
        'telefono': phone,
        'mobile': mobilePhone
      },
      'lista_articoli': items
    }
  };

  // Send the data to the Fatture in Cloud API
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(invoiceData)
  };
  
  try {
    var response = UrlFetchApp.fetch('https://developers.fattureincloud.it', options);
    var responseData = JSON.parse(response.getContentText());
    if (responseData.error) {
      ui.alert("Error: " + responseData.error);
    } else {
      ui.alert("Proforma invoice created successfully.");
    }
  } catch (error) {
    ui.alert("Error creating proforma invoice: " + error.message);
  }
}

// This function loops between all the pages of the response and pushes the results on the sheet
function listInvoices_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  //sheet.clear();
  sheet.appendRow([
    "id",
    "type",
    "numeration",
    "subject",
    "visible_subject",
    "amount_net",
    "amount_vat",
    "date",
    "next_due_date",
    "url",
    "client name",
    "client tax_code",
    "client vat_number",
  ]);
  var fattureUrl = "https://api-v2.fattureincloud.it";
  var endpoint = "/c/1294648/issued_documents";
  var headers = {
    "Content-Type": "application/json",
    Authorization: "Bearer a/eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJyZWYiOiJET2Rja1RYS3Bzd2M0elN2cUhGdm1uNUNaNlJIcjd3VSJ9.1oQAGDpXjx2YsVwpCknBBcvMu9kqYRk1CYxbKrFPoCQ"
  };

  var options = {
    method: "get",
    headers: headers,
  };

  var queryParams = ["type=" + "invoice"];
  var pageKey = "page";
  var pageNum = 1;
  var nextPageUrl;
  var data = {},
    output = [];
  do {
    var currentUrl = fattureUrl + endpoint + "?" + queryParams.join("&");
    currentUrl += "&" + pageKey + "=" + pageNum;

    var response = UrlFetchApp.fetch(currentUrl, options);
    data = JSON.parse(response.getContentText());
    var invoices = data.data;
    for (const index in invoices) {
      sheet.appendRow([
        invoices[index].id,
        invoices[index].type,
        invoices[index].numeration,
        invoices[index].subject,
        invoices[index].visible_subject,
        invoices[index].amount_net,
        invoices[index].amount_vat,
        invoices[index].date,
        invoices[index].next_due_date,
        invoices[index].url,
        invoices[index].entity.name,
        invoices[index].entity.tax_code,
        invoices[index].entity.vat_number,
      ]);
    }
    pageNum++;
    nextPageUrl = data.next_page_url;
  } while (nextPageUrl);
}




function testCreateProforma ()
{
  IMDBCommonLibs.createFICInvoice ("Pipppo Pluto", "client", "person", "Massimo", "Bombino", "", "Via di qui, 32", "Milano", "20123", "PV", "ordini@ilmassimodelbere.it", "+39-348-2639796", "BMBMMSAMMM", "pec@pec.it", "0000000", "BMBMMSAMMM", "Spedier subito", "", "IMDB Acquisto", "IMDB - Acconto per merce da consegnare", "497", "Fattura saldata con pagamento Stripe del ","2025-04-26",  "497", "IMDB - Acconto per merce da consegnare (rif. IMDB VIP Club Diamond)");
}

function testsendWhatsAppCloudTemplateMessage_ ()
{
  sendWhatsAppCloudTemplateMessage_("393482639796", "2025_imdb_anteprima_borgogna_2026_v_1", "it", ["Max"]);
}

/** 
 * @param {string} to - The recipient's phone number in international format (e.g., "393482639796" without the plus sign).
 * @param {string} templateName - The name of the pre-approved template (e.g., "hello_world").
 * @param {string} languageCode - The language code (e.g., "en_US").
 * @param {Array} templateParameters - Optional array of text parameters for the template body.
 * @return {string} The API response.
 */

function sendWhatsAppCloudTemplateMessage_(to, templateName, languageCode, templateParameters) {

  // Replace these with your own credentials.
  var accessToken = 'EAAGkqJyudNsBOxCxoEvfwPqXZBF1asDuZCJAxkHktqz2o9NWuelJhkTaCGdhxodQIZA8pcxvCZCzR0xY00MwdeNwBZCDDuXAiWE1BDN5ITI1klZABfrKYDY9fvO5Nda07YFoGPt0hu9gQ5r6HlKdB0RLLptMiQxPPz7gTmcIl9QnfE1ZCYLRNKusYgr0metLHZA4SQZDZD';
  var phoneNumberId = '620172421177804'; // This is provided by Meta in your WhatsApp Business setup
  
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


function createOrderInvoice()
{
  const ui = SpreadsheetApp.getUi();

  const currentSpreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = SpreadsheetApp.getActiveSheet();
  const currentCell = currentSheet.getActiveCell();
  const currentRow = currentCell.getRow();
  
  var fatturaCol = getEmailRiepilogoColumn_(currentSheet, 'Fattura', 33);
  var fatturaVIPCol = getEmailRiepilogoColumn_(currentSheet, 'Fattura Acconto VIP', 33);
  var pagamentoCol = getEmailRiepilogoColumn_(currentSheet, 'Stato pagamento', 33);
  var totaleEffettivoCol = getEmailRiepilogoColumn_(currentSheet, 'Totale', 33);
  var startColumn = getEmailRiepilogoColumn_(currentSheet, 'Codice bottiglia', 1) + 2;

  // Get the "Indirizzi Spedizione" and "Indirizzi Fatturazione" sheets
  var indirizziSpedizioneSheet = currentSpreadSheet.getSheetByName('Indirizzi Spedizione');
  var indirizziFatturazioneSheet = currentSpreadSheet.getSheetByName('Indirizzi Fatturazione');

  // Get the currently selected range of cells on the active sheet
  var activeRange = currentSheet.getActiveRange();

  // Retrieve all values in the selected range as a 2D array
  var selectedValues = activeRange.getValues();

  // Get the starting row of the active range
  var startingRow = activeRange.getRow();
  
  var activeRow = 1;

  // Iterate over each row in the selected range
  for (var i = 0; i < selectedValues.length; i++) 
  {
    activeRow = startingRow + i;

    var vipStatus = false;

    // VIP Status
    if ((currentSheet.getRange(currentRow, 12).getValue()!= "") && (currentSheet.getRange(currentRow, 12).getValue()!= "N.A."))
      vipStatus = true;

    // Get the value from Column A of the current row (assuming Column A is the first column)
    var colAValue = currentSheet.getRange(activeRow, 1).getValue();
    
    // Check if Column A starts with "Cliente"
    if (!colAValue.startsWith("Cliente")) {
      // If it doesn't start with "Cliente", abort and show an error message
      SpreadsheetApp.getUi().alert("Errore: Controlla di essere in un foglio di ordini nella riga relativa ad un cliente.");
      return; // Abort the script
    }


    if (currentSheet.getRange(activeRow,fatturaCol).getValue())
    {
      var responseDo = ui.alert("Hai già creato una fattura: " + currentSheet.getRange(activeRow,fatturaCol).getValue() + ", vuoi crearla di nuovo?", ui.ButtonSet.YES_NO);
      if (responseDo === ui.Button.NO)
        continue;
    }

    // Scan Row 1 from column AH onwards to get the Range_Prodotti
    var rangeProdotti = [];
    var lastColumn = currentSheet.getLastColumn();
    
    for (var col = startColumn; col <= lastColumn; col++) {
      var value = currentSheet.getRange(1, col).getValue(); // Row 1
      if (value) {
        rangeProdotti.push(col); // Save the column index for products
      } else {
        break; // Stop if we find an empty cell
      }
    }
    
    // Extract the necessary information from the current row in the "Current" sheet
    var codiceCliente = currentSheet.getRange(activeRow, 2).getValue(); // Column B
    var codiceAzienda = currentSheet.getRange(activeRow, 3).getValue(); // Column B
    var partitaIVA = currentSheet.getRange(activeRow, 7).getValue(); // Column B

    // Check HORECA

    var isHORECA = false;
    if (codiceAzienda)
      isHORECA = true;

    var payload = retrieveMautiCustomerData_(codiceCliente);

    if (!payload) continue; // Salta in caso di errori

    var validationError = IMDBCommonLibs.validateShippingDataHeadless(payload);
    if (validationError) {
      ui.alert(validationError);
      continue;
    }

    // 1. Dati cliente
    const clientEntityType        = 'client';
    const clientType              = 'person';
    const clientNome              = payload.firstName;
    const clientCognome           = payload.surname;
    const clientCodice            = codiceCliente;
    const clientAddress           = payload.indirizzo;
    const clientCity              = payload.localita;
    const clientCAP               = payload.zipCode;
    const clientProvincia         = payload.provinciaDestinatario;
    const clientEmail             = payload.email;
    const clientPhone             = payload.telefonoFormatted;
    const clientVatNumber         = partitaIVA;
    const clientPEC               = '';
    const clientSDI               = '';
    const clientCodiceFiscale     = payload.codiceFiscale;
    const clientNoteSpedizione    = '';
    const clientNoteCliente       = '';
    const clienteRagioneSociale   = payload.surname + " " + payload.firstName;
    
    // 2. Dati fattura
    var invoiceSubject;
    var invoiceVisibleSubject;
    var invoiceAmount;

    if (vipStatus)
    {
      invoiceSubject          = 'IMDB - Merce già pagata';
      invoiceVisibleSubject   = 'IMDB - Merce già pagata';
      invoiceAmount = 0;
    }
    else
    {
      invoiceSubject          = 'IMDB - ' + currentSheet.getName();
      invoiceVisibleSubject   = 'IMDB - ' + currentSheet.getName();
      invoiceAmount = currentSheet.getRange(activeRow, 19).getValue();
    }

    // Stripe o bonifico?
    const rawScadenza = String(currentSheet.getRange(activeRow, pagamentoCol).getDisplayValue() || '').trim().slice(0, 10);
    const rawSaldo    = String(currentSheet.getRange(activeRow, pagamentoCol).getDisplayValue() || '').trim().slice(0, 10);

    const clientDataScadenza = rawScadenza ? IMDBCommonLibs.convertDateString(rawScadenza) : new Date();
    const clientDataSaldo    = rawSaldo    ? IMDBCommonLibs.convertDateString(rawSaldo)    : new Date();
    
    var paymentMethod; 
    var paymentEMethod;

    if (vipStatus)
    {
      clientNoteSaldo = 'Merce già pagata ';
      paymentMethod = 'Stripe';
      paymentEMethod = 'MP08'; // Bonifico
    }
    else
    {
      clientNoteSaldo         = 'Fattura saldata con bonifico bancario del ' + clientDataSaldo;
      paymentMethod = 'Bonifico';
      paymentEMethod = 'MP05'; // Bonifico
    }
    
    const clientPaidAmount        = currentSheet.getRange(activeRow, 3).getValue();
    const paymentID = '';
    const invoiceDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    var invoiceItems = [];

    var totaleIMDB = 0.0;

    // Loop through the Range_Prodotti and add new rows for each valid product entry in the current row
    for (var j = 0; j < rangeProdotti.length; j++) 
    {
      var col = rangeProdotti[j];
      
      // Get the quantity to be shipped (actual row and current column in Range_Prodotti)
      var quantitaDaSpedire = currentSheet.getRange(activeRow, col).getValue();
      
      // If the quantity is greater than 0, process the product
      if (quantitaDaSpedire > 0) 
      {
        // Get the article code from Row 1 in the current column
        var codiceArticolo = currentSheet.getRange(1, col).getValue();

        // Get the article code from Row 1 in the current column
        var costoIMDBArticolo = '';
        
        costoIMDBArticolo = Number(currentSheet.getRange(14, col).getValue());
        costoHORECAArticolo = Number(currentSheet.getRange(9, col).getValue());

        totaleIMDB = Number(totaleIMDB + costoIMDBArticolo * quantitaDaSpedire);

        if (vipStatus)
          costoIMDBArticolo = 0;


        // Get the article description from Row 2-3 in the current column
        var nomeArticolo = currentSheet.getRange(3, col).getValue();

        // Get the article description from Row 2-3 in the current column
        var descrizioneArticolo = currentSheet.getRange(2, col).getValue() + " - " + currentSheet.getRange(3, col).getValue();
        

        if (isHORECA)
        {
          // Prepare the new row to append to "Spedizioni"
          var newRow = [
            { 
              code: codiceArticolo, 
              name: nomeArticolo, 
              description: descrizioneArticolo,
              qty: quantitaDaSpedire, 
              net_price: costoHORECAArticolo, 
              vat: { id: 0, value: 22, description: 'IVA 22%' } },
          ];
        }
        else
        {
          // Prepare the new row to append to "Spedizioni"
          var newRow = [
            { 
              code: codiceArticolo, 
              name: nomeArticolo, 
              description: descrizioneArticolo,
              qty: quantitaDaSpedire, 
              gross_price: costoIMDBArticolo, 
              vat: { id: 0, value: 22, description: 'IVA 22%' } },
          ];
        }
        // Append the new row to "Spedizioni"
        invoiceItems.push(newRow);
      }
    }

    var costoSpedizione = Number(currentSheet.getRange(activeRow, 17).getValue());
    var totaleSCONTO = Number(totaleIMDB - currentSheet.getRange(activeRow, 19).getValue() + costoSpedizione);

    if (!vipStatus)
    {
      // Spedizione
      if (costoSpedizione > 0)
      {
        // Prepare the SPEDIZIONE row 
        var newRow = [
          { 
            code: 'SPED', 
            name: 'Spese di spedizione',
            gross_price: costoSpedizione, 
            vat: { id: 0, value: 22, description: 'IVA 22%' } },
        ];
        
        // Append the new row to "Spedizioni"
        invoiceItems.push(newRow);
      }

      // Calcola SCONTO (no HORECA)

      if (isHORECA)
        totaleSCONTO = 0;

      if (totaleSCONTO > 0)
      {
        var newRow = [
          { 
            code: 'SCONTO', 
            name: 'Sconto in base alle condizioni concordate in offerta',
            gross_price: -totaleSCONTO,
            vat: { id: 0, value: 22, description: 'IVA 22%' } },
        ];
        
        // Append the new row to "Spedizioni"
        invoiceItems.push(newRow);
      }
    }
    else
    {
      var newRow = [
      { 
        code: '', 
        name: '', 
        description: `Merce già pagata

        Rif. ft. di acconto ` + currentSheet.getRange(activeRow, fatturaVIPCol).getValue() +  `
        Campagna "` + currentSheet.getName() + `"
        Importo merce: ` + totaleIMDB.toFixed(2) + ` EUR
        Sconto: ` + totaleSCONTO.toFixed(2) + ` EUR
        Spedizione: ` + costoSpedizione.toFixed(2) + ` EUR
        Totale effettivo ordine: ` + currentSheet.getRange(activeRow, 19).getValue().toFixed(2) + ' EUR',
        qty: '', 
        gross_price: ''
      }];

      // Append the new row to "Spedizioni"
      invoiceItems.push(newRow);
    }   

    var paymentItems = [];
    
    if(!vipStatus)
    {
      paymentItems = [
      {
        amount: invoiceAmount,
        due_date: clientDataScadenza,
        paid_date: clientDataSaldo,
        id: null,
        forfettari_revenue: '',
        payment_terms: null,
        status: 'paid',
        payment_account: {
          id: 1232535,
          name: 'Credit Agricole',
          virtual: false
        },
        ei_raw: null
      }
      ];
    }
    
    // 3. Chiamata finale
    var data = IMDBCommonLibs.createFICOrderInvoice(
      'invoice',
      clienteRagioneSociale,
      clientEntityType,
      clientType,
      clientNome,
      clientCognome,
      clientCodice,
      clientAddress,
      clientCity,
      clientCAP,
      clientProvincia,
      clientEmail,
      clientPhone,
      clientVatNumber,
      clientPEC,
      clientSDI,
      clientCodiceFiscale,
      clientNoteSpedizione,
      clientNoteCliente,
      invoiceDate,
      invoiceSubject,
      invoiceVisibleSubject,
      invoiceAmount,
      invoiceItems,
      false,
      paymentID,
      paymentMethod,
      paymentEMethod,
      clientNoteSaldo,
      paymentItems,
      !isHORECA
    );
    
    if (data.data.number > 0)
    {
      Logger.log("createFICInvoice() invocata sulla riga ${row} con fattura: " + data.data.number);
      currentSheet.getRange(activeRow, fatturaCol).setValue(data.data.number + "/2026")
    }
    else
    {
      Logger.log("Errore createFICInvoice(): " + data);
    }
  }
}

/**
 * Per ogni riga selezionata nel foglio attuale:
 * - se in colonna L è presente Silver/Gold/Diamond/Black,
 *   legge Cognome (col. 5, E) e Nome (col. 4, D),
 *   apre lo Spreadsheet esterno (ID indicato),
 *   cerca lo sheet chiamato "Cognome Nome",
 *   legge B8 (saldo) e scrive il valore in colonna N della riga corrente.
 */
function aggiornaSaldoVIPSelezionati(range = null)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet(); // foglio ordini corrente (campagna)
  var rangeList;

  if (!range)
  {
    range = SpreadsheetApp.getActiveRange();
    rangeList = SpreadsheetApp.getActiveRangeList();
  }

  var campaignName = currentSheet.getName(); // nome campagna = nome tab corrente

  var DEST_SPREADSHEET_ID = '19SOEhBqA43lWavEnEERukTBrMdyFTJd6XJLcqvNQkhA';
  var destSS = SpreadsheetApp.openById(DEST_SPREADSHEET_ID);

  function upsertCampaignAndGetI_(destSheet, campaignName, totaleIMDB, totaleCliente)
  {
    var START_ROW = 12;
    var COL_A = 1;
    var COL_F = 6; // Totale IMDB
    var COL_G = 7; // Totale Cliente
    var COL_H = 8; // Campagna
    var COL_I = 9; // Valore/Formula da leggere

    // --- trova fine tabella: prima riga vuota in colonna A
    var lastRow = Math.max(destSheet.getLastRow(), START_ROW);
    var colA = destSheet.getRange(START_ROW, COL_A, lastRow - START_ROW + 1, 1).getValues();

    var tableEndRow = START_ROW - 1;
    for (var i = 0; i < colA.length; i++) {
      if (!String(colA[i][0] || '').trim()) break;
      tableEndRow = START_ROW + i;
    }

    if (tableEndRow < START_ROW) {
      Logger.log('Tabella vuota nel sheet VIP "' + destSheet.getName() + '" (da riga ' + START_ROW + ').');
      return null;
    }

    // --- cerca campagna in colonna H
    var numRows = tableEndRow - START_ROW + 1;
    var colH = destSheet.getRange(START_ROW, COL_H, numRows, 1).getValues();

    var campaignRow = null;
    for (var r = 0; r < colH.length; r++) {
      if (String(colH[r][0] || '').trim() === campaignName) {
        campaignRow = START_ROW + r;
        break;
      }
    }

    var lastTableRow;

    // --- se non c'è, crea una nuova riga e duplica l'ultima riga tabella
    if (!campaignRow) {
      lastTableRow = tableEndRow;
      var newRow = lastTableRow + 1;
       
      // 1) aggiungi riga in fondo alla tabella (sposta giù tutto il resto)
      destSheet.insertRowsAfter(lastTableRow, 1);

      // 2) duplica ultima riga tabella nella nuova
      var lastCol = destSheet.getLastColumn();

      destSheet.getRange(lastTableRow, 1, 1, lastCol)
        .copyTo(destSheet.getRange(newRow, 1, 1, lastCol), { formatOnly: true });

      destSheet.getRange(lastTableRow, COL_I)
        .copyTo(destSheet.getRange(newRow, COL_I), { contentsOnly: false });

      campaignRow = newRow;

      var formulaAbove = destSheet.getRange(lastTableRow, COL_I).getFormula();
      if (formulaAbove) {
        // Copia "intelligente": copia il range (non contentsOnly) così i riferimenti si adattano alla nuova riga
        destSheet.getRange(lastTableRow, COL_I).copyTo(destSheet.getRange(campaignRow, COL_I), { contentsOnly: false });
      } 
      else {
        destSheet.getRange(campaignRow, COL_I).setValue(destSheet.getRange(lastTableRow, COL_I).getValue());
      }
    }

    // 3) scrivi i campi richiesti (A..H) e copia formula colonna I dalla riga sopra
    destSheet.getRange(campaignRow, 1, 1, 8).setValues([[
      new Date(),                        // A: data di oggi
      'Ordine "' + campaignName + '"',   // B: descrizione
      '',                                // C
      '',                                // D
      '',                                // E
      totaleIMDB,                        // F: Totale IMDB
      totaleCliente,                     // G: Totale Cliente
      campaignName                      // H: nome campagna
    ]]);

    // dentro upsertCampaignAndGetI_(), sostituisci il return finale:
    SpreadsheetApp.flush();
    var v = Number(destSheet.getRange(tableEndRow, COL_I).getValue() || 0);
    v = Math.round(v * 100) / 100;
    destSheet.getRange('B8').setValue(v);

    return v;

  } //upsertCampaignAndGetI_

  // Helper per processare un range rettangolare
  function processRange(rg)
  {
    var startRow = rg.getRow();
    var numRows = rg.getNumRows();

    for (var i = 0; i < numRows; i++)
    {
      var r = startRow + i;

      // Colonne: D=4 (Cognome), E=5 (Nome), L=12 (Livello), N=14 (saldo da scrivere)
      var livello = String(currentSheet.getRange(r, 12).getValue() || '').trim();
      if (!/^(silver|gold|diamond|black)$/i.test(livello)) {
        continue;
      }

      var nome = String(currentSheet.getRange(r, 5).getValue() || '').trim();
      var cognome = String(currentSheet.getRange(r, 4).getValue() || '').trim();
      if (!nome || !cognome) {
        Logger.log('Nome/Cognome vuoti alla riga ' + r + ' — salto.');
        continue;
      }

      // >>> QUI metti le colonne reali degli importi (placeholder)
      // Esempio: se Totale IMDB è colonna X e Totale Cliente è colonna Y:
      // var totaleIMDB = Number(sheet.getRange(r, X).getValue() || 0);
      // var totaleCliente = Number(sheet.getRange(r, Y).getValue() || 0);

      var totaleIMDBCol = getEmailRiepilogoColumn_(currentSheet, 'Totale IMDB#', 33);
      var totaleClienteCol = getEmailRiepilogoColumn_(currentSheet, 'Totale#', 33);

      var totaleIMDB = Number(currentSheet.getRange(r, totaleIMDBCol).getValue() || 0);      // <-- DA SOSTITUIRE
      var totaleCliente = Number(currentSheet.getRange(r, totaleClienteCol).getValue() || 0);   // <-- DA SOSTITUIRE

      var sheetName = (cognome + ' ' + nome).replace(/\s+/g, ' ').trim();
      var destSheet = destSS.getSheetByName(sheetName);
      if (!destSheet) {
        Logger.log('Sheet "' + sheetName + '" non trovato nello Spreadsheet di destinazione — riga ' + r);
        continue;
      }

      var saldo = upsertCampaignAndGetI_(destSheet, campaignName, totaleIMDB, totaleCliente);
      if (saldo === null) {
        // non siamo riusciti a calcolare/leggere
        continue;
      }

      // nel chiamante, lascia invariato saldo=... ma scrivi N con il valore arrotondato:
      currentSheet.getRange(r, 14).setValue("VIP: " + saldo.toFixed(2).replace('.', ',') + "€");

      IMDBVIPClubScripts.sendCreditSummaryEmail(destSheet); 

    }
  } //processRange()

  // Gestisce multi-selezioni e singola selezione
  if (rangeList)
  {
    var ranges = rangeList.getRanges();
    for (var k = 0; k < ranges.length; k++) {
      processRange(ranges[k]);
    }
  }
  else if (range)
  {
    processRange(range);
  }
}


/*
function checkGiacenzeSpedizioniOLD()
{
  var ui = SpreadsheetApp.getUi();

  var token = IMDBCommonLibs.getScotToken(false);
  if (token) {
    Logger.log("Received token: " + token);
  } else {
    Logger.log("Failed to retrieve token.");
    return false;
  }

  // 1. Recupera le giacenze dal servizio comune
  var giacenzaData;
  try
  {
    giacenzaData = IMDBCommonLibs.getGiacenze(token, "MDB");
  }
  catch (e)
  {
    Logger.log('Errore chiamando IMDBCommonLibs.getGiacenze(): ' + e);
    ui.alert('Errore nel recupero delle giacenze: ' + e);
    return false;
  }

  // ATTENZIONE: ora usiamo "instock_list" (minuscolo)
  if (!giacenzaData || !giacenzaData.instock_list || !Array.isArray(giacenzaData.instock_list))
  {
    Logger.log('Risposta giacenze non valida o priva di instock_list');
    ui.alert('Errore: la risposta delle giacenze non contiene instock_list valida.');
    return false;
  }

  var ui = SpreadsheetApp.getUi();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy_MM_dd');

  const filesRighe = DriveApp.getFilesByName(today + '_Uscite_Righe');

  // Se non esiste nessun file con quel nome → esci
  if (!filesRighe.hasNext())
  {
    Logger.log('❌ File "' + today + '_Uscite_Righe" non trovato in Drive.');
    SpreadsheetApp.getUi().alert('File "' + today + '_Uscite_Righe" non trovato in Drive.');
    return false; // esce dalla funzione
  }

  const ssRighe   = SpreadsheetApp.open(filesRighe.next());

  // Trova i file per la data corrente
  
  var sheet = ssRighe.getSheetByName('preimpostato_righe_uscite');
  var lastRow = sheet.getLastRow();
  if (lastRow < 2)
  {
    ui.alert('Nessuna riga di dati trovata (serve almeno dalla riga 2 in poi).');
    return false;
  }

  // Colonne: B = 2 (codice), C = 3 (quantità)
  // Leggiamo da riga 2 a lastRow, colonne 2-3
  var numRows = lastRow - 1;
  var dataRange = sheet.getRange(2, 2, numRows, 3);
  var values = dataRange.getValues();

  // 3. Sommarizza le quantità per codice (col B/C)
  var richiestePerCodice = {};
  for (var i = 0; i < values.length; i++)
  {
    var codice = String(values[i][0] || '').trim();  // colonna B
    var qtyRaw = values[i][1];

    if (!codice)
    {
      continue;
    }

    var qty = Number(qtyRaw) || 0;
    if (!richiestePerCodice[codice])
    {
      richiestePerCodice[codice] = 0;
    }
    richiestePerCodice[codice] += qty;
  }

  // Se non ci sono codici, niente da fare
  if (Object.keys(richiestePerCodice).length === 0)
  {
    ui.alert('Nessun codice trovato in colonna B dalla riga 2 in giù.');
    return false;
  }

  // 4. Costruisci mappa giacenze per codice: { codice: { quantity, quantity_in_orders } }
  var giacenze = {};
  giacenzaData.instock_list.forEach(function(item)
  {
    var code = String(item.code || '').trim();
    if (!code)
    {
      return;
    }
    giacenze[code] =
    {
      quantity: Number(item.quantity) || 0,
      quantity_in_orders: Number(item.quantity_in_orders) || 0
    };
  });

  // 5. Verifica disponibilità per ogni codice richiesto
  var errorMessages = [];

  for (var codice in richiestePerCodice)
  {
    if (!Object.prototype.hasOwnProperty.call(richiestePerCodice, codice))
    {
      continue;
    }

    var richiesto = richiestePerCodice[codice];

    if (!giacenze[codice])
    {
      var msgNotFound = 'Codice ' + codice + ': nessuna giacenza trovata (richiesti ' + richiesto + ').';
      Logger.log(msgNotFound);
      errorMessages.push(msgNotFound);
      continue;
    }

    var qty = giacenze[codice].quantity;
    var qtyInOrders = giacenze[codice].quantity_in_orders;
    var disponibile = qty - qtyInOrders;

    if (disponibile < richiesto)
    {
      var msg =
        'Codice ' + codice +
        ': richiesti ' + richiesto +
        ', disponibili ' + disponibile +
        ' (giacenza ' + qty +
        ', già in ordini ' + qtyInOrders + ').';

      Logger.log(msg);
      errorMessages.push(msg);
    }
  }

  // 6. Mostra risultato
  if (errorMessages.length > 0)
  {
    ui.alert(
      'ATTENZIONE: giacenze insufficienti',
      errorMessages.join('\n'),
      ui.ButtonSet.OK
    );
    return false;
  }
  else
  {
    ui.alert('Giacenze sufficienti per tutti i codici trovati.');
    return true;
  }
}*/

function ritrovaGiacenzeSCOT(giacenze)
{
  var token = IMDBCommonLibs.getScotToken(false);
  if (token)
  {
    Logger.log("Received token: " + token);
  }
  else
  {
    Logger.log("Failed to retrieve token.");
    return false;
  }

  var giacenzaData;
  try
  {
    giacenzaData = IMDBCommonLibs.getGiacenze(token, "MDB");
  }
  catch (e)
  {
    Logger.log('Errore chiamando IMDBCommonLibs.getGiacenze(): ' + e);
    ui.alert('Errore nel recupero delle giacenze: ' + e);
    return false;
  }

  if (!giacenzaData || !giacenzaData.instock_list || !Array.isArray(giacenzaData.instock_list))
  {
    Logger.log('Risposta giacenze non valida o priva di instock_list');
    ui.alert('Errore: la risposta delle giacenze non contiene instock_list valida.');
    return false;
  }

  giacenzaData.instock_list.forEach(function(item)
  {
    var code = String(item.code || '').trim();

    if (!code)
    {
      return;
    }

    giacenze[code] =
    {
      quantity: Number(item.quantity) || 0,
      quantity_in_orders: Number(item.quantity_in_orders) || 0
    };
  });

  return giacenze;
}

function checkGiacenzeSpedizioni()
{
  var ui = SpreadsheetApp.getUi();

  var giacenze = null;

  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy_MM_dd');

  var filesTestate = DriveApp.getFilesByName(today + '_Uscite_Testate');
  var filesRighe = DriveApp.getFilesByName(today + '_Uscite_Righe');

  if (!filesTestate.hasNext())
  {
    Logger.log('❌ File "' + today + '_Uscite_Testate" non trovato in Drive.');
    ui.alert('File "' + today + '_Uscite_Testate" non trovato in Drive.');
    return false;
  }

  if (!filesRighe.hasNext())
  {
    Logger.log('❌ File "' + today + '_Uscite_Righe" non trovato in Drive.');
    ui.alert('File "' + today + '_Uscite_Righe" non trovato in Drive.');
    return false;
  }

  var ssTestate = SpreadsheetApp.open(filesTestate.next());
  var ssRighe = SpreadsheetApp.open(filesRighe.next());

  var sheetTestate = ssTestate.getSheetByName('preimpostato_testate_uscite');
  var sheetRighe = ssRighe.getSheetByName('preimpostato_righe_uscite');

  if (!sheetTestate)
  {
    ui.alert('Foglio "preimpostato_testate_uscite" non trovato.');
    return false;
  }

  if (!sheetRighe)
  {
    ui.alert('Foglio "preimpostato_righe_uscite" non trovato.');
    return false;
  }

  var lastRowTestate = sheetTestate.getLastRow();
  var lastColTestate = sheetTestate.getLastColumn();

  if (lastRowTestate < 2)
  {
    ui.alert('Nessuna testata trovata (serve almeno dalla riga 2 in poi).');
    return false;
  }

  var headerTestate = sheetTestate.getRange(1, 1, 1, lastColTestate).getValues()[0];

  var dataSpedizioneCol = getColumnByHeaderName_(headerTestate, 'Data Spedizione');
  if (!dataSpedizioneCol)
  {
    dataSpedizioneCol = lastColTestate + 1;
    sheetTestate.getRange(1, dataSpedizioneCol).setValue('Data Spedizione');
    Logger.log('Colonna "Data Spedizione" creata alla posizione ' + dataSpedizioneCol);
  }

  var giacenzaCheckCol = getColumnByHeaderName_(headerTestate, 'Giacenza?');
  if (!giacenzaCheckCol)
  {
    giacenzaCheckCol = lastColTestate + 1;
    sheetTestate.getRange(1, giacenzaCheckCol).setValue('Giacenza?');
    Logger.log('Colonna "Giacenza?" creata alla posizione ' + giacenzaCheckCol);
  }

  var lastRowRighe = sheetRighe.getLastRow();
  if (lastRowRighe < 2)
  {
    ui.alert('Nessuna riga trovata nel foglio righe.');
    return false;
  }

  var lastColRighe = sheetRighe.getLastColumn();
  var headerRighe = sheetRighe.getRange(1, 1, 1, lastColRighe).getValues()[0];

  var numeroSpedizioneRigheCol = 1;
  var codiceCol = 2;
  var quantitaCol = 3;

  var numeroSpedizioneRigheByHeader = getColumnByHeaderNameMultiple_(headerRighe, ['Numero spedizione', 'Numero Spedizione', 'Spedizione', 'Num spedizione']);
  var codiceByHeader = getColumnByHeaderNameMultiple_(headerRighe, ['Codice', 'Codice articolo', 'Codice Articolo', 'Articolo']);
  var quantitaByHeader = getColumnByHeaderNameMultiple_(headerRighe, ['Quantità', 'Quantita', 'Qta', 'Qtà']);

  if (numeroSpedizioneRigheByHeader)
  {
    numeroSpedizioneRigheCol = numeroSpedizioneRigheByHeader;
  }

  if (codiceByHeader)
  {
    codiceCol = codiceByHeader;
  }

  if (quantitaByHeader)
  {
    quantitaCol = quantitaByHeader;
  }

  var righeValues = sheetRighe.getRange(2, 1, lastRowRighe - 1, lastColRighe).getValues();
  var righePerSpedizione = {};

  for (var i = 0; i < righeValues.length; i++)
  {
    var numeroSpedizione = String(righeValues[i][numeroSpedizioneRigheCol - 1] || '').trim();
    var codice = String(righeValues[i][codiceCol - 1] || '').trim();
    var qty = Number(righeValues[i][quantitaCol - 1]) || 0;

    if (!numeroSpedizione || !codice)
    {
      continue;
    }

    if (!righePerSpedizione[numeroSpedizione])
    {
      righePerSpedizione[numeroSpedizione] = {};
    }

    if (!righePerSpedizione[numeroSpedizione][codice])
    {
      righePerSpedizione[numeroSpedizione][codice] = 0;
    }

    righePerSpedizione[numeroSpedizione][codice] += qty;
  }

  var testateValues = sheetTestate.getRange(2, 1, lastRowTestate - 1, Math.max(lastColTestate, giacenzaCheckCol)).getValues();

  var errorMessages = [];
  var okCount = 0;
  var koCount = 0;
  var checkedCount = 0;

  for (var r = 0; r < testateValues.length; r++)
  {
    var sheetRow = r + 2;

    var numeroSpedizioneTestata = String(testateValues[r][0] || '').trim();
    var dataSpedizioneValue = testateValues[r][dataSpedizioneCol - 1];

    if (!numeroSpedizioneTestata)
    {
      continue;
    }

    if (isValidDateValue_(dataSpedizioneValue))
    {
      Logger.log(
        'Spedizione ' +
        numeroSpedizioneTestata +
        ': saltata perché "Data Spedizione" è valorizzata.'
      );
      continue;
    }

    // Recupera Giacenze se non fatto prima
    if (!giacenze)
    {
      giacenze = {};
      giacenze = ritrovaGiacenzeSCOT(giacenze);
    }
    checkedCount++;

    var richiestePerCodice = righePerSpedizione[numeroSpedizioneTestata];
    var rowErrors = [];

    if (!richiestePerCodice)
    {
      rowErrors.push(
        'Spedizione ' +
        numeroSpedizioneTestata +
        ': nessuna riga trovata in _Uscite_Righe.'
      );
    }
    else
    {
      for (var codiceArticolo in richiestePerCodice)
      {
        if (!Object.prototype.hasOwnProperty.call(richiestePerCodice, codiceArticolo))
        {
          continue;
        }

        var richiesto = richiestePerCodice[codiceArticolo];

        if (!giacenze[codiceArticolo])
        {
          rowErrors.push(
            'Spedizione ' +
            numeroSpedizioneTestata +
            ' - Codice ' +
            codiceArticolo +
            ': nessuna giacenza trovata (richiesti ' +
            richiesto +
            ').'
          );
          continue;
        }

        var qty = giacenze[codiceArticolo].quantity;
        var qtyInOrders = giacenze[codiceArticolo].quantity_in_orders;
        var disponibile = qty - qtyInOrders;

        if (disponibile < richiesto)
        {
          rowErrors.push(
            'Spedizione ' +
            numeroSpedizioneTestata +
            ' - Codice ' +
            codiceArticolo +
            ': richiesti ' +
            richiesto +
            ', disponibili ' +
            disponibile +
            ' (giacenza ' +
            qty +
            ', già in ordini ' +
            qtyInOrders +
            ').'
          );
        }
      }
    }

    if (rowErrors.length > 0)
    {
      sheetTestate.getRange(sheetRow, giacenzaCheckCol).setValue('KO');
      koCount++;

      for (var e = 0; e < rowErrors.length; e++)
      {
        Logger.log(rowErrors[e]);
        errorMessages.push(rowErrors[e]);
      }
    }
    else
    {
      sheetTestate.getRange(sheetRow, giacenzaCheckCol).setValue('OK');
      okCount++;

      Logger.log(
        'Spedizione ' +
        numeroSpedizioneTestata +
        ': giacenza sufficiente.'
      );
    }
  }

  SpreadsheetApp.flush();

  if (checkedCount === 0)
  {
    ui.alert('Nessuna spedizione in partenza da verificare');
    return true;
  }

  if (errorMessages.length > 0)
  {
    ui.alert(
      'ATTENZIONE: giacenze insufficienti',
      'Spedizioni OK: ' + okCount + '\n' +
      'Spedizioni KO: ' + koCount + '\n\n' +
      errorMessages.join('\n'),
      ui.ButtonSet.OK
    );
    return false;
  }
  else
  {
    ui.alert(
      'Controllo giacenze completato',
      'Tutte le spedizioni in partenza hanno giacenza sufficiente.\n' +
      'Spedizioni OK: ' + okCount,
      ui.ButtonSet.OK
    );
    return true;
  }
}

function getColumnByHeaderName_(headerRow, headerName)
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

function getColumnByHeaderNameMultiple_(headerRow, possibleNames)
{
  for (var p = 0; p < possibleNames.length; p++)
  {
    var col = getColumnByHeaderName_(headerRow, possibleNames[p]);

    if (col)
    {
      return col;
    }
  }

  return 0;
}

function isValidDateValue_(value)
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

/**
 * Legge la cella attuale, estrae il codice spedizione a 10 cifre,
 * chiama le API SCOT /api/uscite/stato e mostra il risultato.
 */
function checkSpedizioneCorrente() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell  = sheet.getActiveCell();

  if (!cell) {
    ui.alert("Errore", "Nessuna cella attiva trovata.", ui.ButtonSet.OK);
    return;
  }

  // Prende il valore visualizzato per non perdere eventuali zeri iniziali
  let value = String(cell.getDisplayValue()).trim();

  // Rimuove eventuali prefissi "Spedito: " o "Preparato: "
  value = value.replace(/^Spedito:\s*/i, "");
  value = value.replace(/^Preparato:\s*/i, "");

  const orderId = value;

  // Recupera il token tramite la libreria condivisa
  const token = IMDBCommonLibs.getScotToken(false);
  if (!token) {
    ui.alert("Errore token", "Impossibile ottenere il token SCOT.", ui.ButtonSet.OK);
    return;
  }

  // Costruisce il payload secondo lo schema ordiniUsciteStato
  const payload = {
    order_id: orderId,
    client: 'MDB'
  };

  const url = IMDBCommonLibs.scotBaseURL + "/api/uscite/stato/";

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    headers: {
      "Authorization": "Bearer " + token
    },
    muteHttpExceptions: true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (e) {
    ui.alert(
      "Errore di rete",
      "Chiamata a SCOT fallita:\n" + e,
      ui.ButtonSet.OK
    );
    return;
  }

  const statusCode = response.getResponseCode();
  const bodyText   = response.getContentText() || "";

  if (statusCode !== 200) {
    ui.alert(
      "Errore API",
      "La chiamata ha restituito HTTP " + statusCode + ":\n\n" + bodyText,
      ui.ButtonSet.OK
    );
    return;
  }

  let data;
  try {
    data = JSON.parse(bodyText);
  } catch (e) {
    ui.alert(
      "Errore parsing",
      "Risposta non in JSON valido:\n" + e + "\n\nCorpo risposta:\n" + bodyText,
      ui.ButtonSet.OK
    );
    return;
  }

  // Formattazione risposta ordiniUsciteStatoResponse
  const lines = [];

  lines.push("Ordine: " + (data.order_id || orderId));
  if (typeof data.status !== "undefined") {
    lines.push("Stato: " + data.status);
  }
  if (data.acquisition_date) {
    lines.push("Data acquisizione: " + data.acquisition_date);
  }
  if (data.conclusion_date) {
    lines.push("Data conclusione: " + data.conclusion_date);
  }

  lines.push("");
  lines.push("Righe:");

  if (data.rows && data.rows.length) {
    data.rows.forEach(function (row) {
      const rId   = row.id;
      const rCode = row.code;
      const rNum  = row.row_number;
      const qReq  = row.quantity_required;
      const qProc = row.quantity_processed;

      lines.push(
        "Riga " + rNum +
        " (ID " + rId + "): " +
        rCode +
        " | richiesti: " + qReq +
        " | processati: " + qProc
      );
    });
  } else {
    lines.push("Nessuna riga trovata.");
  }

  const message = lines.join("\n");

  ui.alert(
    "Stato spedizione",
    message,
    ui.ButtonSet.OK
  );
}

/*
// =======================
// FUNZIONE PRINCIPALE: LETTURA CONTATTI SU PIÙ RIGHE
// =======================
function fillMauticCustomerDataOBSOLETE() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const range = sheet.getActiveRange();
  if (!range) {
    ui.alert('Nessuna selezione attiva.');
    return;
  }

  const numRows  = range.getNumRows();
  const startRow = range.getRow();
  const col      = range.getColumn();

  let errorMessages = [];

  for (let offset = 0; offset < numRows; offset++) {
    const rowIndex = startRow + offset;

    if (sheet.isRowHiddenByFilter(rowIndex) || sheet.isRowHiddenByUser(rowIndex)) {
      continue;
    }

    const cell = sheet.getRange(rowIndex, col);
    const rawValue = String(cell.getDisplayValue()).trim();

    if (!rawValue) continue;

    try {
      const mauticData = IMDBCommonLibs.getMauticCustomerData(rawValue);

      if (!mauticData) {
        errorMessages.push('Riga ' + rowIndex + ': nessun dato trovato per "' + rawValue + '".');
        continue;
      }

      // Se più match: chiedi quale usare
      let customerData = mauticData;
      if (Array.isArray(mauticData)) {
        const picked = pickMauticContactFromArray_(ui, mauticData, rawValue);
        if (!picked) {
          errorMessages.push('Riga ' + rowIndex + ': selezione contatto annullata per "' + rawValue + '".');
          continue;
        }
        customerData = picked;
      }

      // La risposta tipica Mautic è { contact: {...} }
      const contact = customerData.contact || customerData || {};
      const fields  = contact.fields || {};
      const all     = fields.all || {};

      // ---------------------
      // ID (colonna B)
      // ---------------------
      const id = (typeof contact.id !== 'undefined' && contact.id !== null)
        ? contact.id
        : (typeof all.id !== 'undefined' ? all.id : '');

      // ---------------------
      // Last name (colonna D)
      // ---------------------
      const lastName = IMDBCommonLibs.getMauticFieldNormalized(contact, 'lastname') || '';

      // ---------------------
      // First name (colonna E)
      // ---------------------
      const firstName = IMDBCommonLibs.getMauticFieldNormalized(contact, 'firstname') || '';

      // ---------------------
      // Codice fiscale (colonna F)
      // ---------------------
      const codiceFiscale = IMDBCommonLibs.getMauticFieldNormalized(contact, 'codice_fiscale') || '';

      // ---------------------
      // Email (colonna H)
      // ---------------------
      const email = IMDBCommonLibs.getMauticFieldNormalized(contact, 'email') || '';

      // ---------------------
      // Telefono (colonna I) -> preferisci mobile, poi phone
      // ---------------------
      const phone =
        IMDBCommonLibs.getMauticFieldNormalized(contact, 'mobile') ||
        IMDBCommonLibs.getMauticFieldNormalized(contact, 'phone')  ||
        '';

      // ---------------------
      // Campagna acquisizione (Breve) (colonna K)
      // ---------------------
      const campagnaBreve = IMDBCommonLibs.getMauticFieldNormalized(contact, 'campagna_breve') || '';

      // ---------------------
      // IMDB VIP Club (colonna L)
      // ---------------------
      const imdbVipClub = IMDBCommonLibs.getMauticFieldNormalized(contact, 'imdb_vip_club') || '';

      // =====================
      // SCRITTURA SUL FOGLIO
      // =====================
      sheet.getRange(rowIndex, 2).setValue(id);
      sheet.getRange(rowIndex, 4).setValue(lastName);
      sheet.getRange(rowIndex, 5).setValue(firstName);
      sheet.getRange(rowIndex, 6).setValue(codiceFiscale);
      sheet.getRange(rowIndex, 8).setValue(email);

      if (phone) {
        sheet.getRange(rowIndex, 9).setValue("'" + phone);
      } else {
        sheet.getRange(rowIndex, 9).clearContent();
      }

      sheet.getRange(rowIndex, 11).setValue(campagnaBreve);
      sheet.getRange(rowIndex, 12).setValue(imdbVipClub);

    } catch (e) {
      // Se è errore di autorizzazione, meglio fermare subito (evita spam di errori su ogni riga)
      if (String(e && e.message || e).indexOf('Mautic non autorizzato') !== -1) {
        ui.alert('Autorizzazione Mautic necessaria', e.message, ui.ButtonSet.OK);
        return;
      }

      Logger.log('Errore Mautic su riga ' + rowIndex + ': ' + e);
      errorMessages.push('Riga ' + rowIndex + ': ' + (e && e.message ? e.message : e));
    }
  }

  if (errorMessages.length) {
    ui.alert('Completato con alcuni errori', errorMessages.join('\n'), ui.ButtonSet.OK);
  }
}
*/

/**
 * Helper UI: dato un array di customerData (risposta /contacts/{id}),
 * fa scegliere all'utente quale contatto usare.
 *
 * Ritorna l'elemento scelto (customerData) oppure null se annullato.
 */
function pickMauticContactFromArray_(ui, customerDataArray, keyLabel) {
  const n = customerDataArray.length;

  let msg = 'Sono stati trovati ' + n + ' contatti per "' + keyLabel + '":\n\n';

  for (let i = 0; i < n; i++) {
    const cd = customerDataArray[i];
    const contact = (cd && cd.contact) ? cd.contact : (cd || {});

    const firstName = IMDBCommonLibs.getMauticFieldNormalized(contact, 'firstname') || '';
    const lastName  = IMDBCommonLibs.getMauticFieldNormalized(contact, 'lastname')  || '';
    const email     = IMDBCommonLibs.getMauticFieldNormalized(contact, 'email')     || '';

    const fields = contact.fields || {};
    const all    = fields.all || {};
    const id     = (typeof contact.id !== 'undefined' && contact.id !== null)
      ? contact.id
      : (typeof all.id !== 'undefined' ? all.id : '');

    msg += (i + 1) + ') ID ' + id + ' - ' + (firstName + ' ' + lastName).trim() +
           (email ? ' <' + email + '>' : '') + '\n';
  }

  msg += '\nInserisci il numero del contatto da usare (1-' + n + '), oppure 0 per annullare:';

  const promptResult = ui.prompt('Seleziona contatto Mautic', msg, ui.ButtonSet.OK_CANCEL);
  const button = promptResult.getSelectedButton();
  if (button !== ui.Button.OK) return null;

  const choiceStr = String(promptResult.getResponseText()).trim();
  const choice = parseInt(choiceStr, 10);

  if (isNaN(choice) || choice < 0 || choice > n) {
    throw new Error('Scelta non valida: "' + choiceStr + '". Operazione annullata.');
  }
  if (choice === 0) return null;

  return customerDataArray[choice - 1];
}


/**
 * SCRIPT 1: CALCOLO COSTI IMPORT FRANCIA
 * Calcola il costo logistico progressivo per l'importazione.
 */

function calcoloImportFrancia(numBottiglie, markupOccultoPerBtg)
{
  // Pulizia input
  var btg = parseFloat(numBottiglie);
  var mkt = parseFloat(markupOccultoPerBtg);

  if (isNaN(btg) || isNaN(mkt) || btg <= 0) {
    return "⚠️ Inserire Q.tà e Markup"; 
  }

  const SCAGLIONI = [
    { limite: 120, tariffa: 1.20 }, { limite: 150, tariffa: 1.10 },
    { limite: 180, tariffa: 0.95 }, { limite: 239, tariffa: 0.81 },
    { limite: 300, tariffa: 0.75 }, { limite: 480, tariffa: 0.68 },
    { limite: 732, tariffa: 0.61 }, { limite: 1092, tariffa: 0.56 },
    { limite: 1452, tariffa: 0.52 }, { limite: 1790, tariffa: 0.49 }
  ];

  let costoTrasportoTotal = 0;
  let btgRimanenti = btg;
  let limitePrecedente = 0;

  for (let i = 0; i < SCAGLIONI.length; i++) {
    let rangeCapienza = SCAGLIONI[i].limite - limitePrecedente;
    let btgInQuestoRange = Math.min(btgRimanenti, rangeCapienza);
    if (btgInQuestoRange > 0) {
      costoTrasportoTotal += btgInQuestoRange * SCAGLIONI[i].tariffa;
      btgRimanenti -= btgInQuestoRange;
      limitePrecedente = SCAGLIONI[i].limite;
    }
    if (btgRimanenti <= 0) break;
  }
  
  if (btgRimanenti > 0) costoTrasportoTotal += btgRimanenti * 0.49;

  const COSTO_BUROCRAZIA_FLAT = 40.00;
  let costoTotaleReale = costoTrasportoTotal + COSTO_BUROCRAZIA_FLAT;
  let incassoDaMarkup = btg * mkt;
  let margineLogistico = incassoDaMarkup - costoTotaleReale;

  return [
    ["REPORT IMPORT FRANCIA", "RISULTATO"],
    ["N. Bottiglie in Input", btg],
    ["Costo Trasporto Netto", formattatoreItaliano(costoTrasportoTotal)],
    ["Costo Burocrazia (Flat)", formattatoreItaliano(COSTO_BUROCRAZIA_FLAT)],
    ["COSTO TOTALE IMPORT", costoTotaleReale], // Valore raw per macro
    ["Incidenza/Bottiglia", formattatoreItaliano(costoTotaleReale / btg)],
    ["MARGINE LOGISTICO", formattatoreItaliano(margineLogistico)],
    ["STATUS", margineLogistico >= 0 ? "✅ OK" : "❌ PERDITA"]
  ];
}

/**
 * FORMATTATORE MANUALE CORRETTO
 */
function formattatoreItaliano(valore) {
  // Forza la conversione a numero per evitare isNaN su stringhe numeriche
  var n = Number(valore);
  
  if (isNaN(n)) return "0,00 €";
  
  // Formattazione manuale: 10.000,24 €
  var str = n.toFixed(2);
  var parti = str.split(".");
  parti[0] = parti[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  
  return parti[0] + "," + parti[1] + " €";
}

/**
 * MACRO OPERATIVA AGGIORNATA
 */
function aggiornaCostiImport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var currentRow = 10; // sheet.getActiveCell().getRow();
  
  // Lettura colonna D (4)
  var cellVal = sheet.getRange(currentRow, 4).getValue();
  var numBottiglie = parseFloat(cellVal);

  var cellVal = sheet.getRange(currentRow, 8).getValue();
  var markupFisso = parseFloat(cellVal);

  if (!isNaN(numBottiglie) && numBottiglie > 0) {
    
    // Calcolo
    var risultatoTable = calcoloImportFrancia(numBottiglie, markupFisso);

    Logger.log(risultatoTable);
    
    // Verifica che risultato sia un array e non la stringa di errore
    if (Array.isArray(risultatoTable)) {
      var costoNumerico = risultatoTable[4][1];
      
      // Formattazione
      var costoTesto = formattatoreItaliano(costoNumerico);
      
      // Scrittura in G (7)
      sheet.getRange(currentRow, 7).setValue(costoTesto);
      SpreadsheetApp.getActiveSpreadsheet().toast("Costo scritto: " + costoTesto);
    } else {
      SpreadsheetApp.getUi().alert(risultatoTable); // Mostra l'errore "⚠️ Inserire Q.tà..."
    }
    
  } else {
    SpreadsheetApp.getUi().alert("Errore: La cella in colonna D deve contenere un numero.");
  }
}

/**
 * PROFIT ENGINE V4.0 - MAPPING DINAMICO E CICLO SU SELEZIONE
 * Usa le funzioni originali ANALISI_ORDINE_SINGOLO e calcoloImportFrancia senza toccarle.
 */
function calcolaLogisticaDaSelezione() 
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  
  // 1. MAPPING DELLE COLONNE (Cerca le intestazioni alla riga 33)
  const headerRow = 33;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  // Sistema i costi import
  
  aggiornaCostiImport();
  
  const colMap = {
    numBtgTotali: findCol(headers, "Bottiglie"),
    numBtgFrancesi: findCol(headers, "Bottiglie"),
    isPallet: findCol(headers, "Pallet?"),
    regione: findCol(headers, "Regione"),
    isMistoCartoneOut: findCol(headers, "Misto?"),
    isZonaDisagiata: findCol(headers, "Disagiata?"),
    isPreavviso: findCol(headers, "Appuntamento Telefonico?"),
    isUrgenza: findCol(headers, "Priority?"),
    incassoSpedEsplicita: findCol(headers, "Spedizione"),
    myCDV: findCol(headers, "CDV"),
    myCDV_Puro: findCol(headers, "CDV Puro"),
    pesoTotale : findCol (headers, "Peso totale"),
    outputCol: findCol (headers, "Spedizione italiana"),
    logisticaCol : findCol (headers, "Logistica OUT"),
  };

  // Validazione Mapping: se una colonna non esiste, interrompe
  for (let key in colMap) {
    if (colMap[key] === -1) {
      SpreadsheetApp.getUi().alert("Errore: Non trovo la colonna '" + key + "' alla riga 33.");
      return;
    }
  }

  pesoTotale = 0;

  // 2. CICLO SULLE RIGHE SELEZIONATE
  for (let i = 0; i < numRows; i++) {
    let currentRow = startRow + i;
    
    // Salta l'esecuzione se siamo sopra o sulla riga delle intestazioni
    if (currentRow <= headerRow) continue; 

    // Recupero dati dalle celle
    let numBtgTotali = parseNum(sheet.getRange(currentRow, colMap.numBtgTotali).getValue());
    let numBtgFrancesi = parseNum(sheet.getRange(currentRow, colMap.numBtgFrancesi).getValue());
    
    let isPallet = parseBool(sheet.getRange(currentRow, colMap.isPallet).getValue());
    let regione = sheet.getRange(currentRow, colMap.regione).getValue().toString();
    let isMisto = parseBool(sheet.getRange(currentRow, colMap.isMistoCartoneOut).getValue());
    let isDisagiata = parseBool(sheet.getRange(currentRow, colMap.isZonaDisagiata).getValue());
    let isPreavviso = parseBool(sheet.getRange(currentRow, colMap.isPreavviso).getValue());
    let isUrgenza = parseBool(sheet.getRange(currentRow, colMap.isUrgenza).getValue());
    let pesoTotale = sheet.getRange(currentRow, colMap.pesoTotale).getValue();
    
    let incassoEsp = parseNum(sheet.getRange(currentRow, colMap.incassoSpedEsplicita).getValue());
    let incassoOcc = parseNum(sheet.getRange(currentRow, colMap.myCDV).getValue()) - parseNum(sheet.getRange(currentRow, colMap.myCDV_Puro).getValue());

    // 3. CHIAMATA ALLE FUNZIONI ORIGINALI (NON TOCCATE)
    
    // Calcolo Logistica Italia tramite ANALISI_ORDINE_SINGOLO
    // Nota: estraiamo il valore numerico del costo reale dalla tabella restituita (riga 5, colonna 2)
    let tabellaItalia = spedizioneOrdineSingolo_(
      numBtgTotali, 
      numBtgFrancesi, 
      isPallet, 
      regione, 
      isMisto, 
      isDisagiata, 
      isPreavviso, 
      isUrgenza, 
      incassoEsp, 
      incassoOcc,
      pesoTotale
    );

    Logger.log(tabellaItalia);
    
    // Assumiamo che ANALISI_ORDINE_SINGOLO restituisca il costo reale in formato numerico alla riga 5 (index 4)
    let costoSpedizione = tabellaItalia[4][1]; 
    let costoLogisticaOUT = tabellaItalia[3][1];

/*    // Calcolo Import Francia tramite calcoloImportFrancia
    // Nota: markupOccultoPerBtg qui non serve per il costo, passiamo 1 come placeholder
    let tabellaImport = calcoloImportFrancia(numBtgFrancesi, 1);
    let costoImportSoloFR = tabellaImport[4][1];
*/
    sheet.getRange(currentRow, colMap.outputCol).setValue(formattatoreItaliano(parseFloat(costoSpedizione)));
    sheet.getRange(currentRow, colMap.logisticaCol).setValue(formattatoreItaliano(parseFloat(costoLogisticaOUT)));
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Costi calcolati su selezione.", "OPERATIONS");
}

/* --- FUNZIONI DI SUPPORTO PER IL MAPPING --- */

function findCol(headers, stringa) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().toLowerCase().trim() === stringa.toLowerCase().trim()) return i + 1;
    // Fallback se la stringa è contenuta
    if (headers[i].toString().toLowerCase().includes(stringa.toLowerCase())) return i + 1;
  }
  return -1;
}

function parseNum(val) {
  if (typeof val === 'string') {
    val = val.replace('.', '').replace(',', '.').replace(/[^\d.-]/g, '');
  }
  let n = parseFloat(val);
  return isNaN(n) ? 0 : n;
}

function parseBool(val) {
  if (!val) return false;
  let s = val.toString().toLowerCase().trim();
  return (s === "si" || s === "sì" || s === "true" || s === "x" || s === "1");
}

/**
 * SCRIPT 2: ANALISI SINGOLA SPEDIZIONE (Last Mile + Quota Import + Pallet Regionale)
 */
function spedizioneOrdineSingolo_(
  numBtgTotali, 
  numBtgFrancesi, 
  isPallet, 
  regione, // Da inserire come testo (es. "PIEMONTE", "SARDEGNA")
  isMistoCartoneOut, 
  isZonaDisagiata, 
  isPreavviso, 
  isUrgenza, 
  incassoSpedEsplicita, 
  incassoOccultoTotale, 
  pesoTotale
) {
  
  // 1. COSTANTI LOGISTICA ITALIA
  const FEE_FISSA = 1.50; // DDT + Minimo Ordine (prudenziale)
  const PICK_BTG = 0.55;
  const PICK_CARTONE = 0.50;
  const NAK_6 = 3.57;
  const PALLET_LEGNO = 4.20;
  const FUEL = 1.09; // +9%
  
  // 2. TABELLA TARIFFE PALLET (Fino a 50kg, 51-100kg, +ogni 100kg dopo i 100)
  const TARIFFE_PALLET = {
    "PIEMONTE": [13.92, 12.58, 11.95],
    "VALLE D'AOSTA": [17.75, 16.58, 14.67],
    "LOMBARDIA": [16.73, 15.59, 13.97],
    "TRENTINO": [17.93, 16.64, 15.92],
    "VENETO": [17.51, 16.90, 15.98],
    "FRIULI": [19.53, 19.10, 18.65],
    "LIGURIA": [19.83, 19.10, 18.52],
    "EMILIA ROMAGNA": [19.92, 18.74, 17.89],
    "TOSCANA": [21.92, 21.58, 20.59],
    "UMBRIA": [22.98, 21.69, 20.95],
    "MARCHE": [22.87, 21.93, 20.76],
    "LAZIO": [22.93, 22.57, 21.98],
    "ABRUZZO": [22.98, 21.87, 20.96],
    "MOLISE": [21.84, 21.87, 20.61],
    "CAMPANIA": [22.97, 21.81, 19.49],
    "PUGLIA": [22.97, 21.81, 19.38],
    "BASILICATA": [25.80, 24.98, 23.82],
    "CALABRIA": [30.09, 28.71, 26.72],
    "SICILIA": [34.21, 33.83, 30.20],
    "SARDEGNA": [34.21, 33.83, 30.20]
  };

  const TARIFFE_ESPRESSO = [
    { kg: 3, c: 6.82 }, { kg: 5, c: 7.88 }, { kg: 10, c: 8.23 },
    { kg: 20, c: 9.55 }, { kg: 30, c: 13.95 }, { kg: 50, c: 22.42 },
    { kg: 70, c: 26.51 }, { kg: 100, c: 29.10 }
  ];

  // 4. CALCOLO MAGAZZINO E IMBALLO
  let costoMagazzino = FEE_FISSA;
  if (isMistoCartoneOut) 
  {
    costoMagazzino += numBtgTotali * PICK_BTG;
  } else {
    costoMagazzino += Math.ceil(numBtgTotali / 6) * PICK_CARTONE;
  }

  if (isPallet) {
    costoMagazzino += PALLET_LEGNO;
  } else {
    costoMagazzino += Math.ceil(numBtgTotali / 6) * NAK_6;
  }

  let costoBaseTrasporto = 0;

  if (isPallet) {
    let t = TARIFFE_PALLET[regione.toUpperCase()] || TARIFFE_PALLET["LOMBARDIA"]; // Default se regione errata
    // Fino a 50kg
    costoBaseTrasporto = t[0];
    // Da 51 a 100kg
    if (pesoTotale > 50) {
      costoBaseTrasporto += t[1];
    }
    // Oltre i 100kg (ogni 100kg)
    if (pesoTotale > 100) {
      let eccedenza = pesoTotale - 100;
      costoBaseTrasporto += (eccedenza / 100) * t[2];
    }
  } else {
    costoBaseTrasporto = 35.00; // Default oltre 100kg espresso
    for (let f of TARIFFE_ESPRESSO) {
      if (pesoTotale <= f.kg) { costoBaseTrasporto = f.c; break; }
    }
  }
  
  if (isUrgenza) costoBaseTrasporto *= 1.30;
  let trasportoNetto = costoBaseTrasporto * FUEL;

  // 6. ACCESSORI
  if (isZonaDisagiata) trasportoNetto += 7.70;
  if (isPreavviso) trasportoNetto += 2.50;

  let costoRealeTotale = costoMagazzino + trasportoNetto;
  let incassoTotale = incassoSpedEsplicita + incassoOccultoTotale;
  let differenza = incassoTotale - costoRealeTotale;

  return [
    ["VOCE COSTO", "VALORE"],
    ["Regione Destinazione", regione.toUpperCase()],
    ["Peso Stimato (kg)", pesoTotale.toFixed(2)],
    ["Magazzino & Imballo", costoMagazzino.toFixed(2) + " €"],
    ["Trasporto (Espresso/Pallet)", trasportoNetto.toFixed(2) + " €"],
    ["COSTO REALE TOTALE", costoRealeTotale.toFixed(2) + " €"],
    ["INCASSO TOTALE", incassoTotale.toFixed(2) + " €"],
    ["MARGINE ORDINE", differenza.toFixed(2) + " €"],
    ["STATUS", differenza > 0 ? "✅ OK" : "❌ PERDITA"]
  ];
}

// =======================
// FUNZIONE PRINCIPALE AGGIORNATA
// =======================
function fillMauticCustomerData() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const range = sheet.getActiveRange();
  if (!range) {
    ui.alert('Nessuna selezione attiva.');
    return;
  }

  const numRows  = range.getNumRows();
  const startRow = range.getRow();
  const col      = range.getColumn();

  // 1. RICERCA DINAMICA COLONNA REGIONE (Riga 33)
  const headerRow = 33;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  let colRegione = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().toLowerCase().includes("regione")) {
      colRegione = i + 1;
      break;
    }
  }

  let errorMessages = [];

  for (let offset = 0; offset < numRows; offset++) {
    const rowIndex = startRow + offset;

    if (sheet.isRowHiddenByFilter(rowIndex) || sheet.isRowHiddenByUser(rowIndex)) {
      continue;
    }

    const cell = sheet.getRange(rowIndex, col);
    const rawValue = String(cell.getDisplayValue()).trim();

    if (!rawValue) continue;

    try {
      const mauticData = IMDBCommonLibs.getMauticCustomerData(rawValue);

      if (!mauticData) {
        errorMessages.push('Riga ' + rowIndex + ': nessun dato trovato per "' + rawValue + '".');
        continue;
      }

      let customerData = mauticData;
      if (Array.isArray(mauticData)) {
        const picked = pickMauticContactFromArray_(ui, mauticData, rawValue);
        if (!picked) {
          errorMessages.push('Riga ' + rowIndex + ': selezione contatto annullata per "' + rawValue + '".');
          continue;
        }
        customerData = picked;
      }

      const contact = customerData.contact || customerData || {};
      
      // --- DATI STANDARD ---
      const id = (typeof contact.id !== 'undefined' && contact.id !== null) ? contact.id : '';
      const lastName = IMDBCommonLibs.getMauticFieldNormalized(contact, 'lastname') || '';
      const firstName = IMDBCommonLibs.getMauticFieldNormalized(contact, 'firstname') || '';
      const codiceFiscale = IMDBCommonLibs.getMauticFieldNormalized(contact, 'codice_fiscale') || '';
      const email = IMDBCommonLibs.getMauticFieldNormalized(contact, 'email') || '';
      const phone = IMDBCommonLibs.getMauticFieldNormalized(contact, 'mobile') || IMDBCommonLibs.getMauticFieldNormalized(contact, 'phone') || '';
      const campagnaBreve = IMDBCommonLibs.getMauticFieldNormalized(contact, 'campagna_breve') || '';
      const imdbVipClub = IMDBCommonLibs.getMauticFieldNormalized(contact, 'imdb_vip_club') || '';

      // 2. LOGICA REGIONE (Provincia -> Acronimo -> Regione)
      const provinciaRaw = IMDBCommonLibs.getMauticFieldNormalized(contact, 'provincia') || '';
      const acronimo = IMDBCommonLibs.getProvinceAcronym(provinciaRaw);
      const regioneNome = IMDBCommonLibs.getRegioneDaAcronimo(acronimo);

      // =====================
      // SCRITTURA SUL FOGLIO
      // =====================
      sheet.getRange(rowIndex, 2).setValue(id);
      sheet.getRange(rowIndex, 4).setValue(lastName);
      sheet.getRange(rowIndex, 5).setValue(firstName);
      sheet.getRange(rowIndex, 6).setValue(codiceFiscale);
      sheet.getRange(rowIndex, 8).setValue(email);

      if (phone) {
        sheet.getRange(rowIndex, 9).setValue("'" + phone);
      } else {
        sheet.getRange(rowIndex, 9).clearContent();
      }

      sheet.getRange(rowIndex, 11).setValue(campagnaBreve);
      sheet.getRange(rowIndex, 12).setValue(imdbVipClub);

      // Scrittura dinamica Regione
      if (colRegione > 0 && regioneNome) {
        sheet.getRange(rowIndex, colRegione).setValue(regioneNome);
      }

    } catch (e) {
      if (String(e && e.message || e).indexOf('Mautic non autorizzato') !== -1) {
        ui.alert('Autorizzazione Mautic necessaria', e.message, ui.ButtonSet.OK);
        return;
      }
      Logger.log('Errore Mautic su riga ' + rowIndex + ': ' + e);
      errorMessages.push('Riga ' + rowIndex + ': ' + (e && e.message ? e.message : e));
    }
  }

  if (errorMessages.length) {
    ui.alert('Completato con alcuni errori', errorMessages.join('\n'), ui.ButtonSet.OK);
  }
}

/* Esempio d'uso */

function metaCapiExample() {
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

  const res = IMDBCommonLibs.metaCapiSend(
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

/**
 * Funzione principale: invia eventi Meta CAPI per le righe selezionate,
 * controllando la colonna "Meta CAPI" trovata in riga 33 tramite:
 *   getEmailRiepilogoColumn_(currentSheet, 'Meta CAPI', 33);
 *
 * - Se la cella "Meta CAPI" della riga è VUOTA -> invia evento e scrive data odierna dd/MM/yyyy
 * - Se NON è vuota -> mostra il valore e chiede se reinviare l’evento (YES/NO)
 *
 * Assunzioni:
 * - Colonna 2: Mautic customer id
 * - Lead ID Meta: campo Mautic "meta_id" (normalizzato in forma "l:123...")
 * - metaCapiSend(eventName, customer, value, options) esiste già nel progetto
 * - IMDBCommonLibs.getMauticCustomerData(...) esiste già nel progetto
 */
function metaCapiSendFromSelectedRows() 
{
  const ui = SpreadsheetApp.getUi();
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const range = currentSheet.getActiveRange();
  if (!range) {
    ui.alert('Nessuna selezione attiva.');
    return;
  }

  const eventName = "Purchase";
  const HEADER_ROW = 33;

  // Colonne via helper esistente
  const metaCapiCol = Number(getEmailRiepilogoColumn_(currentSheet, 'Meta CAPI', HEADER_ROW)) || 0;
  if (!metaCapiCol) {
    ui.alert(`Colonna "Meta CAPI" non trovata (header in riga ${HEADER_ROW}).\nImpossibile continuare.`);
    return;
  }

  const valueCol = Number(getEmailRiepilogoColumn_(currentSheet, 'Ricarico reale', HEADER_ROW)) || 0;

  const dataPagamentoCol = Number(getEmailRiepilogoColumn_(currentSheet, 'Stato pagamento', HEADER_ROW)) || 0;
  if (!dataPagamentoCol) {
    ui.alert(`Colonna "Stato pagamento" non trovata (header in riga ${HEADER_ROW}).\nImpossibile continuare.`);
    return;
  }

  const numRows = range.getNumRows();
  const startRow = range.getRow();

  let ok = 0, skipped = 0, failed = 0, asked = 0, resent = 0;
  const errors = [];

  for (let offset = 0; offset < numRows; offset++) {
    const rowIndex = startRow + offset;

    if (currentSheet.isRowHiddenByFilter(rowIndex) || currentSheet.isRowHiddenByUser(rowIndex)) {
      continue;
    }

    // Gate: controlla colonna Meta CAPI
    const metaCell = currentSheet.getRange(rowIndex, metaCapiCol);
    const alreadySentVal = String(metaCell.getDisplayValue()).trim();

    if (alreadySentVal) {
      asked++;
      const btn = ui.alert(
        'Meta CAPI già valorizzato',
        `Riga ${rowIndex}: "Meta CAPI" = "${alreadySentVal}".\nVuoi reinviare l'evento "${eventName}"?`,
        ui.ButtonSet.YES_NO
      );
      if (btn !== ui.Button.YES) {
        skipped++;
        continue;
      }
      resent++;
    }

    // Colonna 2: Mautic customer id
    const mauticId = String(currentSheet.getRange(rowIndex, 2).getDisplayValue()).trim();
    if (!mauticId) {
      skipped++;
      continue;
    }

    try {
      // 1) fetch Mautic
      let mauticData = IMDBCommonLibs.getMauticCustomerData(mauticId);
      if (!mauticData) {
        skipped++;
        continue;
      }
      if (Array.isArray(mauticData)) mauticData = mauticData[0];

      const contact = mauticData.contact || mauticData || {};

      var codiceFiscale = IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["codice_fiscale"]);
      var imdbVIPClub = IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["imdb_vip_club"]);

      // 2) customer
      const customer = {
        leadId: "",
        email: IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["email"]),
        phone: IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["mobile", "phone"]),
        firstName: IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["firstname", "first_name", "nome"]),
        lastName: IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["lastname", "last_name", "cognome"]),
        city: IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["city", "citta", "città"]),
        state: IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["state", "provincia", "province", "stato"]),
        zip: IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["zipcode", "zip", "cap"]),
        country: (IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["country", "paese"]) || "IT"),
        externalId: (contact.id != null ? String(contact.id) : mauticId),

        // CF-derived (se presente/valido)
        birthDate: "", // "YYYYMMDD"
        gender: ""     // "m" | "f"
      };

      // ✅ Lead ID: Mautic meta_id
      const metaIdRaw = IMDBCommonLibs.getMauticFieldNormalizedSafe(contact, ["meta_id"]);
      const metaLeadId = normalizeMetaLeadId_(metaIdRaw);
      if (metaLeadId) customer.leadId = metaLeadId;

      // ✅ CF -> birthdate + gender (se c’è colonna e valore valido)
      if (codiceFiscale) 
      {
        const cfInfo = parseCodiceFiscaleBirthSex_(codiceFiscale);
        if (cfInfo) 
        {
          customer.birthDate = cfInfo.birthDateYYYYMMDD; // es. "19840209"
          customer.gender = cfInfo.gender;               // "m"|"f"
        }
      }

      // 3) valore
      const value = getNumericValueFromCell_(currentSheet, rowIndex, valueCol);

            // ✅ event time dalla colonna Mail "Stato Pagamento"
      let eventTimeSec = getEventTimeSecFromCell_(currentSheet, rowIndex, dataPagamentoCol);

      if (!eventTimeSec)
      {
        if (imdbVIPClub > 0)
        {
          const now = new Date();
          const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");

          Logger.log(`Riga ${rowIndex}: data pagamento assente, uso data odierna (${todayStr}) perché cliente VIP`);

          eventTimeSec = Math.floor(now.getTime() / 1000);
        }
        else
        {
          skipped++;
          Logger.log(`Riga ${rowIndex}: data mancante/non valida in "Stato pagamento"`);
          errors.push(`Riga ${rowIndex}: data mancante/non valida in "Stato pagamento"`);
          continue;
        }
      }

      // ✅ options PER RIGA (fresh object)
      const optionsRow = {
        actionSource: "system_generated",
        eventSource: "crm",
        leadEventSource: "chatwoot",
        currency: "EUR",
        eventTimeSec: eventTimeSec,
        eventId: buildEventIdFromDateAndMauticId_(mauticId, 6)
      };

      // 4) invio
      const res = IMDBCommonLibs.metaCapiSend(eventName, customer, value, optionsRow);
      Logger.log(`Riga ${rowIndex} ok: ${JSON.stringify(res)}`);

      // 5) marca invio
      metaCell.setValue(formatTodayDDMMYYYY_());
      ok++;

    } catch (e) {
      const msg = (e && e.message) ? e.message : String(e);

      if (msg.indexOf('Mautic non autorizzato') !== -1) {
        ui.alert('Autorizzazione Mautic necessaria', msg, ui.ButtonSet.OK);
        return;
      }

      failed++;
      errors.push(`Riga ${rowIndex}: ${msg}`);
      Logger.log(`Errore riga ${rowIndex}: ${msg}`);
    }
  }

  const summary =
    `Meta CAPI completato\n` +
    `Inviati ok: ${ok}\n` +
    `Saltati: ${skipped}\n` +
    `Richieste reinvio: ${asked}\n` +
    `Reinviati: ${resent}\n` +
    `Errori: ${failed}` +
    (errors.length ? `\n\nDettagli (prime 10):\n- ${errors.slice(0, 10).join('\n- ')}` : '');

  ui.alert(summary);
}


/**
 * Converte la cella data in unix seconds.
 * Supporta:
 * - Date vera di Sheets
 * - stringhe tipo "dd/MM/yyyy" o "dd/MM/yyyy HH:mm"
 */
function getEventTimeSecFromCell_(sheet, rowIndex, colIndex) {
  const cell = sheet.getRange(rowIndex, colIndex);
  const v = cell.getValue();

  let d = null;

  if (v instanceof Date && !isNaN(v.getTime())) {
    d = v;
  } else {
    const s = String(cell.getDisplayValue() || '').substring(0, 10);
    if (!s) return null;

    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      const dd = Number(m[1]);
      const mm = Number(m[2]) - 1;
      const yyyy = Number(m[3].length === 2 ? ('20' + m[3]) : m[3]);
      const hh = m[4] ? Number(m[4]) : 0;
      const mi = m[5] ? Number(m[5]) : 0;
      const ss = m[6] ? Number(m[6]) : 0;
      d = new Date(yyyy, mm, dd, hh, mi, ss);
    } else {
      const ts = Date.parse(s);
      if (!isNaN(ts)) d = new Date(ts);
    }
  }

  if (!d) return null;
  return Math.floor(d.getTime() / 1000);
}


/**
 * Estrae data di nascita e sesso dal Codice Fiscale italiano.
 * Ritorna:
 *  - { birthDateYYYYMMDD: "YYYYMMDD", gender: "m"|"f" }
 * oppure null se non valido.
 *
 * Note:
 * - Supporta CF standard (non gestisce omocodia completa, ma di solito basta).
 * - Inferenza secolo: 00..(annoCorrente%100) => 2000, altrimenti 1900.
 */
function parseCodiceFiscaleBirthSex_(cf) {
  if (!cf) return null;
  const s = String(cf).trim().toUpperCase();
  if (!s) return null;

  // basic sanity
  if (!/^[A-Z0-9]{16}$/.test(s)) return null;

  const yy = s.substring(6, 8);
  const mChar = s.substring(8, 9);
  const ddRaw = s.substring(9, 11);

  const year2 = parseInt(yy, 10);
  const dayNum = parseInt(ddRaw, 10);
  if (!isFinite(year2) || !isFinite(dayNum)) return null;

  const monthMap = {
    A: 1, B: 2, C: 3, D: 4, E: 5, H: 6,
    L: 7, M: 8, P: 9, R: 10, S: 11, T: 12
  };
  const month = monthMap[mChar];
  if (!month) return null;

  let gender = "m";
  let day = dayNum;

  if (dayNum >= 41 && dayNum <= 71) {
    gender = "f";
    day = dayNum - 40;
  }

  if (day < 1 || day > 31) return null;

  const now = new Date();
  const currentYY = now.getFullYear() % 100;
  const fullYear = (year2 <= currentYY) ? (2000 + year2) : (1900 + year2);

  // valida data (evita 31/02 ecc.)
  const d = new Date(fullYear, month - 1, day);
  if (d.getFullYear() !== fullYear || (d.getMonth() + 1) !== month || d.getDate() !== day) return null;

  const birthDateYYYYMMDD =
    String(fullYear) +
    String(month).padStart(2, '0') +
    String(day).padStart(2, '0');

  return { birthDateYYYYMMDD, gender };
}


function buildEventIdFromDateAndMauticId_(mauticId, padLength) {
  if (!mauticId) return null;

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "yyyyMMdd");

  const idNum = String(mauticId).replace(/\D/g, '');
  if (!idNum) return null;

  const padded = idNum.padStart(padLength || 6, '0'); // default 10 cifre

  return today + padded;
}

/* =========================================================
 * Helpers: date dd/MM/yyyy
 * ========================================================= */
function formatTodayDDMMYYYY_() {
  const tz = Session.getScriptTimeZone(); // di solito Europe/Rome
  return Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
}


/* =========================================================
 * Helpers: Meta lead id normalization (meta_id -> "l:123...")
 * ========================================================= */

function normalizeMetaLeadId_(raw) {
  if (raw === null || raw === undefined) return "";
  let s = String(raw).trim();
  if (!s) return "";

  s = s.replace(/\s+/g, "");

  let m = s.match(/^l:(\d{10,20})$/i);
  if (m) return "l:" + m[1];

  m = s.match(/^(\d{10,20})$/);
  if (m) return "l:" + m[1];

  m = s.match(/l:(\d{10,20})/i);
  if (m) return "l:" + m[1];

  m = s.match(/(\d{10,20})/);
  if (m) return "l:" + m[1];

  return "";
}

function getNumericValueFromCell_(sheet, rowIndex, colIndex) {
  if (!colIndex) return null;

  const raw = sheet.getRange(rowIndex, colIndex).getDisplayValue();
  if (!raw) return null;

  let s = String(raw).trim();
  if (!s) return null;

  // rimuovi simboli valuta e spazi
  s = s.replace(/€/g, '').replace(/\s/g, '');

  // caso 1.234,56
  if (s.indexOf('.') !== -1 && s.indexOf(',') !== -1) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else {
    // caso 123,45
    s = s.replace(',', '.');
  }

  // lascia solo cifre, punto e meno
  s = s.replace(/[^0-9.\-]/g, '');
  if (!s) return null;

  const num = Number(s);
  if (!isFinite(num)) return null;

  return Math.round(num * 100) / 100; // 2 decimali
}

// IMDB Web App — Campagne/Clienti (single-file bundle)


// IMDB Web App — Campagne/Clienti (single-file bundle)
// Incolla TUTTO in Apps Script (solo Code.gs, nessun file HTML)

/***************
 * IMPORTANT: Compatibilità
 * - Niente nullish coalescing (??)
 * - Niente optional chaining (?.)
 ***************/

var APP_CFG = {
  headerRow: 33,
  firstClientRow: 34,
  colA: 1,
  metaRows: { code: 1, producer: 2, description: 3 }
};

var ANCHORS = {
  idCliente: "ID Cliente",
  cognome: "Cognome",
  nome: "Nome",
  vip: "VIP",
  email: "Email",
  telefono: "Telefono",
  codiceFiscale: "Codice Fiscale",
  ad: "AD",
  statoPagamento: "Stato Pagamento",
  note: "Note",

  ordineBaseFields: [
    "Prezzo",
    "Spedizione",
    "Voucher",
    "Saldo Logistica",
    "Totale",
    "Totale IMDB",
    "CDV Puro",
    "Ricarico Reale",
    "Ricarico puro"
  ],

  ordineRangeStart: "Anno nascita",
  ordineRangeEnd: "Fattura Acconto VIP",

  bottleLabelInRow1: "Codice bottiglia",
  bottleStartOffset: 2,

  formulaCopyStartHeader: "Totale",
  formulaCopyEndHeader: "Bottiglie"
};

/* =========================
 * LOG buffer (server-side)
 * ========================= */

var LOG_CFG = {
  propKey: "IMDB_WEBAPP_LOG",
  maxChars: 45000
};

function logNowIso_() {
  return new Date().toISOString();
}

function logAppend_(msg) {
  var s = String(msg == null ? "" : msg);
  try { Logger.log(s); } catch (e) {}

  var props = PropertiesService.getScriptProperties();
  var prev = String(props.getProperty(LOG_CFG.propKey) || "");
  var line = "[" + logNowIso_() + "] " + s;
  var next = prev ? (prev + "\n" + line) : line;

  if (next.length > LOG_CFG.maxChars) {
    next = next.slice(next.length - LOG_CFG.maxChars);
  }

  props.setProperty(LOG_CFG.propKey, next);
}

function logClear_() {
  PropertiesService.getScriptProperties().deleteProperty(LOG_CFG.propKey);
}

function logGet_() {
  return String(PropertiesService.getScriptProperties().getProperty(LOG_CFG.propKey) || "");
}

function logError_(context, err, extraObj) {
  var parts = [];
  parts.push("ERROR");
  if (context) parts.push("ctx=" + context);

  if (err) {
    if (err.stack) parts.push("stack=" + err.stack);
    else if (err.message) parts.push("message=" + err.message);
    else parts.push("err=" + String(err));
  }

  if (extraObj) {
    try {
      parts.push("extra=" + JSON.stringify(extraObj));
    } catch (e) {
      parts.push("extra_raw=" + String(extraObj));
    }
  }

  logAppend_(parts.join(" | "));
}

function safe_(ctx, fn) {
  try {
    var out = fn();
    if (out && typeof out === "object" && Object.prototype.hasOwnProperty.call(out, "ok")) return out;
    return { ok: true, data: out };
  } catch (e) {
    logError_(ctx, e, null);
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function syncStripeSuccessfulPayments()
{
  var CONFIG =
  {
    stripeBaseUrl: 'https://api.stripe.com/v1',
    pageLimit: 100,
    secretKeyProperty: 'STRIPE_SECRET_KEY',
    startDateProperty: 'STRIPE_START_DATE',
    lastCreatedProperty: 'STRIPE_LAST_CREATED_UTC',
    sheetNameProperty: 'STRIPE_SHEET_NAME',
    headers:
    [
      'id',
      'Created date (UTC)',
      'Amount',
      'Amount Refunded',
      'Currency',
      'Captured',
      'Converted Amount',
      'Converted Amount Refunded',
      'Converted Currency',
      'Decline Reason',
      'Description',
      'Fee',
      'Refunded date (UTC)',
      'Statement Descriptor',
      'Status',
      'Seller Message',
      'Taxes On Fee',
      'Card ID',
      'Customer ID',
      'Customer Description',
      'Customer Email'
    ]
  };

  var lock = LockService.getScriptLock();

  try
  {
    lock.waitLock(30000);

    var props = PropertiesService.getScriptProperties();

    var sheetName = String(props.getProperty(CONFIG.sheetNameProperty) || '').trim();
    if (!sheetName)
    {
      throw new Error('Manca la Script Property ' + CONFIG.sheetNameProperty);
    }


    var secretKey = String(props.getProperty(CONFIG.secretKeyProperty) || '').trim();
    if (!secretKey)
    {
      throw new Error('Manca la Script Property ' + CONFIG.secretKeyProperty);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ensureStripeSheet_(ss, sheetName, CONFIG.headers);
    var existingIds = getExistingIds_(sheet);

    var lastCreated = Number(props.getProperty(CONFIG.lastCreatedProperty) || 0);
    var maxCreatedSeen = lastCreated;
    var startingAfter = '';
    var hasMore = true;
    var imported = 0;
    var scanned = 0;

    while (hasMore)
    {
      var url = buildStripeChargesUrl_(CONFIG, props, startingAfter);
      var response = stripeGetJson_(url, secretKey);
      var charges = Array.isArray(response.data) ? response.data : [];

      if (!charges.length)
      {
        break;
      }

      var rowsToAppend = [];

      for (var i = 0; i < charges.length; i++)
      {
        var charge = charges[i];
        var created = Number(charge.created || 0);

        scanned++;

        if (created > maxCreatedSeen)
        {
          maxCreatedSeen = created;
        }

        if (String(charge.status || '') !== 'succeeded')
        {
          continue;
        }

        if (existingIds[charge.id])
        {
          continue;
        }

        rowsToAppend.push(buildStripeRow_(charge));
        existingIds[charge.id] = true;
      }

      if (rowsToAppend.length > 0)
      {
        sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, CONFIG.headers.length).setValues(rowsToAppend);
        imported += rowsToAppend.length;
      }

      hasMore = response.has_more === true;

      if (hasMore)
      {
        startingAfter = charges[charges.length - 1].id;
      }
    }

    if (maxCreatedSeen > lastCreated)
    {
      props.setProperty(CONFIG.lastCreatedProperty, String(maxCreatedSeen));
    }

    Logger.log('Stripe sync completata. Scansionati: ' + scanned + ', importati: ' + imported);
  }
  finally
  {
    lock.releaseLock();
  }
}

function resetStripeLastCreated_()
{
  PropertiesService.getScriptProperties().deleteProperty('STRIPE_LAST_CREATED_UTC');
  Logger.log('Property STRIPE_LAST_CREATED_UTC azzerata');
}

function setStripeStartDate_()
{
  PropertiesService.getScriptProperties().setProperty('STRIPE_START_DATE', '01/01/2026');
  Logger.log('Property STRIPE_START_DATE impostata');
}

function buildStripeChargesUrl_(config, props, startingAfter)
{
  var lastCreated = Number(props.getProperty(config.lastCreatedProperty) || 0);
  var startDateStr = String(props.getProperty(config.startDateProperty) || '').trim();
  var createdGte = lastCreated;

  if (!createdGte)
  {
    if (!startDateStr)
    {
      throw new Error('Script property ' + config.startDateProperty + ' mancante (formato dd/MM/yyyy)');
    }

    createdGte = parseDdMmYyyyToUnixStartOfDay_(startDateStr);
  }

  var params = [];
  params.push('limit=' + encodeURIComponent(String(config.pageLimit)));
  params.push('created[gte]=' + encodeURIComponent(String(createdGte)));
  params.push('expand[]=' + encodeURIComponent('data.balance_transaction'));
  params.push('expand[]=' + encodeURIComponent('data.customer'));

  if (startingAfter)
  {
    params.push('starting_after=' + encodeURIComponent(startingAfter));
  }

  return config.stripeBaseUrl + '/charges?' + params.join('&');
}

function stripeGetJson_(url, secretKey)
{
  var response = UrlFetchApp.fetch(url,
  {
    method: 'get',
    muteHttpExceptions: true,
    headers:
    {
      Authorization: 'Bearer ' + secretKey
    }
  });

  var code = response.getResponseCode();
  var text = response.getContentText();

  if (code < 200 || code >= 300)
  {
    throw new Error('Stripe API error HTTP ' + code + ': ' + text);
  }

  return JSON.parse(text);
}

function buildStripeRow_(charge)
{
  var balanceTx = getExpandedObject_(charge.balance_transaction);
  var customerInfo = getCustomerInfoFromCharge_(charge);

  var amount = centsToAmount_(charge.amount, charge.currency);
  var amountRefunded = centsToAmount_(charge.amount_refunded, charge.currency);
  var convertedCurrency = balanceTx ? String(balanceTx.currency || '').toUpperCase() : '';
  var convertedAmount = balanceTx ? centsToAmount_(balanceTx.amount, balanceTx.currency) : '';
  var convertedAmountRefunded = getConvertedRefundedAmount_(charge, balanceTx);
  var fee = balanceTx ? centsToAmount_(balanceTx.fee, balanceTx.currency) : '';
  var taxesOnFee = balanceTx ? getTaxesOnFee_(balanceTx) : '';
  var refundedDateUtc = getRefundedDateUtc_(charge);
  var declineReason = getDeclineReason_(charge);
  var statementDescriptor = String(charge.statement_descriptor || charge.calculated_statement_descriptor || '').trim();
  var sellerMessage = String((charge.outcome && charge.outcome.seller_message) || '').trim();
  var cardId = getCardId_(charge);

  return [
    String(charge.id || ''),
    formatUnixUtc_(charge.created),
    amount,
    amountRefunded,
    String(charge.currency || '').toUpperCase(),
    charge.captured === true ? 'TRUE' : 'FALSE',
    convertedAmount,
    convertedAmountRefunded,
    convertedCurrency,
    declineReason,
    String(charge.description || '').trim(),
    fee,
    refundedDateUtc,
    statementDescriptor,
    String(charge.status || '').trim(),
    sellerMessage,
    taxesOnFee,
    cardId,
    customerInfo.id,
    customerInfo.description,
    customerInfo.email
  ];
}

function getCustomerInfoFromCharge_(charge)
{
  var fallbackDescription = String((charge.billing_details && charge.billing_details.name) || '').trim();
  var fallbackEmail = String((charge.billing_details && charge.billing_details.email) || '').trim();

  if (!charge.customer)
  {
    return {
      id: '',
      description: fallbackDescription,
      email: fallbackEmail
    };
  }

  if (typeof charge.customer === 'object')
  {
    return {
      id: String(charge.customer.id || '').trim(),
      description: String(charge.customer.description || fallbackDescription || '').trim(),
      email: String(charge.customer.email || fallbackEmail || '').trim()
    };
  }

  return {
    id: String(charge.customer || '').trim(),
    description: fallbackDescription,
    email: fallbackEmail
  };
}

function getExpandedObject_(value)
{
  if (!value)
  {
    return null;
  }

  if (typeof value === 'object')
  {
    return value;
  }

  return null;
}

function getTaxesOnFee_(balanceTx)
{
  var feeDetails = Array.isArray(balanceTx.fee_details) ? balanceTx.fee_details : [];
  var total = 0;

  for (var i = 0; i < feeDetails.length; i++)
  {
    var item = feeDetails[i];

    if (String(item.type || '') === 'tax')
    {
      total += Number(item.amount || 0);
    }
  }

  return centsToAmount_(total, balanceTx.currency);
}

function getConvertedRefundedAmount_(charge, balanceTx)
{
  if (!balanceTx)
  {
    return '';
  }

  var amountRefunded = Number(charge.amount_refunded || 0);

  if (!amountRefunded)
  {
    return '';
  }

  var exchangeRate = Number(balanceTx.exchange_rate || 0);

  if (exchangeRate > 0)
  {
    var convertedMinor = Math.round(amountRefunded * exchangeRate);
    return centsToAmount_(convertedMinor, balanceTx.currency);
  }

  if (String(balanceTx.currency || '') === String(charge.currency || ''))
  {
    return centsToAmount_(amountRefunded, balanceTx.currency);
  }

  return '';
}

function getRefundedDateUtc_(charge)
{
  if (!charge || !charge.refunded)
  {
    return '';
  }

  var refunds = (charge.refunds && Array.isArray(charge.refunds.data)) ? charge.refunds.data : [];

  if (!refunds.length)
  {
    return '';
  }

  var latest = 0;

  for (var i = 0; i < refunds.length; i++)
  {
    var created = Number(refunds[i].created || 0);

    if (created > latest)
    {
      latest = created;
    }
  }

  return latest ? formatUnixUtc_(latest) : '';
}

function getDeclineReason_(charge)
{
  if (charge.outcome && charge.outcome.reason)
  {
    return String(charge.outcome.reason || '').trim();
  }

  if (charge.failure_code)
  {
    return String(charge.failure_code || '').trim();
  }

  return '';
}

function getCardId_(charge)
{
  if (charge.payment_method)
  {
    return String(charge.payment_method);
  }

  if (charge.source && charge.source.id)
  {
    return String(charge.source.id);
  }

  return '';
}

function centsToAmount_(minor, currency)
{
  if (minor === null || minor === '' || typeof minor === 'undefined')
  {
    return '';
  }

  var amount = Number(minor);
  var code = String(currency || '').toLowerCase();

  if (isNaN(amount))
  {
    return '';
  }

  if (isZeroDecimalCurrency_(code))
  {
    return amount;
  }

  return amount / 100;
}

function isZeroDecimalCurrency_(currency)
{
  var zeroDecimal =
  {
    bif: true,
    clp: true,
    djf: true,
    gnf: true,
    jpy: true,
    kmf: true,
    krw: true,
    mga: true,
    pyg: true,
    rwf: true,
    ugx: true,
    vnd: true,
    vuv: true,
    xaf: true,
    xof: true,
    xpf: true
  };

  return zeroDecimal[currency] === true;
}

function formatUnixUtc_(unixSec)
{
  var n = Number(unixSec || 0);

  if (!n)
  {
    return '';
  }

  return Utilities.formatDate(new Date(n * 1000), 'UTC', 'yyyy-MM-dd HH:mm:ss');
}

function parseDdMmYyyyToUnixStartOfDay_(value)
{
  var s = String(value || '').trim();
  var parts = s.split('/');

  if (parts.length !== 3)
  {
    throw new Error('Formato data non valido: ' + s + ' (atteso dd/MM/yyyy)');
  }

  var dd = Number(parts[0]);
  var mm = Number(parts[1]);
  var yyyy = Number(parts[2]);

  if (!dd || !mm || !yyyy)
  {
    throw new Error('Data non valida: ' + s);
  }

  var dt = new Date(yyyy, mm - 1, dd, 0, 0, 0, 0);

  return Math.floor(dt.getTime() / 1000);
}

function ensureStripeSheet_(ss, sheetName, headers)
{
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet)
  {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0)
  {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  else
  {
    var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var rewriteHeaders = false;

    for (var i = 0; i < headers.length; i++)
    {
      if (String(currentHeaders[i] || '') !== headers[i])
      {
        rewriteHeaders = true;
        break;
      }
    }

    if (rewriteHeaders)
    {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function getExistingIds_(sheet)
{
  var ids = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2)
  {
    return ids;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (var i = 0; i < values.length; i++)
  {
    var id = String(values[i][0] || '').trim();

    if (id)
    {
      ids[id] = true;
    }
  }

  return ids;
}

function setScriptPropertyPrompt()
{
  var ui = SpreadsheetApp.getUi();

  var respName = ui.prompt(
    'Script Property',
    'Inserisci il NOME della property',
    ui.ButtonSet.OK_CANCEL
  );

  if (respName.getSelectedButton() !== ui.Button.OK)
  {
    return;
  }

  var name = String(respName.getResponseText() || '').trim();

  if (!name)
  {
    ui.alert('Nome property vuoto');
    return;
  }

  var respValue = ui.prompt(
    'Script Property',
    'Inserisci il VALORE della property "' + name + '"',
    ui.ButtonSet.OK_CANCEL
  );

  if (respValue.getSelectedButton() !== ui.Button.OK)
  {
    return;
  }

  var value = String(respValue.getResponseText() || '');

  PropertiesService.getScriptProperties().setProperty(name, value);

  ui.alert('Property salvata:\n' + name);
}


function syncPayPalSuccessfulTransactions()
{
  var CONFIG =
  {
    oauthUrl: 'https://api-m.paypal.com/v1/oauth2/token',
    transactionsUrl: 'https://api-m.paypal.com/v1/reporting/transactions',
    clientIdProperty: 'PAYPAL_CLIENT_ID',
    secretProperty: 'PAYPAL_SECRET',
    sheetNameProperty: 'PAYPAL_SHEET_NAME',
    startDateProperty: 'PAYPAL_START_DATE',
    lastSyncProperty: 'PAYPAL_LAST_SYNC_UTC',
    accountEmailProperty: 'PAYPAL_ACCOUNT_EMAIL',
    pageSize: 500,
    windowDays: 30,
    headers:
    [
      'Unique Key',
      'Transaction ID',
      'Date',
      'Time',
      'TimeZone',
      'Name',
      'Type',
      'Status',
      'Currency',
      'Gross',
      'Fee',
      'Net',
      'From Email Address',
      'To Email Address',
      'Balance Impact'
    ]
  };

  var lock = LockService.getScriptLock();

  try
  {
    lock.waitLock(30000);

    var props = PropertiesService.getScriptProperties();
    var clientId = String(props.getProperty(CONFIG.clientIdProperty) || '').trim();
    var secret = String(props.getProperty(CONFIG.secretProperty) || '').trim();
    var sheetName = String(props.getProperty(CONFIG.sheetNameProperty) || '').trim();
    var accountEmail = String(props.getProperty(CONFIG.accountEmailProperty) || '').trim();

    if (!clientId)
    {
      throw new Error('Manca la Script Property ' + CONFIG.clientIdProperty);
    }

    if (!secret)
    {
      throw new Error('Manca la Script Property ' + CONFIG.secretProperty);
    }

    if (!sheetName)
    {
      throw new Error('Manca la Script Property ' + CONFIG.sheetNameProperty);
    }

    var accessToken = getPayPalAccessToken_(CONFIG, clientId, secret);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ensurePayPalSheet_(ss, sheetName, CONFIG.headers);
    var existingKeys = getExistingUniqueKeys_(sheet);

    var startDate = getPayPalSyncStartDate_(CONFIG, props);
    var now = new Date();
    var maxSeenIso = String(props.getProperty(CONFIG.lastSyncProperty) || '').trim();
    var imported = 0;
    var scanned = 0;

    while (startDate < now)
    {
      var endDate = addDays_(startDate, CONFIG.windowDays);

      if (endDate > now)
      {
        endDate = new Date(now.getTime());
      }

      var page = 1;
      var hasMore = true;

      while (hasMore)
      {
        var url = buildPayPalTransactionsUrl_(CONFIG, startDate, endDate, page);
        var response = paypalGetJson_(url, accessToken);
        var details = Array.isArray(response.transaction_details) ? response.transaction_details : [];
        var rowsToAppend = [];

        for (var i = 0; i < details.length; i++)
        {
          var detail = details[i];
          var row = buildPayPalRow_(detail, accountEmail);

          scanned++;

          if (!row)
          {
            continue;
          }

          var uniqueKey = row[0];
          var txIso = getPayPalTransactionIso_(detail);

          if (txIso && (!maxSeenIso || txIso > maxSeenIso))
          {
            maxSeenIso = txIso;
          }

          if (existingKeys[uniqueKey])
          {
            continue;
          }

          rowsToAppend.push(row);
          existingKeys[uniqueKey] = true;
        }

        if (rowsToAppend.length > 0)
        {
          sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, CONFIG.headers.length).setValues(rowsToAppend);
          imported += rowsToAppend.length;
        }

        hasMore = details.length === CONFIG.pageSize;
        page++;
      }

      startDate = endDate;
    }

    if (maxSeenIso)
    {
      props.setProperty(CONFIG.lastSyncProperty, maxSeenIso);
    }

    Logger.log('PayPal sync completata. Scansionati: ' + scanned + ', importati: ' + imported);
  }
  finally
  {
    lock.releaseLock();
  }
}

function resetPayPalLastSync_()
{
  PropertiesService.getScriptProperties().deleteProperty('PAYPAL_LAST_SYNC_UTC');
  Logger.log('Property PAYPAL_LAST_SYNC_UTC azzerata');
}

function getPayPalAccessToken_(config, clientId, secret)
{
  var basic = Utilities.base64Encode(clientId + ':' + secret);

  var response = UrlFetchApp.fetch(config.oauthUrl,
  {
    method: 'post',
    muteHttpExceptions: true,
    payload: 'grant_type=client_credentials',
    contentType: 'application/x-www-form-urlencoded',
    headers:
    {
      Authorization: 'Basic ' + basic
    }
  });

  var code = response.getResponseCode();
  var text = response.getContentText();

  if (code < 200 || code >= 300)
  {
    throw new Error('PayPal OAuth error HTTP ' + code + ': ' + text);
  }

  var json = JSON.parse(text);
  var token = String(json.access_token || '').trim();

  if (!token)
  {
    throw new Error('Access token PayPal mancante nella risposta OAuth');
  }

  return token;
}

function getPayPalSyncStartDate_(config, props)
{
  var lastSyncIso = String(props.getProperty(config.lastSyncProperty) || '').trim();

  if (lastSyncIso)
  {
    return new Date(lastSyncIso);
  }

  var startDateStr = String(props.getProperty(config.startDateProperty) || '').trim();

  if (!startDateStr)
  {
    throw new Error('Manca la Script Property ' + config.startDateProperty + ' (formato dd/MM/yyyy)');
  }

  return parseDdMmYyyyToDate_(startDateStr);
}

function buildPayPalTransactionsUrl_(config, startDate, endDate, page)
{
  var params = [];
  params.push('start_date=' + encodeURIComponent(toPayPalIsoUtc_(startDate)));
  params.push('end_date=' + encodeURIComponent(toPayPalIsoUtc_(endDate)));
  params.push('fields=' + encodeURIComponent('transaction_info,payer_info'));
  params.push('balance_affecting_records_only=Y');
  params.push('transaction_status=S');
  params.push('page_size=' + encodeURIComponent(String(config.pageSize)));
  params.push('page=' + encodeURIComponent(String(page)));

  return config.transactionsUrl + '?' + params.join('&');
}

function paypalGetJson_(url, accessToken)
{
  var response = UrlFetchApp.fetch(url,
  {
    method: 'get',
    muteHttpExceptions: true,
    headers:
    {
      Authorization: 'Bearer ' + accessToken
    }
  });

  var code = response.getResponseCode();
  var text = response.getContentText();

  if (code < 200 || code >= 300)
  {
    throw new Error('PayPal API error HTTP ' + code + ': ' + text);
  }

  return JSON.parse(text);
}

function buildPayPalRow_(detail, accountEmail)
{
  var info = detail && detail.transaction_info ? detail.transaction_info : null;

  if (!info)
  {
    return null;
  }

  var transactionId = String(info.transaction_id || '').trim();

  if (!transactionId)
  {
    return null;
  }

  var iso = String(info.transaction_initiation_date || '').trim();

  if (!iso)
  {
    return null;
  }

  var parsed = splitIsoForSheet_(iso);
  var amount = getMoneyValue_(info.transaction_amount);
  var fee = getMoneyValue_(info.fee_amount);
  var net = '';

  if (amount !== '' && fee !== '')
  {
    net = roundMoney_(Number(amount) + Number(fee));
  }

  var currency = getMoneyCurrency_(info.transaction_amount);
  var payerInfo = detail.payer_info || {};
  var fromEmail = String(payerInfo.email_address || '').trim();
  var toEmail = String(accountEmail || '').trim();
  var balanceImpact = 'Y';
  var name = getPayPalName_(detail);
  var typeValue = String(info.transaction_event_code || info.transaction_subject || '').trim();
  var statusValue = mapPayPalStatus_(String(info.transaction_status || '').trim());

  var uniqueKey = [
    transactionId,
    iso,
    String(amount),
    String(fee),
    statusValue
  ].join('|');

  return [
    uniqueKey,
    transactionId,
    parsed.date,
    parsed.time,
    parsed.timeZone,
    name,
    typeValue,
    statusValue,
    currency,
    amount,
    fee,
    net,
    fromEmail,
    toEmail,
    balanceImpact
  ];
}

function getPayPalTransactionIso_(detail)
{
  var info = detail && detail.transaction_info ? detail.transaction_info : null;
  return info ? String(info.transaction_initiation_date || '').trim() : '';
}

function getPayPalName_(detail)
{
  var payerInfo = detail && detail.payer_info ? detail.payer_info : {};
  var payerName = payerInfo.payer_name || {};

  var alt = String(payerName.alternate_full_name || '').trim();

  if (alt)
  {
    return alt;
  }

  var given = String(payerName.given_name || '').trim();
  var surname = String(payerName.surname || '').trim();
  var full = (given + ' ' + surname).replace(/\s+/g, ' ').trim();

  if (full)
  {
    return full;
  }

  var email = String(payerInfo.email_address || '').trim();

  if (email)
  {
    return email;
  }

  var info = detail && detail.transaction_info ? detail.transaction_info : {};
  return String(info.transaction_subject || '').trim();
}

function mapPayPalStatus_(statusCode)
{
  if (statusCode === 'S')
  {
    return 'Success';
  }

  if (statusCode === 'P')
  {
    return 'Pending';
  }

  if (statusCode === 'V')
  {
    return 'Reversed';
  }

  if (statusCode === 'D')
  {
    return 'Denied';
  }

  return statusCode;
}

function getMoneyValue_(money)
{
  if (!money || typeof money !== 'object')
  {
    return '';
  }

  var value = String(money.value || '').trim();

  if (value === '')
  {
    return '';
  }

  var n = Number(value);

  if (isNaN(n))
  {
    return value;
  }

  return roundMoney_(n);
}

function getMoneyCurrency_(money)
{
  if (!money || typeof money !== 'object')
  {
    return '';
  }

  return String(money.currency_code || '').trim();
}

function roundMoney_(n)
{
  return Math.round(n * 100) / 100;
}

function splitIsoForSheet_(iso)
{
  var d = new Date(iso);

  if (String(iso).slice(-1) === 'Z')
  {
    return {
      date: Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd'),
      time: Utilities.formatDate(d, 'UTC', 'HH:mm:ss'),
      timeZone: 'UTC'
    };
  }

  return {
    date: Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    time: Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm:ss'),
    timeZone: Session.getScriptTimeZone()
  };
}

function toPayPalIsoUtc_(date)
{
  return Utilities.formatDate(date, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
}

function addDays_(date, days)
{
  return new Date(date.getTime() + (days * 24 * 60 * 60 * 1000));
}

function parseDdMmYyyyToDate_(value)
{
  var s = String(value || '').trim();
  var parts = s.split('/');

  if (parts.length !== 3)
  {
    throw new Error('Formato data non valido: ' + s + ' (atteso dd/MM/yyyy)');
  }

  var dd = Number(parts[0]);
  var mm = Number(parts[1]);
  var yyyy = Number(parts[2]);

  if (!dd || !mm || !yyyy)
  {
    throw new Error('Data non valida: ' + s);
  }

  return new Date(yyyy, mm - 1, dd, 0, 0, 0, 0);
}

function ensurePayPalSheet_(ss, sheetName, headers)
{
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet)
  {
    sheet = ss.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0)
  {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  else
  {
    var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var rewriteHeaders = false;

    for (var i = 0; i < headers.length; i++)
    {
      if (String(currentHeaders[i] || '') !== headers[i])
      {
        rewriteHeaders = true;
        break;
      }
    }

    if (rewriteHeaders)
    {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }

  return sheet;
}

function getExistingUniqueKeys_(sheet)
{
  var keys = {};
  var lastRow = sheet.getLastRow();

  if (lastRow < 2)
  {
    return keys;
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (var i = 0; i < values.length; i++)
  {
    var key = String(values[i][0] || '').trim();

    if (key)
    {
      keys[key] = true;
    }
  }

  return keys;
}

function verificaPagamentiOrdiniSelezionati()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetOrdini = ss.getActiveSheet();
  var headerRowOrdini = 33;
  var ui = SpreadsheetApp.getUi();

  var headersOrdini = getHeaderMap_(sheetOrdini, headerRowOrdini);

  var requiredOrderHeaders = [
    "Mail",
    "Cognome",
    "Nome",
    "Note",
    "Stato pagamento",
    "Fattura",
    "Totale"
  ];

  ensureHeadersExist_(sheetOrdini, headersOrdini, requiredOrderHeaders, "foglio ordini");

  var righeSelezionate = getSelectedVisibleRows_(sheetOrdini, headerRowOrdini);

  if (righeSelezionate.length === 0)
  {
    var msgNoRows = "Nessuna riga ordine selezionata e visibile da elaborare.";
    Logger.log(msgNoRows);
    ui.alert(msgNoRows);
    return;
  }

  var bancaSheet = ss.getSheetByName("Banca 2026");
  var stripeSheet = ss.getSheetByName("Stripe 2026");
  var paypalSheet = ss.getSheetByName("PayPal 2026");

  if (!bancaSheet || !stripeSheet || !paypalSheet)
  {
    throw new Error("Uno o più fogli movimenti non esistono: Banca 2026, Stripe 2026, PayPal 2026.");
  }

  var bancaHeaders = getOrCreateHeaderMap_(
    bancaSheet,
    1,
    ["Descrizione", "Importo", "Data Val.", "Fattura", "Spedizione", "Campagna"]
  );

  var stripeHeaders = getOrCreateHeaderMap_(
    stripeSheet,
    1,
    ["Customer Email", "Amount", "Created date (UTC)", "Fattura", "Spedizione", "Campagna"]
  );

  var paypalHeaders = getOrCreateHeaderMap_(
    paypalSheet,
    1,
    ["Customer From Email Address", "Gross", "Date", "Fattura", "Spedizione", "Campagna"]
  );

  var risultati = {
    esatti: [],
    parziali: [],
    nonTrovati: [],
    errori: [],
    ambigui: [],
    saltati: []
  };

  var nomeCampagna = sheetOrdini.getName();

  righeSelezionate.forEach(
    function(row)
    {
      var email = normalizeString_(sheetOrdini.getRange(row, headersOrdini["Mail"]).getDisplayValue());
      var cognome = normalizeString_(sheetOrdini.getRange(row, headersOrdini["Cognome"]).getDisplayValue());
      var nome = normalizeString_(sheetOrdini.getRange(row, headersOrdini["Nome"]).getDisplayValue());
      var note = normalizeString_(sheetOrdini.getRange(row, headersOrdini["Note"]).getDisplayValue());
      var statoPagamento = normalizeString_(sheetOrdini.getRange(row, headersOrdini["Stato pagamento"]).getDisplayValue());
      var fatturaOrdine = normalizeString_(sheetOrdini.getRange(row, headersOrdini["Fattura"]).getDisplayValue());
      var totaleDisplay = sheetOrdini.getRange(row, headersOrdini["Totale"]).getDisplayValue();
      var totaleOrdine = normalizeAmount_(sheetOrdini.getRange(row, headersOrdini["Totale"]).getValue());

      var spedizioneMatch = note.match(/(?:Preparato|Spedito):\s*([A-Za-z0-9._\-\/]+)/i);

      if (!spedizioneMatch)
      {
        risultati.saltati.push({
          row: row,
          motivo: "Note senza 'Preparato:' o 'Spedito:' con codice ordine",
          dettaglio: buildOrderDetail_(null, cognome, nome, email, totaleDisplay)
        });
        return;
      }

      var codiceSpedizione = spedizioneMatch[1];

      if (!statoPagamento)
      {
        risultati.saltati.push({
          row: row,
          motivo: "Stato pagamento vuoto",
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay)
        });
        return;
      }

      if (totaleOrdine === null)
      {
        risultati.errori.push({
          row: row,
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay),
          motivo: "Totale ordine non valido"
        });
        return;
      }

      var paymentInfo = parseStatoPagamento_(statoPagamento);

      if (!paymentInfo.data)
      {
        risultati.errori.push({
          row: row,
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay),
          motivo: "Data non riconosciuta in Stato pagamento"
        });
        return;
      }

      var esitoRicerca = null;

      if (paymentInfo.provider === "BANCA")
      {
        esitoRicerca = findPaymentMatchInBanca_(
          bancaSheet,
          bancaHeaders,
          totaleOrdine,
          email,
          cognome,
          paymentInfo.data,
          fatturaOrdine
        );
      }
      else if (paymentInfo.provider === "STRIPE")
      {
        esitoRicerca = findPaymentMatchInStripe_(
          stripeSheet,
          stripeHeaders,
          totaleOrdine,
          email,
          cognome,
          paymentInfo.data,
          fatturaOrdine
        );
      }
      else if (paymentInfo.provider === "PAYPAL")
      {
        esitoRicerca = findPaymentMatchInPaypal_(
          paypalSheet,
          paypalHeaders,
          totaleOrdine,
          email,
          cognome,
          paymentInfo.data,
          fatturaOrdine
        );
      }
      else
      {
        risultati.errori.push({
          row: row,
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay),
          motivo: "Provider pagamento non riconosciuto"
        });
        return;
      }

      if (esitoRicerca.ambigui && esitoRicerca.ambigui.length > 1)
      {
        risultati.ambigui.push({
          row: row,
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay),
          foglio: esitoRicerca.sheet.getName(),
          righe: esitoRicerca.ambigui
        });
      }

      if (!esitoRicerca.match)
      {
        risultati.nonTrovati.push({
          row: row,
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay),
          motivo: "Nessun movimento compatibile trovato"
        });
        return;
      }

      if (esitoRicerca.fatturaConflict)
      {
        risultati.errori.push({
          row: row,
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay),
          motivo: "Fattura già presente nel movimento e diversa da quella ordine",
          foglio: esitoRicerca.sheet.getName(),
          movimentoRow: esitoRicerca.match.row,
          fatturaOrdine: fatturaOrdine,
          fatturaMovimento: esitoRicerca.existingFattura
        });
        return;
      }

      writeMovimentoResult_(
        esitoRicerca.sheet,
        esitoRicerca.headers,
        esitoRicerca.match.row,
        fatturaOrdine,
        codiceSpedizione,
        nomeCampagna,
        esitoRicerca.match.tipoMatch === "PARZIALE"
      );

      if (esitoRicerca.match.tipoMatch === "ESATTO")
      {
        risultati.esatti.push({
          row: row,
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay),
          foglio: esitoRicerca.sheet.getName(),
          movimentoRow: esitoRicerca.match.row
        });
      }
      else
      {
        risultati.parziali.push({
          row: row,
          dettaglio: buildOrderDetail_(codiceSpedizione, cognome, nome, email, totaleDisplay),
          foglio: esitoRicerca.sheet.getName(),
          movimentoRow: esitoRicerca.match.row,
          motivo: esitoRicerca.match.partialReason || "Corrispondenza non piena",
          dataOrdine: paymentInfo.data || "",
          dataMovimento: esitoRicerca.match.dataMovimento || ""
        });
      }
    }
  );

  var report = buildFinalReport_(sheetOrdini.getName(), risultati);

  Logger.log(report);
  ui.alert(report);
}

function getSelectedVisibleRows_(sheet, headerRow)
{
  var activeRangeList = sheet.getActiveRangeList();
  var ranges = [];

  if (activeRangeList)
  {
    ranges = activeRangeList.getRanges();
  }
  else
  {
    var activeRange = sheet.getActiveRange();

    if (activeRange)
    {
      ranges = [activeRange];
    }
  }

  var rowsMap = {};

  ranges.forEach(
    function(range)
    {
      var startRow = range.getRow();
      var endRow = startRow + range.getNumRows() - 1;

      for (var r = startRow; r <= endRow; r++)
      {
        if (r <= headerRow)
        {
          continue;
        }

        if (sheet.isRowHiddenByFilter(r))
        {
          continue;
        }

        if (sheet.isRowHiddenByUser(r))
        {
          continue;
        }

        rowsMap[r] = true;
      }
    }
  );

  return Object.keys(rowsMap)
    .map(
      function(r)
      {
        return parseInt(r, 10);
      }
    )
    .sort(
      function(a, b)
      {
        return a - b;
      }
    );
}

function getHeaderMap_(sheet, headerRow)
{
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};

  for (var c = 0; c < headers.length; c++)
  {
    var h = normalizeHeader_(headers[c]);

    if (h)
    {
      map[h] = c + 1;
    }
  }

  return map;
}

function getOrCreateHeaderMap_(sheet, headerRow, requiredHeaders)
{
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};

  for (var c = 0; c < headers.length; c++)
  {
    var h = normalizeHeader_(headers[c]);

    if (h)
    {
      map[h] = c + 1;
    }
  }

  requiredHeaders.forEach(
    function(headerName)
    {
      if (!map[headerName])
      {
        lastCol++;
        sheet.getRange(headerRow, lastCol).setValue(headerName);
        map[headerName] = lastCol;
      }
    }
  );

  return map;
}

function ensureHeadersExist_(sheet, headerMap, requiredHeaders, sheetLabel)
{
  var missing = requiredHeaders.filter(
    function(h)
    {
      return !headerMap[h];
    }
  );

  if (missing.length > 0)
  {
    throw new Error("Nel " + sheetLabel + " mancano queste colonne alla riga intestazioni: " + missing.join(", "));
  }
}

function normalizeHeader_(value)
{
  return String(value || "").trim();
}

function normalizeString_(value)
{
  return String(value || "").trim();
}

function normalizeAmount_(value)
{
  if (value === null || value === undefined || value === "")
  {
    return null;
  }

  if (typeof value === "number")
  {
    return round2_(value);
  }

  var s = String(value).trim();
  s = s.replace(/\s+/g, "");
  s = s.replace(/€/g, "");
  s = s.replace(/\./g, "");
  s = s.replace(/,/g, ".");

  var n = parseFloat(s);

  if (isNaN(n))
  {
    return null;
  }

  return round2_(n);
}

function round2_(n)
{
  return Math.round(n * 100) / 100;
}

function parseStatoPagamento_(statoPagamento)
{
  var s = String(statoPagamento || "").trim();
  var provider = "BANCA";

  if (/stripe/i.test(s))
  {
    provider = "STRIPE";
  }
  else if (/paypal/i.test(s))
  {
    provider = "PAYPAL";
  }

  var data = extractDateFromText_(s);

  return {
    raw: s,
    provider: provider,
    data: data
  };
}

function extractDateFromText_(text)
{
  var s = String(text || "").trim();

  var m = s.match(/(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})/);

  if (!m)
  {
    return null;
  }

  return normalizeDateString_(m[1]);
}

function normalizeDateString_(value)
{
  if (value === null || value === undefined || value === "")
  {
    return null;
  }

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime()))
  {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  var s = String(value).trim();

  var dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);

  if (dmy)
  {
    var day = parseInt(dmy[1], 10);
    var month = parseInt(dmy[2], 10);
    var year = parseInt(dmy[3], 10);

    if (year < 100)
    {
      year += 2000;
    }

    return pad4_(year) + "-" + pad2_(month) + "-" + pad2_(day);
  }

  var iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);

  if (iso)
  {
    return pad4_(parseInt(iso[1], 10)) + "-" + pad2_(parseInt(iso[2], 10)) + "-" + pad2_(parseInt(iso[3], 10));
  }

  var d = new Date(s);

  if (!isNaN(d.getTime()))
  {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  return null;
}

function pad2_(n)
{
  return ("0" + n).slice(-2);
}

function pad4_(n)
{
  var s = String(n);

  while (s.length < 4)
  {
    s = "0" + s;
  }

  return s;
}

function normalizeComparableText_(value)
{
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeFatturaForCompare_(value)
{
  return String(value || "")
    .replace(/\?/g, "")
    .trim()
    .toLowerCase();
}

function containsText_(haystack, needle)
{
  if (!needle)
  {
    return false;
  }

  return normalizeComparableText_(haystack).indexOf(normalizeComparableText_(needle)) !== -1;
}

function buildOrderDetail_(codiceSpedizione, cognome, nome, email, importo)
{
  return (codiceSpedizione || "-") +
    ": " +
    (cognome || "") +
    " " +
    (nome || "") +
    " (" +
    (email || "") +
    ") Importo:" +
    (importo || "");
}

function findPaymentMatchInBanca_(sheet, headers, totaleOrdine, email, cognome, dataOrdine, fatturaOrdine)
{
  return findGenericPaymentMatch_(
    {
      sheet: sheet,
      headers: headers,
      amountHeader: "Importo",
      dateHeader: "Data Val.",
      customerHeader: "Descrizione",
      customerType: "DESCRIZIONE",
      totaleOrdine: totaleOrdine,
      email: email,
      cognome: cognome,
      dataOrdine: dataOrdine,
      fatturaOrdine: fatturaOrdine
    }
  );
}

function findPaymentMatchInStripe_(sheet, headers, totaleOrdine, email, cognome, dataOrdine, fatturaOrdine)
{
  return findGenericPaymentMatch_(
    {
      sheet: sheet,
      headers: headers,
      amountHeader: "Amount",
      dateHeader: "Created date (UTC)",
      customerHeader: "Customer Email",
      customerType: "EMAIL",
      totaleOrdine: totaleOrdine,
      email: email,
      cognome: cognome,
      dataOrdine: dataOrdine,
      fatturaOrdine: fatturaOrdine
    }
  );
}

function findPaymentMatchInPaypal_(sheet, headers, totaleOrdine, email, cognome, dataOrdine, fatturaOrdine)
{
  return findGenericPaymentMatch_(
    {
      sheet: sheet,
      headers: headers,
      amountHeader: "Gross",
      dateHeader: "Date",
      customerHeader: "Customer From Email Address",
      customerType: "EMAIL",
      totaleOrdine: totaleOrdine,
      email: email,
      cognome: cognome,
      dataOrdine: dataOrdine,
      fatturaOrdine: fatturaOrdine
    }
  );
}

function findGenericPaymentMatch_(cfg)
{
  var sheet = cfg.sheet;
  var headers = cfg.headers;
  var lastRow = sheet.getLastRow();

  if (lastRow < 2)
  {
    return {
      sheet: sheet,
      headers: headers,
      match: null,
      fatturaConflict: false,
      existingFattura: "",
      ambigui: []
    };
  }

  var lastCol = sheet.getLastColumn();
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  var candidates = [];

  data.forEach(
    function(row, idx)
    {
      var rowNumber = idx + 2;
      var importo = normalizeAmount_(row[headers[cfg.amountHeader] - 1]);

      if (importo === null)
      {
        return;
      }

      if (round2_(importo) !== round2_(cfg.totaleOrdine))
      {
        return;
      }

      var customerField = normalizeString_(row[headers[cfg.customerHeader] - 1]);
      var dataMovimento = normalizeDateString_(row[headers[cfg.dateHeader] - 1]);

      var emailMatch = false;
      var cognomeMatch = false;

      if (cfg.customerType === "DESCRIZIONE")
      {
        emailMatch = cfg.email ? containsText_(customerField, cfg.email) : false;
        cognomeMatch = cfg.cognome ? containsText_(customerField, cfg.cognome) : false;
      }
      else if (cfg.customerType === "EMAIL")
      {
        emailMatch = cfg.email ? normalizeComparableText_(customerField) === normalizeComparableText_(cfg.email) : false;
        cognomeMatch = cfg.cognome ? containsText_(customerField, cfg.cognome) : false;
      }

      var customerMatch = emailMatch || cognomeMatch;
      var dateMatch = cfg.dataOrdine && dataMovimento ? (cfg.dataOrdine === dataMovimento) : false;

      if (!customerMatch)
      {
        return;
      }

      var tipoMatch = "ESATTO";
      var partialReason = "";

      if (!dateMatch)
      {
        tipoMatch = "PARZIALE";
        partialReason = "Importo e cliente coerenti, data diversa";
      }

      var existingFattura = normalizeString_(row[headers["Fattura"] - 1]);
      var fatturaConflict = false;

      if (existingFattura)
      {
        if (normalizeFatturaForCompare_(existingFattura) !== normalizeFatturaForCompare_(cfg.fatturaOrdine))
        {
          fatturaConflict = true;
        }
      }

      var existingSpedizione = normalizeString_(row[headers["Spedizione"] - 1]);
      var score = 0;

      if (emailMatch)
      {
        score += 100;
      }

      if (cognomeMatch)
      {
        score += 10;
      }

      if (dateMatch)
      {
        score += 1000;
      }

      if (!existingFattura && !existingSpedizione)
      {
        score += 5;
      }

      candidates.push({
        row: rowNumber,
        tipoMatch: tipoMatch,
        partialReason: partialReason,
        emailMatch: emailMatch,
        cognomeMatch: cognomeMatch,
        dateMatch: dateMatch,
        customerField: customerField,
        dataMovimento: dataMovimento,
        fatturaConflict: fatturaConflict,
        existingFattura: existingFattura,
        existingSpedizione: existingSpedizione,
        score: score
      });
    }
  );

  if (candidates.length === 0)
  {
    return {
      sheet: sheet,
      headers: headers,
      match: null,
      fatturaConflict: false,
      existingFattura: "",
      ambigui: []
    };
  }

  var indistinguibili = groupAmbiguousCandidates_(candidates);

  candidates.sort(
    function(a, b)
    {
      if (b.score !== a.score)
      {
        return b.score - a.score;
      }

      if (a.fatturaConflict !== b.fatturaConflict)
      {
        return a.fatturaConflict ? 1 : -1;
      }

      return a.row - b.row;
    }
  );

  var best = candidates[0];

  return {
    sheet: sheet,
    headers: headers,
    match: best,
    fatturaConflict: best.fatturaConflict,
    existingFattura: best.existingFattura,
    ambigui: indistinguibili
  };
}

function groupAmbiguousCandidates_(candidates)
{
  var groups = {};

  candidates.forEach(
    function(c)
    {
      var key = [
        c.dataMovimento || "",
        c.emailMatch ? "EM1" : "EM0",
        c.cognomeMatch ? "CM1" : "CM0",
        c.tipoMatch,
        c.score
      ].join("|");

      if (!groups[key])
      {
        groups[key] = [];
      }

      groups[key].push(c.row);
    }
  );

  var ambigui = [];

  Object.keys(groups).forEach(
    function(key)
    {
      if (groups[key].length > 1)
      {
        ambigui = ambigui.concat(groups[key]);
      }
    }
  );

  ambigui.sort(
    function(a, b)
    {
      return a - b;
    }
  );

  return ambigui;
}

function writeMovimentoResult_(sheet, headers, row, fattura, codiceSpedizione, campagna, isParziale)
{
  var fatturaRange = sheet.getRange(row, headers["Fattura"]);
  var spedizioneRange = sheet.getRange(row, headers["Spedizione"]);
  var campagnaRange = sheet.getRange(row, headers["Campagna"]);

  spedizioneRange.setValue(codiceSpedizione);
  campagnaRange.setValue(campagna);

  if (isParziale)
  {
    var testo = (fattura || "") + "?";
    var normalStyle = SpreadsheetApp.newTextStyle()
      .setBold(false)
      .build();
    var boldStyle = SpreadsheetApp.newTextStyle()
      .setBold(true)
      .build();

    var richText = SpreadsheetApp.newRichTextValue()
      .setText(testo)
      .setTextStyle(0, testo.length - 1, normalStyle)
      .setTextStyle(testo.length - 1, testo.length, boldStyle)
      .build();

    fatturaRange.setRichTextValue(richText);
    fatturaRange.setBackground("#ffcccc");
  }
  else
  {
    fatturaRange.setValue(fattura);
    fatturaRange.setBackground(null);
    fatturaRange.setFontWeight("normal");
  }
}

function buildFinalReport_(nomeFoglioOrdini, risultati)
{
  var lines = [];

  lines.push("Verifica pagamenti - Campagna: " + nomeFoglioOrdini);
  lines.push("");

  lines.push("Corrispondenze esatte: " + risultati.esatti.length);

  if (risultati.esatti.length > 0)
  {
    risultati.esatti.forEach(
      function(x)
      {
        lines.push("- " + x.dettaglio + " [" + x.foglio + " riga " + x.movimentoRow + "]");
      }
    );
  }

  lines.push("");
  lines.push("Corrispondenze parziali: " + risultati.parziali.length);

  if (risultati.parziali.length > 0)
  {
    risultati.parziali.forEach(
      function(x)
      {
        var extraDate = "";

        if (x.dataOrdine || x.dataMovimento)
        {
          extraDate =
            " - data ordine: " + (x.dataOrdine || "(vuota)") +
            ", data movimento: " + (x.dataMovimento || "(vuota)");
        }

        lines.push(
          "- " + x.dettaglio +
          " [" + x.foglio + " riga " + x.movimentoRow + "] - " +
          x.motivo +
          extraDate
        );
      }
    );
  }

  lines.push("");
  lines.push("Non trovati: " + risultati.nonTrovati.length);

  if (risultati.nonTrovati.length > 0)
  {
    risultati.nonTrovati.forEach(
      function(x)
      {
        lines.push("- " + x.dettaglio + (x.motivo ? " [" + x.motivo + "]" : ""));
      }
    );
  }

  lines.push("");
  lines.push("Ambigui: " + risultati.ambigui.length);

  if (risultati.ambigui.length > 0)
  {
    risultati.ambigui.forEach(
      function(x)
      {
        lines.push(
          "- " + x.dettaglio +
          " [" + x.foglio + " righe " + x.righe.join(", ") + "] - movimenti indistinguibili"
        );
      }
    );
  }

  lines.push("");
  lines.push("Errori: " + risultati.errori.length);

  if (risultati.errori.length > 0)
  {
    risultati.errori.forEach(
      function(x)
      {
        var extra = "";

        if (x.foglio && x.movimentoRow)
        {
          extra += " [" + x.foglio + " riga " + x.movimentoRow + "]";
        }

        if (x.motivo === "Fattura già presente nel movimento e diversa da quella ordine")
        {
          extra +=
            " fattura ordine: " + (x.fatturaOrdine || "(vuota)") +
            ", fattura movimento: " + (x.fatturaMovimento || "(vuota)");
        }
        else if (x.fatturaMovimento)
        {
          extra += " fattura movimento: " + x.fatturaMovimento;
        }

        lines.push("- " + x.dettaglio + extra + " [" + x.motivo + "]");
      }
    );
  }

  lines.push("");
  lines.push("Saltati: " + risultati.saltati.length);

  if (risultati.saltati.length > 0)
  {
    risultati.saltati.forEach(
      function(x)
      {
        lines.push("- riga " + x.row + ": " + x.dettaglio + " [" + x.motivo + "]");
      }
    );
  }

  return lines.join("\n");
}



function importBankMovements()
{
  var html = HtmlService.createHtmlOutputFromFile("BankMovements")
    .setWidth(500)
    .setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, "Importa CSV banca");
}

function importaCsvBancaDaUpload(base64Data, fileName, mimeType)
{
  if (!base64Data)
  {
    throw new Error("Nessun file ricevuto.");
  }

  var decoded = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    mimeType || "text/csv",
    fileName || "upload.csv"
  );

  var csvText = decoded.getDataAsString("UTF-8");

  return importaCsvBancaDaTesto_(csvText);
}

function importaCsvBancaDaTesto_(csvText)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Banca 2026");

  if (!sheet)
  {
    throw new Error('Il foglio "Banca 2026" non esiste.');
  }

  var headerRow = 1;
  var requiredHeaders = ["Valuta", "Descrizione", "Importo"];
  var headerMap = getOrCreateHeaderMap_(sheet, headerRow, requiredHeaders);

  var rows = parseBankCsvSemicolon_(csvText);

  if (rows.length === 0)
  {
    return "Nessun movimento trovato nel CSV.";
  }

  var existingKeys = getExistingDescrizioneImportoKeys_(sheet, headerMap, headerRow);
  var newRowsData = [];
  var skipped = 0;

  rows.forEach(
    function(row)
    {
      var key = buildDescrizioneImportoKey_(row.description, row.amount);

      if (existingKeys[key])
      {
        skipped++;
        return;
      }

      existingKeys[key] = true;

      newRowsData.push({
        valutaDate: row.valutaDate,
        description: row.description,
        amount: row.amount
      });
    }
  );

  newRowsData.sort(
    function(a, b)
    {
      return a.valutaDate.getTime() - b.valutaDate.getTime();
    }
  );

  var newRows = newRowsData.map(
    function(item)
    {
      var outputRow = new Array(sheet.getLastColumn()).fill("");

      outputRow[headerMap["Valuta"] - 1] = item.valutaDate;
      outputRow[headerMap["Descrizione"] - 1] = item.description;
      outputRow[headerMap["Importo"] - 1] = item.amount;

      return outputRow;
    }
  );

  if (newRows.length > 0)
  {
    var startRow = sheet.getLastRow() + 1;

    sheet.getRange(startRow, 1, newRows.length, sheet.getLastColumn()).setValues(newRows);

    sheet.getRange(startRow, headerMap["Valuta"], newRows.length, 1).setNumberFormat("dd/MM/yyyy");
    sheet.getRange(startRow, headerMap["Importo"], newRows.length, 1).setNumberFormat("#,##0.00");
  }

  var report = [
    'Foglio: Banca 2026',
    'Movimenti letti dal CSV: ' + rows.length,
    'Nuovi importati: ' + newRows.length,
    'Duplicati saltati: ' + skipped,
    'Regola data Valuta: meno recente tra Txn. Date e Value Date'
  ].join("\n");

  Logger.log(report);
  SpreadsheetApp.getUi().alert(report);

  return report;
}

function parseBankCsvSemicolon_(csvText)
{
  var normalized = String(csvText || "")
    .replace(/^\uFEFF/, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();

  if (!normalized)
  {
    return [];
  }

  var lines = normalized.split("\n");

  if (lines.length < 2)
  {
    return [];
  }

  var header = splitSemicolonCsvLine_(lines[0]);
  var headerMap = {};

  for (var i = 0; i < header.length; i++)
  {
    headerMap[String(header[i] || "").trim()] = i;
  }

  var required = ["Txn. Date", "Value Date", "Description", "Amount"];

  required.forEach(
    function(h)
    {
      if (typeof headerMap[h] === "undefined")
      {
        throw new Error("Colonna mancante nel CSV: " + h);
      }
    }
  );

  var rows = [];

  for (var r = 1; r < lines.length; r++)
  {
    var line = lines[r].trim();

    if (!line)
    {
      continue;
    }

    var cols = splitSemicolonCsvLine_(line);

    var txnDateStr = safeCsvValue_(cols, headerMap["Txn. Date"]);
    var valueDateStr = safeCsvValue_(cols, headerMap["Value Date"]);
    var description = safeCsvValue_(cols, headerMap["Description"]);
    var amountStr = safeCsvValue_(cols, headerMap["Amount"]);

    if (!txnDateStr || !valueDateStr || !description || !amountStr)
    {
      continue;
    }

    var txnDate = parseItalianDate_(txnDateStr);
    var valueDate = parseItalianDate_(valueDateStr);
    var amount = parseItalianAmount_(amountStr);

    if (!txnDate || !valueDate)
    {
      continue;
    }

    if (amount === null)
    {
      continue;
    }

    rows.push({
      txnDate: txnDate,
      valueDate: valueDate,
      valutaDate: getOlderDate_(txnDate, valueDate),
      description: description,
      amount: amount
    });
  }

  return rows;
}

function getOlderDate_(dateA, dateB)
{
  if (!dateA)
  {
    return dateB;
  }

  if (!dateB)
  {
    return dateA;
  }

  return dateA.getTime() <= dateB.getTime() ? dateA : dateB;
}

function splitSemicolonCsvLine_(line)
{
  var result = [];
  var current = "";
  var inQuotes = false;

  for (var i = 0; i < line.length; i++)
  {
    var ch = line.charAt(i);

    if (ch === '"')
    {
      if (inQuotes && i + 1 < line.length && line.charAt(i + 1) === '"')
      {
        current += '"';
        i++;
      }
      else
      {
        inQuotes = !inQuotes;
      }

      continue;
    }

    if (ch === ";" && !inQuotes)
    {
      result.push(current);
      current = "";
      continue;
    }

    current += ch;
  }

  result.push(current);

  return result.map(
    function(v)
    {
      return String(v || "").trim();
    }
  );
}

function safeCsvValue_(cols, index)
{
  if (typeof index === "undefined" || index === null)
  {
    return "";
  }

  if (index < 0 || index >= cols.length)
  {
    return "";
  }

  return String(cols[index] || "").trim();
}

function parseItalianDate_(value)
{
  var s = String(value || "").trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (!m)
  {
    return null;
  }

  var day = parseInt(m[1], 10);
  var month = parseInt(m[2], 10) - 1;
  var year = parseInt(m[3], 10);

  return new Date(year, month, day);
}

function parseItalianAmount_(value)
{
  if (value === null || value === undefined)
  {
    return null;
  }

  var s = String(value).trim();

  if (!s)
  {
    return null;
  }

  s = s.replace(/^'/, "");
  s = s.replace(/\s+/g, "");
  s = s.replace(/\./g, "");
  s = s.replace(/,/g, ".");

  var n = parseFloat(s);

  if (isNaN(n))
  {
    return null;
  }

  return Math.round(n * 100) / 100;
}

function buildDescrizioneImportoKey_(descrizione, importo)
{
  return normalizeKeyText_(descrizione) + "||" + normalizeAmountKey_(importo);
}

function normalizeKeyText_(value)
{
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function normalizeAmountKey_(value)
{
  var n = typeof value === "number" ? value : parseItalianAmount_(value);

  if (n === null)
  {
    return "";
  }

  return n.toFixed(2);
}

function getExistingDescrizioneImportoKeys_(sheet, headerMap, headerRow)
{
  var lastRow = sheet.getLastRow();
  var keys = {};

  if (lastRow <= headerRow)
  {
    return keys;
  }

  var descCol = headerMap["Descrizione"];
  var impCol = headerMap["Importo"];
  var numRows = lastRow - headerRow;

  var descrizioni = sheet.getRange(headerRow + 1, descCol, numRows, 1).getDisplayValues();
  var importi = sheet.getRange(headerRow + 1, impCol, numRows, 1).getValues();

  for (var i = 0; i < numRows; i++)
  {
    var descrizione = String(descrizioni[i][0] || "").trim();
    var importo = importi[i][0];

    if (!descrizione && (importo === "" || importo === null))
    {
      continue;
    }

    var key = buildDescrizioneImportoKey_(descrizione, importo);
    keys[key] = true;
  }

  return keys;
}

function getOrCreateHeaderMap_(sheet, headerRow, requiredHeaders)
{
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};

  for (var c = 0; c < headers.length; c++)
  {
    var h = String(headers[c] || "").trim();

    if (h)
    {
      map[h] = c + 1;
    }
  }

  requiredHeaders.forEach(
    function(headerName)
    {
      if (!map[headerName])
      {
        lastCol++;
        sheet.getRange(headerRow, lastCol).setValue(headerName);
        map[headerName] = lastCol;
      }
    }
  );

  return map;
}
