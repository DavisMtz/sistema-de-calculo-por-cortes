/**
 * @OnlyCurrentDoc
 */

// --- CONSTANTES GLOBALES ---
const SHEET_DATA_NAME = "MesActual";
const DRIVE_CORTES_SHEET_NAME = "DriveCortes";
const CONCENTRADO_SHEET_NAME = "Concentrado";
// IMPORTANTE: Por seguridad, te sugiero rotar (cambiar) esta API Key en Google AI Studio.
const GEMINI_API_KEY = "AQ.Ab8RN6KoTQA2ciTKfxjThzBgOn3BJhAAmitn3Ez-BCRcPfnFRA"; 

// --- CONFIGURACIÓN DE LA INTERFAZ DE GOOGLE SHEETS ---


function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('🏆 Ranking Liverpool')
    .addItem('Subir Reporte de Venta', 'showUploaderDialog')
    .addSeparator()
    .addItem('Actualizar Hoja "Concentrado"', 'updateConcentradoSheet')
    .addToUi();
}

function showUploaderDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Uploader').setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cargar Nuevo Reporte');
}

// --- SERVIDOR WEB ---

function doGet(e) {
  setupSheets();
  return HtmlService.createHtmlOutputFromFile('Ranking').setTitle('Ranking de Ventas Mensual').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// --- LÓGICA DE BACKEND ---

function uploadImageAndProcess(formObject) {
  try {
    setupSheets();
    const blob = Utilities.newBlob(Utilities.base64Decode(formObject.fileData), formObject.mimeType, formObject.fileName);
    const extractedData = callGeminiAPI(blob);
    
    if (!extractedData || !extractedData.agents || extractedData.agents.length === 0) {
      throw new Error("La IA no pudo extraer datos de agentes. Revisa la imagen.");
    }
    
    logDailyCuts(extractedData.agents, extractedData.reportDate);
    updateRankingSheet(extractedData.agents); 
    updateConcentradoSheet();

    return { status: "success", message: "Todas las hojas han sido actualizadas exitosamente." };
  } catch (e) {
    console.error("Error en uploadImageAndProcess: " + e.toString());
    return { status: "error", message: "Error del servidor: " + e.message };
  }
}

function getRankingData() {
  checkAndResetForNewMonth();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATA_NAME);
  if (sheet.getLastRow() < 2) return [];
  const range = sheet.getRange("A2:E" + sheet.getLastRow());
  const values = range.getValues();
  return values.map((row, index) => {
    if (!row[0]) return null;
    const currentRank = index + 1;
    const previousRank = parseInt(row[3]);
    let change = "same";
    if (previousRank === 999) { change = "new"; }
    else if (currentRank < previousRank) { change = "up"; }
    else if (currentRank > previousRank) { change = "down"; }
    return {
      rank: currentRank, name: row[0],
      amount: parseToNumber(row[1]).toLocaleString('es-MX', { style: 'currency', currency: 'MXN' }),
      tickets: parseInt(row[2]) || 0, change: change
    };
  }).filter(Boolean);
}

// --- FUNCIONES DE PROCESAMIENTO ---

/**
 * Función ultra segura para convertir CUALQUIER valor en un número decimal correcto.
 * Garantiza que la suma matemática sea perfecta, ignorando textos extraños o formatos.
 */
function parseToNumber(value) {
  if (typeof value === 'number') return value;
  if (!value) return 0;
  // Convertimos a texto y quitamos TODO lo que no sea dígito, punto o signo menos
  const cleanString = value.toString().replace(/[^0-9.-]+/g, "");
  return parseFloat(cleanString) || 0;
}

/**
 * Función REESCRITA para máxima velocidad y sumas correctas, sin forzar formatos visuales.
 */
function updateRankingSheet(agentsData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATA_NAME);
  const lastRow = sheet.getLastRow();
  
  let currentData = [];
  if (lastRow >= 2) {
    currentData = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  }

  const today = new Date();
  const dataMap = new Map();

  // 1. Cargar datos existentes en memoria (Convirtiendo a números reales)
  currentData.forEach((row, index) => {
    if (row[0]) {
      dataMap.set(row[0].toString().trim().toUpperCase(), {
        name: row[0],
        amount: parseToNumber(row[1]), 
        tickets: parseInt(row[2]) || 0,
        previousRank: index + 1, 
        lastUpdate: row[4]
      });
    }
  });

  // 2. Sumar/Agregar los nuevos datos extraídos
  agentsData.forEach(newAgent => {
    let agentName = newAgent.name ? newAgent.name.trim() : "";
    if (!agentName || agentName.toLowerCase() === 'referencia' || agentName.toLowerCase() === 'referencias') {
      agentName = 'Referencias';
    }

    const mapKey = agentName.toUpperCase();
    const formattedName = toTitleCase(agentName);
    
    // Forzamos la conversión a número puro antes de operar
    const amountToAdd = parseToNumber(newAgent.amount);
    const ticketsToAdd = parseInt(newAgent.tickets) || 0;

    if (dataMap.has(mapKey)) {
      // Actualizar asesor existente con SUMA MATEMÁTICA PURA
      const existing = dataMap.get(mapKey);
      existing.amount += amountToAdd;
      existing.tickets += ticketsToAdd;
      existing.lastUpdate = today;
    } else {
      // Crear asesor nuevo
      dataMap.set(mapKey, {
        name: formattedName,
        amount: amountToAdd,
        tickets: ticketsToAdd,
        previousRank: 999,
        lastUpdate: today
      });
    }
  });

  // 3. Separar y Ordenar (Referencias siempre al final)
  const allAgents = Array.from(dataMap.values());
  const referenciasData = [];
  const otherAdvisorsData = [];

  allAgents.forEach(agent => {
    const rowArray = [agent.name, agent.amount, agent.tickets, agent.previousRank, agent.lastUpdate];
    if (agent.name.toLowerCase() === 'referencias') {
      referenciasData.push(rowArray);
    } else {
      otherAdvisorsData.push(rowArray);
    }
  });

  // Ordenar asesores por monto de venta (de mayor a menor)
  otherAdvisorsData.sort((a, b) => b[1] - a[1]);

  // Juntar todo, dejando referencias al final
  const finalSortedData = otherAdvisorsData.concat(referenciasData);

  // 4. Escribir en la hoja en un solo paso, solo datos puros (Google Sheets pondrá el formato)
  if (finalSortedData.length > 0) {
    sheet.getRange(2, 1, Math.max(lastRow - 1, 1), 5).clearContent();
    const targetRange = sheet.getRange(2, 1, finalSortedData.length, 5);
    targetRange.setValues(finalSortedData);
  }
  SpreadsheetApp.flush();
}

function updateConcentradoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mesActualSheet = ss.getSheetByName(SHEET_DATA_NAME);
  const concentradoSheet = ss.getSheetByName(CONCENTRADO_SHEET_NAME);

  let metaVenta = parseFloat(concentradoSheet.getRange("G1").getValue());
  if (isNaN(metaVenta) || metaVenta <= 0) {
    metaVenta = 1;
  }

  if (mesActualSheet.getLastRow() < 2) {
    if (concentradoSheet.getLastRow() > 1) {
      concentradoSheet.getRange("A2:F" + concentradoSheet.getLastRow()).clearContent();
    }
    return;
  }

  const sourceData = mesActualSheet.getRange("A2:E" + mesActualSheet.getLastRow()).getValues();

  const processedData = sourceData.map((row, index) => {
    const nombre = row[0];
    const montoTotal = parseToNumber(row[1]); 
    const boletasTotales = parseInt(row[2]) || 0;
    const rankingAnterior = row[3];
    const rankingActual = index + 1;

    let cambioSymbol = '➖';
    if (rankingAnterior === 999) cambioSymbol = '⭐';
    else if (rankingActual < rankingAnterior) cambioSymbol = '⬆️';
    else if (rankingActual > rankingAnterior) cambioSymbol = '⬇️';

    const porcentajeMeta = montoTotal / metaVenta;

    return [rankingActual, nombre, boletasTotales, montoTotal, porcentajeMeta, cambioSymbol];
  });

  if (concentradoSheet.getLastRow() > 1) {
    concentradoSheet.getRange("A2:F" + concentradoSheet.getLastRow()).clearContent();
  }

  if (processedData.length > 0) {
    const targetRange = concentradoSheet.getRange(2, 1, processedData.length, 6);
    targetRange.setValues(processedData);
  }

  const now = new Date();
  const formattedDate = Utilities.formatDate(now, "America/Mexico_City", "dd/MM/yyyy HH:mm:ss");
  concentradoSheet.getRange("I1").setValue(formattedDate).setHorizontalAlignment("left");

  SpreadsheetApp.flush();
}

function toTitleCase(str) {
  if (!str) return "";
  return str.toLowerCase().replace(/\b(\w)/g, s => s.toUpperCase());
}

function callGeminiAPI(imageBlob) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-pro-preview:generateContent?key=" + GEMINI_API_KEY;
  
  const prompt = `Analiza la imagen de este reporte de ventas. Extrae la fecha del reporte (por ejemplo, "31 de julio de 2025") y los datos de cada asesor. Ignora "Super Agent". Si un nombre de asesor está en blanco, asígnalo como "Referencias". Devuelve ÚNICAMENTE un objeto JSON válido. El formato debe ser: {"reportDate": "DD/MM/YYYY", "storeType": "Liverpool o Suburbia", "agents": [{"name": "NOMBRE COMPLETO", "tickets": NÚMERO, "amount": MONTO_VENDIDO_COMO_NÚMERO. USA UN PUNTO (.) COMO SEPARADOR DECIMAL. NO INCLUYAS COMAS (,) NI SÍMBOLO DE MONEDA ($). Por ejemplo, $98,621.80 debe ser 98621.80}]}`;
  
  const requestBody = { 
    "contents": [{ 
      "parts": [
        { "text": prompt }, 
        { "inline_data": { "mime_type": imageBlob.getContentType(), "data": Utilities.base64Encode(imageBlob.getBytes()) } }
      ] 
    }],
    "generationConfig": {
      "response_mime_type": "application/json"
    }
  };
  
  const options = { 
    'method': 'post', 
    'contentType': 'application/json', 
    'payload': JSON.stringify(requestBody), 
    'muteHttpExceptions': true 
  };
  
  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) { 
    throw new Error(`Error en API de Gemini. Código: ${response.getResponseCode()}. Respuesta: ${response.getContentText()}`); 
  }
  
  const jsonResponse = JSON.parse(response.getContentText());
  const content = jsonResponse.candidates[0].content.parts[0].text;
  
  return JSON.parse(content.trim());
}

function logDailyCuts(agentsData, reportDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DRIVE_CORTES_SHEET_NAME);
  const rowsToAdd = agentsData.map(agent => [
    reportDate, 
    toTitleCase(agent.name), 
    parseInt(agent.tickets) || 0, 
    parseToNumber(agent.amount) // Aseguramos que sea número real aquí también
  ]);
  
  if (rowsToAdd.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAdd.length, 4).setValues(rowsToAdd);
  }
}

function checkAndResetForNewMonth() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_DATA_NAME);
  if (sheet.getLastRow() < 2) return;
  const lastUpdateValue = sheet.getRange("E2").getValue();
  if (lastUpdateValue && lastUpdateValue instanceof Date) {
    const lastUpdateMonth = lastUpdateValue.getMonth();
    const now = new Date(new Date().toLocaleString("en-US", { timeZone: "America/Mexico_City" }));
    if (lastUpdateMonth !== now.getMonth()) {
      sheet.getRange("A2:E" + sheet.getLastRow()).clearContent();
    }
  }
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let dataSheet = ss.getSheetByName(SHEET_DATA_NAME);
  if (!dataSheet) { dataSheet = ss.insertSheet(SHEET_DATA_NAME); }
  let headers = dataSheet.getRange(1, 1, 1, 5).getValues()[0];
  const expectedHeaders = ["Asesor", "MontoTotal", "BoletasTotales", "RankingAnterior", "FechaActualizacion"];
  if (JSON.stringify(headers) !== JSON.stringify(expectedHeaders)) {
    dataSheet.getRange(1, 1, 1, 5).setValues([expectedHeaders]).setFontWeight("bold").setHorizontalAlignment("center");
    dataSheet.setColumnWidths(1, 5, 120).setColumnWidth(1, 250);
  }

  let cortesSheet = ss.getSheetByName(DRIVE_CORTES_SHEET_NAME);
  if (!cortesSheet) { cortesSheet = ss.insertSheet(DRIVE_CORTES_SHEET_NAME); }
  headers = cortesSheet.getRange(1, 1, 1, 4).getValues()[0];
  const expectedCortesHeaders = ["Fecha", "Agente de Venta", "Boletas", "Monto de venta"];
  if (JSON.stringify(headers) !== JSON.stringify(expectedCortesHeaders)) {
    cortesSheet.getRange(1, 1, 1, 4).setValues([expectedCortesHeaders]).setFontWeight("bold").setHorizontalAlignment("center");
    cortesSheet.setColumnWidth(1, 120).setColumnWidth(2, 250).setColumnWidth(3, 100).setColumnWidth(4, 150);
  }

  let concentradoSheet = ss.getSheetByName(CONCENTRADO_SHEET_NAME);
  if (!concentradoSheet) { concentradoSheet = ss.insertSheet(CONCENTRADO_SHEET_NAME); }
  headers = concentradoSheet.getRange(1, 1, 1, 6).getValues()[0];
  const expectedConcentradoHeaders = ["Nº", "Nombre", "Boletas", "Venta Total", "% Meta", "#"];
  if (JSON.stringify(headers) !== JSON.stringify(expectedConcentradoHeaders)) {
    concentradoSheet.getRange(1, 1, 1, 6).setValues([expectedConcentradoHeaders]).setFontWeight("bold").setHorizontalAlignment("center");
    concentradoSheet.getRange("G1").setValue(700000).setNumberFormat("$#,##0.00");
    concentradoSheet.getRange("H1").setValue("<-- Ingrese la meta de venta aquí").setFontStyle("italic");
    concentradoSheet.getRange("J1").setValue("<-- Última Actualización").setFontStyle("italic");
    concentradoSheet.setColumnWidth(2, 250).setColumnWidth(4, 150).setColumnWidth(9, 150);
  }
}

function processImageOnly(formObject) {
  try {
    setupSheets(); 
    const blob = Utilities.newBlob(Utilities.base64Decode(formObject.fileData), formObject.mimeType, formObject.fileName);
    const extractedData = callGeminiAPI(blob);
    if (!extractedData || !extractedData.agents || extractedData.agents.length === 0) {
      throw new Error("La IA no pudo extraer datos de agentes. Revisa la imagen.");
    }
    return { status: "success", data: extractedData };
  } catch (e) {
    console.error("Error en processImageOnly: " + e.toString());
    return { status: "error", message: "Error del servidor: " + e.toString() };
  }
}

function updateSheetsWithData(processedData) {
  try {
    if (!processedData || !processedData.agents || processedData.agents.length === 0) {
        throw new Error("No se recibieron datos de agentes para procesar.");
    }
    logDailyCuts(processedData.agents, processedData.reportDate);
    updateRankingSheet(processedData.agents);
    updateConcentradoSheet();
    return { status: "success", message: "Todas las hojas han sido actualizadas." };
  } catch (e) {
    console.error("Error en updateSheetsWithData: " + e.toString());
    return { status: "error", message: "Error del servidor al actualizar hojas: " + e.toString() };
  }
}
