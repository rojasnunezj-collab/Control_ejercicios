/**
 * ============================================================================
 * BACKEND: ENTRENADOR EN CASA + CONTROL DE TIEMPOS Y SUPLEMENTOS
 * ============================================================================
 * INSTRUCCIONES:
 * 1. Pega este código en Extensiones -> Apps Script de tu Google Sheets.
 * 2. Ejecuta una vez la función "inicializarHojas" desde el botón "Ejecutar" arriba.
 *    Esto creará y poblará las 3 pestañas automáticamente con todos los 17 ejercicios.
 * 3. Haz clic en "Implementar" -> "Gestionar implementaciones" -> Editar (o Nueva implementación)
 *    -> Tipo: "Aplicación web" -> Versión: "Nueva versión" -> Implementar.
 */

const TELEGRAM_BOT_TOKEN = "8933766777:AAE4_prabw0hjGys8vP3p7Qk4glLnV8ia_M";
const TELEGRAM_CHAT_ID = "8933766777";

// Catálogo completo de los 17 ejercicios de tu rutina para poblar Config_Rutina
const CATALOGO_RUTINA = [
  // Bloque 1: Empuje (Lunes y Jueves)
  { id: "b1_e1", nombre: "Compuesto: Remo vertical agarre abierto + Press militar con barra", peso: "Barra 24 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b1_e2", nombre: "Press en Puente de Glúteos", peso: "2 Mancuernas de 7 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 10, segPorRep: 2.8 },
  { id: "b1_e3", nombre: "Press Hexagonal (Squeeze Press) en Puente de Glúteos", peso: "2 Mancuernas de 7 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 10, segPorRep: 2.8 },
  { id: "b1_e4", nombre: "Flexiones con Manos en Banco (Pecho Bajo)", peso: "Peso Corporal", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b1_e5", nombre: "Extensión de Tríceps tras Nuca / Copa sentado", peso: "Mancuerna 14 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b1_e6", nombre: "Fondos en Banco", peso: "Peso Corporal", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },

  // Bloque 2: Tracción y Laterales (Martes y Viernes)
  { id: "b2_e1", nombre: "Curl de Bíceps con Barra Recta de Pie", peso: "Barra 24 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b2_e2", nombre: "Curl Inverso con Barra / Press Invertido", peso: "Barra 20 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b2_e3", nombre: "Curl Concentrado Sentado con Mancuerna", peso: "Mancuerna 14 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b2_e4", nombre: "Encogimientos de Hombros tras Espalda con Barra", peso: "Barra 24 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b2_e5", nombre: "Superserie Hombros: Elevaciones Laterales + Frontales + Vuelos Posteriores", peso: "2 Mancuernas de 7 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },

  // Bloque 3: Piernas, Oblicuos y Potencia (Miércoles y Sábado)
  { id: "b3_e1", nombre: "Compuesto Funcional: Cargada/Curl Neutro + Press Unilateral", peso: "Mancuerna 14 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b3_e2", nombre: "Sentadilla con Barra", peso: "Barra 24 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 10, segPorRep: 2.8 },
  { id: "b3_e3", nombre: "Zancadas hacia adelante con Mancuernas", peso: "2 Mancuernas de 7 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b3_e4", nombre: "Elevaciones de Pantorrilla con Barra", peso: "Barra 24 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b3_e5", nombre: "Leñador Diagonal (Woodchopper) con Mancuerna", peso: "Mancuerna 14 kg", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 },
  { id: "b3_e6", nombre: "Superserie Oblicuos: Giros Rusos + Plancha Lateral", peso: "Mancuerna 14 kg / Corporal", descansoSerie: 60, descansoArmado: 30, prepSeg: 5, segPorRep: 2.8 }
];

function inicializarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  // 1. Pestaña Config_Rutina
  let sheetConfig = ss.getSheetByName("Config_Rutina");
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet("Config_Rutina");
  }

  // Encabezados
  if (sheetConfig.getLastRow() === 0) {
    sheetConfig.appendRow([
      "ID_Ejercicio",
      "Nombre_Ejercicio",
      "Peso_Texto",
      "Descanso_Serie_Seg",
      "Descanso_PostArmado_Seg",
      "Prep_Seg",
      "Seg_Por_Rep_Base",
      "Ultima_Actualizacion"
    ]);
    sheetConfig.getRange("A1:H1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    sheetConfig.setFrozenRows(1);
  }

  // Poblar o asegurar que todos los ejercicios del catálogo existan en Config_Rutina
  const dataConfig = sheetConfig.getDataRange().getValues();
  const existingIds = new Set();
  for (let i = 1; i < dataConfig.length; i++) {
    if (dataConfig[i][0]) existingIds.add(String(dataConfig[i][0]));
  }

  CATALOGO_RUTINA.forEach(item => {
    if (!existingIds.has(item.id)) {
      sheetConfig.appendRow([
        item.id,
        item.nombre,
        item.peso,
        item.descansoSerie,
        item.descansoArmado,
        item.prepSeg,
        item.segPorRep,
        ahora
      ]);
    }
  });

  // 2. Pestaña Historial_Entrenos
  let sheetHist = ss.getSheetByName("Historial_Entrenos");
  if (!sheetHist) {
    sheetHist = ss.insertSheet("Historial_Entrenos");
    sheetHist.appendRow([
      "Fecha",
      "Hora",
      "Bloque",
      "Ejercicio",
      "Serie_Num",
      "Lado_O_Parte",
      "Reps_Logradas",
      "Peso_Usado",
      "Tiempo_Serie_Seg",
      "Tiempo_Armado_Pesas_Seg",
      "Tiempo_Descanso_Seg"
    ]);
    sheetHist.getRange("A1:K1").setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff");
    sheetHist.setFrozenRows(1);
  }

  // 3. Pestaña Suplementos
  let sheetSup = ss.getSheetByName("Suplementos");
  if (!sheetSup) {
    sheetSup = ss.insertSheet("Suplementos");
    sheetSup.appendRow([
      "Fecha",
      "Hora",
      "Tipo",
      "Detalle",
      "Scoops_Restantes",
      "Timer_Minutos"
    ]);
    sheetSup.getRange("A1:F1").setFontWeight("bold").setBackground("#7c2d12").setFontColor("#ffffff");
    sheetSup.setFrozenRows(1);
  }
}

function doGet(e) {
  inicializarHojas();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfig = ss.getSheetByName("Config_Rutina");
  const data = sheetConfig.getDataRange().getValues();
  const configs = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      configs[row[0]] = {
        peso: row[2],
        descansoSerie: Number(row[3]) || 60,
        descansoArmado: Number(row[4]) || 30,
        prepSeg: Number(row[5]) || 5,
        segPorRep: Number(row[6]) || 2.8
      };
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    configs: configs
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    inicializarHojas();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const body = JSON.parse(e.postData.contents);
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const hora = Utilities.formatDate(ahora, Session.getScriptTimeZone(), "HH:mm:ss");
    const fechaHora = fecha + " " + hora;

    // 1. REGISTRAR SERIE EN HISTORIAL Y ACTUALIZAR CONFIG_RUTINA
    if (body.accion === "guardar_serie") {
      const sheetHist = ss.getSheetByName("Historial_Entrenos");
      sheetHist.appendRow([
        fecha,
        hora,
        body.bloque || "",
        body.ejercicio || "",
        body.serie || 1,
        body.lado || "Completo",
        body.reps || 0,
        body.peso || "",
        body.tiempoSerieSeg || 0,
        body.tiempoArmadoSeg || 0,
        body.tiempoDescansoSeg || 0
      ]);

      // También mantener actualizado el ejercicio en Config_Rutina con el peso y descansos usados
      actualizarOInsertarConfig(ss, {
        idEjercicio: body.idEjercicio || "",
        nombre: body.ejercicio,
        peso: body.peso,
        descansoSerie: body.tiempoDescansoSeg,
        descansoArmado: body.descansoArmado || 30,
        prepSeg: body.prepSeg || 5,
        segPorRep: body.segPorRep || 2.8,
        fechaHora: fechaHora
      });
    }

    // 2. GUARDAR CONFIGURACIÓN DE UN EJERCICIO
    if (body.accion === "guardar_config") {
      actualizarOInsertarConfig(ss, {
        idEjercicio: body.idEjercicio,
        nombre: body.nombre,
        peso: body.peso,
        descansoSerie: body.descansoSerie,
        descansoArmado: body.descansoArmado,
        prepSeg: body.prepSeg,
        segPorRep: body.segPorRep,
        fechaHora: fechaHora
      });
    }

    // 3. REGISTRAR SUPLEMENTO (SCOOP PRE-ENTRENO O VITAMINAS)
    if (body.accion === "guardar_suplemento") {
      const sheetSup = ss.getSheetByName("Suplementos");
      sheetSup.appendRow([
        fecha,
        hora,
        body.tipo || "Pre-Entreno",
        body.detalle || "1 Scoop",
        body.scoopsRestantes || 0,
        body.timerMinutos || 15
      ]);
    }

    // 4. FINALIZAR RUTINA Y NOTIFICAR POR TELEGRAM
    if (body.accion === "finalizar_rutina" && TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) {
      const msg = "🏋️ *ENTRENAMIENTO COMPLETADO*\n" +
                  "📅 *Bloque:* " + body.bloque + "\n" +
                  "⏱️ *Tiempo efectivo en series:* " + body.totalSerieStr + "\n" +
                  "🔧 *Tiempo armando pesas:* " + body.totalArmadoStr + "\n" +
                  "😮‍💨 *Tiempo en descanso:* " + body.totalDescansoStr + "\n" +
                  "⚡ *Pre-entreno tomado:* " + (body.preEntrenoTomado ? "Sí" : "No");
      enviarMensajeTelegram(msg);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function actualizarOInsertarConfig(ss, item) {
  const sheetConfig = ss.getSheetByName("Config_Rutina");
  const data = sheetConfig.getDataRange().getValues();
  let encontrado = false;

  for (let i = 1; i < data.length; i++) {
    const rowId = data[i][0];
    const rowName = data[i][1];
    // Coincidencia por ID o por Nombre del ejercicio
    if ((item.idEjercicio && rowId === item.idEjercicio) || (item.nombre && rowName === item.nombre)) {
      sheetConfig.getRange(i + 1, 2, 1, 7).setValues([[
        item.nombre || rowName,
        item.peso || data[i][2],
        item.descansoSerie || data[i][3],
        item.descansoArmado || data[i][4],
        item.prepSeg || data[i][5],
        item.segPorRep || data[i][6],
        item.fechaHora
      ]]);
      encontrado = true;
      break;
    }
  }

  if (!encontrado && (item.idEjercicio || item.nombre)) {
    sheetConfig.appendRow([
      item.idEjercicio || "",
      item.nombre || "",
      item.peso || "",
      item.descansoSerie || 60,
      item.descansoArmado || 30,
      item.prepSeg || 5,
      item.segPorRep || 2.8,
      item.fechaHora
    ]);
  }
}

function enviarMensajeTelegram(texto) {
  const url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";
  UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      chat_id: TELEGRAM_CHAT_ID,
      text: texto,
      parse_mode: "Markdown"
    })
  });
}
