/**
* @file Jangela.gs
* @description Functions for the Jangela (lunchroom) evaluation system.
*/

/**
 * Función para abrir el modal de Jangela (completamente independiente)
 */
function abrirJangela() {
  try {
    const html = HtmlService.createTemplateFromFile('_modal_jangela');
    const htmlOutput = html.evaluate()
      .setWidth(1300)
      .setHeight(750)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, '🏫 Jangela - Sistema Ebaluazioa');

  } catch (error) {
    console.error('Errorea jangela modala irekitzean:', error);
    SpreadsheetApp.getUi().alert('Errorea  Jangela irekitzen: ' + error.message);
  }
}

/**
 * FUNCIONES PARA JANGELA (COMEDOR)
 */

/**
 * Obtener datos de Jangela para una evaluación específica (HOJA DEL ALUMNO ACTIVO)
 */
function obtenerDatosJangela(evaluacion) {
  try {
    const config = obtenerConfiguracionCompleta();
    const alumnoActivo = obtenerAlumnoActivo();
    if (!alumnoActivo) {
      return { tieneDatos: false };
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAlumno = spreadsheet.getSheetByName(alumnoActivo);

    if (!hojaAlumno) {
      return { tieneDatos: false };
    }

    let celdaJatea, celdaJarrera, celdaAutonomia;
    const jangelaConfig = config.globalSettings.jangela;

    if (evaluacion === 1) {
      celdaJatea = jangelaConfig.eval1.jatea;
      celdaJarrera = jangelaConfig.eval1.jarrera;
      celdaAutonomia = jangelaConfig.eval1.autonomia;
    } else {
      celdaJatea = jangelaConfig.eval2.jatea;
      celdaJarrera = jangelaConfig.eval2.jarrera;
      celdaAutonomia = jangelaConfig.eval2.autonomia;
    }

    const jatea = hojaAlumno.getRange(celdaJatea).getValue();
    const jarrera = hojaAlumno.getRange(celdaJarrera).getValue();
    const autonomia = hojaAlumno.getRange(celdaAutonomia).getValue();

    const tieneDatos = jatea || jarrera || autonomia;

    return {
      tieneDatos: tieneDatos,
      jatea: jatea ? jatea.toString() : '',
      jarrera: jarrera ? jarrera.toString() : '',
      autonomia: autonomia ? autonomia.toString() : ''
    };

  } catch (error) {
    console.error('Error obteniendo datos Jangela:', error);
    return { tieneDatos: false };
  }
}

/**
 * Insertar evaluación de Jangela (EN LA HOJA DEL ALUMNO ACTIVO)
 */
function insertarEvaluacionJangela(evaluacionData, numeroEvaluacion) {
  try {
    const config = obtenerConfiguracionCompleta();
    const alumnoActivo = obtenerAlumnoActivo();
    if (!alumnoActivo) {
      throw new Error('No hay alumno activo seleccionado');
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAlumno = spreadsheet.getSheetByName(alumnoActivo);

    if (!hojaAlumno) {
      throw new Error(`No se encontró la hoja del alumno "${alumnoActivo}"`);
    }

    let celdaJatea, celdaJarrera, celdaAutonomia;
    const jangelaConfig = config.globalSettings.jangela;

    if (numeroEvaluacion === 1) {
      celdaJatea = jangelaConfig.eval1.jatea;
      celdaJarrera = jangelaConfig.eval1.jarrera;
      celdaAutonomia = jangelaConfig.eval1.autonomia;
    } else {
      celdaJatea = jangelaConfig.eval2.jatea;
      celdaJarrera = jangelaConfig.eval2.jarrera;
      celdaAutonomia = jangelaConfig.eval2.autonomia;
    }

    // Insertar valores solo si no están vacíos
    if (evaluacionData.jatea && evaluacionData.jatea.trim() !== '') {
      hojaAlumno.getRange(celdaJatea).setValue(evaluacionData.jatea);
    }

    if (evaluacionData.jarrera && evaluacionData.jarrera.trim() !== '') {
      hojaAlumno.getRange(celdaJarrera).setValue(evaluacionData.jarrera);
    }

    if (evaluacionData.autonomia && evaluacionData.autonomia.trim() !== '') {
      hojaAlumno.getRange(celdaAutonomia).setValue(evaluacionData.autonomia);
    }

    // Actualizar la celda de evaluación en C43
    if (numeroEvaluacion === 1) {
      hojaAlumno.getRange(config.globalSettings.celdaEvaluacion).setValue('Primera');
    } else {
      hojaAlumno.getRange(config.globalSettings.celdaEvaluacion).setValue('Segunda');
    }

    const criteriosInsertados = [
      evaluacionData.jatea ? 'JATEA' : null,
      evaluacionData.jarrera ? 'JARRERA' : null,
      evaluacionData.autonomia ? 'AUTONOMIA' : null
    ].filter(c => c !== null);

    return {
      success: true,
      message: `✅ ${numeroEvaluacion}ª evaluación Jangela de ${alumnoActivo}: ${criteriosInsertados.join(', ')}`
    };

  } catch (error) {
    console.error('Error insertando evaluación Jangela:', error);
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/**
 * MODAL JANGELA - Sistema de comedor
 */
function abrirModalJangela() {
  const html = HtmlService.createTemplateFromFile('_modal_jangela');
  const htmlOutput = html.evaluate()
    .setWidth(600)
    .setHeight(500)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, '🏫 Sistema Jangela');
}

function abrirJangelaParaAlumnoActivo() {
  try {
    const alumnoActivo = obtenerAlumnoActivo();
    if (!alumnoActivo) {
      SpreadsheetApp.getUi().alert('No hay ningún alumno seleccionado');
      return {
        success: false,
        message: 'No hay alumno seleccionado'
      };
    }

    // Cambiar a la hoja del alumno activo
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAlumno = spreadsheet.getSheetByName(alumnoActivo);
    if (hojaAlumno) {
      spreadsheet.setActiveSheet(hojaAlumno);
    }

    // Abrir modal de Jangela ORIGINAL
    const html = HtmlService.createTemplateFromFile('_modal_jangela');
    const htmlOutput = html.evaluate()
      .setWidth(600)
      .setHeight(500)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    SpreadsheetApp.getUi()
      .showModalDialog(htmlOutput, `🏫 Jangela - ${alumnoActivo}`);

    return {
      success: true,
      message: `Modal Jangela abierto para ${alumnoActivo}`
    };

  } catch (error) {
    console.error('Error abriendo Jangela:', error);
    return {
      success: false,
      message: 'Error abriendo modal Jangela: ' + error.message
    };
  }
}
