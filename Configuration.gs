/**
* @file Configuration.gs
* @description Functions for reading configuration from the spreadsheet.
*/

/**
 * Reads the configuration from the 'Hoja 1' sheet.
 * This includes global settings and the list of evaluation scopes.
 * @returns {{globalSettings: object, scopes: Array<object>}}
 */
function obtenerConfiguracionCompleta() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaConfig = spreadsheet.getSheetByName('Hoja 1');
    if (!hojaConfig) {
      throw new Error('No se encontró la hoja de configuración "Hoja 1"');
    }

    const datos = hojaConfig.getDataRange().getValues();
    const headers = datos[0];

    if (datos.length < 2) {
      throw new Error('La hoja de configuración no tiene datos.');
    }

    // --- Validate Headers ---
    const requiredHeaders = [
      'plantilla_txostena', 'celda_nombre_alumno', 'celda_nombre_tutor',
      'celda_curso_escolar', 'celda_evaluacion', 'celda_curso',
      'celda_jangela_jatea_1', 'celda_jangela_jarrera_1', 'celda_jangela_autonomia_1',
      'celda_jangela_jatea_2', 'celda_jangela_jarrera_2', 'celda_jangela_autonomia_2',
      'Nombre de hoja', 'Numero de criterios', 'Celdas que ocupa',
      'Celda 1. evaluacion', 'Celda 2.evaluacion'
    ];
    const missingHeaders = requiredHeaders.filter(h => !headers.includes(h));
    if (missingHeaders.length > 0) {
      throw new Error(`Faltan las siguientes columnas en 'Hoja 1': ${missingHeaders.join(', ')}`);
    }

    // --- Leer Global Settings (from the first data row) ---
    const firstDataRow = datos[1];
    const globalSettings = {
      plantillaTxostena: firstDataRow[headers.indexOf('plantilla_txostena')],
      celdaNombreAlumno: firstDataRow[headers.indexOf('celda_nombre_alumno')],
      celdaNombreTutor: firstDataRow[headers.indexOf('celda_nombre_tutor')],
      celdaCursoEscolar: firstDataRow[headers.indexOf('celda_curso_escolar')],
      celdaEvaluacion: firstDataRow[headers.indexOf('celda_evaluacion')],
      celdaCurso: firstDataRow[headers.indexOf('celda_curso')],
      jangela: {
        eval1: {
          jatea: firstDataRow[headers.indexOf('celda_jangela_jatea_1')],
          jarrera: firstDataRow[headers.indexOf('celda_jangela_jarrera_1')],
          autonomia: firstDataRow[headers.indexOf('celda_jangela_autonomia_1')],
        },
        eval2: {
          jatea: firstDataRow[headers.indexOf('celda_jangela_jatea_2')],
          jarrera: firstDataRow[headers.indexOf('celda_jangela_jarrera_2')],
          autonomia: firstDataRow[headers.indexOf('celda_jangela_autonomia_2')],
        }
      }
    };

    // --- Leer Scopes (from all data rows) ---
    const colNombreHoja = headers.indexOf('Nombre de hoja');
    const colNumeroCriterios = headers.indexOf('Numero de criterios');
    const colCeldasOcupa = headers.indexOf('Celdas que ocupa');
    const colCelda1Eval = headers.indexOf('Celda 1. evaluacion');
    const colCelda2Eval = headers.indexOf('Celda 2.evaluacion');

    const scopes = [];
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const nombreHoja = fila[colNombreHoja];
      if (nombreHoja) { // Process only if the scope has a name
        scopes.push({
          nombreHoja: nombreHoja.toString().trim(),
          numeroCriterios: parseInt(fila[colNumeroCriterios]),
          celdasOcupa: fila[colCeldasOcupa].toString().trim(),
          celda1Eval: fila[colCelda1Eval].toString().trim(),
          celda2Eval: fila[colCelda2Eval].toString().trim()
        });
      }
    }

    return { globalSettings, scopes };

  } catch (error) {
    console.error('Error al obtener la configuración:', error);
    SpreadsheetApp.getUi().alert('Error al leer la configuración: ' + error.message);
    return { globalSettings: {}, scopes: [] };
  }
}

/**
* Función para obtener la configuración de ámbitos desde la Hoja 1
* @deprecated Use obtenerConfiguracionCompleta instead.
*/
function obtenerConfiguracionAmbitos() {
  return obtenerConfiguracionCompleta().scopes;
}
