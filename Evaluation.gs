/**
* @file Evaluation.gs
* @description Functions for handling student evaluations.
*/

/**
 * Parses a spreadsheet range string (e.g., "N11:N14") into its components.
 * @param {string} rangeStr The range string to parse.
 * @returns {{column: string, startRow: number, endRow: number}}
 * @private
 */
function _parseRange(rangeStr) {
  const dosPuntos = rangeStr.indexOf(':');
  const inicio = rangeStr.substring(0, dosPuntos);
  const final = rangeStr.substring(dosPuntos + 1);

  const columna = inicio.replace(/[0-9]/g, '');
  const filaInicio = parseInt(inicio.replace(/[A-Z]/g, ''));
  const filaFinal = parseInt(final.replace(/[A-Z]/g, ''));

  return {
    column: columna,
    startRow: filaInicio,
    endRow: filaFinal,
  };
}

/**
* Función para cargar todos los criterios
*/
function cargarTodosLosCriterios(nombreHoja, numeroCriterios) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Ezin izan da spreadsheet-a lortu');
    }
    const hoja = spreadsheet.getSheetByName(nombreHoja);
    if (!hoja) {
      const hojasDisponibles = spreadsheet.getSheets().map(h => h.getName());
      throw new Error(`Ez da orria aurkitu "${nombreHoja}". Orriak: ${hojasDisponibles.join(', ')}`);
    }
    // Leer TODOS los criterios de la hoja
    const ultimaFila = hoja.getLastRow();
    const rangoTexto = `A2:C${ultimaFila}`;
    const rango = hoja.getRange(rangoTexto);
    const datos = rango.getValues();
    const criterios = [];
    // Procesar TODOS los datos
    for (let i = 0; i < datos.length; i++) {
      const celdaA = datos[i][0];
      const celdaB = datos[i][1];
      const celdaC = datos[i][2];

      if (celdaA && celdaA.toString().trim() !== '') {
        const criterio = {
          id: (criterios.length + 1).toString(),
          texto: celdaA.toString().trim(),
          textoCastellano: celdaB ? celdaB.toString().trim() : '',
          nivel: celdaC ? parseInt(celdaC) : 2
        };

        criterios.push(criterio);
      }
    }
    return criterios;
  } catch (error) {
    console.error(`Itemak kargatzean errorea:`, error);
    throw error;
  }
}


/**
* Función para obtener datos de primera evaluación
*/
function obtenerDatosPrimera(configuracion) {
  try {
    const alumnoActivo = obtenerAlumnoActivo();
    console.log('🔍 LEYENDO DATOS DE ALUMNO:', alumnoActivo);

    if (!alumnoActivo) {
      return { tieneDatos: false, nota: '', criterios: [] };
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAlumno = spreadsheet.getSheetByName(alumnoActivo);

    if (!hojaAlumno) {
      return { tieneDatos: false, nota: '', criterios: [] };
    }

    let nota = '';
    try {
      const valorNota = hojaAlumno.getRange(configuracion.celda1Eval).getValue();
      nota = valorNota ? valorNota.toString().trim() : '';
      console.log('🔍 NOTA LEIDA:', nota);
    } catch (errorNota) {
      console.log('Error leyendo nota:', errorNota);
    }

    const criterios = [];
    try {
      const rango = _parseRange(configuracion.celdasOcupa);

      for (let fila = rango.startRow; fila <= rango.endRow; fila++) {
        const celda = rango.column + fila;
        try {
          const valor = hojaAlumno.getRange(celda).getValue();
          const texto = valor ? valor.toString().trim() : '';

          if (texto && texto.length > 0) {
            const textoCorto = texto.length > 50 ? texto.substring(0, 50) + '...' : texto;
            criterios.push(textoCorto);
            console.log('🔍 CRITERIO LEIDO:', textoCorto);
          }
        } catch (errorCelda) {
          console.log('Error leyendo celda', celda, ':', errorCelda);
        }
      }

    } catch (errorCriterios) {
      console.log('Error general leyendo criterios:', errorCriterios);
    }

    const tieneDatos = (nota !== '') || (criterios.length > 0);
    const resultado = {
      tieneDatos: tieneDatos,
      nota: nota,
      criterios: criterios
    };

    console.log('🔍 RESULTADO LECTURA:', resultado);
    return resultado;
  } catch (error) {
    console.error('❌ ERROR GENERAL LEYENDO:', error);
    return { tieneDatos: false, nota: '', criterios: [] };
  }
}


/**
* Función para insertar criterios
*/
function insertarCriterios(criteriosSeleccionados, notaEvaluacion, numeroEvaluacion, configuracion) {
  try {
    const alumnoActivo = obtenerAlumnoActivo();
    console.log('🔍 INSERTANDO EN ALUMNO:', alumnoActivo);

    if (!alumnoActivo) {
      throw new Error('No hay alumno activo seleccionado. Usa el Selector de Alumnos primero.');
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAlumno = spreadsheet.getSheetByName(alumnoActivo);

    console.log('🔍 HOJA DEL ALUMNO ENCONTRADA:', !!hojaAlumno);

    if (!hojaAlumno) {
      throw new Error('No se encontro la hoja del alumno: ' + alumnoActivo);
    }

    if (!notaEvaluacion || notaEvaluacion.trim() === '') {
      throw new Error('Debe seleccionar una nota para la evaluacion ' + numeroEvaluacion);
    }

    const parsedRange = _parseRange(configuracion.celdasOcupa);
    let targetColumn = parsedRange.column;
    let notaCell = configuracion.celda1Eval;
    let messageClarification = "Primera";

    if (numeroEvaluacion === 2) {
      targetColumn = String.fromCharCode(targetColumn.charCodeAt(0) + 1);
      notaCell = configuracion.celda2Eval;
      messageClarification = "Segunda";
    }

    const criteriaRangeStr = `${targetColumn}${parsedRange.startRow}:${targetColumn}${parsedRange.endRow}`;

    hojaAlumno.getRange(criteriaRangeStr).clearContent();
    hojaAlumno.getRange(notaCell).clearContent();

    let filaActual = parsedRange.startRow;
    for (let i = 0; i < criteriosSeleccionados.length && filaActual <= parsedRange.endRow; i++) {
      const criterio = criteriosSeleccionados[i];
      const celda = targetColumn + filaActual;
      hojaAlumno.getRange(celda).setValue(criterio.texto);
      console.log('✅ INSERTADO EN:', celda, '=', criterio.texto);
      filaActual++;
    }

    hojaAlumno.getRange(notaCell).setValue(notaEvaluacion);
    console.log('✅ NOTA INSERTADA EN:', notaCell, '=', notaEvaluacion);

    return {
      success: true,
      message: `${messageClarification} evaluacion de ${alumnoActivo}: ${criteriosSeleccionados.length} criterios guardados`
    };

  } catch (error) {
    console.error('❌ ERROR AL INSERTAR:', error);
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/**
 * Verificar si una celda está realmente vacía (función auxiliar)
 */
function esCeldaVacia(valor) {
  // Verificación exhaustiva de valores vacíos
  if (valor === null || valor === undefined) return true;
  if (valor === '') return true;
  if (typeof valor === 'string' && valor.trim() === '') return true;
  if (typeof valor === 'number' && valor === 0) return true;

  return false;
}

/**
 * Verificar si los criterios están completos en un rango
 */
function verificarCriteriosCompletos(hoja, rangoStr, esSegundaEvaluacion = false) {
  try {
    let rango = rangoStr;

    // Si es segunda evaluación, ajustar a la siguiente columna
    if (esSegundaEvaluacion) {
      const parsedRange = _parseRange(rangoStr);
      const siguienteColumna = String.fromCharCode(parsedRange.column.charCodeAt(0) + 1);
      rango = `${siguienteColumna}${parsedRange.startRow}:${siguienteColumna}${parsedRange.endRow}`;
    }

    const valores = hoja.getRange(rango).getValues();

    // Verificar si hay al menos un criterio completado
    for (let fila of valores) {
      if (fila[0] && fila[0].toString().trim() !== '') {
        return true;
      }
    }

    return false;

  } catch (error) {
    console.error('Error verificando criterios:', error);
    return false;
  }
}
