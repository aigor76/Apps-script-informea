
/**
* Función para obtener la configuración de ámbitos desde la Hoja 1
*/
function obtenerConfiguracionAmbitos() {
try {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaConfig = spreadsheet.getSheetByName('Hoja 1');
  if (!hojaConfig) {
    console.error('Ez da aurkitu "Hoja 1"');
    return [];
  }
  // Leer datos desde la fila 2 (asumiendo que la fila 1 tiene headers)
  const datos = hojaConfig.getDataRange().getValues();
  const headers = datos[0];
  // Buscar las columnas por nombre
  const colNombreHoja = headers.indexOf('Nombre de hoja');
  const colNumeroCriterios = headers.indexOf('Numero de criterios');
  const colCeldasOcupa = headers.indexOf('Celdas que ocupa');
  const colCelda1Eval = headers.indexOf('Celda 1. evaluacion');
  const colCelda2Eval = headers.indexOf('Celda 2.evaluacion');
  if (colNombreHoja === -1 || colNumeroCriterios === -1 || colCeldasOcupa === -1 ||
      colCelda1Eval === -1 || colCelda2Eval === -1) {
    throw new Error('No se encontraron todas las columnas necesarias en la Hoja 1');
  }
  const configuraciones = [];
  // Procesar cada fila de configuración (saltando la primera que son headers)
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
 
    const nombreHoja = fila[colNombreHoja];
    const numeroCriterios = fila[colNumeroCriterios];
    const celdasOcupa = fila[colCeldasOcupa];
    const celda1Eval = fila[colCelda1Eval];
    const celda2Eval = fila[colCelda2Eval];
 
    // Solo procesar filas que tengan datos válidos
    if (nombreHoja && numeroCriterios && celdasOcupa && celda1Eval && celda2Eval) {
      configuraciones.push({
        nombreHoja: nombreHoja.toString().trim(),
        numeroCriterios: parseInt(numeroCriterios),
        celdasOcupa: celdasOcupa.toString().trim(),
        celda1Eval: celda1Eval.toString().trim(),
        celda2Eval: celda2Eval.toString().trim()
      });
    }
  }
  return configuraciones;
} catch (error) {
  console.error('Konfigurazioa lortzean errorea:', error);
  return [];
}
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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTXOSTENA = spreadsheet.getSheetByName('TXOSTENA');
  if (!hojaTXOSTENA) {
    return { tieneDatos: false, nota: '', criterios: [] };
  }
  // LEER NOTA
  let nota = '';
  try {
    const valorNota = hojaTXOSTENA.getRange(configuracion.celda1Eval).getValue();
    nota = valorNota ? valorNota.toString().trim() : '';
  } catch (errorNota) {
    console.log('Kalifikazioa irakurtzean errorea:', errorNota);
  }
  // LEER CRITERIOS
  const criterios = [];
  try {
    const rango = configuracion.celdasOcupa;
  
    // Parsing del rango N11:N14
    const dosPuntos = rango.indexOf(':');
    const inicio = rango.substring(0, dosPuntos);
    const final = rango.substring(dosPuntos + 1);
  
    // Extraer columna
    let columna = '';
    for (let i = 0; i < inicio.length; i++) {
      const caracter = inicio.charAt(i);
      if (caracter >= 'A' && caracter <= 'Z') {
        columna += caracter;
      } else {
        break;
      }
    }
  
    // Extraer números de fila
    const filaInicio = parseInt(inicio.replace(/[A-Z]/g, ''));
    const filaFinal = parseInt(final.replace(/[A-Z]/g, ''));
  
    // LEER TODAS LAS CELDAS DEL RANGO
    for (let fila = filaInicio; fila <= filaFinal; fila++) {
      const celda = columna + fila;
      try {
        const valor = hojaTXOSTENA.getRange(celda).getValue();
        const texto = valor ? valor.toString().trim() : '';
      
        if (texto && texto.length > 0) {
          const textoCorto = texto.length > 50 ? texto.substring(0, 50) + '...' : texto;
          criterios.push(textoCorto);
        }
      } catch (errorCelda) {
        console.log('Gelaxka irakurtzea errorea', celda, ':', errorCelda);
      }
    }
  
  } catch (errorCriterios) {
    console.log('Itemak lortzean errorea:', errorCriterios);
  }
  // RESULTADO
  const tieneDatos = (nota !== '') || (criterios.length > 0);
  const resultado = {
    tieneDatos: tieneDatos,
    nota: nota,
    criterios: criterios
  };
   return resultado;
} catch (error) {
  console.error('ERRORE OROKORRA:', error);
  return { tieneDatos: false, nota: '', criterios: [] };
}
}


/**
* Función para insertar criterios
*/
function insertarCriterios(criteriosSeleccionados, notaEvaluacion, numeroEvaluacion, configuracion) {
try {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTXOSTENA = spreadsheet.getSheetByName('TXOSTENA');
  if (!hojaTXOSTENA) {
    throw new Error('EZ DA AURKITU "TXOSTENA" ORRIA');
  }
  if (!notaEvaluacion || notaEvaluacion.trim() === '') {
    throw new Error(`Kalifikazio bat aukeratu behar duzu ${numeroEvaluacion}ª ebaluaziorako`);
  }
  const rango = configuracion.celdasOcupa;
  if (numeroEvaluacion === 1) {
    // PRIMERA EVALUACIÓN
    hojaTXOSTENA.getRange(rango).clearContent();
    hojaTXOSTENA.getRange(configuracion.celda1Eval).clearContent();
 
    // Parsing del rango
    const dosPuntos = rango.indexOf(':');
    const inicio = rango.substring(0, dosPuntos);
    const final = rango.substring(dosPuntos + 1);
  
    // Extraer columna
    let columna = '';
    for (let i = 0; i < inicio.length; i++) {
      const caracter = inicio.charAt(i);
      if (caracter >= 'A' && caracter <= 'Z') {
        columna += caracter;
      } else {
        break;
      }
    }
  
    // Extraer números de fila
    const filaInicio = parseInt(inicio.replace(/[A-Z]/g, ''));
    const filaFinal = parseInt(final.replace(/[A-Z]/g, ''));
 
    // Insertar criterios
    let filaActual = filaInicio;
    for (let i = 0; i < criteriosSeleccionados.length && filaActual <= filaFinal; i++) {
      const criterio = criteriosSeleccionados[i];
      const celda = columna + filaActual;
      hojaTXOSTENA.getRange(celda).setValue(criterio.texto);
      filaActual++;
    }
 
    hojaTXOSTENA.getRange(configuracion.celda1Eval).setValue(notaEvaluacion);
 
    return {
      success: true,
      message: `Lehenengo ebaluazioa: ${criteriosSeleccionados.length} criterios en ${rango}`
    };
 
  } else if (numeroEvaluacion === 2) {
    // SEGUNDA EVALUACIÓN
  
    // Extraer primera columna del rango
    let primeraColumna = '';
    for (let i = 0; i < rango.length; i++) {
      const caracter = rango.charAt(i);
      if (caracter >= 'A' && caracter <= 'Z') {
        primeraColumna += caracter;
      } else {
        break;
      }
    }
  
    // Segunda columna será la siguiente
    const segundaColumna = String.fromCharCode(primeraColumna.charCodeAt(0) + 1);
 
    // Parsing del rango para obtener filas
    const dosPuntos = rango.indexOf(':');
    const inicio = rango.substring(0, dosPuntos);
    const final = rango.substring(dosPuntos + 1);
  
    const filaInicio = parseInt(inicio.replace(/[A-Z]/g, ''));
    const filaFinal = parseInt(final.replace(/[A-Z]/g, ''));
  
    const rangoSegunda = segundaColumna + filaInicio + ':' + segundaColumna + filaFinal;
 
    hojaTXOSTENA.getRange(rangoSegunda).clearContent();
    hojaTXOSTENA.getRange(configuracion.celda2Eval).clearContent();
 
    // Insertar criterios
    let filaActual = filaInicio;
    for (let i = 0; i < criteriosSeleccionados.length && filaActual <= filaFinal; i++) {
      const criterio = criteriosSeleccionados[i];
      const celda = segundaColumna + filaActual;
      hojaTXOSTENA.getRange(celda).setValue(criterio.texto);
      filaActual++;
    }
 
    hojaTXOSTENA.getRange(configuracion.celda2Eval).setValue(notaEvaluacion);
 
    return {
      success: true,
      message: `Bigarren ebaluazioa: ${criteriosSeleccionados.length} criterios en ${rangoSegunda}`
    };
  }
} catch (error) {
  console.error('Itemak txertatzean errorea :', error);
  return {
    success: false,
    message: 'Error: ' + error.message
  };
}
}

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
    
    if (evaluacion === 1) {
      // Primera evaluación
      celdaJatea = 'AQ39';
      celdaJarrera = 'AQ42';
      celdaAutonomia = 'AQ45';
    } else {
      // Segunda evaluación
      celdaJatea = 'AS39';
      celdaJarrera = 'AS42';  
      celdaAutonomia = 'AS45';
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
    
    if (numeroEvaluacion === 1) {
      // Primera evaluación
      celdaJatea = 'AQ39';
      celdaJarrera = 'AQ42';
      celdaAutonomia = 'AQ45';
    } else {
      // Segunda evaluación  
      celdaJatea = 'AS39';
      celdaJarrera = 'AS42';
      celdaAutonomia = 'AS45';
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
      hojaAlumno.getRange('C43').setValue('Primera');
    } else {
      hojaAlumno.getRange('C43').setValue('Segunda');
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
 * FUNCIONES DE COMPATIBILIDAD - Para que el _modal.html existente funcione
 * Estas funciones mantienen los nombres originales pero usan la hoja del alumno activo
 */

/**
 * Función para obtener datos de primera evaluación (COMPATIBLE con modal original)
 */
function obtenerDatosPrimera_ORIGINAL(configuracion) {
  // Redirigir a la nueva función que usa alumno activo
  return obtenerDatosPrimera(configuracion);
}

/**
 * Función para insertar criterios (COMPATIBLE con modal original)
 */
function insertarCriterios_ORIGINAL(criteriosSeleccionados, notaEvaluacion, numeroEvaluacion, configuracion) {
  // Redirigir a la nueva función que usa alumno activo
  return insertarCriterios(criteriosSeleccionados, notaEvaluacion, numeroEvaluacion, configuracion);
}

/**
 * SOBRESCRIBIR las funciones originales para que usen alumno activo
 */

// Esta es la función que el modal _modal.html llama para insertar
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
    
    const rango = configuracion.celdasOcupa;
    console.log('🔍 INSERTANDO EN RANGO:', rango, 'EVALUACION:', numeroEvaluacion);
    
    if (numeroEvaluacion === 1) {
      // PRIMERA EVALUACIÓN
      hojaAlumno.getRange(rango).clearContent();
      hojaAlumno.getRange(configuracion.celda1Eval).clearContent();
      
      const dosPuntos = rango.indexOf(':');
      const inicio = rango.substring(0, dosPuntos);
      const final = rango.substring(dosPuntos + 1);
      
      let columna = '';
      for (let i = 0; i < inicio.length; i++) {
        const caracter = inicio.charAt(i);
        if (caracter >= 'A' && caracter <= 'Z') {
          columna += caracter;
        } else {
          break;
        }
      }
      
      const filaInicio = parseInt(inicio.replace(/[A-Z]/g, ''));
      const filaFinal = parseInt(final.replace(/[A-Z]/g, ''));
      
      let filaActual = filaInicio;
      for (let i = 0; i < criteriosSeleccionados.length && filaActual <= filaFinal; i++) {
        const criterio = criteriosSeleccionados[i];
        const celda = columna + filaActual;
        hojaAlumno.getRange(celda).setValue(criterio.texto);
        console.log('✅ INSERTADO EN:', celda, '=', criterio.texto);
        filaActual++;
      }
      
      hojaAlumno.getRange(configuracion.celda1Eval).setValue(notaEvaluacion);
      console.log('✅ NOTA INSERTADA EN:', configuracion.celda1Eval, '=', notaEvaluacion);
      
      return {
        success: true,
        message: 'Primera evaluacion de ' + alumnoActivo + ': ' + criteriosSeleccionados.length + ' criterios guardados'
      };
      
    } else if (numeroEvaluacion === 2) {
      // SEGUNDA EVALUACIÓN
      let primeraColumna = '';
      for (let i = 0; i < rango.length; i++) {
        const caracter = rango.charAt(i);
        if (caracter >= 'A' && caracter <= 'Z') {
          primeraColumna += caracter;
        } else {
          break;
        }
      }
      
      const segundaColumna = String.fromCharCode(primeraColumna.charCodeAt(0) + 1);
      const dosPuntos = rango.indexOf(':');
      const inicio = rango.substring(0, dosPuntos);
      const final = rango.substring(dosPuntos + 1);
      const filaInicio = parseInt(inicio.replace(/[A-Z]/g, ''));
      const filaFinal = parseInt(final.replace(/[A-Z]/g, ''));
      const rangoSegunda = segundaColumna + filaInicio + ':' + segundaColumna + filaFinal;
      
      hojaAlumno.getRange(rangoSegunda).clearContent();
      hojaAlumno.getRange(configuracion.celda2Eval).clearContent();
      
      let filaActual = filaInicio;
      for (let i = 0; i < criteriosSeleccionados.length && filaActual <= filaFinal; i++) {
        const criterio = criteriosSeleccionados[i];
        const celda = segundaColumna + filaActual;
        hojaAlumno.getRange(celda).setValue(criterio.texto);
        console.log('✅ INSERTADO EN:', celda, '=', criterio.texto);
        filaActual++;
      }
      
      hojaAlumno.getRange(configuracion.celda2Eval).setValue(notaEvaluacion);
      console.log('✅ NOTA INSERTADA EN:', configuracion.celda2Eval, '=', notaEvaluacion);
      
      return {
        success: true,
        message: 'Segunda evaluacion de ' + alumnoActivo + ': ' + criteriosSeleccionados.length + ' criterios guardados'
      };
    }
  } catch (error) {
    console.error('❌ ERROR AL INSERTAR:', error);
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

// También sobrescribir la función de obtener datos
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
      const rango = configuracion.celdasOcupa;
      const dosPuntos = rango.indexOf(':');
      const inicio = rango.substring(0, dosPuntos);
      const final = rango.substring(dosPuntos + 1);
      
      let columna = '';
      for (let i = 0; i < inicio.length; i++) {
        const caracter = inicio.charAt(i);
        if (caracter >= 'A' && caracter <= 'Z') {
          columna += caracter;
        } else {
          break;
        }
      }
      
      const filaInicio = parseInt(inicio.replace(/[A-Z]/g, ''));
      const filaFinal = parseInt(final.replace(/[A-Z]/g, ''));
      
      for (let fila = filaInicio; fila <= filaFinal; fila++) {
        const celda = columna + fila;
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
function cerrarModalCriteriosYVolverSelector() {
  // Esta función se llamará desde el JavaScript del modal
  try {
    // Primero cerrar el modal actual
    google.script.host.close();
    
    // Luego abrir el selector
    setTimeout(() => {
      abrirSelectorAlumnos();
    }, 300);
    
  } catch (error) {
    console.error('Error cerrando modal y volviendo al selector:', error);
  }
}/**
 * SISTEMA DE TXOSTENAK - CÓDIGO PRINCIPAL
 * Gestión completa de informes individualizados por alumno
 */

// ID del documento público de Google Sheets
const DOCUMENTO_PUBLICO_ID = '1w8VORB8b9mYv0UWJ2aIsiCRJTKJo71ReNHjX4laEx0A';

/**
 * Crear menú personalizado al abrir el documento
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 Sistema Txostenak')
    .addItem('⚙️ Configuración inicial', 'abrirConfiguracionInicial')
    .addItem('📚 Selector de alumnos', 'abrirSelectorAlumnos')
    .addItem('🏫 Comedor (Jangela)', 'abrirModalJangela')
    .addToUi();
}

/**
 * CONFIGURACIÓN INICIAL - Modal para configurar el sistema
 */
function abrirConfiguracionInicial() {
  const html = HtmlService.createTemplateFromFile('_configuracion_inicial');
  const htmlOutput = html.evaluate()
    .setWidth(800)
    .setHeight(600)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, '⚙️ Configuración del Sistema Txostenak');
}

/**
 * SELECTOR DE ALUMNOS - Modal principal para elegir alumno
 */
function abrirSelectorAlumnos() {
  const html = HtmlService.createTemplateFromFile('_selector_alumnos');
  const htmlOutput = html.evaluate()
    .setWidth(900)
    .setHeight(700)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, '📚 Selector de Alumnos');
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

/**
 * Obtener clases disponibles del documento público
 */
function obtenerClasesDisponibles() {
  try {
    const documentoPublico = SpreadsheetApp.openById(DOCUMENTO_PUBLICO_ID);
    const sheets = documentoPublico.getSheets();
    const clases = [];
    
    // Mostrar TODAS las hojas del documento público
    sheets.forEach(sheet => {
      const nombre = sheet.getName();
      clases.push({
        nombre: nombre,
        activa: true
      });
    });
    
    // Ordenar alfabéticamente
    clases.sort((a, b) => a.nombre.localeCompare(b.nombre));
    
    return {
      success: true,
      clases: clases,
      totalHojas: clases.length
    };
    
  } catch (error) {
    console.error('Error obteniendo clases del documento público:', error);
    return {
      success: false,
      message: 'Error al conectar con el documento público: ' + error.message
    };
  }
}

/**
 * Obtener alumnos de una clase específica del documento público
 */
function obtenerAlumnosClase(nombreClase) {
  try {
    const documentoPublico = SpreadsheetApp.openById(DOCUMENTO_PUBLICO_ID);
    const hoja = documentoPublico.getSheetByName(nombreClase);
    
    if (!hoja) {
      throw new Error(`No se encontró la clase "${nombreClase}" en el documento público`);
    }
    
    // Obtener datos desde la columna B (Izen abizenak), fila 2 hasta el final
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return {
        success: true,
        alumnos: []
      };
    }
    
    const datos = hoja.getRange(2, 2, ultimaFila - 1, 1).getValues(); // Columna B
    const alumnosRaw = [];
    
    // Procesar nombres y eliminar duplicados
    datos.forEach(fila => {
      const nombre = fila[0];
      if (nombre && nombre.toString().trim() !== '') {
        alumnosRaw.push(nombre.toString().trim());
      }
    });
    
    // Limpiar nombres duplicados (quitar "aita" y "ama")
    const alumnosLimpios = limpiarNombresDuplicados(alumnosRaw);
    
    return {
      success: true,
      alumnos: alumnosLimpios,
      duplicadosEliminados: alumnosRaw.length - alumnosLimpios.length
    };
    
  } catch (error) {
    console.error('Error obteniendo alumnos del documento público:', error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Limpiar nombres duplicados quitando "aita" y "ama"
 */
function limpiarNombresDuplicados(nombres) {
  const nombresLimpios = new Set();
  const nombresFinales = [];
  
  nombres.forEach(nombre => {
    let nombreLimpio = nombre;
    
    // Quitar "aita" o "ama" del final
    if (nombre.toLowerCase().endsWith(' aita')) {
      nombreLimpio = nombre.substring(0, nombre.length - 5).trim();
    } else if (nombre.toLowerCase().endsWith(' ama')) {
      nombreLimpio = nombre.substring(0, nombre.length - 4).trim();
    }
    
    // Solo agregar si no existe ya
    if (!nombresLimpios.has(nombreLimpio)) {
      nombresLimpios.add(nombreLimpio);
      nombresFinales.push(nombreLimpio);
    }
  });
  
  return nombresFinales;
}

/**
 * Crear hojas individuales para cada alumno
 */
function crearHojasAlumnos(nombreClase, datosConfiguracion) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaPlantilla = spreadsheet.getSheetByName('TXOSTENA');
    
    if (!hojaPlantilla) {
      throw new Error('No se encontró la hoja plantilla "TXOSTENA"');
    }
    
    // Obtener alumnos de la clase
    const resultadoAlumnos = obtenerAlumnosClase(nombreClase);
    if (!resultadoAlumnos.success) {
      throw new Error(resultadoAlumnos.message);
    }
    
    const alumnos = resultadoAlumnos.alumnos;
    let hojasCreadas = 0;
    let hojasExistentes = 0;
    const errores = [];
    
    // Crear una hoja por cada alumno
    alumnos.forEach(alumno => {
      try {
        const nombreHoja = alumno;
        
        // Verificar si ya existe la hoja
        const hojaExistente = spreadsheet.getSheetByName(nombreHoja);
        if (hojaExistente) {
          hojasExistentes++;
          return;
        }
        
        // Crear nueva hoja copiando la plantilla
        const nuevaHoja = hojaPlantilla.copyTo(spreadsheet);
        nuevaHoja.setName(nombreHoja);
        
        // Configurar datos básicos del alumno
        configurarDatosBasicosAlumno(nuevaHoja, alumno, datosConfiguracion);
        
        hojasCreadas++;
        
      } catch (error) {
        errores.push(`Error con ${alumno}: ${error.message}`);
      }
    });
    
    return {
      success: true,
      hojasCreadas: hojasCreadas,
      hojasExistentes: hojasExistentes,
      errores: errores.length,
      detallesErrores: errores,
      totalAlumnos: alumnos.length,
      clase: nombreClase
    };
    
  } catch (error) {
    console.error('Error creando hojas:', error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Configurar datos básicos en la hoja del alumno
 */
function configurarDatosBasicosAlumno(hoja, nombreAlumno, datos) {
  try {
    // C37: Nombre y apellido del alumno
    hoja.getRange('C37').setValue(nombreAlumno);
    
    // C39: Nombre del tutor (viene de la configuración)
    hoja.getRange('C39').setValue(datos.nombreTutor || '');
    
    // C41: Curso escolar (viene de la configuración)
    hoja.getRange('C41').setValue(datos.cursoEscolar || '');
    
    // C43: Evaluación (viene de la configuración)
    hoja.getRange('C43').setValue(datos.evaluacion || 'Primera');
    
    // C45: Curso (viene de la configuración)
    hoja.getRange('C45').setValue(datos.curso || '');
    
  } catch (error) {
    console.error('Error configurando datos básicos:', error);
    throw error;
  }
}

/**
 * Obtener lista de alumnos con sus estados de progreso (SOLO HOJAS VISIBLES)
 */
function obtenerAlumnosConEstado() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const alumnos = [];
    
    // Buscar solo hojas VISIBLES que no sean las principales
    sheets.forEach(sheet => {
      const nombre = sheet.getName();
      
      // IGNORAR hojas ocultas
      if (sheet.isSheetHidden()) {
        console.log(`Ignorando hoja oculta: ${nombre}`);
        return;
      }
      
      // IGNORAR hojas del sistema
      if (nombre === 'TXOSTENA' || nombre === 'Hoja 1' || esHojaDelSistema(nombre)) {
        console.log(`Ignorando hoja del sistema: ${nombre}`);
        return;
      }
      
      console.log(`Procesando hoja de alumno: ${nombre}`);
      const estado = analizarEstadoAlumno(sheet);
      alumnos.push({
        nombre: nombre,
        estado: estado.estado,
        descripcion: estado.descripcion,
        color: estado.color,
        icono: estado.icono
      });
    });
    
    // Ordenar alfabéticamente
    alumnos.sort((a, b) => a.nombre.localeCompare(b.nombre));
    
    console.log(`Total alumnos encontrados (visibles): ${alumnos.length}`);
    
    return {
      success: true,
      alumnos: alumnos
    };
    
  } catch (error) {
    console.error('Error obteniendo alumnos con estado:', error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Verificar si una hoja pertenece al sistema (no es de alumno)
 */
function esHojaDelSistema(nombreHoja) {
  // Patrones de hojas del sistema a ignorar
  const patronesSistema = [
    /^LH[0-9]+$/i,           // LH3, LH4, LH5, etc.
    /criterios?/i,           // Cualquier cosa con "criterio" o "criterios"
    /^config/i,              // Configuración
    /^setup/i,               // Setup
    /^plantilla/i,           // Plantillas
    /^template/i,            // Templates
    /^sistema/i,             // Sistema
    /^admin/i,               // Admin
    /^_/,                    // Empiezan con guión bajo
    /^#/,                    // Empiezan con #
  ];
  
  return patronesSistema.some(patron => patron.test(nombreHoja));
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
 * Analizar el estado de progreso de un alumno (OPTIMIZADO - SOLO NOTAS)
 */
function analizarEstadoAlumno(hojaAlumno) {
  try {
    console.log(`\n🔍 ANALIZANDO ALUMNO: ${hojaAlumno.getName()}`);
    
    // Obtener configuración de ámbitos UNA sola vez
    const configuracion = obtenerConfiguracionAmbitos();
    if (!configuracion || configuracion.length === 0) {
      console.log('❌ No hay configuración de ámbitos');
      return {
        estado: 'sin_empezar',
        descripcion: 'Sin configuración',
        color: '#9aa0a6',
        icono: '⚠️'
      };
    }
    
    console.log(`📊 Total ámbitos configurados: ${configuracion.length}`);
    
    let ambitosCon1aEvaluacion = 0;
    let ambitosCon2aEvaluacion = 0;
    const totalAmbitos = configuracion.length;
    
    // Analizar ámbito por ámbito para mejor debugging
    for (let i = 0; i < configuracion.length; i++) {
      const config = configuracion[i];
      console.log(`\n📋 Ámbito ${i + 1}: ${config.nombreHoja}`);
      console.log(`   1ª eval celda: ${config.celda1Eval}`);
      console.log(`   2ª eval celda: ${config.celda2Eval}`);
      
      try {
        // Leer 1ª evaluación
        const valor1a = hojaAlumno.getRange(config.celda1Eval).getValue();
        const esta1aVacia = esCeldaVacia(valor1a);
        console.log(`   1ª eval valor: "${valor1a}" (tipo: ${typeof valor1a}) → Vacía: ${esta1aVacia}`);
        
        // Leer 2ª evaluación  
        const valor2a = hojaAlumno.getRange(config.celda2Eval).getValue();
        const esta2aVacia = esCeldaVacia(valor2a);
        console.log(`   2ª eval valor: "${valor2a}" (tipo: ${typeof valor2a}) → Vacía: ${esta2aVacia}`);
        
        // Contar solo si NO están vacías
        if (!esta1aVacia) {
          ambitosCon1aEvaluacion++;
          console.log(`   ✅ Sumando 1ª evaluación`);
        }
        
        if (!esta2aVacia) {
          ambitosCon2aEvaluacion++;
          console.log(`   ✅ Sumando 2ª evaluación`);
        }
        
      } catch (error) {
        console.log(`   ❌ Error leyendo ámbito ${i + 1}:`, error.message);
      }
    }
    
    console.log(`\n📊 RESUMEN ${hojaAlumno.getName()}:`);
    console.log(`   1ª evaluaciones completas: ${ambitosCon1aEvaluacion}/${totalAmbitos}`);
    console.log(`   2ª evaluaciones completas: ${ambitosCon2aEvaluacion}/${totalAmbitos}`);
    
    // Determinar estado con lógica MUY clara
    let estado, descripcion, color, icono;
    
    if (ambitosCon1aEvaluacion === 0 && ambitosCon2aEvaluacion === 0) {
      // CASO 1: Sin ninguna nota
      estado = 'sin_empezar';
      descripcion = 'Sin empezar';
      color = '#9aa0a6';
      icono = '⭕';
      console.log(`   🎯 ESTADO: SIN EMPEZAR (0 notas en total)`);
      
    } else if (ambitosCon2aEvaluacion === totalAmbitos) {
      // CASO 2: Todas las 2ª evaluaciones completas
      estado = 'terminado';
      descripcion = 'Ambas evaluaciones';
      color = '#34a853';
      icono = '✅';
      console.log(`   🎯 ESTADO: TERMINADO (todas las 2ª completas)`);
      
    } else if (ambitosCon1aEvaluacion === totalAmbitos) {
      // CASO 3: Todas las 1ª pero no todas las 2ª
      estado = 'parcial';
      descripcion = 'Solo 1ª evaluación';
      color = '#fbbc04';
      icono = '🟡';
      console.log(`   🎯 ESTADO: PARCIAL (todas las 1ª, no todas las 2ª)`);
      
    } else {
      // CASO 4: Algunas evaluaciones pero no todas las 1ª
      estado = 'parcial';
      descripcion = 'En progreso';
      color = '#fbbc04';
      icono = '🟡';
      console.log(`   🎯 ESTADO: EN PROGRESO (algunas notas pero incompleto)`);
    }
    
    return { estado, descripcion, color, icono };
    
  } catch (error) {
    console.error(`❌ Error analizando estado del alumno ${hojaAlumno.getName()}:`, error);
    return {
      estado: 'error',
      descripcion: 'Error al analizar',
      color: '#ea4335',
      icono: '❌'
    };
  }
}

/**
 * Verificar si los criterios están completos en un rango
 */
function verificarCriteriosCompletos(hoja, rangoStr, esSegundaEvaluacion = false) {
  try {
    let rango = rangoStr;
    
    // Si es segunda evaluación, ajustar a la siguiente columna
    if (esSegundaEvaluacion) {
      const dosPuntos = rangoStr.indexOf(':');
      const inicio = rangoStr.substring(0, dosPuntos);
      const final = rangoStr.substring(dosPuntos + 1);
      
      let columna = '';
      for (let i = 0; i < inicio.length; i++) {
        const caracter = inicio.charAt(i);
        if (caracter >= 'A' && caracter <= 'Z') {
          columna += caracter;
        } else {
          break;
        }
      }
      
      const siguienteColumna = String.fromCharCode(columna.charCodeAt(0) + 1);
      const filaInicio = parseInt(inicio.replace(/[A-Z]/g, ''));
      const filaFinal = parseInt(final.replace(/[A-Z]/g, ''));
      
      rango = `${siguienteColumna}${filaInicio}:${siguienteColumna}${filaFinal}`;
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

/**
 * Establecer alumno activo SIN abrir modal (para el botón comedor)
 */
function establecerAlumnoActivo(nombreAlumno) {
  try {
    // Guardar alumno activo en propiedades del script
    PropertiesService.getScriptProperties().setProperty('ALUMNO_ACTIVO', nombreAlumno);
    
    // Cambiar a la hoja del alumno
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAlumno = spreadsheet.getSheetByName(nombreAlumno);
    
    if (!hojaAlumno) {
      throw new Error(`No se encontró la hoja del alumno "${nombreAlumno}"`);
    }
    
    spreadsheet.setActiveSheet(hojaAlumno);
    
    return {
      success: true,
      message: `Alumno "${nombreAlumno}" establecido correctamente`
    };
    
  } catch (error) {
    console.error('Error estableciendo alumno activo:', error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * Obtener alumno activo actual
 */
function obtenerAlumnoActivo() {
  try {
    const alumnoActivo = PropertiesService.getScriptProperties().getProperty('ALUMNO_ACTIVO');
    return alumnoActivo || null;
  } catch (error) {
    console.error('Error obteniendo alumno activo:', error);
    return null;
  }
}

/**
 * Establecer alumno activo y abrir modal de evaluación (para el botón Evaluar)
 */
function establecerAlumnoActivoYAbrir(nombreAlumno) {
  try {
    // Primero establecer el alumno
    const resultado = establecerAlumnoActivo(nombreAlumno);
    
    if (!resultado.success) {
      return resultado;
    }
    
    // Luego abrir modal de evaluación
    abrirModalCriterios();
    
    return {
      success: true,
      message: `Alumno "${nombreAlumno}" establecido y modal abierto`
    };
    
  } catch (error) {
    console.error('Error estableciendo alumno activo y abriendo modal:', error);
    return {
      success: false,
      message: error.message
    };
  }
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
{Ui()
      .showModalDialog(htmlOutput, `🏫 Jangela - ${alumnoActivo}`);
    
    return {
      success: true,
      message: `Modal Jangela abierto para ${alumnoActivo}`
    };
}
}


// ============ FUNCIONES EXISTENTES ADAPTADAS ============

/**
 * Función principal para abrir el modal de selección (adaptada para alumno activo)
 */
function abrirModalCriterios() {
  // Verificar si hay un alumno activo
  const alumnoActivo = obtenerAlumnoActivo();
  if (!alumnoActivo) {
    SpreadsheetApp.getUi().alert('Selecciona un alumno primero desde el Selector de Alumnos');
    return;
  }
  
  // Obtener configuración de la Hoja 1
  const configuracion = obtenerConfiguracionAmbitos();
  if (!configuracion || configuracion.length === 0) {
    SpreadsheetApp.getUi().alert('No se encontró configuración válida en la Hoja 1');
    return;
  }

  const html = HtmlService.createTemplateFromFile('_modal');
  // Pasar la configuración al modal
  html.configuracion = JSON.stringify(configuracion);
  const htmlOutput = html.evaluate()
    .setWidth(1300)
    .setHeight(750)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'Evaluacion - ' + alumnoActivo);
}
