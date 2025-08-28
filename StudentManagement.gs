/**
* @file StudentManagement.gs
* @description Functions for managing students, classes, and their sheets.
*/

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
    const config = obtenerConfiguracionCompleta();
    const hojaPlantilla = spreadsheet.getSheetByName(config.globalSettings.plantillaTxostena);

    if (!hojaPlantilla) {
      throw new Error('No se encontró la hoja plantilla "' + config.globalSettings.plantillaTxostena + '"');
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
        configurarDatosBasicosAlumno(nuevaHoja, alumno, datosConfiguracion, config.globalSettings);

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
function configurarDatosBasicosAlumno(hoja, nombreAlumno, datos, globalSettings) {
  try {
    hoja.getRange(globalSettings.celdaNombreAlumno).setValue(nombreAlumno);
    hoja.getRange(globalSettings.celdaNombreTutor).setValue(datos.nombreTutor || '');
    hoja.getRange(globalSettings.celdaCursoEscolar).setValue(datos.cursoEscolar || '');
    hoja.getRange(globalSettings.celdaEvaluacion).setValue(datos.evaluacion || 'Primera');
    hoja.getRange(globalSettings.celdaCurso).setValue(datos.curso || '');

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
