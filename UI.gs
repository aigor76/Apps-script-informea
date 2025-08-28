/**
* @file UI.gs
* @description Functions for creating menus, opening modals, and other UI-related tasks.
*/

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
  const scopes = obtenerConfiguracionAmbitos();
  if (!scopes || scopes.length === 0) {
    SpreadsheetApp.getUi().alert('No se encontró configuración válida en la Hoja 1');
    return;
  }

  // Pre-load all criteria for all scopes
  const configuracionConCriterios = scopes.map(scope => {
    const criterios = cargarTodosLosCriterios(scope.nombreHoja, scope.numeroCriterios);
    return { ...scope, criterios: criterios };
  });

  const html = HtmlService.createTemplateFromFile('_modal');
  // Pasar la configuración con los criterios pre-cargados al modal
  html.configuracion = JSON.stringify(configuracionConCriterios);
  const htmlOutput = html.evaluate()
    .setWidth(1300)
    .setHeight(750)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'Evaluacion - ' + alumnoActivo);
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
}
