// Parámetros de control de tiempo de ejecución
const MILISECONDS_FOR_ONE_MINUTE = 60000;
const GOOGLE_SCRIPTS_MAX_RUNTIME = 6 * MILISECONDS_FOR_ONE_MINUTE // 6 minutos

const SAFETY_MARGIN = 0.25; // 25% de margen sobre el máximo histórico
const ABSOLUTE_MIN_TIME = 30000; // 30s mínimo por archivo (para primeras ejecuciones)
const MAX_RUN_TIME = 5.8 * MILISECONDS_FOR_ONE_MINUTE; // 5.8 minutos (348s)
const REPAIR_THRESHOLD = 15 * MILISECONDS_FOR_ONE_MINUTE; // 15 minutos
const SLEEP_TIME = 2000; // 2 segundos
const ABORTING_EXECUTION_MESSAGE = "El tiempo de ejecución máximo está a punto de alcanzarse. Abortando la operación.";

// Parámetros de procesamiento de archivos
const LOG_TITLES = ['Fecha', 'Archivo', 'Filas', 'Estado'];
const LOG_FILENAME_COL = 1;
const LOG_STATE_COL = 3;
const PROCESSED_FOLDER_NAME = 'procesados';
const BACKUP_FOLDER_NAME = 'respaldos';
const SUCCESS_MESSAGE = "Éxito";
const PROCESSING_FAILURE_MESSAGE = "Error procesando archivo";
const SEARCHING_ERROR_MESSAGE = "Error obteniendo los archivos";
const BACKUP_IN_DESTINATION = false;
const KEEP_PROCESSED_IN_SOURCE = false;
const UPDATE_EXISTING_ROWS = false;
const IGNORE_EMPTY_ROWS = true;
const ALLOW_FILE_REPROCESSING = false;

/**
 * @typedef {'string' | 'float' | 'integer' | 'boolean' | 'date'} CellValueType
 */

/**
 * @typedef {Object} Normalizers
 * @property {function(any): any=} string
 * @property {function(any): any=} float
 * @property {function(any): any=} integer
 * @property {function(any): any=} boolean
 * @property {function(any): any=} date
 */

/**
 * @type {Normalizers}
 */
const normalizers = {
  string: normalizeString,
  float: normalizeFloat,
  integer: normalizeInteger,
  boolean: normalizeBoolean,
  date: normalizeDate,
};

/**
 * Función de procesamiento general.
 *
 * @callback ProcessorFunction
 * @param {GoogleAppsScript.DriveApp.File} file Archivo a procesar.
 * @param {ProcessingConfig} config Objeto de configuración de procesamiento.
 * @param {number} keyColumnIndex Índice en base 0 de la columna llave.
 * @param {CellValueType} keyColumnType Tipo de la columna llave.
 * @param {any[]} existingKeys Arreglo de valores de llaves existentes en destino. 
 * @returns {{ newRows: any[][], updatedRows: { key: any, row: any[] }[] }} Objeto que contiene las filas a anexar o actualizar.
 */

/**
 * Mapa de procesadores disponibles
 * @typedef {Map<string, ProcessorFunction>} ProcessorsMap Mapa de funciones de procesamiento general por MimeType.
 */

/**
 * @type {ProcessorsMap}
 */
const processors = new Map(); 

processors.set(MimeType.MICROSOFT_EXCEL, processXlsxFile);
processors.set(MimeType.CSV, processCsvFile);
processors.set(MimeType.GOOGLE_SHEETS, processGoogleSheetFile);

/**
 * Función de búsqueda de los archivos a procesar.
 *
 * @callback FileSearchingFunction
 * @property {(string)[]} mimeType Valor del MimeType de los archivos a buscar.
 * @returns {GoogleAppsScript.DriveApp.File[]} files Archivos a procesar.
 */

/**
 * Función de procesamiento personalizado.
 *
 * @callback FileProcessingFunction
 * @property {any[][]} rows Arreglo de líneas de datos a procesar.
 * @returns {any[][]} Arreglo de líneas de datos procesados.
 */

/**
 * Objeto de configuración del procesamiento personalizado para un tipo de archivo en particular.
 *
 * @typedef {Object} ProcessingConfig
 * @property {(number|string)[]} columns Índices de columnas o fórmulas a utilizar.
 * @property {FileSearchingFunction} searchingFn Función que devuelve un arreglo de objetos de los archivos a procesar.
 * @property {FileProcessingFunction} [processingFn] Función que realiza el procesamiento personalizado común a cada archivo.
 * @property {string} [sourceSheetName] Nombre de la hoja a procesar del archivo fuente. Se ignora en CSVs.
 * @property {boolean} [updateExistingRows] Si las filas de datos existentes en el destino se actualizan. De forma predeterminada sólo se insertan filas nuevas. 
 * @property {boolean} [ignoreEmptyRows] Si se ignoran las filas vacías en el archivo fuente. De forma predeterminada no se copian al archivo destino. 
 * @property {boolean} [allowFileReprocessing] Si se permite que un mismo archivo pueda procesarse más de una vez. De forma predeterminada no se permite.
 * @property {SortCriterion[]} [sortCriteria] Arreglo de criterios de ordenamiento para el archivo de destino.
 */

/**
 * Mapa de configuraciones de procesamiento personalizado.
 * 
 * @typedef {Map} ProcessingConfigMap
 * @param {Map<string, ProcessingConfig>} configs Mapa de configuración de procesamiento por mime type.
 */

/**
 * Objeto de opciones de procesamiento.
 * 
 * @typedef {Object} ProcessingOptions
 * @property {GoogleAppsScript.Spreadsheet} logSpreadsheet Archivo de log.
 * @property {string} logSheetName Nombre de la hoja a utilizar en el archivo de log.
 * @property {string[]} logColumnTitles Títulos de las columnas del log.
 * @property {string} successMessage Mensaje que se registra o muestra en el log al completar con éxito.
 * @property {string} failureMessage Mensaje que se registra o muestra en el log al completar con fallo.
 * @property {Folder} processingFolderID ID de la carpeta que contiene los archivos a procesar.
 * @property {string} processedFolderName Nombre de la subcarpeta donde mover archivos procesados.
 * @property {string} backupFolderName Nombre de la subcarpeta donde guardar respaldos.
 * @property {boolean} backupInDestination Si la carpeta de respaldos se crea en la misma carpeta del archivo destino. De forma predeterminada se crea en la carpeta de procesamiento. 
 * @property {boolean} keepProcessedInSource Si los archivos procesados se mantienen en la carpeta fuente. De forma predeterminada se mueven hacia la subcarpeta de procesados de la carpeta de procesamiento.
 */

/**
 * Procesa los archivos de una carpeta y copia sus datos a una hoja de cálculo de destino.
 *
 * @param {string} targetFileID ID del archivo de destino.
 * @param {string} targetFileSheetName Nombre de la hoja de destino.
 * @param {string[]} targetColumnTitles Títulos de columnas a insertar si la hoja está vacía.
 * @param {number} targetKeyColumnIndex Índice de la columna llave única.
 * @param {CellValueType} targetKeyColumnType Tipo de datos de la columna llave única.
 * @param {ProcessingConfigMap} configs Mapa de configuración de procesamiento por mime type.
 * @param {ProcessingOptions} options Opciones de procesamiento.
 */
function processFolderFilesAndCopyTo(
  targetFileID, targetFileSheetName, targetColumnTitles, targetKeyColumnIndex, targetKeyColumnType, configs, options) {

  // --- Inicialización segura
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  const startTime = new Date();
  const executionId = Utilities.getUuid();
  // ---

  try {
    // --- Gestión de estado e inicio seguro
    if (!validateExecutionState(props)) return;

    const timeStats = safeStartup(startTime, executionId, lock, props);

    if (!timeStats) return;

    let processedCount = 0;
    // ---

    // Destination file
    let spreadsheet = null; let targetSheet = null; let targetFile = null;

    try {
      const { 
        spreadsheet, 
        sheet, 
        file } = getSpreadsheetObjectsWithFallback(targetFileID, targetFileSheetName);
      targetSheet = sheet;
      targetFile = file;
    } catch (error) {
      throw new Error(`Error en el archivo de destino: ${error.message}`)
    }

    // Log file
    const { 
      logSheet, 
      logFile, 
      logColumnTitles, 
      successMessage, 
      failureMessage } = getLogOptionsWithFallback(options);

    // Folders
    const { 
      processingFolder, 
      processedFolder, 
      backupFolder, 
      keepProcessedInSource } = getFoldersOptionsWithFallback(options, targetFile);

    /**
     * Obtener los archivos ya procesados con éxito desde el archivo de log.
     */
    function getAlreadyProcessedFiles() {
      return getUniqueValuesFromColumn(
        logSheet, 
        1, 
        {
          filterRowsFn: (row, index) => {
            const name = row[LOG_FILENAME_COL];
            const state = row[LOG_STATE_COL];
            return (index > 0 && name && state === successMessage);
          }
        }) || new Set();
    }

    const alreadyProcessedFiles = getAlreadyProcessedFiles();

    /**
     * Obtener las claves de las filas de la hoja destino.
     */
    function getAllKeysFromTargetSheet() {
      return getUniqueValuesFromColumn(
        targetSheet, 
        targetKeyColumnIndex,
        {
          formatCellValueFn: (cell) => normalizers[targetKeyColumnType](cell)
        }) || new Set();
    }

    /**
     * Procesa los archivos de un *MimeType*, siguiendo la configuración de procesamiento indicado.
     *
     * @param {string} mimeType *MimeType* de los archivos a procesar.
     * @param {ProcessingConfig} config Objeto de configuración del procesamiento.
     */
    function processFiles(mimeType, config) {
      let files = [];

      if (!config.searchingFn()) {
        console.error(`No se encontró la función de búsqueda de archivos para el tipo MIME: ${mimeType}`);
        return;
      }          

      try {
        files = config.searchingFn(mimeType);
      } catch (e) {
        console.error(`${SEARCHING_ERROR_MESSAGE}: ${e.message}`);
        return;
      }

      if (files.length === 0) {
        console.warn(`No existen archivos del tipo ${mimeType} en la carpeta ${processingFolder.getName()}`);
        return;
      }

      backupFileTo(targetFile, backupFolder);

      files.forEach(file => {
        // --- Chequeo de límite de tiempo de ejecución
        if (!shouldContinueProcessing(startTime, timeStats)) {
          throw new Error(ABORTING_EXECUTION_MESSAGE);
        }
        // ---

        try {
          // --- Marca de inicio de ejecución del archivo actual
          const fileStartTime = new Date();
          // ---

          const fileName = file.getName();

          if (file.getId() === logFile.getId() || (
            !config.allowFileReprocessing && alreadyProcessedFiles.has(fileName))) return;

          const existingKeys = getAllKeysFromTargetSheet();

          const processorFn = processors.get(mimeType);

          if (!processorFn) {
            throw new Error(`No se encontró procesador para el tipo MIME: ${mimeType}`);
          }          

          const processedData = processorFn(file, config, targetKeyColumnIndex, targetKeyColumnType, existingKeys);

          if (config.updateExistingRows) {
            updateRowsInSheet(targetSheet, targetKeyColumnIndex, targetKeyColumnType, processedData.updatedRows);
          }

          appendRowsToSheet(targetSheet, processedData.newRows, targetColumnTitles);

          if (config.sortCriteria && Array.isArray(config.sortCriteria) && config.sortCriteria.length > 0) {
            sortSheet(targetSheet, config.sortCriteria);
          }
          
          if (!keepProcessedInSource) moveFileToFolder(file, processedFolder);

          // --- Adicionar el tiempo de ejecución del archivo actual a las estadísticas
          updateTimeStats(timeStats, fileStartTime, processedCount);
          processedCount++;
          // ---
          
          const newRowsCount = processedData.newRows.length;
          const updatedRowsCount = processedData.updatedRows.length;
          const totalCount = newRowsCount + updatedRowsCount;
          const summary = `${totalCount} [${newRowsCount}; ${updatedRowsCount}]`;

          writeLog(logSheet, fileName, summary, successMessage, logColumnTitles);
          console.info(`Se procesaron ${totalCount} filas de la hoja ${targetSheet.getName() || 'Primera'} del archivo ${fileName}. Cambios realizados en la hoja ${targetFileSheetName || 'primera'} del archivo ${targetFile.getName()}:${newRowsCount === 0 ? '' : `\n\t- Adicionadas ${newRowsCount} líneas.`}${updatedRowsCount === 0 ? '' : `\n\t- Actualizadas ${updatedRowsCount} líneas.`}`);

        } catch (e) {
          writeLog(logSheet, fileName, "", e.message, logColumnTitles);
          console.error(failureMessage + ': ' + fileName + ' - ' + e.message);
        }
      });
    }

    const normalizedConfigs = normalizeProcessingConfigMap(configs);

    try {
      Object.entries(normalizedConfigs).forEach(([mimeType, config]) => {
        validateProcessingConfig(config);
        processFiles(mimeType, config);
      });
    } catch (error) {
      if (error.message === ABORTING_EXECUTION_MESSAGE) {
        console.warn(abortError.message);
      } else {
        throw error;
      }
    }

    // --- Actualización del estado de la ejecución
    saveExecutionState(props, {
      executionId,
      processedCount,
      timeStats
    });
    // ---

  } catch (unexpectedError) {
    console.error("Error inesperado:", unexpectedError);
  } finally {
    // --- Finalización segura
    safeCleanup(props, lock, executionId);
    // ---
  }

}

/**
 * Normaliza una configuración de procesamiento individual aplicando valores por defecto
 * 
 * @param {ProcessingConfig} config - Configuración normalizar
 * @returns {ProcessingConfig} Configuración normalizada completa
 */
function normalizeProcessingConfig(config = {}) {
  return /** @type {ProcessingConfig} */ ({
    columns: config?.columns || [],
    searchingFn: config?.searchingFn,
    processingFn: config?.processingFn,
    sourceSheetName: config?.sourceSheetName,
    updateExistingRows: config?.updateExistingRows || UPDATE_EXISTING_ROWS,
    ignoreEmptyRows: config?.ignoreEmptyRows || IGNORE_EMPTY_ROWS,
    allowFileReprocessing: config?.allowFileReprocessing || ALLOW_FILE_REPROCESSING,
    sortCriteria: config?.sortCriteria,
  })
}

/**
 * Normaliza todo el mapa de configuraciones asegurando que cada entrada esté completa
 * 
 * @param {ProcessingConfigMap} configMap Mapa de configuraciones de procesamiento.
 * @returns {ProcessingConfigMap}  Mapa de configuraciones de procesamiento normalizadas.
 */
function normalizeProcessingConfigMap(configMap){
  const normalizedMap = {};
  
  for (const [mimeType, config] of Object.entries(configMap)) {
    normalizedMap[mimeType] = normalizeProcessingConfig(config);
  }
  
  return /** @type {ProcessingConfigMap} */ (normalizedMap);
}

/**
 * Valida una configuración normalizada.
 * 
 * @param {ProcessingConfig} config 
 * @throws {Error} Si la configuración no es válida
 */
function validateProcessingConfig(config) {
  if (!config.searchingFn || typeof config.searchingFn !== "function") {
    throw new Error("Debe proporcionar una función para obtener los archivos a procesar.");
  }
}

/**
 * @typedef {Object} SpreadsheetObjects
 * @property {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet El objeto Spreadsheet de destino.
 * @property {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja de cálculo de destino.
 * @property {GoogleAppsScript.DriveApp.File} file El archivo de DriveApp asociado.
 */

/**
 * Valida y devuelve los objetos del archivo de hoja de cálculo.
 *
 * @param {string} spreadsheetId ID del archivo de hoja de cálculo.
 * @param {string} sheetName Nombre de la hoja.
 * @returns {SpreadsheetObjects} Objetos de archivo de hoja de cálculo configurados.
 * @throws {Error} Cuando el spreadsheetId no es válido o la hoja no existe.
 */
function getSpreadsheetObjectsWithFallback(spreadsheetId, sheetName) {
  if (!spreadsheetId){
    throw new Error(`El valor del identificador del archivo no ha sido proporcionado.`);
  }

  let spreadsheet = null;

  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error(`El valor de identificador del archivo no es válido.`);
  }

  const sheet = sheetName 
    ? spreadsheet.getSheetByName(sheetName) 
    : spreadsheet.getSheets()[0];
  
  if (!sheet) {
    throw new Error(`Hoja no encontrada: ${sheetName || 'Primera hoja'}`);
  }

  const file = DriveApp.getFileById(spreadsheetId);

  return {spreadsheet, sheet, file};
}

/**
 * @typedef {Object} LogOptions
 * @property {GoogleAppsScript.Spreadsheet.Sheet} logSheet La hoja de cálculo de log.
 * @property {GoogleAppsScript.DriveApp.File} logFile El archivo de DriveApp asociado.
 * @property {string[]} logColumnTitles Arreglo que contiene los títulos de las columnas del log.
 * @property {string} successMessage Mensaje de éxito.
 * @property {string} failureMessage Mensaje de fallo.
 */

/**
 * Valida y devuelve las opciones correspondientes al archivo log.
 * 
 * @param {ProcessingOptions} options Opciones de procesamiento.
 * @returns {LogOptions} Objeto de opciones correspondientes al archivo log.
 */
function getLogOptionsWithFallback(options) {
  const logSpreadsheet = options?.logSpreadsheet

  if (!logSpreadsheet) {
    throw new Error("El archivo de hoja de cálculo a utilizar como log no ha sido proporcionado");
  }

  const logSheet = options?.logSheetName 
    ? logSpreadsheet.getSheetByName(options.logSheetName) 
    : logSpreadsheet.getSheets()[0];
  
  if (!logSheet) {
    throw new Error(`Hoja no encontrada en log: ${options?.logSheetName || 'Primera hoja'}`);
  }

  const logFile = DriveApp.getFileById(logSpreadsheet.getId());

  const logColumnTitles = options?.logColumnTitles || LOG_TITLES;
  const successMessage = options?.successMessage || SUCCESS_MESSAGE;
  const failureMessage = options?.failureMessage || PROCESSING_FAILURE_MESSAGE;

  return { logSheet, logFile, logColumnTitles, successMessage, failureMessage };
}

/**
 * Valida y devuelve las opciones correspondientes a las carpetas de trabajo.
 * 
 * @param {ProcessingOptions} options Opciones de procesamiento.
 * @param {GoogleAppsScript.DriveApp.File} targetFile El archivo de destino. Necesario
 * si se desea realizar el respaldo en el destino.
 * @returns {FoldersOptions} Objeto de opciones correspondientes a las carpetas de trabajo.
 */
function getFoldersOptionsWithFallback(options, targetFile) {
  const logFile = DriveApp.getFileById(options.logSpreadsheet.getId());

  const processingFolder = options?.processingFolderID
    ? DriveApp.getFolderById(options.processingFolderID)
    : logFile.getParents().next();

  const processedFolderName = options?.processedFolderName || PROCESSED_FOLDER_NAME;
  const backupFolderName = options?.backupFolderName || BACKUP_FOLDER_NAME;
  const backupInDestination = options?.backupInDestination || BACKUP_IN_DESTINATION;
  const keepProcessedInSource = options?.keepProcessedInSource || KEEP_PROCESSED_IN_SOURCE;
  
  const processedFolder = getOrCreateSubfolderFrom(processingFolder, processedFolderName);

  let targetFolder = null;

  if (backupInDestination) {
    // Chequear permisos
    try {
      targetFolder = targetFile.getParents().next();
    } catch {
      console.warn("No se tienen los permisos suficientes en la carpeta de destino. El respaldo se creará en la carpeta de procesamiento.");
    }
  }
  
  const backupFolder = getOrCreateSubfolderFrom(
    backupInDestination && targetFolder ? targetFolder : processingFolder, 
    backupFolderName);

  return { processingFolder, processedFolder, backupFolder, keepProcessedInSource };
}

/**
 * Devuelve el valor proporcionado normalizado como cadena. Si el valor
 * no posee el formato de cadena, se devuelve una cadena vacía.
 * 
 * @param {any} value El valor de la celda a convertir.
 * @returns {string} Valor de cadena. Es una cadena vacía si no cumple con el formato.
 */
function normalizeString(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

/**
 * Devuelve el valor proporcionado normalizado como entero. Si el valor
 * no posee el formato de número, se devuelve *null*.
 * 
 * @param {any} value El valor de la celda a convertir.
 * @returns {number|null} Valor numérico o *null*, si no cumple con el formato.
 */
function normalizeInteger(value) {
  const n = parseInt(value);
  return isNaN(n) ? null : n;
}

/**
 * Devuelve el valor proporcionado normalizado como decimal. Si el valor
 * no posee el formato decimal válido, se devuelve *null*.

 * @param {any} value El valor de la celda a convertir.
 * @returns {number|null} Valor numérico o *null*, si no cumple con el formato.
 */
function normalizeFloat(value) {
  if (value === null || value === undefined) return null;

  if (typeof value === 'number') return value;

  if (typeof value === 'string') {
    value = value.trim();

    // Si contiene tanto ',' como '.', intentamos detectar el formato
    if (value.includes(',') && value.includes('.')) {
      // Si la coma está antes del punto, probablemente es formato inglés: "1,234.56"
      if (value.indexOf(',') < value.indexOf('.')) {
        value = value.replace(/,/g, '');
      } else {
        // Si el punto está antes de la coma, probablemente es formato español: "1.234,56"
        value = value.replace(/\./g, '').replace(',', '.');
      }
    } else if (value.includes(',')) {
      // Si solo tiene coma, asumimos que es decimal en formato español
      value = value.replace(',', '.');
    } else if (value.includes('.')) {
      // Solo punto: formato inglés estándar, no se transforma
    }

    const n = parseFloat(value);
    return isNaN(n) ? null : n;
  }

  return null;
}

/**
 * Indica si el valor de cadena tiene el formato ISO: *yyyy-MM-dd*.
 * 
 * @param {string} value Cadena en formato *yyyy-MM-dd*.
 * @returns {boolean} *true* si cumple con el formato.
 */
function isISODateString(value) {
  if (typeof value !== 'string') return false;

  // Expresión regular para ISO 8601 básica: YYYY-MM-DD o YYYY-MM-DDTHH:MM:SS(.sss)?(Z|±HH:MM)?
  const isoRegex = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(?:\.\d+)?(Z|[+-]\d{2}:\d{2})?)?$/;

  return isoRegex.test(value);
}

/**
 * Indica si el valor de cadena tiene el formato *dd/MM/yyyy*.
 * 
 * @param {string} value Cadena en formato *dd/MM/yyyy*.
 * @returns {boolean} *true* si cumple con el formato.
 */
function isDDMMYYYY(value) {
  if (typeof value !== 'string') return false;

  const regex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
  const match = value.match(regex);
  if (!match) return false;

  const day = parseInt(match[1], 10);
  const month = parseInt(match[2], 10);
  const year = parseInt(match[3], 10);

  // Validar rangos básicos
  if (month < 1 || month > 12 || day < 1 || day > 31) return false;

  // Verificar fecha real usando el constructor de Date
  const date = new Date(year, month - 1, day);
  return (
    date.getFullYear() === year &&
    date.getMonth() === month - 1 &&
    date.getDate() === day
  );
}

/**
 * Convierte una cadena con formato *dd/MM/yyyy* en un valor de fecha. Si el
 * valor suministrado no cumple con el formato, devuelve null.
 * 
 * @param {string} value Valor de cadena en formato *dd/MM/yyyy*;
 * @returns {Date|null} Objeto *Date* o *null* si no es válido.
 */
function parseDDMMYYYY(value) {
  const match = value.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;

  const day = parseInt(match[1], 10);
  const month = parseInt(match[2], 10);
  const year = parseInt(match[3], 10);

  const date = new Date(year, month - 1, day);
  return (date.getDate() === day && date.getMonth() === month - 1 && date.getFullYear() === year)
    ? date
    : null;
}

/**
 * Devuelve el valor proporcionado normalizado como date. Si el valor
 * no posee el formato de date, se devuelve null. Cuidado al 
 * establecer valores en una planilla puesto que la fecha devuelta
 * es UTF; primero debe convertirse al formato local de la planilla.
 * 
 * @param {any} value El valor de la celda a convertir.
 * @returns {Date|null} Objeto Date UTF o null si no es válido.
 */
function normalizeDate(value) {
  let date = null;

  if (value instanceof Date) {
    date = new Date(value); // Crear copia para no modificar el original
  } else if (typeof value === 'string') {
    if (isISODateString(value)) {
      const newDate = new Date(value);
      if (!isNaN(newDate.getTime())) {
        date = newDate;
      }
    } else if (isDDMMYYYY(value)) {
      date = parseDDMMYYYY(value);
    }
  } else if (typeof value === 'number') {
    // // Google Sheets puede representar fechas como número serial (días desde 1899-12-30)
    // // Convertir a milisegundos desde esa fecha base.
    // const base = new Date('1899-12-30T00:00:00Z');
    // const ms = value * 24 * 60 * 60 * 1000;
    // date = new Date(base.getTime() + ms);

    const days = Math.floor(value);
    const millisPerDay = 24 * 60 * 60 * 1000;
    const newDate = new Date(1899, 11, 30); // 1899-12-30 en zona horaria local
    newDate.setDate(newDate.getDate() + days);
    date = newDate;
  }

  if (date) date.setHours(0, 0, 0, 0);

  return date;
}

/**
 * Devuelve el valor proporcionado normalizado como boolean. Si el valor
 * no posee el formato de boolean, se devuelve null.
 * 
 * @param {any} value El valor de la celda a convertir.
 * @returns {boolean|null} Objeto boolean o null si no es válido.
 */
function normalizeBoolean(value) {
  if (typeof value === 'boolean') return value;

  if (typeof value === 'string') {
    const v = value.trim().toLowerCase();
    if (['true', '1', 'sí', 'si', 'yes', 'y'].includes(v)) return true;
    if (['false', '0', 'no', 'n'].includes(v)) return false;
  }

  if (typeof value === 'number') {
    if (value === 1) return true;
    if (value === 0) return false;
  }

  return null;
}

/**
 * Devuelve un arreglo con los valores de una columna. Permite
 * procesar los datos mediante handlers opcionales. El orden de ejecución 
 * de los handlers es el siguiente:
 *  1. Se ejecuta handlers.filterRowFn sobre todas las filas de la sheet.
 *     Si no se proporciona, se devuelven todas las filas
 *  2. Se extraen los valores de las celdas correspondientes a la columna 
 *     columnNumber de las filas resultantes del paso 1 y se ejecuta 
 *     handlers.formatCellValueFn sobre cada uno. Si no se proporciona,
 *     se devuelven los valores originales.
 *  3. Se ejecuta handlers.filterCellFn sobre las celdas resultantes
 *     del paso 2. Si no se proporciona, se devuelven todas las celdas.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja.
 * @param {number} columnNumber Índice de la columna (base 0).
 * @param {{
 *   filterRowsFn?: (row: any[]) => boolean,
 *   formatCellValueFn?: (cell: any) => any,
 *   filterCellsFn?: (cell: any) => boolean
 * }} handlers Objeto con funciones de procesamiento opcionales.
 * @returns {any[]} Arreglo de valores de la columna después de ser formateados y filtrados.
 */
function getValuesFromColumn(sheet, columnNumber, handlers = {}) {
  try {
    if (!sheet || !sheet.getDataRange) {
      console.warn('Objeto sheet inválido');
      return null;
    }
    
    if (typeof columnNumber !== 'number' || columnNumber < 0) {
      console.warn('Número de columna inválido');
      return null;
    }

    const dataRange = sheet.getDataRange();
    const rowsCount = dataRange.getNumRows();
    const colsCount = dataRange.getNumColumns();
    if (rowsCount === 0 || colsCount === 0 ||
        // La hoja es nueva
        (rowsCount === 1 && colsCount === 1)) {
      return [];
    }

    const rows = dataRange.getValues();
    
    // Verificar que la columna existe
    if (columnNumber >= rows[0].length) {
      console.warn(`La columna ${columnNumber} no existe en la hoja`);
      return null;
    }

    // Procesamiento con funciones opcionales
    const filteredRows = handlers.filterRowsFn
      ? rows.filter((row, index, array) => handlers.filterRowsFn(row, index, array)) 
      : rows;

    const cells = filteredRows.map(row => handlers.formatCellValueFn
      ? handlers.formatCellValueFn(row[columnNumber]) 
      : row[columnNumber]
    );

    const filteredCells = handlers.filterCellsFn
      ? cells.filter((cell, index, array) => handlers.filterCellsFn(cell, index, array)) 
      : cells;

    return filteredCells;
  } catch (error) {
    console.warn('Error en getValuesFromColumn:', error);
    return null;
  }
}

/**
 * Devuelve el conjunto de valores de una columna, sin duplicados. Permite
 * procesar los datos mediante handlers opcionales. El orden de ejecución 
 * de los handlers es el siguiente:
 *  1. Se ejecuta handlers.filterRowFn sobre todas las filas de la sheet.
 *     Si no se proporciona, se devuelven todas las filas
 *  2. Se extraen los valores de las celdas correspondientes a la columna 
 *     columnNumber de las filas resultantes del paso 1 y se ejecuta 
 *     handlers.formatCellValueFn sobre cada uno. Si no se proporciona,
 *     se devuelven los valores originales.
 *  3. Se ejecuta handlers.filterCellFn sobre las celdas resultantes
 *     del paso 2. Si no se proporciona, se devuelven todas las celdas.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja.
 * @param {number} columnNumber Índice de la columna. Comienza en 0.
 * @param {{
 *   filterRowsFn?: (row: any[]) => boolean,
 *   formatCellValueFn?: (cell: any) => any,
 *   filterCellsFn?: (cell: any) => boolean
 * }} handlers Objeto con funciones de manejo.
 * @returns {Set} Conjunto de valores de la columna después de ser formateados y filtrados.
 */
function getUniqueValuesFromColumn(sheet, columnNumber, handlers) {
  const result = getValuesFromColumn(sheet, columnNumber, handlers);
  return result ? new Set(result) : null;
}

/**
 * Devuelve la subcarpeta con el nombre proporcionado dentro de la carpeta padre. Si no la
 * encuentra, la crea.
 * 
 * @param {DriveApp.Folder} parentFolder Carpeta padre.
 * @param {string} subfolderName Nombre de la subcarpeta a buscar.
 * @returns {DriveApp.folder} Subcarpeta.
 */
function getOrCreateSubfolderFrom(parentFolder, subfolderName) {
  const subfolders = parentFolder.getFoldersByName(subfolderName);
  return subfolders.hasNext() ? subfolders.next() : parentFolder.createFolder(subfolderName);
}

/**
 * Realiza una copia de respaldo del archivo a la carpeta indicada.
 * 
 * @param {GoogleAppsScript.DriveApp.File} file Archivo a respaldar.
 * @param {Folder} backupFolder Carpeta de respaldo.
 * @returns {string} Nombre del archivo de respaldo creado.
 */
function backupFileTo(file, backupFolder) {
  const date = new Date().toISOString().replace(/[:.]/g, '-');
  const backupFileName = `Respaldo ${file.getName()} ${date}`;
  file.makeCopy(backupFileName, backupFolder);
  return backupFileName;
}

/**
 * Pipeline de procesamiento común a todos los archivos.
 * 
 * @param {any[][]} data Arreglo de filas de datos a procesar.
 * @param {ProcessingConfig} config Objeto de configuración de procesamiento.
 * @param {number} keyColumnIndex Índice de la columna llave (0-based).
 * @param {CellValueType} keyColumnType Tipo de datos de la columna llave.
 * @param {Set<any>} [existingKeys] Valores de llaves existentes en destino (opcional).
 * @returns {{ newRows: any[][], updatedRows: { key: any, row: any[] }[] }} Objeto que contiene las filas a anexar o actualizar.
 */
function processDataPipeline(data, config, keyColumnIndex, keyColumnType, existingKeys) {
  // 1. Procesamiento personalizado
  const processedData = config.processingFn ? config.processingFn(data) : data;

  // 2. Filtrar filas no vacías
  const filledRows = processedData.filter(row => 
    row.some(cell => cell !== '' && cell !== null && cell !== undefined)
  );

  // 3. Normalizador para clave
  const normalizeKey = normalizers[keyColumnType] || (v => v);

  if (config.updateExistingRows) {
    // 4. Separar filas por si existe o no su llave en existingKeys
    const newRows = [];
    const updatedRows = [];

    filledRows.forEach(row => {
      const keyValue = normalizeKey(row[config.columns[keyColumnIndex]]);
      const mappedRow = config.columns.map(col => 
        Number.isSafeInteger(col) && col >= 0 && col < row.length 
          ? row[col] 
          : col
      );

      if (existingKeys?.has(keyValue)) {
        updatedRows.push({ key: keyValue, row: mappedRow });
      } else {
        newRows.push(mappedRow);
      }
    });

    return { newRows, updatedRows };

  } else {
    // 4. Excluir las filas cuyas llaves existen en existingKeys
    const filteredData = existingKeys?.size > 0
      ? filledRows.filter(row => {
          const keyValue = normalizeKey(row[config.columns[keyColumnIndex]]);
          return !existingKeys.has(keyValue);
        })
      : filledRows;
    
    const mappedData = filteredData.map(row => 
      config.columns.map(col => 
        Number.isSafeInteger(col) && col >= 0 && col < row.length 
          ? row[col] 
          : col
      )
    );

    return { newRows: mappedData, updatedRows: [] };
  }
}

/**
 * Función de procesamiento de un archivo con formato google sheet. Realiza tareas de procesamiento
 * comunes a todos los archivos google sheet utilizando el objeto ProcessingConfig suministrado en el 
 * parámetro config: 
 *  Lee todas las filas de la primera hoja del archivo.
 *  - Realiza el procesamiento específico mediante la función config.processingFn, si es
 *    proporcionada.
 *  - Elimina de la lista de filas del paso anterior las filas en blanco y las que contienen 
 *    una clave presente en existingKeys.
 *  - Extrae de las filas resultantes del paso anterior las columnas indicadas en 
 *    config.columns.
 * 
 * @param {GoogleAppsScript.DriveApp.File} file Archivo de Google Sheets a procesar
 * @param {ProcessingConfig} config Objeto de configuración de procesamiento
 * @param {number} keyColumnIndex Índice de la columna llave (0-based)
 * @param {CellValueType} keyColumnType Tipo de datos de la columna llave
 * @param {Set<any>} [existingKeys] Valores de llaves existentes en destino (opcional)
 * @returns {{ newRows: any[][], updatedRows: { key: any, row: any[] }[] }} Objeto que contiene las filas a anexar o actualizar.
 * @throws {Error} Si falla el procesamiento
 */
function processGoogleSheetFile(file, config, keyColumnIndex, keyColumnType, existingKeys) {
  if (!file || !file.getId || file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    throw new Error('El archivo debe ser un Google Sheet válido');
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(file.getId());

    const sheet = config.sourceSheetName 
      ? spreadsheet.getSheetByName(config.sourceSheetName) 
      : spreadsheet.getSheets()[0];

    if (!sheet) {
      throw new Error(`No se encontró la hoja: ${config.sourceSheetName || 'Primera Hoja'}`);
    }

    const values = sheet.getDataRange().getValues();

    return processDataPipeline(values, config, keyColumnIndex, keyColumnType, existingKeys);

  } catch (error) {
    throw new Error(`Error procesando el archivo Google Sheet: (${error})`);
  }
}

/**
 * Convierte un archivo a formato Google Sheet.
 * 
 * @param {GoogleAppsScript.DriveApp.File} file Archivo a convertir.
 * @returns {GoogleAppsScript.DriveApp.File} Archivo convertido.
 * @throws {Error} Si falla la conversión.
 */
function convertFileToGoogleSheet(file) {
  try {
    const convertedFile = Drive.Files.create(
      {
        name: file.getName().replace(/\.xlsx?$/, ''),
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [file.getParents().next().getId()]
      },
      file.getBlob(),
      {
        convert: true
      }
    );
    
    return DriveApp.getFileById(convertedFile.id);

  } catch (error) {
    throw new Error(`Error convirtiendo el archivo a Google Sheet: (${error})`);
  }
}

/**
 * Función de procesamiento de un archivo con formato Xlsx. Realiza tareas de procesamiento
 * comunes a todos los archivos Xlsx utilizando el objeto ProcessingConfig suministrado en el 
 * parámetro config: 
 *  Lee todas las filas de la primera hoja del archivo.
 *  - Realiza el procesamiento específico mediante la función config.processingFn, si es
 *    proporcionada.
 *  - Elimina de la lista de filas del paso anterior las filas en blanco y las que contienen 
 *    una clave presente en existingKeys.
 *  - Extrae de las filas resultantes del paso anterior las columnas indicadas en 
 *    config.columns.
 * 
 * @param {GoogleAppsScript.DriveApp.File} file Archivo XLSX a procesar
 * @param {ProcessingConfig} config Configuración de procesamiento
 * @param {number} keyColumnIndex Índice de la columna clave (0-based)
 * @param {CellValueType} keyColumnType Tipo de datos de la columna clave
 * @param {Set<any>} [existingKeys] Valores de llaves existentes en destino (opcional)
 * @returns {{ newRows: any[][], updatedRows: { key: any, row: any[] }[] }} Objeto que contiene las filas a anexar o actualizar.
 * @throws {Error} Si falla la conversión o procesamiento
 */
function processXlsxFile(file, config, keyColumnIndex, keyColumnType, existingKeys) {
  if (!file || !file.getId || file.getMimeType() !== MimeType.MICROSOFT_EXCEL) {
    throw new Error('El archivo debe ser un XLSX válido');
  }

  let convertedFile = null;
  let spreadsheet= null;

  try {
    convertedFile = convertFileToGoogleSheet(file);

    try {
      spreadsheet = SpreadsheetApp.openById(convertedFile.getId());

      const sheet = config.sourceSheetName 
        ? spreadsheet.getSheetByName(config.sourceSheetName) 
        : spreadsheet.getSheets()[0];

      if (!sheet) {
        throw new Error(`No se encontró la hoja: ${config.sourceSheetName || 'Primera hoja'}`);
      }

      const values = sheet.getDataRange().getValues();

      return processDataPipeline(values, config, keyColumnIndex, keyColumnType, existingKeys);

    } catch (error) {
      throw new Error(`Error procesando el archivo XLSX: (${error})`);
    }
  } finally {
    if (convertedFile) {
      try {
        if (spreadsheet) {
          SpreadsheetApp.flush();
        }
        convertedFile.setTrashed(true);
      } catch (cleanError) {
        console.warn(`Error limpiando archivo temporal: (${cleanError})`);
      }
    }
  }
}

/**
 * Detecta el caracter utilizado con mayor probabilidad como separador en una línea de archivo CSV.
 *  
 * @param {string} line Línea de un archivo CSV.
 * @returns {string} Caracter utilizado como separador de valores.
 */
function detectCSVSeparator(line) {
  const separators = [';', ',', '\t', '|'];
  const separatorsCount = separators.map(separator => ({
    separator,
    count: (line.match(new RegExp(`\\${separator}`, 'g')) || []).length
  }));

  separatorsCount.sort((a, b) => b.count - a.count);
  return separatorsCount[0].separator;
}

/**
 * Función auxiliar para parsear líneas CSV complejas.
 * Maneja casos con comillas y separadores dentro de campos.
 * 
 * @param {string} line Línea de texto.
 * @param {string} separator Separador.
 * @returns {any[]} Arreglo de valores de columnas.
 */
function parseCSVLine(line, separator) {
  const pattern = new RegExp(`(?:^|${separator})(?:(?:([^"${separator}]*)|\"((?:[^\"]|\"\")*)\"))`, 'g');
  const fields = [];
  let match;
  
  while ((match = pattern.exec(line)) !== null) {
    fields.push((match[2] !== undefined) ? 
      match[2].replace(/""/g, '"') : // Quitar escapes de comillas
      match[1]);
  }
  
  return fields;
}

/**
 * Función de procesamiento de un archivo con formato CSV. Realiza tareas de procesamiento
 * comunes a todos los archivos CSV utilizando el objeto ProcessingConfig suministrado en el 
 * parámetro config: 
 *  - Lee todo el contenido del archivo.
 *  - Realiza el procesamiento específico mediante la función config.processingFn, si es
 *    proporcionada.
 *  - Elimina de la lista de filas del paso anterior las filas en blanco y las que contienen 
 *    una clave presente en existingKeys.
 *  - Extrae de las filas resultantes del paso anterior las columnas indicadas en 
 *    config.columns.
 * 
 * @param {GoogleAppsScript.DriveApp.File} file Archivo CSV
 * @param {ProcessingConfig} config Configuración de procesamiento
 * @param {number} keyColumnIndex Índice de columna clave (0-based)
 * @param {CellValueType} keyColumnType Tipo de dato de columna clave
 * @param {Set<any>} [existingKeys] Valores de llaves existentes en destino (opcional)
 * @returns {{ newRows: any[][], updatedRows: { key: any, row: any[] }[] }} Objeto que contiene las filas a anexar o actualizar.
 * @throws {Error} Si el archivo no es válido o falla el procesamiento
 */
function processCsvFile(file, config, keyColumnIndex, keyColumnType, existingKeys) {
  if (!file || file.getMimeType() !== 'text/csv') {
    throw new Error('El archivo debe ser un CSV válido');
  }

  try {
    const contents = file.getBlob().getDataAsString('UTF-8');
    if (!contents.trim()) return [];
    
    const lines = contents.split(/\r?\n/).filter(line => line.trim() !== '');
    if (lines.length === 0) return [];
    
    const separator = detectCSVSeparator(lines[0]);
    const rows = lines.map(line => parseCSVLine(line, separator));

    return processDataPipeline(rows, config, keyColumnIndex, keyColumnType, existingKeys);
    
  } catch (error) {
    throw new Error(`Error procesando un archivo CSV: ${error}`);
  }
}

/**
 * Anexa filas a una planilla destino. Si la planilla está vacía, adiciona títulos.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Hoja de cálculo destino.
 * @param {number} keyColumn Índice (0-based) de la columna clave.
 * @param {CellValueType} keyType Tipo de datos de la columna llave.
 * @param {{ key: any, row: any[] }[]} updatedRows Objeto que contiene las filas a actualizar con su clave correspondiente.
 */
function updateRowsInSheet(targetSheet, keyColumn, keyType, updatedRows) {
  if (!updatedRows || updatedRows.length === 0) return;

  const sheetLastRow = targetSheet.getLastRow();

  if (sheetLastRow < 2) return; // No hay datos para actualizar

  const keyColumnIndex = keyColumn + 1;
  const numColumns = updatedRows[0].row.length;

  // Obtener todas las claves de la hoja (asumiendo títulos en la fila 1)
  const keyRange = targetSheet.getRange(2, keyColumnIndex, sheetLastRow - 1);
  const keyValues = keyRange.getValues().flat();

  // Normalizador para clave
  const normalizeKey = normalizers[keyType] || (v => v);

  // Mapa de clave -> número de fila (real en hoja)
  const keyToRowMap = new Map();
  keyValues.forEach((key, i) => {
    if (key !== '' && key !== null && key !== undefined) {
      keyToRowMap.set(normalizeKey(key), i + 2); // +2 porque empieza en fila 2
    }
  });

  // Recorre y actualiza solo las filas existentes
  updatedRows.forEach(({ key, row }) => {
    const targetRow = keyToRowMap.get(key);
    if (!targetRow) return;

    // Reemplazo de "__ROW__" si aplica
    const processedRow = row.map(cell =>
      typeof cell === 'string' && cell.includes('__ROW__')
        ? cell.replace(/__ROW__/g, targetRow)
        : cell
    );

    targetSheet.getRange(targetRow, 1, 1, numColumns).setValues([processedRow]);
  });
}

/**
 * Anexa filas a una planilla destino. Si la planilla está vacía, adiciona títulos.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet Hoja de cálculo destino.
 * @param {any[][]} rows Filas a anexar.
 * @param {string[]} targetColumnTitles Títulos de columnas a insertar si la hoja está vacía.
 */
function appendRowsToSheet(targetSheet, rows, targetColumnTitles) {
  if (!rows || rows.length === 0) return;

  let lastRow = targetSheet.getLastRow();

  if (lastRow === 0 && Array.isArray(targetColumnTitles) && targetColumnTitles.length > 0) {
    targetSheet.appendRow(targetColumnTitles);
    lastRow = 1;
  }

  // Reemplazar "__ROW__" por el número de fila real antes de insertar
  const processedRows = rows.map((row, i) => {
    const rowIndex = lastRow + 1 + i;
    return row.map(cell => {
      return (typeof cell === 'string' && cell.includes('__ROW__'))
        ? cell.replace(/__ROW__/g, rowIndex)
        : cell;
    });
  });

  targetSheet.getRange(lastRow + 1, 1, processedRows.length, processedRows[0].length)
             .setValues(processedRows);
}

/**
 * Mueve un archivo a una carpeta.
 * 
 * @param {GoogleAppsScript.DriveApp.File} file Archivo a mover.
 * @param {GoogleAppsScript.DriveApp.Folder} targetFolder Carpeta destino.
 */
function moveFileToFolder(file, targetFolder) {
  try {
    const parents = file.getParents();
    const previousParents = [];
    while (parents.hasNext()) {
      previousParents.push(parents.next().getId());
    }

    Drive.Files.update(
      {},
      file.getId(),
      null, {
        addParents: targetFolder.getId(),
        removeParents: previousParents.join(',')
      },
    );
  } catch {
    console.warn(`El archivo ${file.getName()} no puede moverse a la carpeta ${targetFolder.getName()}. No se tienen los permisos suficientes.`);
  }
}

/**
 * Escribe un mensaje en la hoja de log.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet Hoja de log.
 * @param {string} fileName Nombre del archivo.
 * @param {string} summary Resumen de las filas procesadas (Total [Nuevas; Modificadas]).
 * @param {string} state Estado de la operación.
 * @param {string[]} logTitles Títulos de las columnas de la hoja de log.
 * @returns {Date} Fecha y hora de escritura del mensaje.
 */
function writeLog(logSheet, fileName, summary, state, logTitles = []) {
  if (logSheet.getLastRow() === 0 && logTitles.length) {
    logSheet.appendRow(logTitles);
  }
  const date = new Date();
  logSheet.appendRow([date, fileName, summary, state]);
  return date;
}

/**
 * Obtiene archivos de una carpeta filtrados por tipo MIME y opcionalmente ordenados.
 * 
 * @param {GoogleAppsScript.DriveApp.Folder} folder - Carpeta donde buscar los archivos.
 * @param {string} mimeType - Tipo MIME de los archivos a buscar (ej. 'application/vnd.google-apps.spreadsheet').
 * @param {function(GoogleAppsScript.DriveApp.File, GoogleAppsScript.DriveApp.File): number} [orderingFn] - Función personalizada para ordenar (opcional).
 * @returns {GoogleAppsScript.DriveApp.File[]} - Array de archivos encontrados (vacío si no hay resultados)
 */
function getFilesFromFolder(folder, mimeType, orderingFn = undefined) {
  if (!folder || !folder.getFilesByType) {
    console.warn('Parámetro "folder" no es una carpeta válida');
    return [];
  }

  if (typeof mimeType !== 'string') {
    console.warn('El tipo MIME debe ser una cadena de texto');
    return [];
  }

  try {
    const files = [];
    const fileIterator = folder.getFilesByType(mimeType);

    if (!fileIterator.hasNext()) return files;

    while (fileIterator.hasNext()) {
      files.push(fileIterator.next());
    }

    // Ordenar los archivos
    if (orderingFn && typeof orderingFn === 'function') {
      files.sort(orderingFn);
    } else {
      // Orden por defecto (alfabético por nombre)
      files.sort((a, b) => a.getName().localeCompare(b.getName()));
    }

    return files;
  } catch (error) {
    console.error('Error al buscar archivos:', error);
    return [];
  }
}

/**
 * @typedef {Object} SortCriterion
 * @property {number} column Número de columna (comenzando en 1; A=1, B=2, etc.).
 * @property {boolean} ascending `true` para ascendente, `false` para descendente.
 */

/**
 * Ordena una hoja por múltiples columnas. Usa el filtro si está presente;
 * de lo contrario, ordena el rango manualmente (excluyendo el encabezado).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Objeto de la hoja a ordenar.
 * @param {SortCriterion[]} sortCriteria Arreglo de criterios de ordenamiento.
 * @returns {void}
 */
function sortSheet(sheet, sortCriteria) {
  if (!sheet || typeof sheet.getRange !== "function") {
    console.warn("No se proporcionó una hoja válida a ordenar.");
    return;
  }

  if (!Array.isArray(sortCriteria) || sortCriteria.length === 0) {
    console.warn("No se proporcionó un arreglo de criterios de ordenamiento.");
    return;
  }

  const columnCount = sheet.getLastColumn();

  for (let i = 0; i < sortCriteria.length; i++) {
    const criterion = sortCriteria[i];

    if (
      typeof criterion.column !== "number" ||
      criterion.column < 1 ||
      criterion.column > columnCount
    ) {
      console.warn(`Criterio de ordenamiento inválido en la posición ${i}: número de columna fuera de rango.`);
      return;
    }

    if (typeof criterion.ascending !== "boolean") {
      console.warn(`Criterio de ordenamiento inválido en la posición ${i}: el valor 'ascending' debe ser booleano.`);
      return;
    }
  }

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    console.warn("No hay suficientes filas para ordenar.");
    return;
  }

  const hadFilter = sheet.getFilter() !== null;

  if (hadFilter) {
    sheet.getFilter().remove();
  }

  try {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, columnCount);
    dataRange.sort(sortCriteria);
  } catch (e) {
    console.warn(`Error al ordenar la hoja '${sheet.getName()}': ${e.message}`);
  } finally {
    if (hadFilter) {
      sheet.getRange(1, 1, 1, columnCount).createFilter();
    }
  }
}

/**
 * Valida el estado actual de ejecución y repara estados huérfanos.
 * 
 * @param {GoogleAppsScript.Properties.Properties} props Instancia de PropertiesService
 * @return {boolean} True si la ejecución puede continuar, False si hay otra ejecución en curso
 * @example
 * if (!validateExecutionState(PropertiesService.getScriptProperties())) return;
 */
function validateExecutionState(props) {
  const status = props.getProperty("EXECUTION_STATUS");
  const lastStart = props.getProperty("LAST_START");
  
  if (status === "RUNNING" && lastStart) {
    const elapsed = new Date() - new Date(lastStart);
    if (elapsed > REPAIR_THRESHOLD) {
      console.warn("Reseteando estado huérfano");
      props.setProperty("EXECUTION_STATUS", "IDLE");
      return true;
    }
    console.warn("Ejecución ya en curso");
    return false;
  }
  return true;
}

/**
 * Intenta establecer el bloqueo de la ejecución del script; si no lo consigue devuelve null.
 * Si consigue establecer el bloqueo, actualiza las propiedades del estado de
 * ejecución con los valores de la actual y devuelve el objeto de las estadísticas
 * de tiempo históricas.
 * 
 * @param {number} startTime Fecha y hora en la que comenzó a ejecutarse el script.
 * @param {string} executionId ID único de esta ejecución
 * @param {GoogleAppsScript.Lock.Lock} lock Instancia de LockService
 * @param {GoogleAppsScript.Properties.Properties} props Instancia de PropertiesService
 * @returns {Object | null} timeStats Estadísticas de tiempo históricas o null si no se pudo establecer el bloqueo.
 * @returns {number} stats.maxTime Tiempo máximo registrado
 * @returns {number} stats.movingAvg Promedio móvil actual
 * @returns {number[]} stats.lastFiles Lista de últimos tiempos
 */
function safeStartup(startTime, executionId, lock, props) {
  if (!lock.tryLock(0)) {
    console.warn("Bloqueo activo. Cancelando...");
    return null;
  }

  props.setProperties({
    "EXECUTION_STATUS": "RUNNING",
    "LAST_START": startTime.toISOString(),
    "CURRENT_EXECUTION_ID": executionId
  }, true);

  const stats = JSON.parse(props.getProperty("TIME_STATS") || '{}');

  return {
    maxTime: stats.maxTime || ABSOLUTE_MIN_TIME,
    movingAvg: stats.movingAvg || ABSOLUTE_MIN_TIME,
    lastFiles: stats.lastFiles || [],
  };
}

/**
 * Determina si el script debe continuar procesando más archivos.
 * 
 * @param {number} startTime Fecha y hora en la que comenzó a ejecutarse el script.
 * @param {Object} timeStats Estadísticas de tiempo históricas
 * @param {number} timeStats.maxTime Tiempo máximo registrado para procesar un archivo
 * @param {number} timeStats.movingAvg Promedio móvil de tiempos de procesamiento
 * @param {number[]} timeStats.lastFiles Últimos tiempos registrados
 * @return {boolean} True si hay tiempo suficiente para procesar el próximo archivo
 * @example
 * const shouldContinue = shouldContinueProcessing(
 *   new Date(), 
 *   {maxTime: 45000, movingAvg: 30000, lastFiles: [35000, 40000]},
 *   0.25
 * );
 */
function shouldContinueProcessing(startTime, timeStats) {
  const remainingTime = MAX_RUN_TIME - (new Date() - startTime);
  // Estimación basada en el MÁXIMO histórico + margen de seguridad
  const estimatedTime = timeStats.maxTime * (1 + SAFETY_MARGIN);
  
  console.log(`Tiempo restante: ${(remainingTime/1000).toFixed(1)}s | ` +
              `Estimado: ${(estimatedTime/1000).toFixed(1)}s (Máx: ${(timeStats.maxTime/1000).toFixed(1)}s)`);
  
  // Condiciones para continuar:
  return remainingTime > estimatedTime * 1.5; // 50% más del estimado como buffer
}

/**
 * Actualiza las estadísticas de tiempo con el último archivo procesado.
 * 
 * @param {Object} stats Objeto de estadísticas a actualizar
 * @param {number} stats.maxTime Tiempo máximo registrado
 * @param {number} stats.movingAvg Promedio móvil actual
 * @param {number[]} stats.lastFiles Lista de últimos tiempos
 * @param {number} fileStartTime Fecha y hora de comienzo de procesamiento del último archivo
 * @param {number} processedCount Total de archivos procesados en esta ejecución
 * @return {void}
 * @example
 * updateTimeStats(
 *   {maxTime: 30000, movingAvg: 25000, lastFiles: []},
 *   new Date(),
 *   3
 * );
 */
function updateTimeStats(stats, fileStartTime, processedCount) {
  const fileTime = new Date() - fileStartTime;

  // Actualiza el máximo histórico
  stats.maxTime = Math.max(stats.maxTime, fileTime);
  
  // Media móvil ponderada (más peso a los últimos archivos)
  stats.movingAvg = processedCount < 5 ? 
    (stats.movingAvg * processedCount + fileTime) / (processedCount + 1) :
    stats.movingAvg * 0.8 + fileTime * 0.2;
  
  // Mantener registro de últimos tiempos
  stats.lastFiles = [fileTime, ...stats.lastFiles].slice(0, 10);
}

/**
 * Guarda el estado de ejecución y estadísticas de forma segura.
 * 
 * @param {GoogleAppsScript.Properties.Properties} props Instancia de PropertiesService
 * @param {Object} params Parámetros de ejecución
 * @param {string} params.executionId ID único de esta ejecución
 * @param {number} params.processedCount Archivos procesados en esta ejecución
 * @param {Object} params.timeStats Estadísticas de tiempo actualizadas
 * @return {void}
 * @example
 * saveExecutionState(PropertiesService.getScriptProperties(), {
 *   executionId: '123e4567-e89b-12d3-a456-426614174000',
 *   processedCount: 5,
 *   timeStats: {maxTime: 45000, movingAvg: 30000, lastFiles: [35000, 40000]}
 * });
 */
function saveExecutionState(props, {executionId, processedCount, timeStats}) {
  const batchUpdates = {};
  
  if (processedCount > 0) {
    batchUpdates["TIME_STATS"] = JSON.stringify(timeStats);
  }
  
  // Solo actualizar estado si somos la ejecución activa
  if (props.getProperty("CURRENT_EXECUTION_ID") === executionId) {
    batchUpdates["EXECUTION_STATUS"] = "IDLE";
    batchUpdates["CURRENT_EXECUTION_ID"] = null;
  }
  
  props.setProperties(batchUpdates, true);
}

/**
 * Realiza limpieza segura de locks y estados
 * 
 * @param {GoogleAppsScript.Properties.Properties} props Instancia de PropertiesService
 * @param {GoogleAppsScript.Lock.Lock} lock Instancia de LockService
 * @param {string} executionId ID único de esta ejecución
 * @return {void}
 * @example
 * safeCleanup(
 *   PropertiesService.getScriptProperties(),
 *   LockService.getScriptLock(),
 *   '123e4567-e89b-12d3-a456-426614174000'
 * );
 */
function safeCleanup(props, lock, executionId) {
  try {
    if (props.getProperty("CURRENT_EXECUTION_ID") === executionId) {
      props.setProperties({
        "EXECUTION_STATUS": "IDLE",
        "CURRENT_EXECUTION_ID": null
      }, true);
    }
  } catch (cleanupError) {
    console.error("Error limpiando estado:", cleanupError);
      
    Utilities.sleep(SLEEP_TIME);
    props.setProperties({
      "EXECUTION_STATUS": "IDLE",
      "CURRENT_EXECUTION_ID": null
    }, true);
  }
  
  try {
    if (lock.hasLock()) lock.releaseLock();
  } catch (lockError) {
    console.error("Error liberando el bloqueo del script:", lockError);
  }
}
