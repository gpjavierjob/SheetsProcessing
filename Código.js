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
      return Utils.getUniqueValuesFromColumn(
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
      return Utils.getUniqueValuesFromColumn(
        targetSheet, 
        targetKeyColumnIndex,
        {
          formatCellValueFn: (cell) => Utils.getNormalizer(targetKeyColumnType)(cell)
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

      Utils.backupFileTo(targetFile, backupFolder);

      let processedLines = 0;
      let addedLines = 0;
      let updatedLines = 0;

      files.forEach(file => {
        // --- Chequeo de límite de tiempo de ejecución
        if (!shouldContinueProcessing(startTime, timeStats)) {
          throw new Error(ABORTING_EXECUTION_MESSAGE);
        }
        // ---

        const fileName = file.getName();

        try {
          // --- Marca de inicio de ejecución del archivo actual
          const fileStartTime = new Date();
          // ---

          if (file.getId() === logFile.getId() || (
            !config.allowFileReprocessing && alreadyProcessedFiles.has(fileName))) return;

          const existingKeys = getAllKeysFromTargetSheet();

          const processorFn = processors.get(mimeType);

          if (!processorFn) {
            throw new Error(`No se encontró procesador para el tipo MIME: ${mimeType}`);
          }          

          const processedData = processorFn(file, config, targetKeyColumnIndex, targetKeyColumnType, existingKeys);

          let updatedRowsCount = 0;
          if (config.updateExistingRows) {
            updatedRowsCount = Utils.updateRowsInSheet(
              targetSheet, targetKeyColumnIndex, targetKeyColumnType, processedData.updatedRows);
          }

          Utils.appendRowsToSheet(targetSheet, processedData.newRows, targetColumnTitles);

          if (config.sortCriteria && Array.isArray(config.sortCriteria) && config.sortCriteria.length > 0) {
            Utils.sortSheet(targetSheet, config.sortCriteria);
          }
          
          if (!keepProcessedInSource) Utils.moveFileToFolder(file, processedFolder);

          // --- Adicionar el tiempo de ejecución del archivo actual a las estadísticas
          updateTimeStats(timeStats, fileStartTime, processedCount);
          processedCount++;
          // ---
          
          const newRowsCount = processedData.newRows.length;
          const totalCount = newRowsCount + updatedRowsCount;
          const summary = `${totalCount} [${newRowsCount}; ${updatedRowsCount}]`;

          writeLog(logSheet, fileName, summary, successMessage, logColumnTitles);
          console.info(`Se procesaron ${totalCount} filas de la hoja ${targetSheet.getName() || 'Primera'} del archivo ${fileName}.${totalCount === 0 ? '' : `\nCambios realizados en la hoja ${targetFileSheetName || 'primera'} del archivo ${targetFile.getName()}:${newRowsCount === 0 ? '' : `\n\t- Adicionadas ${newRowsCount} líneas.`}${updatedRowsCount === 0 ? '' : `\n\t- Actualizadas ${updatedRowsCount} líneas.`}`}`);

          processedLines =+ totalCount;
          addedLines =+ newRowsCount;
          updatedLines =+ updatedRowsCount;

        } catch (e) {
          writeLog(logSheet, fileName, "", e.message, logColumnTitles);
          console.error(failureMessage + ': ' + fileName + ' - ' + e.message);
        }
      });

      return { processed: processedLines, added: addedLines, updated: updatedLines };
    }

    const normalizedConfigs = normalizeProcessingConfigMap(configs);

    let processed = 0;
    let added = 0;
    let updated = 0;

    try {
      Object.entries(normalizedConfigs).forEach(([mimeType, config]) => {
        validateProcessingConfig(config);
        const result = processFiles(mimeType, config);

        if (!result) return;

        processed = processed + result.processed
        added = added + result.added
        updated = updated + result.updated
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

    return { processed, added, updated };

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
  
  const processedFolder = Utils.getOrCreateSubfolderFrom(processingFolder, processedFolderName);

  let targetFolder = null;

  if (backupInDestination) {
    // Chequear permisos
    try {
      targetFolder = targetFile.getParents().next();
    } catch {
      console.warn("No se tienen los permisos suficientes en la carpeta de destino. El respaldo se creará en la carpeta de procesamiento.");
    }
  }
  
  const backupFolder = Utils.getOrCreateSubfolderFrom(
    backupInDestination && targetFolder ? targetFolder : processingFolder, 
    backupFolderName);

  return { processingFolder, processedFolder, backupFolder, keepProcessedInSource };
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
  const normalizeKey = Utils.getNormalizer(keyColumnType) || (v => v);

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
    convertedFile = Utils.convertFileToGoogleSheet(file);

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
  
  console.info(`Tiempo restante: ${(remainingTime/1000).toFixed(1)}s | ` +
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
