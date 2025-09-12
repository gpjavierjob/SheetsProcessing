var SheetsProcessing = (function () {
  'use strict';

  // Parámetros de control de tiempo de ejecución
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

  const ALLOWED_MIME_TYPES = [MimeType.MICROSOFT_EXCEL, MimeType.CSV, MimeType.GOOGLE_SHEETS];
  const ALLOWED_COLUMN_TYPES = ['string', 'float', 'integer', 'boolean', 'date'];

  /**
   * Función de procesamiento general.
   *
   * @callback ProcessorFunction
   * @param {GoogleAppsScript.DriveApp.File} file Archivo a procesar.
   * @param {SourceConfig} config Objeto de configuración de procesamiento.
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
   * Normaliza una configuración de procesamiento individual aplicando valores por defecto
   * 
   * @param {SourceConfig} config - Configuración normalizar
   * @returns {SourceConfig} Configuración normalizada completa
   */
  function normalizeProcessingConfig(config = {}) {
    return /** @type {SourceConfig} */ ({
      columns: config?.columns ? [...config.columns] : [],
      searchingFn: config?.searchingFn,
      processingFn: config?.processingFn,
      sourceSheetName: config?.sourceSheetName,
      updateExistingRows: config?.updateExistingRows ?? UPDATE_EXISTING_ROWS,
      ignoreEmptyRows: config?.ignoreEmptyRows ?? IGNORE_EMPTY_ROWS,
      allowFileReprocessing: config?.allowFileReprocessing ?? ALLOW_FILE_REPROCESSING,
      sortCriteria: config?.sortCriteria,
    })
  }

  /**
   * Normaliza todo el mapa de configuraciones asegurando que cada entrada esté completa
   * 
   * @param {SourceConfigMap} configMap Mapa de configuraciones de procesamiento.
   * @returns {SourceConfigMap}  Mapa de configuraciones de procesamiento normalizadas.
   */
  function normalizeProcessingConfigMap(configMap){
    const normalizedMap = {};

    if (configMap) {    
      for (const [mimeType, config] of Object.entries(configMap)) {
        normalizedMap[mimeType] = normalizeProcessingConfig(config);
      }
    }
    
    return /** @type {SourceConfigMap} */ (normalizedMap);
  }

  /**
   * Valida un MimeType.
   * 
   * @param {string} value 
   * @throws {Error} Si el valor de MimeType no es válido
   */
  function validateMimeType(value) {
    if (!ALLOWED_MIME_TYPES.includes(value)) {
      throw new Error(`El valor de MimeType: ${value} no es válido.`)
    }
  }

  /**
   * Valida una configuración normalizada.
   * 
   * @param {SourceConfig} config 
   * @throws {Error} Si la configuración no es válida
   */
  function validateProcessingConfig(config) {
    if (!Utils.isPlainObject(config)) {
      throw new Error("La configuración debe ser un objeto válido.");
    }
  
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
      throw new Error(`El valor del identificador del archivo de destino no ha sido proporcionado.`);
    }

    let spreadsheet = null;

    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      throw new Error(`El valor de identificador del archivo de destino no es válido.`);
    }

    const sheet = sheetName 
      ? spreadsheet.getSheetByName(sheetName) 
      : spreadsheet.getSheets()[0];
    
    if (!sheet) {
      throw new Error(`Hoja de destino no encontrada: ${sheetName || 'Primera hoja'}`);
    }

    const file = DriveApp.getFileById(spreadsheetId);

    return { spreadsheet, sheet, file };
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
   * Opciones de carpetas de trabajo.
   * 
   * @typedef {Object} FoldersOptions
   * @property {GoogleAppsScript.DriveApp.Folder} processingFolder Objeto carpeta de archivos a procesar
   * @property {GoogleAppsScript.DriveApp.Folder} processedFolder Objeto carpeta de archivos procesados
   * @property {GoogleAppsScript.DriveApp.Folder} backupFolder Objeto carpeta de respalddo
   * @property {boolean} keepProcessedInSource Si es `true`, mantiene el archivo procesado en la carpeta fuente
   */

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

    const processedFolderName = options?.processedFolderName ?? PROCESSED_FOLDER_NAME;
    const backupFolderName = options?.backupFolderName ?? BACKUP_FOLDER_NAME;
    const backupInDestination = options?.backupInDestination ?? BACKUP_IN_DESTINATION;
    const keepProcessedInSource = options?.keepProcessedInSource ?? KEEP_PROCESSED_IN_SOURCE;
    
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
   * @param {SourceConfig} config Objeto de configuración de procesamiento.
   * @param {number} keyColumnIndex Índice de la columna llave (0-based).
   * @param {CellValueType} keyColumnType Tipo de datos de la columna llave.
   * @param {Set<any>} [existingKeys] Valores de llaves existentes en destino (opcional).
   * @returns {{ newRows: any[][], updatedRows: { key: any, row: any[] }[] }} Objeto que contiene las filas a anexar o actualizar.
   */
  function processDataPipeline(data, config, keyColumnIndex, keyColumnType, existingKeys) {
    if (!data) throw new Error("El parámetro 'data' no es válido.");

    if (!config || typeof config !== "object") throw new Error("El parámetro 'config' no es válido.");

    if (!Number.isInteger(keyColumnIndex) || keyColumnIndex < 0) {
      throw new Error("El parámetro 'keyColumnIndex' no es válido.");
    }

    if (typeof keyColumnType !== "string" || !ALLOWED_COLUMN_TYPES.includes(keyColumnType)) {
      throw new Error("El parámetro 'keyColumnType' no es válido.");
    }

    const processedData = config.processingFn ? config.processingFn(data) : data;

    const filledRows = (config.ignoreEmptyRows
      ? processedData.filter(row => row && row.some(cell => cell !== '' && cell !== null && cell !== undefined))
      : processedData
    );

    const normalizeKey = Utils.getNormalizer(keyColumnType) || (v => v);

    if (config.updateExistingRows) {
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
   * comunes a todos los archivos google sheet utilizando el objeto SourceConfig suministrado en el 
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
   * @param {SourceConfig} config Objeto de configuración de procesamiento
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

    if (!config || typeof config !== "object") throw new Error("El parámetro 'config' no es válido.");

    if (!Number.isInteger(keyColumnIndex) || keyColumnIndex < 0) {
      throw new Error("El parámetro 'keyColumnIndex' no es válido.");
    }

    if (typeof keyColumnType !== "string" || !ALLOWED_COLUMN_TYPES.includes(keyColumnType)) {
      throw new Error("El parámetro 'keyColumnType' no es válido.");
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
   * comunes a todos los archivos Xlsx utilizando el objeto SourceConfig suministrado en el 
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
   * @param {SourceConfig} config Configuración de procesamiento
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

    if (!config || typeof config !== "object") throw new Error("El parámetro 'config' no es válido.");

    if (!Number.isInteger(keyColumnIndex) || keyColumnIndex < 0) {
      throw new Error("El parámetro 'keyColumnIndex' no es válido.");
    }

    if (typeof keyColumnType !== "string" || !ALLOWED_COLUMN_TYPES.includes(keyColumnType)) {
      throw new Error("El parámetro 'keyColumnType' no es válido.");
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
   * Maneja casos con comillas, comillas dobles escapadas 
   * y separadores dentro de campos.
   * 
   * @param {string} line Línea de texto.
   * @param {string} separator Separador.
   * @returns {string[]} Arreglo de valores de columnas.
   */
  function parseCSVLine(line, separator) {
    if (!separator) throw new Error("El parámetro 'separator' no es válido.");

    const fields = [];
    let currentField = '';
    let inQuotes = false;
    let i = 0;
    
    while (i < line.length) {
        if (!inQuotes && line.substring(i, i + separator.length) === separator) {
            // Encontramos un separador fuera de comillas - fin del campo
            fields.push(currentField);
            currentField = '';
            i += separator.length;
            continue;
        }
        
        if (line[i] === '"') {
            if (inQuotes) {
                // Estamos dentro de comillas - verificar si es comilla escapada
                if (i + 1 < line.length && line[i + 1] === '"') {
                    // Comilla escapada: agregar una comilla
                    currentField += '"';
                    i += 2;
                } else {
                    // Fin de campo entrecomillado
                    currentField += '"';
                    inQuotes = false;
                    i++;
                }
            } else if (currentField === '') {
                // Inicio de campo entrecomillado (solo si es el primer carácter del campo)
                inQuotes = true;
                currentField += '"';
                i++;
            } else {
                // Comilla dentro de un campo no entrecomillado: tratar como carácter normal
                currentField += '"';
                i++;
            }
        } else {
            // Carácter normal
            currentField += line[i];
            i++;
        }
    }
    
    // Agregar el último campo
    fields.push(currentField);

    return fields.map(field => {
        const trimmed = field.trim();

        // Verificar si es un campo entrecomillado válido
        if (trimmed.startsWith('"') && trimmed.endsWith('"') && trimmed.length > 1) {
            const content = trimmed.slice(1, -1);
            // Reemplazar comillas escapadas
            return content.replace(/""/g, '"');
        }
        
        return field;
    });
  }

  /**
   * Función de procesamiento de un archivo con formato CSV. Realiza tareas de procesamiento
   * comunes a todos los archivos CSV utilizando el objeto SourceConfig suministrado en el 
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
   * @param {SourceConfig} config Configuración de procesamiento
   * @param {number} keyColumnIndex Índice de columna clave (0-based)
   * @param {CellValueType} keyColumnType Tipo de dato de columna clave
   * @param {Set<any>} [existingKeys] Valores de llaves existentes en destino (opcional)
   * @returns {{ newRows: any[][], updatedRows: { key: any, row: any[] }[] }} Objeto que contiene las filas a anexar o actualizar.
   * @throws {Error} Si el archivo no es válido o falla el procesamiento
   */
  function processCsvFile(file, config, keyColumnIndex, keyColumnType, existingKeys) {
    if (!file || file.getMimeType() !== MimeType.CSV) {
      throw new Error('El archivo debe ser un CSV válido');
    }

    if (!config || typeof config !== "object") throw new Error("El parámetro 'config' no es válido.");

    if (!Number.isInteger(keyColumnIndex) || keyColumnIndex < 0) {
      throw new Error("El parámetro 'keyColumnIndex' no es válido.");
    }

    if (typeof keyColumnType !== "string" || !ALLOWED_COLUMN_TYPES.includes(keyColumnType)) {
      throw new Error("El parámetro 'keyColumnType' no es válido.");
    }

    try {
      const contents = file.getBlob().getDataAsString('UTF-8');
      if (!contents.trim()) return { newRows: [], updatedRows: [] };
      
      const lines = contents.split(/\r?\n|\r/).filter(line => line.trim() !== '');
      if (lines.length === 0) return { newRows: [], updatedRows: [] };

      const separator = detectCSVSeparator(lines[0]);
      const rows = lines.map(line => parseCSVLine(line, separator));

      return processDataPipeline(rows, config, keyColumnIndex, keyColumnType, existingKeys);
      
    } catch (error) {
      throw new Error(`Error procesando un archivo CSV: ${error}`);
    }
  }

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
   * @typedef {Object} SortCriterion
   * @property {number} column Número de columna (comenzando en 1; A=1, B=2, etc.).
   * @property {boolean} ascending `true` para ascendente, `false` para descendente.
   */

  /**
   * Objeto de configuración del archivo destino.
   *
   * @typedef {Object} TargetConfig
   * @property {string} fileID ID del archivo de destino.
   * @property {string} fileSheetName Nombre de la hoja de destino.
   * @property {string[]} columnTitles Títulos de columnas a insertar si la hoja está vacía.
   * @property {number} keyColumnIndex Índice de la columna llave única.
   * @property {CellValueType} keyColumnType Tipo de datos de la columna llave única.
   * @property {SortCriterion[]} [sortCriteria] Arreglo de criterios de ordenamiento para el archivo de destino.
   */

  /**
   * Objeto de configuración del procesamiento personalizado para un tipo de archivo en particular.
   *
   * @typedef {Object} SourceConfig
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
   * 
   * @typedef {Object.<string, SourceConfig>} SourceConfigMap
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
   * @param {Ctx} context Contexto de ejecución.
   * @param {TargetConfig} targetConfig Configuración de archivo destino.
   * @param {SourceConfigMap} sourcesConfig Mapa de configuración de archivos fuente por mime type.
   * @param {ProcessingOptions} options Opciones de procesamiento.
   */
  function mergeSheetsDataTo(ctx, targetConfig, sourcesConfig, options) {

    if (!ctx) throw new Error("El parámetro 'ctx' no es válido.");

    if (!targetConfig) throw new Error("El parámetro 'targetConfig' no es válido.");

    const { 
      fileID: targetFileID, 
      fileSheetName: targetFileSheetName, 
      columnTitles: targetColumnTitles, 
      keyColumnIndex: targetKeyColumnIndex, 
      keyColumnType: targetKeyColumnType,
      sortCriteria: targetSortCriteria } = targetConfig;

    const {
       spreadsheet: targetSpreadsheet, sheet: targetSheet, file: targetFile
    } = getSpreadsheetObjectsWithFallback(targetFileID, targetFileSheetName);

    if (!Array.isArray(targetColumnTitles) || targetColumnTitles.some(r => typeof r !== 'string')) {
      throw new Error("El parámetro 'targetConfig.columnTitles' no es válido.");
    }

    if (!Number.isInteger(targetKeyColumnIndex) || 
          targetKeyColumnIndex < 0 || 
          targetKeyColumnIndex >= targetColumnTitles.length) {
      throw new Error("El parámetro 'targetConfig.keyColumnIndex' no es válido.");
    }

    if (typeof targetKeyColumnType !== "string" || !ALLOWED_COLUMN_TYPES.includes(targetKeyColumnType)) {
      throw new Error("El parámetro 'targetConfig.keyColumnType' no es válido.");
    }

    // Log file
    const { 
      logSheet, logFile, logColumnTitles, successMessage, failureMessage 
    } = getLogOptionsWithFallback(options);

    // Folders
    const { 
      processingFolder, processedFolder, backupFolder, keepProcessedInSource 
    } = getFoldersOptionsWithFallback(options, targetFile);

    let processed = 0, added = 0, updated = 0;
    let backupDone = false;

    try {
      /** Asegura que se realice el backup del archivo destino una sola vez */
      function ensureBackupOnce() {
        if (!backupDone) {
          Utils.backupFileTo(targetFile, backupFolder);
          backupDone = true;
        }
      }

      /** Obtener los archivos ya procesados con éxito desde el archivo de log. */
      const alreadyProcessedFiles = (function getAlreadyProcessedFiles() {
        return Utils.getUniqueValuesFromColumn(
          logSheet, 1,
          { filterRowsFn: (row, idx) => (idx > 0 && row[LOG_FILENAME_COL] && row[LOG_STATE_COL] === successMessage) }
        ) || new Set();
      })();

      /** Obtener las claves de las filas de la hoja destino. */
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
       * @param {SourceConfig} config Objeto de configuración del procesamiento.
       */
      function processFiles(mimeType, config) {
        if (!config.searchingFn()) {
          console.error(`No se encontró la función de búsqueda de archivos para el tipo MIME: ${mimeType}`);
          return;
        }          

        let files = [];

        try {
          files = config.searchingFn(mimeType) || [];
        } catch (e) {
          console.error(`${SEARCHING_ERROR_MESSAGE}: ${e.message}`);
          return;
        }

        if (files.length === 0) {
          console.warn(`No existen archivos del tipo ${mimeType} en la carpeta ${processingFolder.getName()}`);
          return;
        }

        // cache de keys que se va actualizando en caliente
        let currentKeys = getAllKeysFromTargetSheet();
        const normalizeKey = Utils.getNormalizer(targetKeyColumnType) || (v => v);

        const result = EC.ExecutionController.runLoop(ctx, files, ({ item: file }) => {
          const fileName = file.getName();

          if (file.getId() === logFile.getId() || (
            !config.allowFileReprocessing && alreadyProcessedFiles.has(fileName))) {
            return { processed: 0, added: 0, updated: 0, fileName, date: new Date() };
          }

          const processorFn = processors.get(mimeType);
          if (!processorFn) {
            throw new Error(`No se encontró procesador para el tipo MIME: ${mimeType}`);
          }          

          const processedData = processorFn(file, config, targetKeyColumnIndex, targetKeyColumnType, currentKeys);

          let updatedRowsCount = 0;
          if (config.updateExistingRows && processedData.updatedRows.length > 0) {
            ensureBackupOnce();
            updatedRowsCount = Utils.updateRowsInSheet(
              targetSheet, targetKeyColumnIndex, targetKeyColumnType, processedData.updatedRows);
          }

          if (processedData.newRows.length > 0) {
            ensureBackupOnce();
            Utils.appendRowsToSheet(targetSheet, processedData.newRows, targetColumnTitles);

            // actualizar cache de keys con lo recién insertado
            for (const r of processedData.newRows) {
              const k = normalizeKey(r[targetKeyColumnIndex]);
              currentKeys.add(k);
            }
          }
          
          if (!keepProcessedInSource) Utils.moveFileToFolder(file, processedFolder);

          if (config.sortCriteria && Array.isArray(config.sortCriteria) && config.sortCriteria.length > 0) {
            Utils.sortSheet(targetSheet, config.sortCriteria);
          }

          const newRowsCount = processedData.newRows.length;
          const totalCount = newRowsCount + updatedRowsCount;

          return { processed: totalCount, added: newRowsCount, updated: updatedRowsCount, fileName, date: new Date() };
        });

        if (result.invalidCtx) throw new Error('El contexto de ejecución no es válido.');

        const sheetLogs = [];
        const consoleLogs = new Map();

        const successes = Array.isArray(result.successes) ? result.successes : [];
        const fails = Array.isArray(result.fails) ? result.fails : [];

        const [processed, added, updated] = successes.reduce(
          ([processed, added, updated], item) => {
            const summary = `${item.processed} [${item.added}; ${item.updated}]`;
            sheetLogs.push([item.date, item.fileName, summary, successMessage]);
            consoleLogs.set(
              item.fileName,
              `Se procesaron ${item.processed} filas de la hoja ${targetSheet.getName() || 'Primera'} del archivo ${item.fileName}.${item.processed === 0 ? '' : `\nCambios realizados en la hoja ${targetFileSheetName || 'primera'} del archivo ${targetFile.getName()}:${item.added === 0 ? '' : `\n\t- Adicionadas ${item.added} líneas.`}${item.updated === 0 ? '' : `\n\t- Actualizadas ${item.updated} líneas.`}`}`
            );
            return [processed + item.processed, added + item.added, updated + item.updated];
          }, 
          [0, 0, 0]
        );

        fails.forEach(fail => {
          sheetLogs.push([fail.date, fail.item.fileName, "", fail.message]);
          consoleLogs.set(fail.item.fileName, `${failureMessage}: ${fail.item.fileName} - ${fail.message}`)
        });

        sheetLogs.sort((a, b) => a[0].getTime() - b[0].getTime());

        Utils.appendRowsToSheet(logSheet, sheetLogs, logColumnTitles);

        sheetLogs.forEach(([ , fileName, summary ]) => {
          const msg = consoleLogs.get(fileName);
          if (summary) console.info(msg); else console.error(msg);
        });

        return { processed, added, updated, stoppedEarly: result.stoppedEarly };
      }

      const normalizedConfigs = normalizeProcessingConfigMap(sourcesConfig);

      let stoppedEarly = false;

      try {
        Object.entries(normalizedConfigs).forEach(([mimeType, config]) => {
          validateMimeType(mimeType)
          validateProcessingConfig(config);
          const result = processFiles(mimeType, config);

          if (!result) return;

          processed += result.processed; added += result.added; updated += result.updated

          if (result.stoppedEarly) throw new Error(ABORTING_EXECUTION_MESSAGE);

        });
      } catch (error) {
        if (error.message === ABORTING_EXECUTION_MESSAGE) {
          console.warn(error.message);
          stoppedEarly = true;
        } else {
          throw error;
        }
      }

      if (!stoppedEarly && targetSortCriteria && Array.isArray(targetSortCriteria) && targetSortCriteria.length > 0) {
        Utils.sortSheet(targetSheet, targetSortCriteria);
      }

      return { processed, added, updated, success: true };

    } catch (unexpectedError) {
      console.error("Error inesperado:", unexpectedError);
      return { processed, added, updated, success: false, error: unexpectedError.message };
    }

  }

  // --- API de testing ---
  const _test = {
    normalizeProcessingConfig,
    normalizeProcessingConfigMap,
    validateMimeType,
    validateProcessingConfig,
    getSpreadsheetObjectsWithFallback,
    getLogOptionsWithFallback,
    getFoldersOptionsWithFallback,
    processDataPipeline,
    processGoogleSheetFile,
    processXlsxFile,
    detectCSVSeparator,
    parseCSVLine,
    processCsvFile,
    mergeSheetsDataTo,
  };

  // --- API pública ---
  var api = {
    mergeSheetsDataTo,
    _test,
  };

  return api;

})()
