let TEST_SPREADSHEET = null;
let TEST_FOLDER = null;

/**
 * Ejecuta todas las pruebas.
 */
function runAllTests() {
  createTestSpreadsheet();

  try {
    test_normalizeDate();
    // test_normalizeInteger();
    // test_normalizeString();
    // test_normalizeFloat();
    // test_normalizeBoolean();
    // test_getValuesFromColumn();
    // test_getOrCreateSubfolderFrom();
    // test_backupFileTo();
    // test_appendRowsToSheet();
    // test_updateRowsInSheet();
    // test_moveFileToFolder();
    // test_writeLog();
    test_sortSheet();
    // test_processCSVFile();
    // test_processXlsxFile();
    // test_processGoogleSheetFile();
    // test_processFolderFilesAndCopyTo();
  } finally {
    // deleteTestSpreadsheet();
  }
}

/**
 * Devuelve el objeto Folder de la carpeta que contiene a este script.
 */
function getScriptParentFolder() {
  const scriptId = ScriptApp.getScriptId();

  const files = Drive.Files.list({
    q: "mimeType='application/vnd.google-apps.script' and trashed = false"
  }).files;

  for (const file of files) {
    if (file.id === scriptId) {
      const driveFile = DriveApp.getFileById(file.id);
      const parentId = driveFile.getParents().next().getId();
      return DriveApp.getFolderById(parentId);
    }
  }

  return null;
}

/**
 * Crea una hoja de cálculo temporal para pruebas.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function createTestSpreadsheet() {
  const folder = getScriptParentFolder();

  const spreadsheet = SpreadsheetApp.create('TestSheet_' + new Date().toISOString());
  const fileId = spreadsheet.getId();

  const folderId = folder.getId();
  const file = DriveApp.getFileById(fileId);

  const parents = file.getParents();
  const previousParents = [];
  while (parents.hasNext()) {
    previousParents.push(parents.next().getId());
  }

  Drive.Files.update(
    {},
    fileId,
    null, {
      addParents: folderId,
      removeParents: previousParents.join(',')
    },
  );

  TEST_SPREADSHEET = spreadsheet;
  TEST_FOLDER = folder;

  return spreadsheet;
}

function deleteTestSpreadsheet() {
  if (TEST_SPREADSHEET) {
    DriveApp.getFileById(TEST_SPREADSHEET.getId()).setTrashed(true);
  }
}

/**
 * Test para la función normalizeDate
 */
function test_normalizeDate() {
  const inputs = [
    new Date(2023, 0, 1, 12, 30),      // Date con hora
    '2023-01-01T15:45:00Z',            // ISO string
    '01/01/2023',                      // dd/MM/yyyy string
    44927,                             // Número serial de Google Sheets (2023-01-01)
    'invalid-date',                    // Cadena no válida
    {},                                // Objeto no fecha
    null                               // null
  ];

  const expected = [
    new Date(2023, 0, 1),              // 2023-01-01 00:00
    new Date(2023, 0, 1),              // 2023-01-01 00:00
    new Date(2023, 0, 1),              // 2023-01-01 00:00
    new Date(2023, 0, 1),              // 2023-01-01 00:00
    null,
    null,
    null
  ];  
  
  const results = inputs.map(normalizeDate);

  Logger.log(`Test: normalizeDate: ${JSON.stringify(expected) === JSON.stringify(results)}`);
  // Logger.log(`Esperados: ${JSON.stringify(expected)}`);
  // Logger.log(`Resultados: ${JSON.stringify(results)}`);
}


/**
 * Test para la función normalizeInteger
 */
function test_normalizeInteger() {
  const inputs = [42, '42', '  42  ', 'abc', '', null, undefined, 42.8];
  const esperados = [42, 42, 42, null, null, null, null, 42];

  const resultados = inputs.map(normalizeInteger);

  Logger.log(`Test: normalizeInteger: ${JSON.stringify(esperados) === JSON.stringify(resultados)}`);
  // Logger.log('Esperado: ' + JSON.stringify(esperados));
  // Logger.log('Resultado: ' + JSON.stringify(resultados));
}


/**
 * Test para la función normalizeString
 */
function test_normalizeString() {
  const inputs = ['  hola  ', '', null, undefined, 123, true];
  const esperados = ['hola', '', '', '', '123', 'true'];

  const resultados = inputs.map(normalizeString);

  Logger.log(`Test: normalizeString: ${JSON.stringify(esperados) === JSON.stringify(resultados)}`);
  // Logger.log('Esperado: ' + JSON.stringify(esperados));
  // Logger.log('Resultado: ' + JSON.stringify(resultados));
}

/**
 * Test para la función normalizeFloat
 */
function test_normalizeFloat() {
  const valores = ['1234.56', '1,234.56', '1.234,56', '12,34', '12.34', 'abc', '', null, 45];
  const esperados = [1234.56, 1234.56, 1234.56, 12.34, 12.34, null, null, null, 45.0];
  
  const resultados = valores.map(normalizeFloat);
  
  Logger.log(`Test: normalizeFloat: ${JSON.stringify(esperados) === JSON.stringify(resultados)}`);
  // Logger.log('Esperado: ' + JSON.stringify(esperados));
  // Logger.log('Resultado: ' + JSON.stringify(resultados));
}


/**
 * Test para la función normalizeBoolean
 */
function test_normalizeBoolean() {
  const valores = ['true', 'false', '1', '0', true, false, 'sí', '', null];
  const esperados = [true, false, true, false, true, false, true, null, null];
  
  const resultados = valores.map(normalizeBoolean);
  
  Logger.log(`Test: normalizeBoolean: ${JSON.stringify(esperados) === JSON.stringify(resultados)}`);
  // Logger.log('Esperado: ' + JSON.stringify(esperados));
  // Logger.log('Resultado: ' + JSON.stringify(resultados));
}

function test_getValuesFromColumn() {
  const hoja = TEST_SPREADSHEET.insertSheet('getValuesFromColumn');
  hoja.getRange('A1:B5').setValues([
    ['ID', 'Nombre'],
    [1, 'Ana'],
    [2, 'Luis'],
    [3, 'Ana'],
    [4, 'Carlos']
  ]);

  const handlers = {
    filterRowsFn: row => row[1] !== 'Ana' && row[1] !== 'Nombre',
    formatCellValueFn: value => value.toString().toUpperCase(),
    filterCellsFn: cell => cell !== 'LUIS'
  };

  const result = getValuesFromColumn(hoja, 1, handlers); // Columna B (índice 1)

  Logger.log(`Test: getValuesFromColumn: ${Array.from(result).join("") === 'CARLOS'}`);
  // Logger.log(Array.from(result)); // Debería contener ['CARLOS']
}

function test_getOrCreateSubfolderFrom() {
  const carpetaPadre = DriveApp.getFileById(TEST_SPREADSHEET.getId()).getParents().next();
  const nombreSubcarpeta = 'Carpeta de Prueba';

  const oldFoldersIterator = carpetaPadre.getFoldersByName(nombreSubcarpeta);

  while (oldFoldersIterator.hasNext()) oldFoldersIterator.next().setTrashed(true)

  const subcarpeta = getOrCreateSubfolderFrom(carpetaPadre, nombreSubcarpeta);

  const creadaCarpeta = carpetaPadre.getFoldersByName(nombreSubcarpeta).hasNext();

  Logger.log(`Test: getOrCreateSubfolderFrom: ${creadaCarpeta}`);
  // Logger.log(`¿Se creó la carpeta?: ${creadaCarpeta}`); // Debería ser true

  subcarpeta.setTrashed(true);
}

function test_backupFileTo() {
  const parentFolders = DriveApp.getFileById(TEST_SPREADSHEET.getId()).getParents();
  const parentFolder = parentFolders.next();

  const archivoOriginal = parentFolder.createFile('archivo_prueba.txt', 'Contenido de prueba');
  const carpetaBackup = parentFolder.createFolder('CarpetaBackupTest');

  const backupFileName = backupFileTo(archivoOriginal, carpetaBackup);

  const archivos = carpetaBackup.getFilesByName(backupFileName);
  const hayRespaldo = archivos.hasNext();

  Logger.log(`Test: backupFileTo: ${hayRespaldo}`);
  // Logger.log(`¿Se creó el respaldo?: ${hayRespaldo}`); // Debería ser true

  // Limpieza
  archivoOriginal.setTrashed(true);
  carpetaBackup.setTrashed(true);
}

/**
 * Test para la función appendRowsToSheet
 */
function test_appendRowsToSheet() {
  const hoja = TEST_SPREADSHEET.insertSheet('appendRowsToSheet');

  const TITULOS_DESTINO = ['Nombre', 'Fecha', 'Monto'];
  const datos = [
    ['Juan', new Date('2024-01-01'), 100],
    ['Ana', new Date('2024-02-01'), 200],
  ];

  appendRowsToSheet(hoja, datos, TITULOS_DESTINO);

  const valores = hoja.getDataRange().getValues();
  Logger.log(`Test: appendRowsToSheet: ${JSON.stringify([TITULOS_DESTINO, ...datos]) === JSON.stringify(valores)}`);
  // Logger.log('Contenido datos: ' + JSON.stringify([TITULOS_DESTINO, ...datos]));
  // Logger.log('Contenido hoja: ' + JSON.stringify(valores));
}

/**
 * Test para la función updateRowsInSheet
 */
function test_updateRowsInSheet() {
  const hoja = TEST_SPREADSHEET.insertSheet('updateRowsInSheet');

  const TITULOS_DESTINO = ['Nombre', 'Fecha', 'Monto'];
  const datos = [
    ['Ana', new Date('2024-02-01'), 200],
    ['Juan', new Date('2024-01-01'), 100],
    ['María', null, null],
    ['Pedro', '', ''],
  ];

  appendRowsToSheet(hoja, datos, TITULOS_DESTINO);

  const nuevosDatos = [
    ['Pedro', new Date('2024-10-03'), 4000],
    ['María', new Date('2024-10-04'), 3000],
    ['Juan', new Date('2024-10-01'), 2000],
    ['Ana', new Date('2024-10-02'), 1000],
  ];

  const updatedRows = nuevosDatos.map(row => ({ key: row[0], row }));

  updateRowsInSheet(hoja, 0, 'string', updatedRows);

  const valores = hoja.getDataRange().getValues();
  const ordenados = nuevosDatos.sort((a, b) => a[0].localeCompare(b[0]));
  Logger.log(`Test: updateRowsInSheet: ${JSON.stringify([TITULOS_DESTINO, ...ordenados]) === JSON.stringify(valores)}`);
  // Logger.log('Contenido datos: ' + JSON.stringify([TITULOS_DESTINO, ...ordenados]));
  // Logger.log('Contenido hoja: ' + JSON.stringify(valores));
}

function test_moveFileToFolder() {
  const parentFolders = DriveApp.getFileById(TEST_SPREADSHEET.getId()).getParents();
  const parentFolder = parentFolders.next();

  const carpetaOrigen = parentFolder.createFolder('CarpetaOrigenTest');
  const carpetaDestino = parentFolder.createFolder('CarpetaDestinoTest');

  // Crear archivo de prueba
  const nombreArchivo = 'archivo_test_move.txt';
  const archivo = carpetaOrigen.createFile(nombreArchivo, 'Contenido de prueba');

  // Ejecutar función
  moveFileToFolder(archivo, carpetaDestino);

  // Comprobar que ya no está en la carpeta de origen y sí en destino
  const estaEnDestino = carpetaDestino.getFilesByName(nombreArchivo).hasNext();
  const estaEnOrigen = carpetaOrigen.getFilesByName(nombreArchivo).hasNext();

  Logger.log(`Test: moveFileToFolder: ${!estaEnOrigen && estaEnDestino}`);
  // Logger.log(`¿Está en destino?: ${estaEnDestino}`); // true
  // Logger.log(`¿Está en origen?: ${estaEnOrigen}`);   // false

  // Limpieza
  archivo.setTrashed(true);
  carpetaOrigen.setTrashed(true);
  carpetaDestino.setTrashed(true);
}

function test_writeLog() {
  const sheetName = "writeLog";
  const sheet = TEST_SPREADSHEET.insertSheet(sheetName);
  
  const titulos = ['Fecha', 'Archivo', 'Cantidad', 'Estado'];

  const mensajes = [
    ['archivo_test.xlsx', 10, 'Procesado OK'],
    ['archivo_test2.xlsx', 0, 'Vacío']
  ]

  mensajes.forEach((mensaje,) => mensaje.unshift(writeLog(sheet, ...mensaje, titulos)));

  const datos = sheet.getDataRange().getValues();

  Logger.log(`Test: writeLog: ${JSON.stringify(datos) === JSON.stringify([titulos, ...mensajes])}`);
  // Logger.log('Mensajes: ' + JSON.stringify([titulos, ...mensajes]));
  // Logger.log('En log: ' + JSON.stringify(datos));
}

function test_sortSheet() {
  const hoja = TEST_SPREADSHEET.insertSheet('sortSheet');

  const TITULOS_DESTINO = ['Nombre', 'Fecha', 'Monto'];
  const datos = [
    ['Pedro', new Date('2024-10-03'), 4000],
    ['María', new Date('2024-10-04'), 3000],
    ['Juan', new Date('2024-10-01'), 2000],
    ['Ana', new Date('2024-10-02'), 1000],
  ];

  appendRowsToSheet(hoja, datos, TITULOS_DESTINO);

  const datosOrdenados = datos.sort((a, b) => {
    const dateA = normalizeDate(a[1]);
    const dateB = normalizeDate(b[1]);
    return (dateA > dateB) ? 1 : (dateA < dateB) ? -1 : 0;
  });

  sortSheet(hoja, [{column: 2, ascending: true}]);

  const valores = hoja.getDataRange().getValues();
  Logger.log(`Test: sortSheet: ${JSON.stringify([TITULOS_DESTINO, ...datosOrdenados]) === JSON.stringify(valores)}`);
  Logger.log('Contenido datos: ' + JSON.stringify([TITULOS_DESTINO, ...datosOrdenados]));
  Logger.log('Contenido hoja: ' + JSON.stringify(valores));
}

function test_processCSVFile() {
  const sheetName = "CSV";
  const sheet = TEST_SPREADSHEET.insertSheet(sheetName);

  const iterator = TEST_FOLDER.getFilesByType(MimeType.CSV);
  while (iterator.hasNext()) {
    const processedRows = processCsvFile(
      iterator.next(), 
      {
        columns: [0, 7, 8, 1, 2, 17, 18],
        processingFn: (rows) => 
          rows
            .slice(1)
            .sort((a, b) => {
              const valA = Number(a[0]); 
              const valB = Number(b[0]);
              return (valA > valB) ? 1 : (valA < valB) ? -1 : 0;
            })
        ,
      },
      0,
      "integer",
    );
    appendRowsToSheet(
      sheet, 
      processedRows.newRows, 
      ['Servicio', 'Id Vehículo', 'Vehículo', 'Fecha', 'Hora', 'Producto', 'Litros']);
    Logger.log(`Test: processCSVFile: adicionadas ${processedRows.newRows.length} líneas.`);
  }
}

function test_processXlsxFile() {
  const sheetName = "Excel";
  const sheet = TEST_SPREADSHEET.insertSheet(sheetName);

  const iterator = TEST_FOLDER.getFilesByType(MimeType.MICROSOFT_EXCEL);
  while (iterator.hasNext()) {
    const processedRows = processXlsxFile(
      iterator.next(), 
      {
        columns: [2, 0, 1, 3, 4, 7, 10],
        processingFn: (rows) => 
          rows
            .slice(8)
            .sort((a, b) => {
              const valA = Number(a[2]); 
              const valB = Number(b[2]);
              return (valA > valB) ? 1 : (valA < valB) ? -1 : 0;
            })
        ,
      },
      0,
      "integer",
    );
    appendRowsToSheet(
      sheet, 
      processedRows.newRows, 
      ['Servicio', 'Id Vehículo', 'Vehículo', 'Fecha', 'Hora', 'Producto', 'Litros']);
    Logger.log(`Test: processXlsxFile: adicionadas ${processedRows.newRows.length} líneas.`);
  }
}

function test_processGoogleSheetFile() {
  const sheetName = "GoogleSheet";
  const sheet = TEST_SPREADSHEET.insertSheet(sheetName);

  const iterator = TEST_FOLDER.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iterator.hasNext()) {
    const processedRows = processGoogleSheetFile(
      iterator.next(), 
      {
        columns: [2, 0, 1, 3, 4, 7, 10],
        processingFn: (rows) => 
          rows
            .slice(8)
            .sort((a, b) => {
              const valA = Number(a[2]); 
              const valB = Number(b[2]);
              return (valA > valB) ? 1 : (valA < valB) ? -1 : 0;
            })
        ,
      },
      0,
      "integer",
    );
    appendRowsToSheet(
      sheet, 
      processedRows.newRows, 
      ['Servicio', 'Id Vehículo', 'Vehículo', 'Fecha', 'Hora', 'Producto', 'Litros']);
    Logger.log(`Test: processGoogleSheetFile: adicionadas ${processedRows.newRows.length} líneas.`);
  }
}

function test_processFolderFilesAndCopyTo() {
  const logSheetName = "SURTIDOR";
  
  TEST_SPREADSHEET.insertSheet(logSheetName);

  const targetFileID = "1hAuyzRozJe-Ec-ircp0TR3RHQ2Fect5hvX-QfuczLEg";
  const targetFileSheetName = "SURTIDOR";
  const targetColumnTitles = ['Servicio', 'Id Vehículo', 'Vehículo', 'Fecha', 'Hora', 'Producto', 'Litros'];
  const targetKeyColumnIndex = 0;
  const targetKeyColumnType = "integer";
  /** @type ProcessingConfig */
  const configs = {
    [MimeType.MICROSOFT_EXCEL]: {
      columns: [2, 0, 1, 3, 4, 7, 10],
      searchingFn: () => getFilesFromFolder(
        TEST_FOLDER, 
        MimeType.MICROSOFT_EXCEL
      ), 
      processingFn: (rows) => 
        rows
          .slice(8)
          .sort((a, b) => {
            const valA = Number(a[2]); 
            const valB = Number(b[2]);
            return (valA > valB) ? 1 : (valA < valB) ? -1 : 0;
          })
    },
    [MimeType.GOOGLE_SHEETS]: {
      columns: [2, 0, 1, 3, 4, 7, 10],
      // Solo google sheet: Copia de CONSUMOS 01-03 23-04
      searchingFn: () => [DriveApp.getFileById("1IcsTKir_k1us9FbMCt-iR42gFg_OwyfyI2oFxJGU_Zc")], 
      processingFn: (rows) => 
        rows
          .slice(8)
          .sort((a, b) => {
            const valA = Number(a[2]); 
            const valB = Number(b[2]);
            return (valA > valB) ? 1 : (valA < valB) ? -1 : 0;
          })
    },
    [MimeType.CSV]: {
      columns: [0, 7, 8, 1, 2, 17, 18],
      searchingFn: () => getFilesFromFolder(
        TEST_FOLDER, 
        MimeType.CSV,
        (a, b) => a.getName() === "Copia de SERVICIOS.CSV" ? -1 : 1,
      ),
      processingFn: (rows) => 
        rows
          .slice(1)
          .sort((a, b) => {
            const valA = Number(a[0]); 
            const valB = Number(b[0]);
            return (valA > valB) ? 1 : (valA < valB) ? -1 : 0;
          })
    },
  };

  /** type ProcessingOptions */
  const options = {
    logSpreadsheet: TEST_SPREADSHEET,
    logSheetName: logSheetName,
    keepProcessedInSource: true,
  };

  processFolderFilesAndCopyTo(
    targetFileID, targetFileSheetName, targetColumnTitles, targetKeyColumnIndex, targetKeyColumnType, configs, options);
}
