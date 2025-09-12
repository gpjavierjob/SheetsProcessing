function _runSheetsProcessingTests() {
  const lib = SheetsProcessing._test;

  //----------------------------------------------
  //
  // Tests para validateMimeType
  //
  //----------------------------------------------

  Utils.describe("validateMimeType", () => {
    
    const ALLOWED_MIME_TYPES = [MimeType.MICROSOFT_EXCEL, MimeType.CSV, MimeType.GOOGLE_SHEETS];

    Utils.it("no lanza error para MIME types válidos", () => {
      ALLOWED_MIME_TYPES.forEach(mimeType => {
        Utils.assertFunctionParams(
          lib.validateMimeType,
          [mimeType],
          false, 
          `MIME type ${mimeType} debería ser válido`
        );
      });
    });

    Utils.it("lanza error para MIME types no válidos", () => {
      const invalidMimeTypes = [
        "application/pdf",
        "application/json",
        "application/xml",
        "application/zip",
        "application/vnd.google-apps.document",
        "application/vnd.ms-word",
        "application/vnd.ms-excel",
        "image/png",
        "image/jpeg",
        "text/plain",
        "invalid/mime-type",
        "",
      ];

      invalidMimeTypes.forEach(mimeType => {
        Utils.assertFunctionParams(
          lib.validateMimeType,
          [mimeType],
          true, 
          `El valor de MimeType: ${mimeType} no es válido.`
        );
      });
    });

    Utils.it("lanza error con mensaje descriptivo", () => {
      const invalidMimeType = "application/pdf";
      
      Utils.assertFunctionParams(
        lib.validateMimeType,
        [invalidMimeType],
        true,
        `El valor de MimeType: ${invalidMimeType} no es válido`
      );
    });

    Utils.it("lanza error para valores null", () => {
      Utils.assertFunctionParams(
        lib.validateMimeType,
        [null],
        true,
        "El valor de MimeType: null no es válido"
      );
    });

    Utils.it("lanza error para valores undefined", () => {
      Utils.assertFunctionParams(
        lib.validateMimeType,
        [undefined],
        true,
        "El valor de MimeType: undefined no es válido"
      );
    });

    Utils.it("lanza error para valores no string", () => {
      const nonStringValues = [
        123,
        true,
        false,
        {},
        [],
        function() {}
      ];

      nonStringValues.forEach(value => {
        Utils.assertFunctionParams(
          lib.validateMimeType,
          [value],
          true,
          `El valor de MimeType: ${value} no es válido`
        );
      });
    });

    Utils.it("lanza error para string vacío", () => {
      Utils.assertFunctionParams(
        lib.validateMimeType,
        [""],
        true,
        "El valor de MimeType:  no es válido"
      );
    });

    Utils.it("lanza error para string con solo espacios", () => {
      Utils.assertFunctionParams(
        lib.validateMimeType,
        ["   "],
        true,
        "El valor de MimeType:     no es válido"
      );
    });

    Utils.it("es case sensitive si ALLOWED_MIME_TYPES lo es", () => {
      const upperCaseMimeType = "TEXT/CSV";
      const mixedCaseMimeType = "Application/Vnd.Google-Apps.Spreadsheet";
      
      // Asumiendo que la validación es case sensitive (lo usual con MIME types)
      Utils.assertFunctionParams(
        lib.validateMimeType,
        [upperCaseMimeType],
        true,
        `El valor de MimeType: ${upperCaseMimeType} no es válido`
      );

      Utils.assertFunctionParams(
        lib.validateMimeType,
        [mixedCaseMimeType],
        true,
        `El valor de MimeType: ${mixedCaseMimeType} no es válido`
      );
    });

    Utils.it("maneja MIME types con parámetros correctamente", () => {
      const mimeTypeWithParams = "text/csv; charset=utf-8";
      
      Utils.assertFunctionParams(
        lib.validateMimeType,
        [mimeTypeWithParams],
        true, 
        `El valor de MimeType: ${mimeTypeWithParams} no es válido`
      );
    });
  });

  //----------------------------------------------
  //
  // Tests para validateProcessingConfig
  //
  //----------------------------------------------

  Utils.describe("validateProcessingConfig", () => {
    
    function createValidConfig() {
      return {
        columns: [0, 1, 2],
        searchingFn: function() { return []; },
        processingFn: function(data) { return data; },
        updateExistingRows: false,
        ignoreEmptyRows: true,
        allowFileReprocessing: false
      };
    }

    Utils.it("no lanza error para configuración válida", () => {
      const validConfig = createValidConfig();
      
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [validConfig],
        false, 
        "Configuración válida no debería lanzar error"
      );
    });

    Utils.it("lanza error cuando searchingFn está missing", () => {
      const configWithoutSearchingFn = createValidConfig();
      delete configWithoutSearchingFn.searchingFn;
      
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [configWithoutSearchingFn],
        true, 
        "Debe proporcionar una función para obtener los archivos a procesar"
      );
    });

    Utils.it("lanza error cuando searchingFn es null", () => {
      const configWithNullSearchingFn = {
        ...createValidConfig(),
        searchingFn: null
      };
      
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [configWithNullSearchingFn],
        true,
        "Debe proporcionar una función para obtener los archivos a procesar"
      );
    });

    Utils.it("lanza error cuando searchingFn es undefined", () => {
      const configWithUndefinedSearchingFn = {
        ...createValidConfig(),
        searchingFn: undefined
      };
      
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [configWithUndefinedSearchingFn],
        true,
        "Debe proporcionar una función para obtener los archivos a procesar"
      );
    });

    Utils.it("lanza error cuando searchingFn no es función", () => {
      const nonFunctionValues = [
        "not-a-function",
        123,
        true,
        false,
        {},
        [],
        null,
        undefined
      ];

      nonFunctionValues.forEach(value => {
        const invalidConfig = {
          ...createValidConfig(),
          searchingFn: value
        };
        
        Utils.assertFunctionParams(
          lib.validateProcessingConfig,
          [invalidConfig],
          true,
          "Debe proporcionar una función para obtener los archivos a procesar"
        );
      });
    });

    Utils.it("no valida otras propiedades cuando searchingFn es válido", () => {
      const configWithInvalidColumns = {
        ...createValidConfig(),
        columns: "invalid" 
      };

      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [configWithInvalidColumns],
        false,
        "Solo debería validar searchingFn, no otras propiedades"
      );
    });

    Utils.it("funciona con configuraciones mínimas válidas", () => {
      const minimalConfig = {
        searchingFn: function() { return []; }
        // otras propiedades son opcionales
      };
      
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [minimalConfig],
        false,
        "Configuración mínima con searchingFn debería ser válida"
      );
    });

    Utils.it("lanza error para config null", () => {
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [null],
        true,
        "La configuración debe ser un objeto válido"
      );
    });

    Utils.it("lanza error para config undefined", () => {
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [undefined],
        true,
        "La configuración debe ser un objeto válido"
      );
    });

    Utils.it("lanza error para config no objeto plano", () => {
      class ConfigClass {
        constructor() {
          this.searchingFn = function() {};
          this.columns = [0, 1, 2];
        }
      }

      const nonObjectValues = [
        "string",
        123,
        true,
        false,
        function() {},
        [],
        new Map(),
        new Date(),
        /test/,
        new ConfigClass(),
      ];

      nonObjectValues.forEach(value => {
        Utils.assertFunctionParams(
          lib.validateProcessingConfig,
          [value],
          true,
          "La configuración debe ser un objeto válido"
        );
      });
    });

    Utils.it("acepta funciones con diferentes firmas", () => {
      const functionSignatures = [
        function() { return []; },
        function(mimeType) { return []; },
        () => [],
        function namedFunction() { return []; }
      ];

      functionSignatures.forEach(searchingFn => {
        const config = {
          ...createValidConfig(),
          searchingFn: searchingFn
        };
        
        Utils.assertFunctionParams(
          lib.validateProcessingConfig,
          [config],
          false,
          `Función con firma ${searchingFn.name || 'anonymous'} debería ser válida`
        );
      });
    });

    Utils.it("maneja funciones que devuelven diferentes tipos", () => {
      const functionsWithDifferentReturns = [
        function() { return []; },
        function() { return null; },
        function() { return undefined; },
        function() { return "invalid"; }
      ];

      functionsWithDifferentReturns.forEach(searchingFn => {
        const config = {
          ...createValidConfig(),
          searchingFn: searchingFn
        };
        
        Utils.assertFunctionParams(
          lib.validateProcessingConfig,
          [config],
          false,
          "La validación solo debería chequear que sea función, no el return value"
        );
      });
    });
  });

  Utils.describe("validateProcessingConfig - Edge Cases", () => {

    Utils.it("rechaza objetos que tienen searchingFn como propiedad heredada", () => {
      function ConfigWithInheritedFn() {}
      ConfigWithInheritedFn.prototype.searchingFn = function() { return []; };
      
      const configInstance = new ConfigWithInheritedFn();
      configInstance.columns = [0, 1, 2];
  
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [configInstance],
        true,
        "La configuración debe ser un objeto válido"
      );
    });

    Utils.it("rechaza objetos que tienen toString como searchingFn", () => {
      const trickyObject = {
        searchingFn: "I'm not a function, I'm a string!",
        toString: function() { return []; } 
      };
      
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [trickyObject],
        true,
        "Debe proporcionar una función para obtener los archivos a procesar"
      );
    });

    Utils.it("mantiene el mensaje de error original", () => {
      const configWithoutSearchingFn = {
        columns: [0, 1, 2],
        // searchingFn missing
      };
      
      Utils.assertFunctionParams(
        lib.validateProcessingConfig,
        [configWithoutSearchingFn],
        true,
        "Debe proporcionar una función para obtener los archivos a procesar."
      );
    });
  });

  //----------------------------------------------
  //
  // Tests para normalizeProcessingConfig
  //
  //----------------------------------------------

  Utils.describe("normalizeProcessingConfig", () => {
    
    const UPDATE_EXISTING_ROWS = false;
    const IGNORE_EMPTY_ROWS = true;
    const ALLOW_FILE_REPROCESSING = false;

    const sampleConfig = {
      columns: [0, 1, 2],
      searchingFn: function() { return []; },
      processingFn: function(data) { return data; },
      sourceSheetName: "Data",
      updateExistingRows: true,
      ignoreEmptyRows: false,
      allowFileReprocessing: true,
      sortCriteria: [{ column: 1, ascending: true }]
    };

    Utils.it("aplica valores por defecto cuando config es undefined", () => {
      const result = lib.normalizeProcessingConfig(undefined);
      
      Utils.assertEquals(result.columns, []);
      Utils.assertEquals(result.updateExistingRows, UPDATE_EXISTING_ROWS);
      Utils.assertEquals(result.ignoreEmptyRows, IGNORE_EMPTY_ROWS);
      Utils.assertEquals(result.allowFileReprocessing, ALLOW_FILE_REPROCESSING);
      Utils.assertEquals(result.sourceSheetName, undefined);
      Utils.assertEquals(result.processingFn, undefined);
      Utils.assertEquals(result.searchingFn, undefined);
      Utils.assertEquals(result.sortCriteria, undefined);
    });

    Utils.it("aplica valores por defecto cuando config es null", () => {
      const result = lib.normalizeProcessingConfig(null);
      
      Utils.assertEquals(result.columns, []);
      Utils.assertEquals(result.updateExistingRows, UPDATE_EXISTING_ROWS);
      Utils.assertEquals(result.ignoreEmptyRows, IGNORE_EMPTY_ROWS);
      Utils.assertEquals(result.allowFileReprocessing, ALLOW_FILE_REPROCESSING);
      Utils.assertEquals(result.sourceSheetName, undefined);
      Utils.assertEquals(result.processingFn, undefined);
      Utils.assertEquals(result.searchingFn, undefined);
      Utils.assertEquals(result.sortCriteria, undefined);
    });

    Utils.it("aplica valores por defecto cuando config es objeto vacío", () => {
      const result = lib.normalizeProcessingConfig({});
      
      Utils.assertEquals(result.columns, []);
      Utils.assertEquals(result.updateExistingRows, UPDATE_EXISTING_ROWS);
      Utils.assertEquals(result.ignoreEmptyRows, IGNORE_EMPTY_ROWS);
      Utils.assertEquals(result.allowFileReprocessing, ALLOW_FILE_REPROCESSING);
      Utils.assertEquals(result.sourceSheetName, undefined);
      Utils.assertEquals(result.processingFn, undefined);
      Utils.assertEquals(result.searchingFn, undefined);
      Utils.assertEquals(result.sortCriteria, undefined);
    });

    Utils.it("mantiene valores proporcionados cuando están definidos", () => {
      const result = lib.normalizeProcessingConfig(sampleConfig);
      
      Utils.assertEquals(result.columns, [0, 1, 2]);
      Utils.assertEquals(result.updateExistingRows, true);
      Utils.assertEquals(result.ignoreEmptyRows, false);
      Utils.assertEquals(result.allowFileReprocessing, true);
      Utils.assertEquals(result.sourceSheetName, "Data");
      Utils.assertEquals(result.sortCriteria, [{ column: 1, ascending: true }]);
      Utils.assertTrue(typeof result.searchingFn === 'function');
      Utils.assertTrue(typeof result.processingFn === 'function');
    });

    Utils.it("maneja propiedades parcialmente definidas", () => {
      const partialConfig = {
        columns: [3, 4],
        updateExistingRows: true,
        // omitir otras propiedades
      };

      const result = lib.normalizeProcessingConfig(partialConfig);
      
      Utils.assertEquals(result.columns, [3, 4]);
      Utils.assertEquals(result.updateExistingRows, true);
      Utils.assertEquals(result.ignoreEmptyRows, IGNORE_EMPTY_ROWS);
      Utils.assertEquals(result.allowFileReprocessing, ALLOW_FILE_REPROCESSING); 
      Utils.assertEquals(result.sourceSheetName, undefined);
      Utils.assertEquals(result.processingFn, undefined);
      Utils.assertEquals(result.searchingFn, undefined);
      Utils.assertEquals(result.sortCriteria, undefined);
    });

    Utils.it("maneja propiedades con valores falsy correctamente", () => {
      const falsyConfig = {
        columns: [],
        updateExistingRows: false,
        ignoreEmptyRows: false,
        allowFileReprocessing: false,
        sourceSheetName: "",
        sortCriteria: []
      };

      const result = lib.normalizeProcessingConfig(falsyConfig);
      
      Utils.assertEquals(result.columns, []);
      Utils.assertEquals(result.updateExistingRows, false);
      Utils.assertEquals(result.ignoreEmptyRows, false);
      Utils.assertEquals(result.allowFileReprocessing, false);
      Utils.assertEquals(result.sourceSheetName, "");
      Utils.assertEquals(result.sortCriteria, []);
      Utils.assertEquals(result.processingFn, undefined);
      Utils.assertEquals(result.searchingFn, undefined);
    });

    Utils.it("maneja funciones null/undefined correctamente", () => {
      const configWithNullFunctions = {
        columns: [1],
        searchingFn: null,
        processingFn: undefined,
        updateExistingRows: true
      };

      const result = lib.normalizeProcessingConfig(configWithNullFunctions);
      
      Utils.assertEquals(result.columns, [1]);
      Utils.assertEquals(result.updateExistingRows, true);
      Utils.assertEquals(result.ignoreEmptyRows, IGNORE_EMPTY_ROWS);
      Utils.assertEquals(result.allowFileReprocessing, ALLOW_FILE_REPROCESSING);
      Utils.assertEquals(result.searchingFn, null);
      Utils.assertEquals(result.processingFn, undefined);
    });

    Utils.it("preserva la identidad de las funciones", () => {
      const testFunction = function() { return "test"; };
      
      const configWithFunction = {
        columns: [1],
        processingFn: testFunction
      };

      const result = lib.normalizeProcessingConfig(configWithFunction);
      
      Utils.assertEquals(result.processingFn, testFunction);
      Utils.assertEquals(result.processingFn(), "test");
    });

    Utils.it("maneja arrays complejos en columns", () => {
      const complexColumnsConfig = {
        columns: [0, "fixed_value", 2, "=A1+B1"]
      };

      const result = lib.normalizeProcessingConfig(complexColumnsConfig);
      
      Utils.assertEquals(result.columns, [0, "fixed_value", 2, "=A1+B1"]);
    });

    Utils.it("maneja criterios de ordenamiento complejos", () => {
      const complexSortConfig = {
        columns: [0, 1],
        sortCriteria: [
          { column: 1, ascending: true },
          { column: 2, ascending: false },
          { column: 3, ascending: true }
        ]
      };

      const result = lib.normalizeProcessingConfig(complexSortConfig);

      Utils.assertEquals(result.sortCriteria.length, 3);
      Utils.assertEquals(result.sortCriteria[0], { column: 1, ascending: true });
      Utils.assertEquals(result.sortCriteria[1], { column: 2, ascending: false });
      Utils.assertEquals(result.sortCriteria[2], { column: 3, ascending: true });
    });

    Utils.it("devuelve nuevo objeto y no modifica el original", () => {
      const originalConfig = { columns: [1, 2, 3] };
      const result = lib.normalizeProcessingConfig(originalConfig);
      
      Utils.assertFalse(result === originalConfig);
      
      // Modificar el resultado no debería afectar el original
      result.columns.push(4);

      Utils.assertEquals(originalConfig.columns, [1, 2, 3]);
      Utils.assertEquals(result.columns, [1, 2, 3, 4]);
    });

    Utils.it("maneja propiedades undefined explícitamente", () => {
      const configWithUndefined = {
        columns: undefined,
        updateExistingRows: undefined,
        ignoreEmptyRows: undefined,
        allowFileReprocessing: undefined,
        sourceSheetName: undefined,
        processingFn: undefined,
        searchingFn: undefined,
        sortCriteria: undefined
      };

      const result = lib.normalizeProcessingConfig(configWithUndefined);

      Utils.assertEquals(result.columns, []);
      Utils.assertEquals(result.updateExistingRows, UPDATE_EXISTING_ROWS);
      Utils.assertEquals(result.ignoreEmptyRows, IGNORE_EMPTY_ROWS);
      Utils.assertEquals(result.allowFileReprocessing, ALLOW_FILE_REPROCESSING);
      Utils.assertEquals(result.sourceSheetName, undefined);
      Utils.assertEquals(result.processingFn, undefined);
      Utils.assertEquals(result.searchingFn, undefined);
      Utils.assertEquals(result.sortCriteria, undefined);
    });
  });

  Utils.describe("normalizeProcessingConfig - Edge Cases", () => {
    
    const UPDATE_EXISTING_ROWS = false;
    const IGNORE_EMPTY_ROWS = true;
    const ALLOW_FILE_REPROCESSING = false;

    Utils.it("maneja propiedades con valores por defecto personalizados", () => {
      const emptyConfig = {};
      const result = lib.normalizeProcessingConfig(emptyConfig);
      
      Utils.assertEquals(result.updateExistingRows, UPDATE_EXISTING_ROWS);
      Utils.assertEquals(result.ignoreEmptyRows, IGNORE_EMPTY_ROWS);
      Utils.assertEquals(result.allowFileReprocessing, ALLOW_FILE_REPROCESSING);
    });

    Utils.it("maneja propiedades con valores null específicos", () => {
      const configWithNulls = {
        columns: null,
        sortCriteria: null
      };

      const result = lib.normalizeProcessingConfig(configWithNulls);
      
      Utils.assertEquals(result.columns, []); 
      Utils.assertEquals(result.sortCriteria, null); 
    });

    Utils.it("mantiene referencias a objetos complejos", () => {
      const complexObject = { complex: "object" };
      const configWithComplex = {
        columns: [1],
        sortCriteria: complexObject
      };

      const result = lib.normalizeProcessingConfig(configWithComplex);
      
      Utils.assertEquals(result.sortCriteria, complexObject);
      Utils.assertTrue(result.sortCriteria === complexObject); 
    });
  });


  //----------------------------------------------
  //
  // Tests para normalizeProcessingConfigMap
  //
  //----------------------------------------------

  Utils.describe("normalizeProcessingConfigMap", () => {
    
    function createValidConfig() {
      return {
        columns: [0, 1, 2],
        searchingFn: function() { return []; }
      };
    }

    Utils.it("maneja configMap null", () => {
      const result = lib.normalizeProcessingConfigMap(null);
      Utils.assertEquals(result, {});
    });

    Utils.it("maneja configMap undefined", () => {
      const result = lib.normalizeProcessingConfigMap(undefined);
      Utils.assertEquals(result, {});
    });

    Utils.it("maneja configMap vacío", () => {
      const result = lib.normalizeProcessingConfigMap({});
      Utils.assertEquals(result, {});
    });

    Utils.it("normaliza configMap con una entrada", () => {
      const configMap = {
        "text/csv": createValidConfig()
      };

      const result = lib.normalizeProcessingConfigMap(configMap);
      
      Utils.assertEquals(Object.keys(result).length, 1);
      Utils.assertTrue("text/csv" in result);
      Utils.assertTrue(result["text/csv"].hasOwnProperty("columns"));
      Utils.assertTrue(result["text/csv"].hasOwnProperty("searchingFn"));
      Utils.assertEquals(result["text/csv"].ignoreEmptyRows, true); 
      Utils.assertEquals(result["text/csv"].updateExistingRows, false); 
    });

    Utils.it("normaliza configMap con múltiples entradas", () => {
      const configMap = {
        "text/csv": createValidConfig(),
        "application/vnd.google-apps.spreadsheet": createValidConfig(),
        "application/vnd.ms-excel": createValidConfig()
      };

      const result = lib.normalizeProcessingConfigMap(configMap);
      
      Utils.assertEquals(Object.keys(result).length, 3);
      Utils.assertTrue("text/csv" in result);
      Utils.assertTrue("application/vnd.google-apps.spreadsheet" in result);
      Utils.assertTrue("application/vnd.ms-excel" in result);
      
      Object.values(result).forEach(config => {
        Utils.assertTrue(config.hasOwnProperty("ignoreEmptyRows"));
        Utils.assertTrue(config.hasOwnProperty("updateExistingRows"));
        Utils.assertTrue(config.hasOwnProperty("allowFileReprocessing"));
      });
    });

    Utils.it("mantiene las claves originales del configMap", () => {
      const configMap = {
        "text/csv": createValidConfig(),
        "application/vnd.google-apps.spreadsheet": createValidConfig(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": createValidConfig()
      };

      const result = lib.normalizeProcessingConfigMap(configMap);
      
      const expectedKeys = [
        "text/csv",
        "application/vnd.google-apps.spreadsheet",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ];
      
      Utils.assertEquals(Object.keys(result).sort(), expectedKeys.sort());
    });

    Utils.it("aplica valores por defecto a configuraciones parciales", () => {
      const partialConfig = {
        columns: [0, 1] 
        // searchingFn y otros no definidos
      };

      const configMap = {
        "text/csv": partialConfig
      };

      const result = lib.normalizeProcessingConfigMap(configMap);
      const normalizedConfig = result["text/csv"];
      
      Utils.assertEquals(normalizedConfig.ignoreEmptyRows, true);
      Utils.assertEquals(normalizedConfig.updateExistingRows, false);
      Utils.assertEquals(normalizedConfig.allowFileReprocessing, false);
      Utils.assertEquals(normalizedConfig.columns, [0, 1]);
    });

    Utils.it("devuelve un nuevo objeto, no una referencia", () => {
      const originalConfig = createValidConfig();
      const configMap = {
        "text/csv": originalConfig
      };

      const result = lib.normalizeProcessingConfigMap(configMap);
      
      Utils.assertFalse(result === configMap);
      Utils.assertFalse(result["text/csv"] === originalConfig);

      result["newKey"] = "test";
      Utils.assertFalse("newKey" in configMap);
      
      result["text/csv"].newProperty = "test";
      Utils.assertFalse("newProperty" in originalConfig);
    });

    Utils.it("maneja claves de MIME type inválidas", () => {
      const configMap = {
        "": createValidConfig(),
        "invalid-mime-type": createValidConfig(),
        "text/csv": createValidConfig()
      };

      const result = lib.normalizeProcessingConfigMap(configMap);
      
      Utils.assertEquals(Object.keys(result).length, 3);
      Utils.assertTrue("" in result);
      Utils.assertTrue("invalid-mime-type" in result);
      Utils.assertTrue("text/csv" in result);
    });

    Utils.it("preserva funciones personalizadas", () => {
      const customFunction = function() { return "custom"; };
      const configWithCustomFn = {
        columns: [0],
        searchingFn: customFunction,
        processingFn: customFunction
      };

      const configMap = {
        "text/csv": configWithCustomFn
      };

      const result = lib.normalizeProcessingConfigMap(configMap);
      
      Utils.assertEquals(result["text/csv"].searchingFn, customFunction);
      Utils.assertEquals(result["text/csv"].processingFn, customFunction);
      Utils.assertEquals(result["text/csv"].searchingFn(), "custom");
    });
  });

  Utils.describe("normalizeProcessingConfigMap - Edge Cases", () => {
    
    Utils.it("maneja objetos con prototipo null como configMap", () => {
      const nullProtoMap = Object.create(null);
      nullProtoMap["text/csv"] = { 
        columns: [0], 
        searchingFn: function() {} 
      };
      nullProtoMap["application/vnd.ms-excel"] = { 
        columns: [0], 
        searchingFn: function() {} 
      };

      const result = lib.normalizeProcessingConfigMap(nullProtoMap);
      
      Utils.assertEquals(Object.keys(result).length, 2);
      Utils.assertTrue("text/csv" in result);
      Utils.assertTrue("application/vnd.ms-excel" in result);
      Utils.assertTrue(result["text/csv"].hasOwnProperty("ignoreEmptyRows"));
      Utils.assertTrue(result["application/vnd.ms-excel"].hasOwnProperty("updateExistingRows"));
    });

    Utils.it("maneja configMap con propiedades no enumerables", () => {
      const configMap = {
        "text/csv": { 
          columns: [0], 
          searchingFn: function() {} 
        }
      };

      Object.defineProperty(configMap, "nonEnumerable", {
        value: { columns: [1], searchingFn: function() {} },
        enumerable: false
      });

      const result = lib.normalizeProcessingConfigMap(configMap);
      
      Utils.assertEquals(Object.keys(result).length, 1);
      Utils.assertTrue("text/csv" in result);
      Utils.assertFalse("nonEnumerable" in result);
    });

    Utils.it("maneja configuraciones con valores null/undefined", () => {
      const configMap = {
        "type1": { columns: [0], searchingFn: function() {} },
        "type2": null,
        "type3": undefined,
        "type4": { columns: [0], searchingFn: function() {} }
      };

      const result = lib.normalizeProcessingConfigMap(configMap);
      
      Utils.assertEquals(Object.keys(result).length, 4);
      Utils.assertTrue("type1" in result);
      Utils.assertTrue("type2" in result);
      Utils.assertTrue("type3" in result);
      Utils.assertTrue("type4" in result);
      
      Utils.assertTrue(result["type2"].hasOwnProperty("columns"));
      Utils.assertTrue(result["type2"].hasOwnProperty("searchingFn"));
      Utils.assertTrue(result["type3"].hasOwnProperty("columns"));
      Utils.assertTrue(result["type3"].hasOwnProperty("searchingFn"));
    });
  });

  //----------------------------------------------
  //
  // Tests para detectCSVSeparator
  //
  //----------------------------------------------

  Utils.describe("detectCSVSeparator", () => {
    
    Utils.it("detecta punto y coma como separador principal", () => {
      const line = "nombre;apellido;edad;ciudad";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ";");
    });

    Utils.it("detecta coma como separador principal", () => {
      const line = "nombre,apellido,edad,ciudad";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ",");
    });

    Utils.it("detecta tabulador como separador principal", () => {
      const line = "nombre\tapellido\tedad\tciudad";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, "\t");
    });

    Utils.it("detecta pipe como separador principal", () => {
      const line = "nombre|apellido|edad|ciudad";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, "|");
    });

    Utils.it("prefiere el separador con mayor conteo", () => {
      const line = "nombre,apellido;edad;ciudad;pais";
      // 3 punto y coma vs 1 coma - debería preferir ;
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ";");
    });

    Utils.it("funciona con mixed separators", () => {
      const line = "field1,field2;field3,field4|field5";
      // Conteos: , = 2, ; = 1, | = 1 - debería preferir ,
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ",");
    });

    Utils.it("maneja empates correctamente (entonces primero en array)", () => {
      const line = "field1,field2;field3,field4;field5";
      // Conteos: , = 2, ; = 2 - debería preferir ; porque viene primero en el array
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ";");
    });

    Utils.it("funciona con líneas que contienen los separadores como datos", () => {
      const line = 'nombre;"apellido;con punto y coma";edad;ciudad';
      // Los ; dentro de comillas no deberían contar, pero la función actual los cuenta
      // Este test muestra el comportamiento actual
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ";");
    });

    Utils.it("maneja líneas vacías", () => {
      const line = "";
      const result = lib.detectCSVSeparator(line);
      // Debería devolver el primer separador del array (;)
      Utils.assertEquals(result, ";");
    });

    Utils.it("maneja líneas sin separadores", () => {
      const line = "solouncampo";
      const result = lib.detectCSVSeparator(line);
      // Debería devolver el primer separador del array (;)
      Utils.assertEquals(result, ";");
    });

    Utils.it("maneja líneas con espacios alrededor de separadores", () => {
      const line = "nombre , apellido ; edad , ciudad";
      // Debería contar los separadores correctamente ignorando espacios
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ",");
    });

    Utils.it("funciona con separadores al inicio/final de línea", () => {
      const line = ";nombre;apellido;edad;";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ";");
    });

    Utils.it("maneja separadores repetidos consecutivamente", () => {
      const line = "nombre;;apellido;;;edad";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ";");
    });

    Utils.it("es case sensitive con los separadores", () => {
      const line = "field1,field2,FIELD3";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ",");
    });

    Utils.it("funciona con números que contienen separadores", () => {
      const line = "1000,2000;3000|4000";
      // Conteos: , = 1, ; = 1, | = 1 - debería preferir ; (primero en array)
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ";");
    });

    Utils.it("maneja líneas muy largas", () => {
      const longLine = "field1,field2,field3,field4,field5,field6,field7,field8,field9,field10";
      const result = lib.detectCSVSeparator(longLine);
      Utils.assertEquals(result, ",");
    });
  });

  Utils.describe("detectCSVSeparator - Edge Cases", () => {
    
    Utils.it("maneja caracteres especiales regex correctamente", () => {
      // Los separadores como . * + ? ^ $ { } ( ) | [ ] \ necesitan escape en regex
      const line = "field1|field2|field3";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, "|");
    });

    Utils.it("funciona con strings que contienen regex special chars", () => {
      const line = "field1.field2;field3*field4,field5?field6";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ";"); // ; tiene 1, , tiene 1 - ; gana por orden array
    });

    Utils.it("maneja unicode y caracteres especiales", () => {
      const line = "nombreñ,apellidó;edadé,ciudadú";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, ","); 
    });

    Utils.it("funciona con solo un tipo de separador presente", () => {
      const line = "just|one|separator|type";
      const result = lib.detectCSVSeparator(line);
      Utils.assertEquals(result, "|");
    });

    Utils.it("maneja líneas con solo separadores", () => {
      const line = ";;,|,;";
      const result = lib.detectCSVSeparator(line);
      // Conteos: ; = 3, , = 2, | = 1 - debería preferir ;
      Utils.assertEquals(result, ";");
    });
  });

  Utils.describe("detectCSVSeparator - Orden de preferencia", () => {
    
    Utils.it("respeta el orden del array en caso de empate", () => {
      const separators = [';', ',', '\t', '|'];
      
      const line = "a;b,c\td|e";
      const result = lib.detectCSVSeparator(line);
      // Todos tienen count = 1, debería preferir ; (primero en array)
      Utils.assertEquals(result, ";");
    });

    Utils.it("el orden del array define la prioridad", () => {
      const testLine1 = "a;b,c"; // ; = 1, , = 1 - ; debería ganar (primero en array)
      const testLine2 = "a,b;c"; // , = 1, ; = 1 - ; debería ganar (primero en array)
      
      Utils.assertEquals(lib.detectCSVSeparator(testLine1), ";");
      Utils.assertEquals(lib.detectCSVSeparator(testLine2), ";");
    });

    Utils.it("funciona con el orden actual de separators", () => {
      // Orden: [';', ',', '\t', '|']
      const line1 = "a;b,c"; 
      const line2 = "a,b|c"; 
      const line3 = "a\tb|c"; 
      
      Utils.assertEquals(lib.detectCSVSeparator(line1), ";");
      Utils.assertEquals(lib.detectCSVSeparator(line2), ",");
      Utils.assertEquals(lib.detectCSVSeparator(line3), "\t");
    });
  });

  //----------------------------------------------
  //
  // Tests para parseCSVLine
  //
  //----------------------------------------------

  Utils.describe("parseCSVLine", () => {
    
    Utils.it("lanza error para separator inválido", () => {
      Utils.assertFunctionParams(
        lib.parseCSVLine,
        ["test", null],
        true,
        "El parámetro 'separator' no es válido"
      );
      
      Utils.assertFunctionParams(
        lib.parseCSVLine,
        ["test", undefined],
        true,
        "El parámetro 'separator' no es válido"
      );
      
      Utils.assertFunctionParams(
        lib.parseCSVLine,
        ["test", ""],
        true,
        "El parámetro 'separator' no es válido"
      );
    });

    Utils.it("parsea línea simple con coma", () => {
      const line = "nombre,apellido,edad";
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["nombre", "apellido", "edad"]);
    });

    Utils.it("parsea línea simple con punto y coma", () => {
      const line = "nombre;apellido;edad";
      const result = lib.parseCSVLine(line, ";");
      Utils.assertEquals(result, ["nombre", "apellido", "edad"]);
    });

    Utils.it("parsea línea simple con tabulador", () => {
      const line = "nombre\tapellido\tedad";
      const result = lib.parseCSVLine(line, "\t");
      Utils.assertEquals(result, ["nombre", "apellido", "edad"]);
    });

    Utils.it("parsea línea simple con pipe", () => {
      const line = "nombre|apellido|edad";
      const result = lib.parseCSVLine(line, "|");
      Utils.assertEquals(result, ["nombre", "apellido", "edad"]);
    });

    Utils.it("maneja campos entre comillas simples", () => {
      const line = '"nombre","apellido","edad"';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["nombre", "apellido", "edad"]);
    });

    Utils.it("maneja campos con separadores dentro de comillas", () => {
      const line = '"nombre,completo","apellido;con;puntos","edad"';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["nombre,completo", "apellido;con;puntos", "edad"]);
    });

    Utils.it("maneja comillas dobles escapadas", () => {
      const line = '"nombre ""especial""","apellido","edad"';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ['nombre "especial"', "apellido", "edad"]);
    });

    Utils.it("maneja mezcla de campos con y sin comillas", () => {
      const line = 'nombre,"apellido,completo",edad';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["nombre", "apellido,completo", "edad"]);
    });

    Utils.it("maneja campos vacíos", () => {
      const line = "nombre,,edad";
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["nombre", "", "edad"]);
    });

    Utils.it("maneja campos vacíos entre comillas", () => {
      const line = 'nombre,"",edad';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["nombre", "", "edad"]);
    });

    Utils.it("maneja línea vacía", () => {
      const line = "";
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, [""]);
    });

    Utils.it("maneja sólo separadores", () => {
      const line = ",,,";
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["", "", "", ""]);
    });

    Utils.it("maneja espacios alrededor de campos", () => {
      const line = " nombre , apellido , edad ";
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, [" nombre ", " apellido ", " edad "]);
    });

    Utils.it("maneja comillas no cerradas", () => {
      const line = '"nombre,apellido,edad';
      const result = lib.parseCSVLine(line, ",");
      // Comportamiento esperado de recuperación:
      // Al no encontrar una comilla de cierre, toma toda la cadena como un solo valor 
      Utils.assertEquals(result, ['"nombre,apellido,edad']);
    });

    Utils.it("maneja comillas no abiertas", () => {
      const line = 'nombre",apellido,edad';
      const result = lib.parseCSVLine(line, ",");
      // Comportamiento esperado de recuperación:
      // Al no existir una comilla de apertura, la toma como parte del valor del campo 
      Utils.assertEquals(result, ['nombre"','apellido','edad']);
    });

    Utils.it("maneja comillas al inicio pero no al final", () => {
      const line = '"nombre,apellido",edad';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ['nombre,apellido', 'edad']);
    });

    Utils.it("maneja comillas al final pero no al inicio", () => {
      const line = 'nombre,"apellido"';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ['nombre', 'apellido']);
    });

    Utils.it("maneja múltiples comillas dentro de campos", () => {
      const line = '"nombre""test""","apellido","edad"';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ['nombre"test"', "apellido", "edad"]);
    });

    Utils.it("maneja separadores que son regex special chars", () => {
      const line = "a.b.c";
      const result = lib.parseCSVLine(line, ".");
      Utils.assertEquals(result, ["a", "b", "c"]);
    });

    Utils.it("maneja separadores con múltiples caracteres", () => {
      const line = "a::b::c";
      const result = lib.parseCSVLine(line, "::");
      Utils.assertEquals(result, ["a", "b", "c"]);
    });
  });

  Utils.describe("parseCSVLine - Casos Complejos", () => {
    
    Utils.it("maneja campos con comillas y separadores mezclados", () => {
      const line = 'nombre,"apellido,con;varios|separadores",edad';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["nombre", "apellido,con;varios|separadores", "edad"]);
    });

    Utils.it("maneja comillas dentro de campos no entrecomillados", () => {
      const line = 'nombre,apellido"con comilla",edad';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["nombre", 'apellido"con comilla"', "edad"]);
    });

    Utils.it("maneja línea con sólo comillas", () => {
      const line = '""';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, [""]);
    });

    Utils.it("maneja múltiples comillas consecutivas - formato válido (par)", () => {
        // 6 comillas dobles: representan "" (dos comillas) dentro de un campo entrecomillado
        const line = '"""""","test"';
        const result = lib.parseCSVLine(line, ",");
        Utils.assertEquals(result, ['"', "test"]);
    });

    Utils.it("maneja múltiples comillas consecutivas - formato inválido (impar)", () => {
        // 5 comillas dobles: formato inválido pero la función se recupera
        const line = '""""","test"';
        const result = lib.parseCSVLine(line, ",");
        
        // Comportamiento esperado de recuperación:
        // - La primera comilla marca el inicio de un campo. 
        // - Las 4 comillas siguientes se interpretan como 2 comillas escapadas → ""
        // - La coma se toma como parte del campo porque todavía no se ha encontrado
        //   la comilla de cierre
        // - Como en el resto de la cadena no existe otra coma con comilla previa
        //   que indique un nuevo campo, se adiciona el resto como parte del campo
        //   que quedaría como un campo entrecomillado '"","test"'
        // - Al eliminar las comillas envolventes quedaría '","test'
        // Resultado: ['","test']
        Utils.assertEquals(result, ['","test']);
    });

    Utils.it("maneja diversos casos de comillas consecutivas", () => {
        // 1 comilla
        // Comportamiento esperado de recuperación:
        // - Toma la comilla como parte del valor pues no encuentra la de cierre
        Utils.assertEquals(lib.parseCSVLine('"', ","), ['"']);
        
        // 2 comillas
        // Comportamiento esperado de recuperación:
        // - Toma el campo como un campo entrecomillado, elimina las comillas
        Utils.assertEquals(lib.parseCSVLine('""', ","), ['']);
        
        // 3 comillas
        // Comportamiento esperado de recuperación:
        // - Toma la primera comilla como la de apertura del valor (valor: '"')
        // - La siguiente comilla la toma como una comilla interna, verifica si 
        //   es una comilla escapada analizando el siguiente caracter que también
        //   es una comilla y las sustituye por una sola (valor: '""')
        // - Elimina las comillas por ser un campo entrecomillado (valor: '')
        Utils.assertEquals(lib.parseCSVLine('"""', ","), ['']);
        
        // 4 comillas
        // Comportamiento esperado de recuperación:
        // - Toma la primera comilla como la de apertura del valor (valor: '"')
        // - La siguiente comilla la toma como una comilla interna, verifica si 
        //   es una comilla escapada analizando el siguiente caracter que también
        //   es una comilla y las sustituye por una sola (valor: '""')
        // - La siguiente comilla la toma como la de cierre (valor: '"""')
        // - Elimina las comillas por ser un campo entrecomillado (valor: '"')
        Utils.assertEquals(lib.parseCSVLine('""""', ","), ['"']);
        
        // 5 comillas
        // Comportamiento esperado de recuperación:
        // - Toma la primera comilla como la de apertura del valor (valor: '"')
        // - La siguiente comilla la toma como una comilla interna, verifica si 
        //   es una comilla escapada analizando el siguiente caracter que también
        //   es una comilla y las sustituye por una sola (valor: '""')
        // - La siguiente comilla la toma como interna, verifica si 
        //   es una comilla escapada analizando el siguiente caracter que también
        //   es una comilla y las sustituye por una sola (valor: '"""')
        // - No encuentra la comilla del cierre (valor: '"""')
        // - Elimina las comillas por ser un campo entrecomillado (valor: '"')
        Utils.assertEquals(lib.parseCSVLine('"""""', ","), ['"']);
        
        // 6 comillas
        // Comportamiento esperado de recuperación:
        // - Toma la primera comilla como la de apertura del valor (valor: '"')
        // - La siguiente comilla la toma como una comilla interna, verifica si 
        //   es una comilla escapada analizando el siguiente caracter que también
        //   es una comilla y las sustituye por una sola (valor: '""')
        // - La siguiente comilla la toma como interna, verifica si 
        //   es una comilla escapada analizando el siguiente caracter que también
        //   es una comilla y las sustituye por una sola (valor: '"""')
        // - La siguiente comilla la toma como la de cierre (valor: '""""')
        // - Elimina las comillas por ser un campo entrecomillado (valor: '""')
        Utils.assertEquals(lib.parseCSVLine('""""""', ","), ['"']);
    });

    Utils.it("maneja comillas al inicio y final con contenido", () => {
      const line = '"texto"';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["texto"]);
    });

    Utils.it("maneja campos con comillas y espacios", () => {
      const line = '" nombre ", " apellido " , " edad "';
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, [" nombre ", " apellido ", " edad "]);
    });

    Utils.it("maneja campos con espacios solo a la izquierda", () => {
        const line = '   "test",other';
        const result = lib.parseCSVLine(line, ",");
        Utils.assertEquals(result, ["test", "other"]);
    });

    Utils.it("maneja campos con espacios solo a la derecha", () => {
        const line = '"test"   ,other';
        const result = lib.parseCSVLine(line, ",");
        Utils.assertEquals(result, ["test", "other"]);
    });

    Utils.it("mantiene espacios en campos no entrecomillados", () => {
        const line = '  field1  ,  field2  ';
        const result = lib.parseCSVLine(line, ",");
        Utils.assertEquals(result, ["  field1  ", "  field2  "]);
    });
  });

  Utils.describe("parseCSVLine - Rendimiento", () => {
    
    Utils.it("maneja líneas largas", () => {
      const longField = "a".repeat(1000);
      const line = `${longField},${longField},${longField}`;
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, [longField, longField, longField]);
    });

    Utils.it("maneja muchos campos", () => {
      const fields = Array(100).fill("test");
      const line = fields.join(",");
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result.length, 100);
      Utils.assertEquals(result, fields);
    });

    Utils.it("maneja campos con muchos caracteres especiales", () => {
      const complexField = 'a"b"c"d"e"f"g"h"i"j"k"l"m"n"o"p"q"r"s"t"u"v"w"x"y"z';
      const line = `normal,${complexField},simple`;
      const result = lib.parseCSVLine(line, ",");
      Utils.assertEquals(result, ["normal", complexField, "simple"]);
    });
  });

  Utils.describe("parseCSVLine - Escape de Regex", () => {
    
    Utils.it("maneja separadores con caracteres regex especiales", () => {
      const regexSpecialChars = ['.', '*', '+', '?', '^', '$', '[', ']', '(', ')', '{', '}', '|', '\\'];
      
      regexSpecialChars.forEach(char => {
        const line = `a${char}b${char}c`;
        const result = lib.parseCSVLine(line, char);
        Utils.assertEquals(result, ["a", "b", "c"], `Failed with separator: ${char}`);
      });
    });

    Utils.it("maneja separadores con múltiples caracteres especiales", () => {
      const line = "a.*+?^$[](){}|\\b.*+?^$[](){}|\\c";
      const separator = ".*+?^$[](){}|\\";
      const result = lib.parseCSVLine(line, separator);
      Utils.assertEquals(result, ["a", "b", "c"]);
    });
  });

  //----------------------------------------------
  //
  // Tests para getSpreadsheetObjectsWithFallback
  //
  //----------------------------------------------

  Utils.describe("getSpreadsheetObjectsWithFallback", () => {
    Utils.it("lanza error si no se pasa spreadsheetId", () => {
      Utils.assertFunctionParams(lib.getSpreadsheetObjectsWithFallback, [null, null], true, /no ha sido proporcionado/);
    });
    Utils.it("lanza error si spreadsheetId no es válido", () => {
      const oldOpenById = SpreadsheetApp.openById;
      SpreadsheetApp.openById = (id) => {
        if (id !== "TEST_ID") throw new Error("id inesperado");
        return fakeSpreadsheet;
      };

      Utils.assertFunctionParams(lib.getSpreadsheetObjectsWithFallback, ["WRONG_ID", null], true, /no es válido/);

      SpreadsheetApp.openById = oldOpenById; 
    });
    Utils.it("lanza error si sheetName no existe", () => {
      const fakeSheet = { getName: () => "Hoja1" };
      const fakeSheets = [fakeSheet];
      const fakeSpreadsheet = { 
        getSheets: () => fakeSheets, 
        getSheetByName: (name) => fakeSheets.find(sheet => sheet.getName() === name) };
      const oldOpenById = SpreadsheetApp.openById;
      SpreadsheetApp.openById = (id) => {
        if (id !== "TEST_ID") throw new Error("id inesperado");
        return fakeSpreadsheet;
      };

      Utils.assertFunctionParams(
        lib.getSpreadsheetObjectsWithFallback, ["TEST_ID", "dummy"], true, /Hoja de destino no encontrada/);

      SpreadsheetApp.openById = oldOpenById; 
    });
    Utils.it("retorna spreadsheet, sheet y file válidos", () => {
      const fakeSheet1 = { getName: () => "Hoja1" };
      const fakeSheet2 = { getName: () => "Hoja2" };
      const fakeSheets = [fakeSheet1, fakeSheet2];
      const fakeSpreadsheet = { 
        getName: () => "FAKE_SPREADSHEET",
        getSheets: () => fakeSheets, 
        getSheetByName: (name) => fakeSheets.find(sheet => sheet.getName() === name) };
      const oldOpenById = SpreadsheetApp.openById;
      SpreadsheetApp.openById = (id) => {
        if (id !== "TEST_ID") throw new Error("id inesperado");
        return fakeSpreadsheet;
      };
      const oldGetFileById = DriveApp.getFileById;
      DriveApp.getFileById = (id) => fakeSpreadsheet;

      const result = lib.getSpreadsheetObjectsWithFallback("TEST_ID", null);
      Utils.assertEquals(result.spreadsheet, fakeSpreadsheet);
      Utils.assertEquals(result.sheet.getName(), "Hoja1");
      Utils.assertEquals(result.file.getName(), "FAKE_SPREADSHEET");
      const result2 = lib.getSpreadsheetObjectsWithFallback("TEST_ID", "Hoja2");
      Utils.assertEquals(result2.spreadsheet, fakeSpreadsheet);
      Utils.assertEquals(result2.sheet.getName(), "Hoja2");
      Utils.assertEquals(result2.file.getName(), "FAKE_SPREADSHEET");

      DriveApp.getFileById = oldGetFileById; 
      SpreadsheetApp.openById = oldOpenById; 
    });
  });

  //----------------------------------------------
  //
  // Tests para getLogOptionsWithFallback
  //
  //----------------------------------------------

  Utils.describe("getLogOptionsWithFallback", () => {
    Utils.it("lanza error si no se pasa logSpreadsheet", () => {
      Utils.assertFunctionParams(lib.getLogOptionsWithFallback, [{}], true, /no ha sido proporcionado/);
    });
    Utils.it("lanza error si logSheetName no existe", () => {
      const fakeSheet = { getName: () => "Hoja1" };
      const fakeSheets = [fakeSheet];
      const fakeSpreadsheet = { 
        getId: () => "TEST_ID",
        getSheets: () => fakeSheets, 
        getSheetByName: (name) => fakeSheets.find(sheet => sheet.getName() === name) };
      const oldGetFileById = DriveApp.getFileById;
      DriveApp.getFileById = (id) => fakeSpreadsheet;

      Utils.assertFunctionParams(lib.getLogOptionsWithFallback, [{ logSpreadsheet: fakeSpreadsheet, logSheetName: "dummy" }], true, /Hoja no encontrada/);

      DriveApp.getFileById = oldGetFileById; 
    });
    Utils.it("devuelve opciones predeterminadas", () => {
      const fakeSheet1 = { getName: () => "Hoja1" };
      const fakeSheet2 = { getName: () => "Hoja2" };
      const fakeSheets = [fakeSheet1, fakeSheet2];
      const fakeSpreadsheet = { 
        getId: () => "TEST_ID",
        getSheets: () => fakeSheets, 
        getSheetByName: (name) => fakeSheets.find(sheet => sheet.getName() === name) };
      const oldGetFileById = DriveApp.getFileById;
      DriveApp.getFileById = (id) => fakeSpreadsheet;

      const result = lib.getLogOptionsWithFallback({ logSpreadsheet: fakeSpreadsheet });
      Utils.assertEquals(result.logSheet.getName(), "Hoja1");
      Utils.assertEquals(result.logFile.getId(), "TEST_ID");
      Utils.assertEquals(JSON.stringify(result.logColumnTitles), JSON.stringify(['Fecha', 'Archivo', 'Filas', 'Estado']));
      Utils.assertEquals(result.successMessage, "Éxito");
      Utils.assertEquals(result.failureMessage, "Error procesando archivo");
      const result2 = lib.getLogOptionsWithFallback({ logSpreadsheet: fakeSpreadsheet, logSheetName: "Hoja2" });
      Utils.assertEquals(result2.logSheet.getName(), "Hoja2");
      Utils.assertEquals(result2.logFile.getId(), "TEST_ID");
      Utils.assertEquals(JSON.stringify(result.logColumnTitles), JSON.stringify(['Fecha', 'Archivo', 'Filas', 'Estado']));
      Utils.assertEquals(result2.successMessage, "Éxito");
      Utils.assertEquals(result2.failureMessage, "Error procesando archivo");

      DriveApp.getFileById = oldGetFileById; 
    });
    Utils.it("respeta opciones suministradas", () => {
      const fakeSheet1 = { getName: () => "Hoja1" };
      const fakeSheet2 = { getName: () => "Hoja2" };
      const fakeSheets = [fakeSheet1, fakeSheet2];
      const fakeSpreadsheet = { 
        getId: () => "TEST_ID",
        getSheets: () => fakeSheets, 
        getSheetByName: (name) => fakeSheets.find(sheet => sheet.getName() === name) };
      const oldGetFileById = DriveApp.getFileById;
      DriveApp.getFileById = (id) => fakeSpreadsheet;

      const result = lib.getLogOptionsWithFallback({ 
        logSpreadsheet: fakeSpreadsheet, 
        logColumnTitles: ['A', 'B', 'C', 'D'],
        successMessage: "SUCCESS",
        failureMessage: "FAILURE",
      });
      Utils.assertEquals(result.logSheet.getName(), "Hoja1");
      Utils.assertEquals(result.logFile.getId(), "TEST_ID");
      Utils.assertEquals(JSON.stringify(result.logColumnTitles), JSON.stringify(['A', 'B', 'C', 'D']));
      Utils.assertEquals(result.successMessage, "SUCCESS");
      Utils.assertEquals(result.failureMessage, "FAILURE");
      const result2 = lib.getLogOptionsWithFallback({ 
        logSpreadsheet: fakeSpreadsheet, 
        logSheetName: "Hoja2",
        logColumnTitles: ['A', 'B', 'C', 'D'],
        successMessage: "SUCCESS",
        failureMessage: "FAILURE",
      });
      Utils.assertEquals(result2.logSheet.getName(), "Hoja2");
      Utils.assertEquals(result2.logFile.getId(), "TEST_ID");
      Utils.assertEquals(JSON.stringify(result2.logColumnTitles), JSON.stringify(['A', 'B', 'C', 'D']));
      Utils.assertEquals(result2.successMessage, "SUCCESS");
      Utils.assertEquals(result2.failureMessage, "FAILURE");

      DriveApp.getFileById = oldGetFileById; 
    });
  });

  //----------------------------------------------
  //
  // Tests para getFoldersOptionsWithFallback
  //
  //----------------------------------------------

  Utils.describe("getFoldersOptionsWithFallback", () => {
    Utils.it("retorna opciones predeterminadas", () => {
      const fakeFolder = {
        id: null,
        name: null,
        getId: function() { return this.id },
        getName: function() { return this.name },
      }
      const fakeProcessingFolder = { ...fakeFolder, id: "processing-folder-id", name: "processing-folder" };
      const fakeLogFolder = { ...fakeFolder, id: "log-folder-id", name: "log-folder" };
      const fakeLogSpreadsheet = { 
        getId: () => "log-file-id",
        getParents: () => ({
          next: () => fakeLogFolder
        })
      };

      const oldGetFileById = DriveApp.getFileById;
      DriveApp.getFileById = (_) => fakeLogSpreadsheet;
      const oldGetFolderById = DriveApp.getFolderById;
      DriveApp.getFolderById = (_) => fakeProcessingFolder;
      const oldGetOrCreate = Utils.getOrCreateSubfolderFrom;
      Utils.getOrCreateSubfolderFrom = (parent, name) => {
        return { ...fakeFolder, id: `${parent.getId()}_${name}`, name };
      };

      const fakeTargetFolder = { ...fakeFolder, id: "target-folder-id", name: "target-folder" };
      const fakeTargetFile = { 
        getId: () => "target-file-id",
        getParents: () => ({
          next: () => fakeTargetFolder
        })
      };

      const result = lib.getFoldersOptionsWithFallback(
        {
          logSpreadsheet: fakeLogSpreadsheet, 
        }, 
        fakeTargetFile
      );

      Utils.assertEquals(result.processingFolder.getId(), "log-folder-id");
      Utils.assertEquals(result.processedFolder.getId(), "log-folder-id_procesados");
      Utils.assertEquals(result.backupFolder.getId(), "log-folder-id_respaldos");
      Utils.assertFalse(result.keepProcessedInSource);

      Utils.getOrCreateSubfolderFrom = oldGetOrCreate; 
      DriveApp.getFolderById = oldGetFolderById; 
      DriveApp.getFileById = oldGetFileById; 
    });
    Utils.it("respeta las opciones suministradas", () => {
      const fakeFolder = {
        id: null,
        name: null,
        getId: function() { return this.id },
        getName: function() { return this.name },
      }
      const fakeProcessingFolder = { ...fakeFolder, id: "processing-folder-id", name: "processing-folder" };
      const fakeLogFolder = { ...fakeFolder, id: "log-folder-id", name: "log-folder" };
      const fakeLogSpreadsheet = { 
        getId: () => "log-file-id",
        getParents: () => ({
          next: () => fakeLogFolder
        })
      };
      const oldGetFileById = DriveApp.getFileById;
      DriveApp.getFileById = (_) => fakeLogSpreadsheet;
      const oldGetFolderById = DriveApp.getFolderById;
      DriveApp.getFolderById = (_) => fakeProcessingFolder;

      const oldGetOrCreate = Utils.getOrCreateSubfolderFrom;
      Utils.getOrCreateSubfolderFrom = (parent, name) => {
        return { ...fakeFolder, id: `${parent.getId()}_${name}`, name };
      };

      const fakeTargetFolder = { ...fakeFolder, id: "target-folder-id", name: "target-folder" };
      const fakeTargetFile = { 
        getId: () => "target-file-id",
        getParents: () => ({
          next: () => fakeTargetFolder
        })
      };

      const result = lib.getFoldersOptionsWithFallback(
        {
          logSpreadsheet: fakeLogSpreadsheet,
          processingFolderID: "processing-folder-id",
          processedFolderName: "A",
          backupFolderName: "B",
          backupInDestination: true,
          keepProcessedInSource: true,
        }, 
        fakeTargetFile
      );
      Utils.assertEquals(result.processingFolder.getId(), "processing-folder-id");
      Utils.assertEquals(result.processedFolder.getId(), "processing-folder-id_A");
      Utils.assertEquals(result.backupFolder.getId(), "target-folder-id_B");
      Utils.assertTrue(result.keepProcessedInSource);

      Utils.getOrCreateSubfolderFrom = oldGetOrCreate; 
      DriveApp.getFolderById = oldGetFolderById; 
      DriveApp.getFileById = oldGetFileById; 
    });
    Utils.it("permisos insuficientes en la carpeta destino -> se crea el respaldo en la de procesamiento", () => {
      const fakeFolder = {
        id: null,
        name: null,
        getId: function() { return this.id },
        getName: function() { return this.name },
      }
      const fakeProcessingFolder = { ...fakeFolder, id: "processing-folder-id", name: "processing-folder" };
      const fakeLogFolder = { ...fakeFolder, id: "log-folder-id", name: "log-folder" };
      const fakeLogSpreadsheet = { 
        getId: () => "log-file-id",
        getParents: () => ({
          next: () => fakeLogFolder
        })
      };
      const oldGetFileById = DriveApp.getFileById;
      DriveApp.getFileById = (_) => fakeLogSpreadsheet;
      const oldGetFolderById = DriveApp.getFolderById;
      DriveApp.getFolderById = (_) => fakeProcessingFolder;

      const oldGetOrCreate = Utils.getOrCreateSubfolderFrom;
      Utils.getOrCreateSubfolderFrom = (parent, name) => {
        return { ...fakeFolder, id: `${parent.getId()}_${name}`, name };
      };

      const fakeTargetFolder = { ...fakeFolder, id: "target-folder-id", name: "target-folder" };
      const fakeTargetFile = { 
        getId: () => "target-file-id",
        getParents: () => ({
          next: () => null
        })
      };

      const result = lib.getFoldersOptionsWithFallback(
        {
          logSpreadsheet: fakeLogSpreadsheet,
          processingFolderID: "processing-folder-id",
          processedFolderName: "A",
          backupFolderName: "B",
          backupInDestination: true,
          keepProcessedInSource: true,
        }, 
        fakeTargetFile
      );
      Utils.assertEquals(result.processingFolder.getId(), "processing-folder-id");
      Utils.assertEquals(result.processedFolder.getId(), "processing-folder-id_A");
      Utils.assertEquals(result.backupFolder.getId(), "processing-folder-id_B");
      Utils.assertTrue(result.keepProcessedInSource);

      Utils.getOrCreateSubfolderFrom = oldGetOrCreate; 
      DriveApp.getFolderById = oldGetFolderById; 
      DriveApp.getFileById = oldGetFileById; 
    });
  });

  //----------------------------------------------
  //
  // Tests para processDataPipeline
  //
  //----------------------------------------------

  Utils.describe("processDataPipeline", () => {
    
    const baseConfig = {
      columns: [0, 1, 2], 
      searchingFn: () => [],
      ignoreEmptyRows: true,
      updateExistingRows: false
    };

    const sampleData = [
      [1, "Alice", "alice@email.com"],
      [2, "Bob", "bob@email.com"],
      [3, "Charlie", "charlie@email.com"]
    ];

    Utils.it("procesa datos básicos sin actualización", () => {
      const result = lib.processDataPipeline(
        sampleData,
        baseConfig,
        0, 
        "integer" 
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.updatedRows.length, 0);
      Utils.assertEquals(result.newRows[0], [1, "Alice", "alice@email.com"]);
    });

    Utils.it("filtra filas existentes cuando se proporciona existingKeys", () => {
      const existingKeys = new Set([1, 3]); 

      const result = lib.processDataPipeline(
        sampleData,
        baseConfig,
        0,
        "integer",
        existingKeys
      );

      Utils.assertEquals(result.newRows.length, 1);
      Utils.assertEquals(result.updatedRows.length, 0);
      Utils.assertEquals(result.newRows[0], [2, "Bob", "bob@email.com"]);
    });

    Utils.it("actualiza filas existentes cuando updateExistingRows es true", () => {
      const configWithUpdate = {
        ...baseConfig,
        updateExistingRows: true
      };

      const existingKeys = new Set([1, 3]);

      const result = lib.processDataPipeline(
        sampleData,
        configWithUpdate,
        0,
        "integer",
        existingKeys
      );

      Utils.assertEquals(result.newRows.length, 1);
      Utils.assertEquals(result.updatedRows.length, 2);
      Utils.assertEquals(result.newRows[0], [2, "Bob", "bob@email.com"]);
      
      Utils.assertEquals(result.updatedRows[0].key, 1);
      Utils.assertEquals(result.updatedRows[0].row, [1, "Alice", "alice@email.com"]);
    });

    Utils.it("aplica processingFn si está definido", () => {
      const configWithProcessing = {
        ...baseConfig,
        columns: [0, 1, 2, 3],
        processingFn: (data) => data.map(row => [...row, "processed"])
      };

      const result = lib.processDataPipeline(
        sampleData,
        configWithProcessing,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.newRows[0], [1, "Alice", "alice@email.com", "processed"]);
    });

    Utils.it("filtra filas vacías cuando ignoreEmptyRows es true", () => {
      const dataWithEmptyRows = [
        [1, "Alice", "alice@email.com"],
        ["", "", ""], 
        [2, "Bob", "bob@email.com"],
        [null, null, null] 
      ];

      const result = lib.processDataPipeline(
        dataWithEmptyRows,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);
      Utils.assertEquals(result.newRows[0], [1, "Alice", "alice@email.com"]);
      Utils.assertEquals(result.newRows[1], [2, "Bob", "bob@email.com"]);
    });

    Utils.it("mantiene filas vacías cuando ignoreEmptyRows es false", () => {
      const configWithoutIgnore = {
        ...baseConfig,
        ignoreEmptyRows: false
      };

      const dataWithEmptyRows = [
        [1, "Alice", "alice@email.com"],
        ["", "", ""], 
        [2, "Bob", "bob@email.com"]
      ];

      const result = lib.processDataPipeline(
        dataWithEmptyRows,
        configWithoutIgnore,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);
    });

    Utils.it("maneja mapeo complejo de columnas", () => {
      const configComplexColumns = {
        ...baseConfig,
        columns: [2, 0, "fixed_value", 1] 
      };

      const result = lib.processDataPipeline(
        sampleData,
        configComplexColumns,
        1, 
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.newRows[0], ["alice@email.com", 1, "fixed_value", "Alice"]);
    });

    Utils.it("normaliza keys según el tipo especificado", () => {
      const dataWithStringNumbers = [
        ["1", "Alice", "alice@email.com"],
        ["2", "Bob", "bob@email.com"]
      ];

      const result = lib.processDataPipeline(
        dataWithStringNumbers,
        baseConfig,
        0,
        "integer" 
      );

      const existingKeys = new Set([1]); 
      const resultWithFilter = lib.processDataPipeline(
        dataWithStringNumbers,
        baseConfig,
        0,
        "integer",
        existingKeys
      );

      Utils.assertEquals(resultWithFilter.newRows.length, 1);
      Utils.assertEquals(resultWithFilter.newRows[0], ["2", "Bob", "bob@email.com"]);
    });

    Utils.it("maneja empty existingKeys set", () => {
      const result = lib.processDataPipeline(
        sampleData,
        baseConfig,
        0,
        "integer",
        new Set() 
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.updatedRows.length, 0);
    });

    Utils.it("maneja undefined existingKeys", () => {
      const result = lib.processDataPipeline(
        sampleData,
        baseConfig,
        0,
        "integer",
        undefined 
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.updatedRows.length, 0);
    });

    Utils.it("maneja línea vacía", () => {
      const emptyData = [[]];

      const result = lib.processDataPipeline(
        emptyData,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 0);
      Utils.assertEquals(result.updatedRows.length, 0);
    });

    Utils.it("valida parámetros requeridos", () => {
      Utils.assertFunctionParams(
        lib.processDataPipeline,
        [null, baseConfig, 0, "integer"],
        true, 
        "'data' no es válido"
      );

      Utils.assertFunctionParams(
        lib.processDataPipeline,
        [sampleData, null, 0, "integer"],
        true, 
        "'config' no es válido"
      );

      Utils.assertFunctionParams(
        lib.processDataPipeline,
        [sampleData, baseConfig, -1, "integer"],
        true, 
        "'keyColumnIndex' no es válido"
      );

      Utils.assertFunctionParams(
        lib.processDataPipeline,
        [sampleData, baseConfig, 0, "number"],
        true, 
        "'keyColumnType' no es válido"
      );
    });

    Utils.it("maneja diferentes tipos de keys", () => {
      const mixedData = [
        [1, "Number key"],
        ["key1", "String key"],
        [true, "Boolean key"]
      ];

      const resultString = lib.processDataPipeline(
        mixedData,
        baseConfig,
        0,
        "string"
      );

      const resultBoolean = lib.processDataPipeline(
        [mixedData[2]],
        baseConfig,
        0,
        "boolean"
      );

      Utils.assertEquals(resultString.newRows.length, 3);
      Utils.assertEquals(resultBoolean.newRows.length, 1);
    });

    Utils.it("maneja índices de columna fuera de rango", () => {
      const configWithInvalidIndex = {
        ...baseConfig,
        columns: [0, 5, 1] // Índice 5 fuera de rango
      };

      const result = lib.processDataPipeline(
        [[1, "A", "B", "C"]], // Solo 4 columnas (0-3)
        configWithInvalidIndex,
        0,
        "integer"
      );

      // Debería usar el valor literal del índice inválido
      Utils.assertEquals(result.newRows[0], [1, 5, "A"]);
    });

    Utils.it("maneja processingFn que modifica estructura", () => {
      const transformingConfig = {
        ...baseConfig,
        processingFn: (data) => data.flatMap(row => [row, [row[0] * 10, row[1] + "_copy"]])
      };

      const result = lib.processDataPipeline(
        [[1, "Original"]],
        transformingConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);
      Utils.assertEquals(result.newRows[1], [10, "Original_copy", 2]);
    });
  });

  //----------------------------------------------
  //
  // Tests para processGoogleSheetFile
  //
  //----------------------------------------------

  Utils.describe("processGoogleSheetFile", () => {
    
    function createMockFile(mimeType) {
      return {
        getId: () => "mock-file-id",
        getMimeType: () => mimeType,
        getName: () => "test-file"
      };
    }

    function createMockSheet(values) {
      return {
        getDataRange: () => ({
          getValues: () => values
        })
      };
    }

    const baseConfig = {
      columns: [0, 1, 2],
      searchingFn: () => [],
      ignoreEmptyRows: true,
      updateExistingRows: false
    };

    const sampleData = [
      [1, "Alice", "alice@email.com"],
      [2, "Bob", "bob@email.com"],
      [3, "Charlie", "charlie@email.com"]
    ];

    const mockSheet = createMockSheet(sampleData);
    const mockEmptySheet = createMockSheet([[]]);
    
    const mockSpreadsheet = {
      getSheetByName: function(name) {
        return name === "TestSheet" ? mockSheet : null;
      },
      getSheets: function() {
        return [mockSheet];
      }
    };

    const mockSpreadsheetWithNamedSheet = {
      getSheetByName: function(name) {
        return name === "Data" ? mockSheet : null;
      },
      getSheets: function() {
        return [mockSheet];
      }
    };

    const mockSpreadsheetEmpty = {
      getSheetByName: function(name) {
        return mockEmptySheet;
      },
      getSheets: function() {
        return [mockEmptySheet];
      }
    };

    const originalOpenById = SpreadsheetApp.openById;

    Utils.it("procesa archivo Google Sheet válido correctamente", () => {
      SpreadsheetApp.openById = function(id) {
        Utils.assertEquals(id, "mock-file-id");
        return mockSpreadsheet;
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      const result = lib.processGoogleSheetFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.updatedRows.length, 0);
      Utils.assertEquals(result.newRows[0], [1, "Alice", "alice@email.com"]);

      
      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("lanza error para archivo no Google Sheet", () => {
      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      
      Utils.assertFunctionParams(
        lib.processGoogleSheetFile,
        [mockFile, baseConfig, 0, "integer"],
        true,
        "El archivo debe ser un Google Sheet válido"
      );
    });

    Utils.it("lanza error para archivo null/undefined", () => {
      Utils.assertFunctionParams(
        lib.processGoogleSheetFile,
        [null, baseConfig, 0, "integer"],
        true,
        "El archivo debe ser un Google Sheet válido"
      );

      Utils.assertFunctionParams(
        lib.processGoogleSheetFile,
        [undefined, baseConfig, 0, "integer"],
        true,
        "El archivo debe ser un Google Sheet válido"
      );
    });

    Utils.it("lanza error cuando config, keyColumnIndex o keyColumnType no son válidos", () => {
      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);

      Utils.assertFunctionParams(
        lib.processGoogleSheetFile,
        [mockFile, null, 0, "integer"],
        true,
        "'config' no es válido"
      );

      Utils.assertFunctionParams(
        lib.processGoogleSheetFile,
        [mockFile, baseConfig, -1, "integer"],
        true,
        "'keyColumnIndex' no es válido"
      );

      Utils.assertFunctionParams(
        lib.processGoogleSheetFile,
        [mockFile, baseConfig, 0, "number"],
        true,
        "'keyColumnType' no es válido"
      );
    });

    Utils.it("usa hoja específica cuando se especifica sourceSheetName", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheetWithNamedSheet;
      };

      const configWithSheetName = {
        ...baseConfig,
        sourceSheetName: "Data"
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      const result = lib.processGoogleSheetFile(
        mockFile,
        configWithSheetName,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.newRows[0], [1, "Alice", "alice@email.com"]);

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("lanza error cuando no encuentra la hoja especificada", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };

      const configWithInvalidSheet = {
        ...baseConfig,
        sourceSheetName: "NonExistentSheet"
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      Utils.assertFunctionParams(
        lib.processGoogleSheetFile,
        [mockFile, configWithInvalidSheet, 0, "integer"],
        true,
        "No se encontró la hoja"
      );

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("maneja hoja vacía correctamente", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheetEmpty;
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      const result = lib.processGoogleSheetFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 0);
      Utils.assertEquals(result.updatedRows.length, 0);

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("pasa existingKeys correctamente a processDataPipeline", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };

      const existingKeys = new Set([1, 3]);
      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);

      const result = lib.processGoogleSheetFile(
        mockFile,
        baseConfig,
        0,
        "integer",
        existingKeys
      );

      Utils.assertEquals(result.newRows.length, 1);
      Utils.assertEquals(result.newRows[0], [2, "Bob", "bob@email.com"]);

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("maneja modo updateExistingRows correctamente", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };

      const configWithUpdate = {
        ...baseConfig,
        updateExistingRows: true
      };

      const existingKeys = new Set([1]);
      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);

      const result = lib.processGoogleSheetFile(
        mockFile,
        configWithUpdate,
        0,
        "integer",
        existingKeys
      );

      Utils.assertEquals(result.newRows.length, 2);
      Utils.assertEquals(result.updatedRows.length, 1);
      Utils.assertEquals(result.updatedRows[0].key, 1);

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("maneja errores de apertura del spreadsheet", () => {
      SpreadsheetApp.openById = function(id) {
        throw new Error("Cannot open spreadsheet");
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      Utils.assertFunctionParams(
        lib.processGoogleSheetFile,
        [mockFile, baseConfig, 0, "integer"],
        true,
        "Error procesando el archivo Google Sheet"
      );

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("usa primera hoja cuando sourceSheetName no está definido", () => {
      let firstSheetAccessed = false;
      const mockSpreadsheetFirstSheet = {
        getSheetByName: function(name) {
          throw new Error("getSheetByName no debería ser llamado");
        },
        getSheets: function() {
          firstSheetAccessed = true;
          return [mockSheet];
        }
      };

      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheetFirstSheet;
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      const result = lib.processGoogleSheetFile(
        mockFile,
        baseConfig, 
        0,
        "integer"
      );

      Utils.assertTrue(firstSheetAccessed);
      Utils.assertEquals(result.newRows.length, 3);

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("aplica processingFn del config correctamente", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };

      const configWithProcessing = {
        ...baseConfig,
        processingFn: function(data) {
          return data.filter(row => row[0] !== 2); 
        }
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      const result = lib.processGoogleSheetFile(
        mockFile,
        configWithProcessing,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);
      Utils.assertEquals(result.newRows[0], [1, "Alice", "alice@email.com"]);
      Utils.assertEquals(result.newRows[1], [3, "Charlie", "charlie@email.com"]);

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });
  });

  Utils.describe("processGoogleSheetFile - Edge Cases", () => {
    
    function createMockFile(mimeType) {
      return {
        getId: () => "mock-file-id",
        getMimeType: () => mimeType,
        getName: () => "test-file"
      };
    }

    function createMockSheet(values) {
      return {
        getDataRange: () => ({
          getValues: () => values
        })
      };
    }

    const mockSheetWithEmptyRows = createMockSheet([
      [1, "Alice", "alice@email.com"],
      ["", "", ""], 
      [2, "Bob", "bob@email.com"],
      [null, null, null] 
    ]);

    const mockSpreadsheet = {
      getSheetByName: function(name) {
        return mockSheetWithEmptyRows;
      },
      getSheets: function() {
        return [mockSheetWithEmptyRows];
      }
    };

    const baseConfig = {
      columns: [0, 1, 2],
      searchingFn: () => [],
      ignoreEmptyRows: true,
      updateExistingRows: false
    };

    Utils.it("filtra filas vacías cuando ignoreEmptyRows es true", () => {
      const originalOpenById = SpreadsheetApp.openById;
      
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      const result = lib.processGoogleSheetFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });

    Utils.it("mantiene filas vacías cuando ignoreEmptyRows es false", () => {
      const originalOpenById = SpreadsheetApp.openById;
      
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };

      const configWithoutIgnore = {
        ...baseConfig,
        ignoreEmptyRows: false
      };

      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      const result = lib.processGoogleSheetFile(
        mockFile,
        configWithoutIgnore,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 4);

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
    });
  });

  //----------------------------------------------
  //
  // Tests para processXlsxFile
  //
  //----------------------------------------------

  Utils.describe("processXlsxFile", () => {
    
    function createMockFile(mimeType, id) {
      return {
        getId: () => id || "mock-xlsx-file-id",
        getMimeType: () => mimeType,
        getName: () => "test-file.xlsx",
        setTrashed: () => console.info("Eliminado el mock file."),
      };
    }

    function createMockSheet(values) {
      return {
        getDataRange: () => ({
          getValues: () => values
        })
      };
    }

    const baseConfig = {
      columns: [0, 1, 2],
      searchingFn: () => [],
      ignoreEmptyRows: true,
      updateExistingRows: false
    };

    const sampleData = [
      [1, "Alice", "alice@email.com"],
      [2, "Bob", "bob@email.com"],
      [3, "Charlie", "charlie@email.com"]
    ];

    const mockSheet = createMockSheet(sampleData);
    const mockEmptySheet = createMockSheet([[]]);
    
    const mockSpreadsheet = {
      getSheetByName: function(name) {
        return name === "TestSheet" ? mockSheet : null;
      },
      getSheets: function() {
        return [mockSheet];
      }
    };

    const mockConvertedFile = createMockFile(MimeType.GOOGLE_SHEETS, "converted-file-id");

    const originalOpenById = SpreadsheetApp.openById;
    const originalConvertFile = Utils.convertFileToGoogleSheet;
    const originalFlush = SpreadsheetApp.flush;

    Utils.it("procesa archivo XLSX válido correctamente", () => {
      SpreadsheetApp.openById = function(id) {
        Utils.assertEquals(id, "converted-file-id");
        return mockSpreadsheet;
      };
      
      Utils.convertFileToGoogleSheet = function(file) {
        Utils.assertEquals(file.getMimeType(), MimeType.MICROSOFT_EXCEL);
        return mockConvertedFile;
      };
      
      SpreadsheetApp.flush = function() {}; 

      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      const result = lib.processXlsxFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.updatedRows.length, 0);
      Utils.assertEquals(result.newRows[0], [1, "Alice", "alice@email.com"]);

      
      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
      if (originalFlush) SpreadsheetApp.flush = originalFlush;
    });

    Utils.it("lanza error para archivo no XLSX", () => {
      const mockFile = createMockFile(MimeType.GOOGLE_SHEETS);
      
      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, baseConfig, 0, "integer"],
        true,
        "El archivo debe ser un XLSX válido"
      );
    });

    Utils.it("lanza error para parámetros inválidos", () => {
      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      
      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, null, 0, "integer"],
        true,
        "El parámetro 'config' no es válido"
      );
      
      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, baseConfig, -1, "integer"],
        true,
        "El parámetro 'keyColumnIndex' no es válido"
      );
      
      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, baseConfig, 0, "invalid-type"],
        true,
        "El parámetro 'keyColumnType' no es válido"
      );
    });

    Utils.it("lanza error cuando config, keyColumnIndex o keyColumnType no son válidos", () => {
      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);

      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, null, 0, "integer"],
        true,
        "'config' no es válido"
      );

      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, baseConfig, -1, "integer"],
        true,
        "'keyColumnIndex' no es válido"
      );

      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, baseConfig, 0, "number"],
        true,
        "'keyColumnType' no es válido"
      );
    });

    Utils.it("usa hoja específica cuando se especifica sourceSheetName", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };
      
      Utils.convertFileToGoogleSheet = function(file) {
        return mockConvertedFile;
      };
      
      SpreadsheetApp.flush = function() {};

      const configWithSheetName = {
        ...baseConfig,
        sourceSheetName: "TestSheet"
      };

      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      const result = lib.processXlsxFile(
        mockFile,
        configWithSheetName,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.newRows[0], [1, "Alice", "alice@email.com"]);

      
      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
      if (originalFlush) SpreadsheetApp.flush = originalFlush;
    });

    Utils.it("lanza error cuando no encuentra la hoja especificada", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };
      
      Utils.convertFileToGoogleSheet = function(file) {
        return mockConvertedFile;
      };
      
      SpreadsheetApp.flush = function() {};

      const configWithInvalidSheet = {
        ...baseConfig,
        sourceSheetName: "NonExistentSheet"
      };

      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, configWithInvalidSheet, 0, "integer"],
        true,
        "No se encontró la hoja"
      );

      
      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
      if (originalFlush) SpreadsheetApp.flush = originalFlush;
    });

    Utils.it("maneja errores de conversión de archivo", () => {
      Utils.convertFileToGoogleSheet = function(file) {
        throw new Error("Error de conversión");
      };

      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, baseConfig, 0, "integer"],
        true,
        "Error de conversión"
      );

      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
    });

    Utils.it("maneja errores de apertura del spreadsheet convertido", () => {
      Utils.convertFileToGoogleSheet = function(file) {
        return mockConvertedFile;
      };
      
      SpreadsheetApp.openById = function(id) {
        throw new Error("Cannot open converted spreadsheet");
      };

      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      Utils.assertFunctionParams(
        lib.processXlsxFile,
        [mockFile, baseConfig, 0, "integer"],
        true,
        "Error procesando el archivo XLSX"
      );

      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
    });

    Utils.it("limpia archivos temporales en bloque finally", () => {
      let flushed = false;
      let trashed = false;
      
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };
      
      Utils.convertFileToGoogleSheet = function(file) {
        return {
          getId: () => "temp-file-id",
          getMimeType: () => MimeType.GOOGLE_SHEETS,
          setTrashed: function() { trashed = true; }
        };
      };
      
      SpreadsheetApp.flush = function() { flushed = true; };

      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      const result = lib.processXlsxFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertTrue(flushed, "Debería llamar a SpreadsheetApp.flush()");
      Utils.assertTrue(trashed, "Debería llamar a setTrashed(true)");

      
      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
      if (originalFlush) SpreadsheetApp.flush = originalFlush;
    });

    Utils.it("maneja errores durante la limpieza de archivos temporales", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };
      
      Utils.convertFileToGoogleSheet = function(file) {
        return {
          getId: () => "temp-file-id",
          getMimeType: () => MimeType.GOOGLE_SHEETS,
          setTrashed: function() { throw new Error("Error al eliminar"); }
        };
      };
      
      SpreadsheetApp.flush = function() { throw new Error("Error al flush"); };

      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      
      const result = lib.processXlsxFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);

      
      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
      if (originalFlush) SpreadsheetApp.flush = originalFlush;
    });

    Utils.it("pasa existingKeys correctamente a processDataPipeline", () => {
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };
      
      Utils.convertFileToGoogleSheet = function(file) {
        return mockConvertedFile;
      };
      
      SpreadsheetApp.flush = function() {};

      const existingKeys = new Set([1, 3]);
      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);

      const result = lib.processXlsxFile(
        mockFile,
        baseConfig,
        0,
        "integer",
        existingKeys
      );

      Utils.assertEquals(result.newRows.length, 1);
      Utils.assertEquals(result.newRows[0], [2, "Bob", "bob@email.com"]);

      
      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
      if (originalFlush) SpreadsheetApp.flush = originalFlush;
    });
  });

  Utils.describe("processXlsxFile - Edge Cases", () => {
    
    function createMockFile(mimeType) {
      return {
        getId: () => "mock-xlsx-file-id",
        getMimeType: () => mimeType,
        getName: () => "test-file.xlsx"
      };
    }

    function createMockSheet(values) {
      return {
        getDataRange: () => ({
          getValues: () => values
        })
      };
    }

    const mockSheetWithEmptyRows = createMockSheet([
      [1, "Alice", "alice@email.com"],
      ["", "", ""], 
      [2, "Bob", "bob@email.com"],
      [null, null, null] 
    ]);

    const mockSpreadsheet = {
      getSheetByName: function(name) {
        return mockSheetWithEmptyRows;
      },
      getSheets: function() {
        return [mockSheetWithEmptyRows];
      }
    };

    const baseConfig = {
      columns: [0, 1, 2],
      searchingFn: () => [],
      ignoreEmptyRows: true,
      updateExistingRows: false
    };

    Utils.it("filtra filas vacías cuando ignoreEmptyRows es true", () => {
      const originalOpenById = SpreadsheetApp.openById;
      const originalConvertFile = Utils.convertFileToGoogleSheet;
      const originalFlush = SpreadsheetApp.flush;
      
      SpreadsheetApp.openById = function(id) {
        return mockSpreadsheet;
      };
      
      Utils.convertFileToGoogleSheet = function(file) {
        return {
          getId: () => "temp-file-id",
          getMimeType: () => MimeType.GOOGLE_SHEETS,
          setTrashed: function() {}
        };
      };
      
      SpreadsheetApp.flush = function() {};

      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL);
      const result = lib.processXlsxFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);

      
      if (originalOpenById) SpreadsheetApp.openById = originalOpenById;
      if (originalConvertFile) Utils.convertFileToGoogleSheet = originalConvertFile;
      if (originalFlush) SpreadsheetApp.flush = originalFlush;
    });
  });

  //----------------------------------------------
  //
  // Tests para processCsvFile
  //
  //----------------------------------------------

  Utils.describe("processCsvFile", () => {
    
    function createMockFile(mimeType, contents) {
      return {
        getId: () => "mock-csv-file-id",
        getMimeType: () => mimeType,
        getName: () => "test-file.csv",
        getBlob: () => ({
          getDataAsString: (encoding) => {
            Utils.assertEquals(encoding, 'UTF-8');
            return contents;
          }
        })
      };
    }

    const baseConfig = {
      columns: [0, 1, 2],
      searchingFn: () => [],
      ignoreEmptyRows: true,
      updateExistingRows: false
    };

    Utils.it("procesa archivo CSV válido correctamente", () => {
      const csvContents = `1,Alice,alice@email.com\n2,Bob,bob@email.com\n3,Charlie,charlie@email.com`;
      const mockFile = createMockFile(MimeType.CSV, csvContents);
      
      const result = lib.processCsvFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 3);
      Utils.assertEquals(result.updatedRows.length, 0);
      Utils.assertEquals(result.newRows[0], ["1", "Alice", "alice@email.com"]);
    });

    Utils.it("lanza error para archivo no CSV", () => {
      const mockFile = createMockFile(MimeType.MICROSOFT_EXCEL, "content");
      
      Utils.assertFunctionParams(
        lib.processCsvFile,
        [mockFile, baseConfig, 0, "integer"],
        true,
        "El archivo debe ser un CSV válido"
      );
    });

    Utils.it("lanza error para parámetros inválidos", () => {
      const mockFile = createMockFile(MimeType.CSV, "content");
      
      Utils.assertFunctionParams(
        lib.processCsvFile,
        [mockFile, null, 0, "integer"],
        true,
        "El parámetro 'config' no es válido"
      );
      
      Utils.assertFunctionParams(
        lib.processCsvFile,
        [mockFile, baseConfig, -1, "integer"],
        true,
        "El parámetro 'keyColumnIndex' no es válido"
      );
      
      Utils.assertFunctionParams(
        lib.processCsvFile,
        [mockFile, baseConfig, 0, "invalid-type"],
        true,
        "El parámetro 'keyColumnType' no es válido"
      );
    });

    Utils.it("maneja archivo CSV vacío", () => {
      const mockFile = createMockFile(MimeType.CSV, "");
      
      const result = lib.processCsvFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 0);
      Utils.assertEquals(result.updatedRows.length, 0);
    });

    Utils.it("maneja archivo CSV con solo espacios en blanco", () => {
      const mockFile = createMockFile(MimeType.CSV, "   \n  \t\n  ");
      
      const result = lib.processCsvFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 0);
      Utils.assertEquals(result.updatedRows.length, 0);
    });

    Utils.it("detecta correctamente diferentes separadores", () => {
      const csvContents = "1;Alice;alice@email.com\n2;Bob;bob@email.com";
      const mockFile = createMockFile(MimeType.CSV, csvContents);
      
      const result = lib.processCsvFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);
      Utils.assertEquals(result.newRows[0], ["1", "Alice", "alice@email.com"]);
    });

    Utils.it("maneja líneas vacías en el CSV", () => {
      const csvContents = "1,Alice,alice@email.com\n\n2,Bob,bob@email.com\n\t\n";
      const mockFile = createMockFile(MimeType.CSV, csvContents);
      
      const result = lib.processCsvFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);
      Utils.assertEquals(result.newRows[0], ["1", "Alice", "alice@email.com"]);
      Utils.assertEquals(result.newRows[1], ["2", "Bob", "bob@email.com"]);
    });

    Utils.it("pasa existingKeys correctamente a processDataPipeline", () => {
      const csvContents = "1,Alice,alice@email.com\n2,Bob,bob@email.com\n3,Charlie,charlie@email.com";
      const mockFile = createMockFile(MimeType.CSV, csvContents);
      
      const existingKeys = new Set([1, 3]);
      const result = lib.processCsvFile(
        mockFile,
        baseConfig,
        0,
        "integer",
        existingKeys
      );

      Utils.assertEquals(result.newRows.length, 1);
      Utils.assertEquals(result.newRows[0], ["2", "Bob", "bob@email.com"]);
    });

    Utils.it("maneja CSV con comillas y caracteres especiales", () => {
      const csvContents = '1,"Alice, Smith","alice@email.com"\n2,"Bob""Brown","bob@email.com"';
      const mockFile = createMockFile(MimeType.CSV, csvContents);
      
      const result = lib.processCsvFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);
      Utils.assertEquals(result.newRows[0], ["1", "Alice, Smith", "alice@email.com"]);
      Utils.assertEquals(result.newRows[1], ["2", 'Bob"Brown', "bob@email.com"]);
    });

    Utils.it("maneja diferentes tipos de salto de línea", () => {
      const csvContents = "1,Alice,alice@email.com\r\n2,Bob,bob@email.com\r3,Charlie,charlie@email.com\n4,Diana,diana@email.com";
      const mockFile = createMockFile(MimeType.CSV, csvContents);
      
      const result = lib.processCsvFile(
        mockFile,
        baseConfig,
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 4);
      Utils.assertEquals(result.newRows[0], ["1", "Alice", "alice@email.com"]);
      Utils.assertEquals(result.newRows[3], ["4", "Diana", "diana@email.com"]);
    });
  });

  Utils.describe("processCsvFile - Edge Cases", () => {
    
    function createMockFile(mimeType, contents) {
      return {
        getId: () => "mock-csv-file-id",
        getMimeType: () => mimeType,
        getName: () => "test-file.csv",
        getBlob: () => ({
          getDataAsString: (encoding) => contents
        })
      };
    }

    Utils.it("maneja encoding UTF-8 correctamente", () => {
      const csvContents = "1,José,jose@email.com\n2,María,maria@email.com";
      const mockFile = createMockFile(MimeType.CSV, csvContents);
      
      const result = lib.processCsvFile(
        mockFile,
        {
          columns: [0, 1, 2],
          searchingFn: () => [],
          ignoreEmptyRows: true,
          updateExistingRows: false
        },
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 2);
      Utils.assertEquals(result.newRows[0], ["1", "José", "jose@email.com"]);
    });

    Utils.it("maneja archivo CSV con solo headers", () => {
      const csvContents = "id,name,email";
      const mockFile = createMockFile(MimeType.CSV, csvContents);

      const result = lib.processCsvFile(
        mockFile,
        {
          columns: [0, 1, 2],
          searchingFn: () => [],
          ignoreEmptyRows: true,
          updateExistingRows: false
        },
        0,
        "integer"
      );

      Utils.assertEquals(result.newRows.length, 1);
      Utils.assertEquals(result.newRows[0], ["id", "name", "email"]);
    });
  });

  //----------------------------------------------
  //
  // Tests para mergeSheetsDataTo
  //
  //----------------------------------------------

  Utils.describe("mergeSheetsDataTo", () => {
    
    function createMockContext() {
      return {
        shouldStop: false,
        remainingTime: () => 300000
      };
    }

    function createMockFile(id, name, mimeType, parent=null) {
      return {
        getId: () => id,
        getName: () => name,
        getMimeType: () => mimeType,
        getBlob: () => ({ getDataAsString: () => "1,test,test@email.com" }),
        getParents: () => ({ next: () => parent }),
      };
    }

    function createMockFolder(id, name, parent) {
      return {
        getId: () => id,
        getName: () => name,
        getParents: () => ({ next: () => parent }),
      };
    }

    function createMockSpreadsheet(id, name) {
      const fakeSheet1 = { getName: () => "Hoja1" };
      const fakeSheet2 = { getName: () => "Hoja2" };
      const fakeSheets = [fakeSheet1, fakeSheet2];
      return {
        getId: () => id,
        getName: () => name,
        getSheets: () => fakeSheets, 
        getSheetByName: (name) => fakeSheets.find(sheet => sheet.getName() === name),
      };
    }

    const baseConfig = {
      columns: [0, 1, 2],
      searchingFn: function(mimeType) {
        if (mimeType === MimeType.CSV) {
          return [createMockFile("csv1", "test.csv", MimeType.CSV)];
        }
        return [];
      },
      ignoreEmptyRows: true,
      updateExistingRows: false,
      allowFileReprocessing: true 
    };

    const configs = {
      [MimeType.CSV]: baseConfig
    };

    const options = {
      logColumnTitles: ["Fecha", "Archivo", "Resumen", "Estado"],
      successMessage: "Éxito",
      failureMessage: "Error",
      keepProcessedInSource: true 
    };

    const oldOpenById = SpreadsheetApp.openById;
    const oldGetFileById = DriveApp.getFileById;
    const oldGetFolderById = DriveApp.getFolderById;
    const oldGetOrCreate = Utils.getOrCreateSubfolderFrom;
    const originalBackupFileTo = Utils.backupFileTo;
    const originalGetUniqueValues = Utils.getUniqueValuesFromColumn;
    const originalUpdateRows = Utils.updateRowsInSheet;
    const originalAppendRows = Utils.appendRowsToSheet;
    const originalMoveFile = Utils.moveFileToFolder;
    const originalSortSheet = Utils.sortSheet;
    const originalGetNormalizer = Utils.getNormalizer;
    const originalExecutionController = EC.ExecutionController;

    const mockContext = createMockContext();
    const mockProcessingFolder = createMockFolder("processing-folder-id", "processing-folder", null);
    const mockLogSpreadsheet = createMockSpreadsheet("log-ss-id", "log-ss");
    const mockLogFolder = createMockFolder("log-folder-id", "log-folder", null);
    const mockLogFile = createMockFile("log-ss-id", "log-ss", MimeType.GOOGLE_SHEETS, mockLogFolder);
    const mockTargetSpreadsheet = createMockSpreadsheet("target-ss-id", "target-ss");
    const mockTargetFolder = createMockFolder("target-folder-id", "target-folder", null);
    const mockTargetFile = createMockFile("target-ss-id", "target-ss", MimeType.GOOGLE_SHEETS, mockTargetFolder);

    SpreadsheetApp.openById = (id) => {
      if (id === 'target-ss-id') return mockTargetSpreadsheet;
      if (id === 'log-ss-id') return mockLogSpreadsheet;
      console.warn(`SpreadsheetId inválido: ${id}`)
      throw new Error('SpreadsheetId inválido');
    }
    DriveApp.getFileById = (id) => {
      if (id === 'target-ss-id') return mockTargetFile;
      if (id === 'log-ss-id') return mockLogFile;
      console.warn(`FileId inválido: ${id}`)
      throw new Error("FileId inválido")
    }
    DriveApp.getFolderById = (id) => {
      if (id === 'processing-folder-id') return mockProcessingFolder;
      console.warn(`FolderId inválido: ${id}`)
      throw new Error("FolderId inválido")
    }
    Utils.getOrCreateSubfolderFrom = (parent, name) => {
      return createMockFolder(`${name}_id`, name, parent);
    };
    
    Utils.it("valida parámetros requeridos correctamente", () => {
      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo, 
        [
          null, 
          { 
            fileID: "target-ss-id", 
            fileSheetName: "Hoja2", 
            columnTitles: ["col1"], 
            keyColumnIndex: 0, 
            keyColumnType: "string",
          },
          configs, 
          options,
        ],
        true,
        "'ctx' no es válido"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo, 
        [
          mockContext, 
          null,
          configs, 
          options,
        ],
        true,
        "'targetConfig' no es válido"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo,
        [
          mockContext, 
          { 
            fileID: null, 
            fileSheetName: "Hoja2", 
            columnTitles: ["col1"], 
            keyColumnIndex: 0, 
            keyColumnType: "string",
          },
          configs, 
          options,
        ],
        true,
        "identificador del archivo de destino no ha sido proporcionado"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo,
        [
          mockContext, 
          { 
            fileID: "ss2", 
            fileSheetName: "Hoja2", 
            columnTitles: ["col1"], 
            keyColumnIndex: 0, 
            keyColumnType: "string",
          },
          configs, 
          options,
        ],
        true,
        "identificador del archivo de destino no es válido"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo,
        [
          mockContext, 
          { 
            fileID: "target-ss-id", 
            fileSheetName: "invalid-name", 
            columnTitles: ["col1"], 
            keyColumnIndex: 0, 
            keyColumnType: "string",
          },
          configs, 
          options,
        ],
        true,
        "Hoja de destino no encontrada"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo,
        [
          mockContext, 
          { 
            fileID: "target-ss-id", 
            fileSheetName: "Hoja2", 
            columnTitles: "invalid", 
            keyColumnIndex: 0, 
            keyColumnType: "string",
          },
          configs, 
          options,
        ],
        true,
        "'targetConfig.columnTitles' no es válido"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo,
        [
          mockContext, 
          { 
            fileID: "target-ss-id", 
            fileSheetName: "Hoja2", 
            columnTitles: ["col1"], 
            keyColumnIndex: "a", 
            keyColumnType: "string",
          },
          configs, 
          options,
        ],
        true,
        "'targetConfig.keyColumnIndex' no es válido"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo,
        [
          mockContext, 
          { 
            fileID: "target-ss-id", 
            fileSheetName: "Hoja2", 
            columnTitles: ["col1"], 
            keyColumnIndex: 0, 
            keyColumnType: "invalid-type",
          },
          configs, 
          options,
        ],
        true,
        "'targetConfig.keyColumnType' no es válido"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo,
        [
          mockContext, 
          { 
            fileID: "target-ss-id", 
            fileSheetName: "Hoja2", 
            columnTitles: ["col1"], 
            keyColumnIndex: 0, 
            keyColumnType: "string",
          },
          configs, 
          { ...options, logSpreadsheet: null },
        ],
        true,
        "log no ha sido proporcionado"
      );

      Utils.assertFunctionParams(
        lib.mergeSheetsDataTo,
        [
          mockContext, 
          { 
            fileID: "target-ss-id", 
            fileSheetName: "Hoja2", 
            columnTitles: ["col1"], 
            keyColumnIndex: 0, 
            keyColumnType: "string",
          },
          configs, 
          { ...options, logSpreadsheet: mockLogSpreadsheet, logSheetName: "invalid-name" },
        ],
        true,
        "Hoja no encontrada en log"
      );
    });

    Utils.it("maneja configs vacías o inválidas", () => {
      Utils.backupFileTo = function() {};
      Utils.getUniqueValuesFromColumn = function() { return new Set(); };
      
      const resultEmptyConfigs = lib.mergeSheetsDataTo(
        mockContext, 
        { 
          fileID: "target-ss-id", 
          fileSheetName: "Hoja2", 
          columnTitles: ["col1"], 
          keyColumnIndex: 0, 
          keyColumnType: "string",
        },
        {}, 
        { ...options, logSpreadsheet: mockLogSpreadsheet, logSheetName: "Hoja1" }
      );
      
      Utils.assertEquals(resultEmptyConfigs.processed, 0);
      Utils.assertEquals(resultEmptyConfigs.added, 0);
      Utils.assertEquals(resultEmptyConfigs.updated, 0);

      const resultNullConfigs = lib.mergeSheetsDataTo(
        mockContext,
        { 
          fileID: "target-ss-id", 
          fileSheetName: "Hoja2", 
          columnTitles: ["col1"], 
          keyColumnIndex: 0, 
          keyColumnType: "string",
        },
        null, 
        { ...options, logSpreadsheet: mockLogSpreadsheet, logSheetName: "Hoja1" }
      );
      
      Utils.assertEquals(resultNullConfigs.processed, 0);
      Utils.assertEquals(resultNullConfigs.added, 0);
      Utils.assertEquals(resultNullConfigs.updated, 0);

      
      if (originalBackupFileTo) Utils.backupFileTo = originalBackupFileTo;
      if (originalGetUniqueValues) Utils.getUniqueValuesFromColumn = originalGetUniqueValues;
    });

    Utils.it("ejecuta el loop de procesamiento correctamente", () => {
      let runLoopCalled = false;
      EC.ExecutionController = {
        runLoop: function(ctx, files, processor) {
          runLoopCalled = true;
          Utils.assertEquals(files.length, 1);
          Utils.assertEquals(files[0].getMimeType(), MimeType.CSV);
          
          return {
            stoppedEarly: false,
            invalidCtx: false,
            successes: [{ processed: 1, added: 1, updated: 0, fileName: "test.csv", date: new Date() }],
            fails: []
          };
        }
      };

      Utils.backupFileTo = function() {};
      Utils.getUniqueValuesFromColumn = function() { return new Set(); };
      Utils.updateRowsInSheet = function() { return 0; };
      Utils.appendRowsToSheet = function() {};
      Utils.getNormalizer = function() { return v => v; };

      const result = lib.mergeSheetsDataTo(
        mockContext,
        { 
          fileID: "target-ss-id", 
          fileSheetName: "Hoja2", 
          columnTitles: ["ID", "Name", "Email"],
          keyColumnIndex: 0, 
          keyColumnType: "integer",
        },
        configs,
        { ...options, logSpreadsheet: mockLogSpreadsheet, logSheetName: "Hoja1" }
      );

      Utils.assertTrue(runLoopCalled, "runLoop debería haberse llamado");
      Utils.assertEquals(result.processed, 1);
      Utils.assertEquals(result.added, 1);
      Utils.assertEquals(result.updated, 0);

      
      if (originalExecutionController) EC.ExecutionController = originalExecutionController;
      if (originalBackupFileTo) Utils.backupFileTo = originalBackupFileTo;
      if (originalGetUniqueValues) Utils.getUniqueValuesFromColumn = originalGetUniqueValues;
      if (originalUpdateRows) Utils.updateRowsInSheet = originalUpdateRows;
      if (originalAppendRows) Utils.appendRowsToSheet = originalAppendRows;
      if (originalGetNormalizer) Utils.getNormalizer = originalGetNormalizer;
    });

    Utils.it("maneja errores en el procesamiento de archivos", () => {
      EC.ExecutionController = {
        runLoop: function(ctx, files, processor) {
          return {
            stoppedEarly: false,
            invalidCtx: false,
            successes: [],
            fails: [{ 
              item: files[0], 
              message: "Error de procesamiento",
              date: new Date()
            }]
          };
        }
      };

      let logEntries = [];
      Utils.appendRowsToSheet = function(sheet, rows, headers) {
        logEntries = logEntries.concat(rows);
      };

      Utils.backupFileTo = function() {};
      Utils.getUniqueValuesFromColumn = function() { return new Set(); };

      const result = lib.mergeSheetsDataTo(
        mockContext,
        { 
          fileID: "target-ss-id", 
          fileSheetName: "Hoja1", 
          columnTitles: ["ID", "Name", "Email"],
          keyColumnIndex: 0, 
          keyColumnType: "integer",
        },
        configs,
        { ...options, logSpreadsheet: mockLogSpreadsheet, logSheetName: "Hoja1" }
      );

      Utils.assertEquals(result.processed, 0);
      Utils.assertEquals(result.added, 0);
      Utils.assertEquals(result.updated, 0);
      Utils.assertEquals(logEntries.length, 1);
      Utils.assertEquals(logEntries[0][3], "Error de procesamiento");

      
      if (originalExecutionController) EC.ExecutionController = originalExecutionController;
      if (originalAppendRows) Utils.appendRowsToSheet = originalAppendRows;
      if (originalBackupFileTo) Utils.backupFileTo = originalBackupFileTo;
      if (originalGetUniqueValues) Utils.getUniqueValuesFromColumn = originalGetUniqueValues;
    });

    Utils.it("maneja múltiples tipos MIME en configs", () => {
      const multiConfigs = {
        [MimeType.CSV]: {
          ...baseConfig,
          searchingFn: function(mimeType) {
            if (mimeType === MimeType.CSV) return [createMockFile("csv1", "test1.csv", MimeType.CSV)];
            return [];
          }
        },
        [MimeType.GOOGLE_SHEETS]: {
          ...baseConfig,
          searchingFn: function(mimeType) {
            if (mimeType === MimeType.GOOGLE_SHEETS) return [createMockFile("gs1", "test2.xlsx", MimeType.GOOGLE_SHEETS)];
            return [];
          }
        }
      };

      let processedTypes = new Set();
      EC.ExecutionController = {
        runLoop: function(ctx, files, processor) {
          processedTypes.add(files[0].getMimeType());
          return {
            stoppedEarly: false,
            invalidCtx: false,
            successes: [{ processed: 1, added: 1, updated: 0, fileName: "test", date: new Date() }],
            fails: []
          };
        }
      };

      Utils.backupFileTo = function() {};
      Utils.getUniqueValuesFromColumn = function() { return new Set(); };
      Utils.appendRowsToSheet = function() {};

      const result = lib.mergeSheetsDataTo(
        mockContext,
        { 
          fileID: "target-ss-id", 
          fileSheetName: "Hoja1", 
          columnTitles: ["ID", "Name", "Email"],
          keyColumnIndex: 0, 
          keyColumnType: "integer",
        },
        multiConfigs,
        { ...options, logSpreadsheet: mockLogSpreadsheet, logSheetName: "Hoja1" }
      );

      Utils.assertEquals(result.processed, 2);
      Utils.assertEquals(processedTypes.size, 2);
      Utils.assertTrue(processedTypes.has(MimeType.CSV));
      Utils.assertTrue(processedTypes.has(MimeType.GOOGLE_SHEETS));

      
      if (originalExecutionController) EC.ExecutionController = originalExecutionController;
      if (originalBackupFileTo) Utils.backupFileTo = originalBackupFileTo;
      if (originalGetUniqueValues) Utils.getUniqueValuesFromColumn = originalGetUniqueValues;
      if (originalAppendRows) Utils.appendRowsToSheet = originalAppendRows;
    });

    Utils.it("maneja execution controller detenido temprano", () => {
      EC.ExecutionController = {
        runLoop: function(ctx, files, processor) {
          return {
            stoppedEarly: true,
            invalidCtx: false,
            successes: [{ processed: 1, added: 1, updated: 0, fileName: "test.csv", date: new Date() }],
            fails: []
          };
        }
      };

      Utils.backupFileTo = function() {};
      Utils.getUniqueValuesFromColumn = function() { return new Set(); };
      Utils.appendRowsToSheet = function() {};

      const result = lib.mergeSheetsDataTo(
        mockContext,
        { 
          fileID: "target-ss-id", 
          fileSheetName: "Hoja1", 
          columnTitles: ["ID", "Name", "Email"],
          keyColumnIndex: 0, 
          keyColumnType: "integer",
        },
        configs,
        { ...options, logSpreadsheet: mockLogSpreadsheet, logSheetName: "Hoja1" }
      );

      Utils.assertEquals(result.processed, 1);

      
      if (originalExecutionController) EC.ExecutionController = originalExecutionController;
      if (originalBackupFileTo) Utils.backupFileTo = originalBackupFileTo;
      if (originalGetUniqueValues) Utils.getUniqueValuesFromColumn = originalGetUniqueValues;
      if (originalAppendRows) Utils.appendRowsToSheet = originalAppendRows;
    });

    Utils.it("utiliza correctamente los normalizadores de keys", () => {
      const mockContext = { shouldStop: false, remainingTime: () => 300000 };
      
      EC.ExecutionController = {
        runLoop: function(ctx, files, processor) {
          return {
            stoppedEarly: false,
            invalidCtx: false,
            successes: [{ processed: 1, added: 1, updated: 0, fileName: "test.csv", date: new Date() }],
            fails: []
          };
        }
      };

      let normalizerType = "";
      Utils.getNormalizer = function(type) {
        normalizerType = type;
        if (type === "integer") return v => Number(v);
        if (type === "float") return v => Number(v);
        if (type === "boolean") return v => Boolean(v);
        if (type === "date") return v => new Date(v);
        if (type === "string") return v => String(v).toLowerCase();
        return v => v;
      };

      Utils.backupFileTo = function() {};
      Utils.getUniqueValuesFromColumn = function() { return new Set(); };
      Utils.appendRowsToSheet = function() {};

      const configs = {
        [MimeType.CSV]: {
          columns: [0, 1, 2],
          searchingFn: function(mimeType) {
            if (mimeType === MimeType.CSV) return [{
              getId: () => "csv1",
              getName: () => "test.csv", 
              getMimeType: () => MimeType.CSV,
              getBlob: () => ({ getDataAsString: () => "1,test,test@email.com" })
            }];
            return [];
          },
          ignoreEmptyRows: true,
          updateExistingRows: false,
          allowFileReprocessing: true
        }
      };

      lib.mergeSheetsDataTo(
        mockContext,
        { 
          fileID: "target-ss-id", 
          fileSheetName: "Hoja1", 
          columnTitles: ["ID", "Name", "Email"],
          keyColumnIndex: 0, 
          keyColumnType: "string",
        },
        configs,
        { ...options, logSpreadsheet: mockLogSpreadsheet, logSheetName: "Hoja1", keepProcessedInSource: true  }
      );

      Utils.assertEquals(normalizerType, "string");

      
      if (originalExecutionController) EC.ExecutionController = originalExecutionController;
      if (originalGetNormalizer) Utils.getNormalizer = originalGetNormalizer;
      if (originalBackupFileTo) Utils.backupFileTo = originalBackupFileTo;
      if (originalGetUniqueValues) Utils.getUniqueValuesFromColumn = originalGetUniqueValues;
      if (originalAppendRows) Utils.appendRowsToSheet = originalAppendRows;
    });

    
    if (oldOpenById) SpreadsheetApp.openById = oldOpenById;
    if (oldGetFileById) DriveApp.getFileById = oldGetFileById;
    if (oldGetFolderById) DriveApp.getFolderById = oldGetFolderById;
    if (oldGetOrCreate) Utils.getOrCreateSubfolderFrom = oldGetOrCreate;
  });

  Utils.endTests();
}
