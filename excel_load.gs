/* =========== Main CIM Load File Creation function =========== */

function createExcelLoad(e) {
  setupLog_();
  var i, config, configName;
  log_('Running on: ' + now);

  var configs = getConfigs_(getOrCreateSheet_(CONFIG_SHEET));
  var sysConfigs = getSysConfigs_(getOrCreateSheet_(CONFIG_SHEET));

  if (!configs.length) {
    log_('No Excel Load configurations found');
  } else {
    log_('Found ' + configs.length + ' Excel Load configurations.');

    for (i = 0; config = configs[i]; ++i) {
      //Logger.log(config);
      configName = config.file_name;
      //Logger.log(configName);
      //Logger.log(config['sheet_name']);
      if (config['sheet_name']) {
        try {
          log_('Creating Excel Load for: ' + configName);
          switch(configName) {
            case 'supplier_items_VIS000_<YYYYMMDD>':
              loadTemplateSI(sysConfigs[0],config);
              break;
            case 'supplier_pricelist_VIS000_<YYYYMMDD>':
              loadTemplateSPL(sysConfigs[0],config);
              break;
            case 'item_site_costs_VIS000_<YYYYMMDD>':
              loadTemplateISC(sysConfigs[0],config);
              break;
            default:
              continue;
          }
        } catch (error) {
          log_('Error executing ' + configName + ': ' + error.message);
        }
      } else {
        log_('No sheet-name found: ' + configName);
      }
    }
  }
  now = new Date();
  log_('Script done: ' + now);

  // Update the user about the status of the queries.
  if( e === undefined ) {
    showLogDialog_();
    dumpLog_(getOrCreateSheet_(LOG_SHEET));
    dumpError_(getOrCreateSheet_(ERROR_SHEET));
  }
}

/* =========== Load AUX Template functions =========== */

/* == Supplier Items ================================= */
function loadTemplateSI(sysConfig,config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(config.worksheet);
  var sourceRange = sheet.getRange(2, 1, sheet.getLastRow(),config.wksheet_cols);
  var sheetData = sourceRange.getValues();
  var numRows = sourceRange.getNumRows();
  //  Logger.log(numRows);
  var startRow = config.row_start;
  var data = [];
  /*
  var options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
  sysConfigs.quote_date.toLocaleDateString("en-GB");
  spreadsheet_date
  expire_date
  var funcString = '(function(){ return 4+4 })'
  var result = eval(funcString)()
  var fn = Function("alert('hello there')");
  fn();
  var addition = Function("a", "b", "return a + b;");
  alert(addition(5, 3)); // shows '8'
  */

  for (var i=0; i < numRows; i++) {
    if (sheetData[i][0] === '' || sheetData[i][20] === '') {
      continue;
    }
    else {
      var v_delete                  = "No"                    //                   /           /Delete Record
      var v_DomainCode              = sysConfig.domain_code   // Supplier Item Keys/           /Domain [Mandatory]
      var v_ItemCode                = sheetData[i][20]        //                   /           /Item [Mandatory]
      var v_SupplierCode            = sysConfig.supplier      //                   /           /Supplier [Mandatory]
      var v_SupplierItem            = String(sheetData[i][3]) //                   /           /Supplier Item [Mandatory]
      var v_rowData                 = ""                      //                   /           /Row Data (a formula)
      var v_ItemDescription         = ""                      // Main              /           /Description
      var v_SupplierName            = ""                      //                   /           /Name
      var v_UnitOfMeasure           = sysConfig.um            //                   /           /Unit of Measure [Mandatory]
      var v_SupplierLeadTime        = "0"                     //                   /           /Supplier Lead Time
      var v_IsUseSOReductionPrice   = "No"                    // Price             /           /Use SO Reduction Price
      var v_SOPriceReduction        = "0.00%"                 //                   /           /SO Price Reduction
      var v_PriceList               = sysConfig.price_list    //                   /           /Price List
      var v_CurrencyCode            = sysConfig.currency_code //                   /           /Currency [Mandatory]
      var v_CurrencyDescription     = ""                      //                   /           /(currencyDescription)
      var v_QuotePrice              = sheetData[i][14]        //                   /Quote Price/Quote Price
      var v_QuoteDate               = "30/06/2024"            //                   /           /Quote Date
      var v_QuoteQuantity           = "1000"                  //                   /           /Quote Quantity
      var v_PositivePlanVariance    = ""                      // Release Scheduling/           /Positive Plan Variance
      var v_NegativePlanVariance    = ""                      //                   /           /Negative Plan Variance
      var v_PositiveShipVariance    = ""                      //                   /           /Positive Ship Variance
      var v_NegativeShipVariance    = ""                      //                   /           /Negative Ship Variance
      var v_Manufacturer            = ""                      // Manufacturing     /           /Manufacturer
      var v_ManufacturerDescription = ""                      //                   /           /(manufacturerDescription)
      var v_ManufacturerItem        = ""                      //                   /           /Manufacturer Item
      var v_Comment                 = ""                      //                   /           /Comment
      var v_IsControlSupply         = ""                      // Subcontract       /           /Control Supply
      var v_IsBulkSupply            = ""                      //                   /           /Bulk Supply
      var v_AutoUpdateHorizon       = ""                      //                   /           /Auto Update Horizon      
      
      data.push([v_delete,v_DomainCode,v_ItemCode,v_SupplierCode,v_SupplierItem,v_rowData,v_ItemDescription,v_SupplierName,v_UnitOfMeasure,v_SupplierLeadTime,v_IsUseSOReductionPrice,v_SOPriceReduction,v_PriceList,v_CurrencyCode,v_CurrencyDescription,v_QuotePrice,v_QuoteDate,v_QuoteQuantity,v_PositivePlanVariance,v_NegativePlanVariance,v_PositiveShipVariance,v_NegativeShipVariance,v_Manufacturer,v_ManufacturerDescription,v_ManufacturerItem,v_Comment,v_IsControlSupply,v_IsBulkSupply,v_AutoUpdateHorizon]);;
    }
  };
  /*
  supplier_items_template_v1
  https://docs.google.com/spreadsheets/d/1LQwbbjTt4LmWhZcEyCmELRiP-1c_t5tQVsRuXYh24mw/edit?gid=319957131#gid=319957131
  */
  /*
  var ss = SpreadsheetApp.openById("MYSHEETKEY");
  var newSS = ss.copy("Copy of " + ss.getName());
  // Move to original folder
  var originalFolder = DriveApp.getFileById("MYSHEETKEY").getParents().next();
  var newSSFile = DriveApp.getFileById(newSS.getId());
  originalFolder.addFile(newSSFile);
  DriveApp.getRootFolder().removeFile(newSSFile);
  var ts = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1LQwbbjTt4LmWhZcEyCmELRiP-1c_t5tQVsRuXYh24mw/edit?gid=319957131#gid=319957131");
  */
  var ts = SpreadsheetApp.openByUrl(config.template);
  var tsheet = ts.getSheetByName(config.sheet_name);
  var targetRange = tsheet.getRange(startRow,1,data.length,data[0].length);
  targetRange.setValues(data);

  var fdata = [];
  // populate the array with the formulas.
  for (var i=0; i < data.length; i++)
  {
    fdata[i] = ['=IF(COUNTA(B' + (i+7).toString() + ':E' + (i+7).toString() + ',G' + (i+7).toString() + ':AC' + (i+7).toString() + ')>0,"Supplier Item","")' ];
  }
  // set the column values.
  tsheet.getRange(startRow,config.func_col,data.length,1).setFormulas(fdata);
}

/* == Supplier Price List ============================ */
function loadTemplateSPL(sysConfig,config) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(config.worksheet);
  var sourceRange = sheet.getRange(2, 1, sheet.getLastRow(),config.wksheet_cols);
  var sheetData = sourceRange.getValues();
  var numRows = sourceRange.getNumRows();
  //  Logger.log(numRows);

  var startRow = config.row_start;
  var active = true;
  var v_type = sysConfig.list_type;
  var data = [];

  for (var i=0; i < numRows; i++) {
    /* breqk/end of list condition */
    if (sheetData[i][9] === '') { break; }
    if (sheetData[i][0] !== '' ) {
      if (sheetData[i][23] === 'ACTIVE') {
        active = true;
      }
      else {
        active = false;
      }
    }
    if (active) {
      if (sheetData[i][20] === '' ) {
        if (v_type === "List Price") {
          continue;
        }
        else {
          var v_DomainCode               = ""                       // Supplier Price List Keys/          /          /Domain
          var v_PriceListCode            = ""                       //                         /          /          /Price List
          var v_CurrencyCode             = ""                       //                         /          /          /Currency
          var v_ProductLine              = ""                       //                         /          /          /Product Line
          var v_ItemCode                 = ""                       //                         /          /          /Item
          var v_UnitOfMeasure            = ""                       //                         /          /          /Unit of Measure
          var v_StartDate                = ""                       //                         /          /          /Start Date
          var v_rowData                  = ""                       //                         /          /          /Row Data
          var v_PriceListDescription     = ""                       // Main                    /          /          /Description
          var v_ProdLineDescription      = ""                       //                         /          /          /Product Line Description
          var v_ItemDescription          = ""                       //                         /          /          /Item Description
          var v_AmountType               = ""                       //                         /          /          /Amount Type
          var v_CurrencyDescription      = ""                       //                         /          /          /(currencyDescription)
          var v_UnitOfMeasureDescription = ""                       //                         /          /          /(unitOfMeasureDescription)
          var v_ExpireDate               = ""                       //                         /          /          /Expiration Date
          var v_IsTemporary              = ""                       //                         /          /          /Temporary
          var v_ItemListPrice            = ""                       //                         /Item      /          /Item Master List Price
          var v_DomainCurrency           = ""                       //                         /          /          /(domainCurrency)
          var v_ThisLevelGLCost          = ""                       //                         /          /          /Total This Level GL Cost
          var v_TotalGLCost              = ""                       //                         /          /          /Total GL Cost
          var v_StockUOM                 = ""                       //                         /          /          /Stock UM
          var v_StockUOMDescription      = ""                       //                         /          /          /(stockUOMDescription)
          var v_SiteCode                 = ""                       //                         /          /          /Site
          var v_SiteDescription          = ""                       //                         /          /          /(siteDescription)
          var vs_DomainCode              = sysConfig.domain_code    //                         /Price List/Price List/(domainCode)
          var vs_PriceListCode           = sysConfig.price_list     //                         /          /          /Price List
          var vs_CurrencyCode            = sysConfig.currency_code  //                         /          /          /Currency
          var vs_ProductLine             = ""                       //                         /          /          /Product Line
          var vs_ItemCode                = sheetData[i][20]         //                         /          /          /Item
          var vs_UnitOfMeasure           = sysConfig.um             //                         /          /          /Unit of Measure
          var vs_StartDate               = "01/07/2024"             //                         /          /          /Start Date
          var v_MinQuantity              = "0"                      //                         /          /          /Minimum Quantity
          var v_PriceAmount              = sheetData[i][16]         //                         /          /          /Amount
          var v_ListPrice                = ""                       //                         /Price List/          /List Price
          var v_MinimumPrice             = ""                       //                         /          /          /Minimum Price
          var v_MaximumPrice             = ""                       //                         /          /          /Maximum Price

          data.push([v_DomainCode,v_PriceListCode,v_CurrencyCode,v_ProductLine,v_ItemCode,v_UnitOfMeasure,v_StartDate,v_rowData,v_PriceListDescription,v_ProdLineDescription,v_ItemDescription,v_AmountType,v_CurrencyDescription,v_UnitOfMeasureDescription,v_ExpireDate,v_IsTemporary,v_ItemListPrice,v_DomainCurrency,v_ThisLevelGLCost,v_TotalGLCost,v_StockUOM,v_StockUOMDescription,v_SiteCode,v_SiteDescription,vs_DomainCode,vs_PriceListCode,vs_CurrencyCode,vs_ProductLine,vs_ItemCode,vs_UnitOfMeasure,vs_StartDate,v_MinQuantity,v_PriceAmount,v_ListPrice,v_MinimumPrice,v_MaximumPrice]);
        }
      }

      else {
        v_DomainCode               = sysConfig.domain_code    // Supplier Price List Keys/          /          /Domain
        v_PriceListCode            = sysConfig.price_list     //                         /          /          /Price List
        v_CurrencyCode             = sysConfig.currency_code  //                         /          /          /Currency
        v_ProductLine              = ""                       //                         /          /          /Product Line
        v_ItemCode                 = String(sheetData[i][20]) //                         /          /          /Item
        v_UnitOfMeasure            = sysConfig.um             //                         /          /          /Unit of Measure
        v_StartDate                = "01/07/2024"             //                         /          /          /Start Date
        v_rowData                  = ""                       //                         /          /          /Row Data
        v_PriceListDescription     = sheetData[i][4]          // Main                    /          /          /Description
        v_ProdLineDescription      = ""                       //                         /          /          /Product Line Description
        v_ItemDescription          = sheetData[i][21]         //                         /          /          /Item Description
        v_AmountType               = v_type                   //                         /          /          /Amount Type
        v_CurrencyDescription      = ""                       //                         /          /          /(currencyDescription)
        v_UnitOfMeasureDescription = ""                       //                         /          /          /(unitOfMeasureDescription)
        v_ExpireDate               = "30/06/2025"             //                         /          /          /Expiration Date
        v_IsTemporary              = "No"                     //                         /          /          /Temporary
        v_ItemListPrice            = sheetData[i][16]         //                         /Item      /          /Item Master List Price
        v_DomainCurrency           = sysConfig.currency_code  //                         /          /          /(domainCurrency)
        v_ThisLevelGLCost          = sheetData[i][16]         //                         /          /          /Total This Level GL Cost
        v_TotalGLCost              = sheetData[i][16]         //                         /          /          /Total GL Cost
        v_StockUOM                 = sysConfig.um             //                         /          /          /Stock UM
        v_StockUOMDescription      = ""                       //                         /          /          /(stockUOMDescription)
        v_SiteCode                 = sysConfig.site           //                         /          /          /Site
        v_SiteDescription          = ""                       //                         /          /          /(siteDescription)
        vs_DomainCode              = ""                       //                         /Price List/Price List/(domainCode)
        vs_PriceListCode           = ""                       //                         /          /          /Price List
        vs_CurrencyCode            = ""                       //                         /          /          /Currency
        vs_ProductLine             = ""                       //                         /          /          /Product Line
        vs_ItemCode                = ""                       //                         /          /          /Item
        vs_UnitOfMeasure           = ""                       //                         /          /          /Unit of Measure
        vs_StartDate               = ""                       //                         /          /          /Start Date
        v_MinQuantity              = ""                       //                         /          /          /Minimum Quantity
        v_PriceAmount              = ""                       //                         /          /          /Amount
        v_ListPrice                = sheetData[i][16]         //                         /Price List/          /List Price
        v_MinimumPrice             = sheetData[i][16]         //                         /          /          /Minimum Price
        v_MaximumPrice             = sheetData[i][16]         //                         /          /          /Maximum Price

        data.push([v_DomainCode,v_PriceListCode,v_CurrencyCode,v_ProductLine,v_ItemCode,v_UnitOfMeasure,v_StartDate,v_rowData,v_PriceListDescription,v_ProdLineDescription,v_ItemDescription,v_AmountType,v_CurrencyDescription,v_UnitOfMeasureDescription,v_ExpireDate,v_IsTemporary,v_ItemListPrice,v_DomainCurrency,v_ThisLevelGLCost,v_TotalGLCost,v_StockUOM,v_StockUOMDescription,v_SiteCode,v_SiteDescription,vs_DomainCode,vs_PriceListCode,vs_CurrencyCode,vs_ProductLine,vs_ItemCode,vs_UnitOfMeasure,vs_StartDate,v_MinQuantity,v_PriceAmount,v_ListPrice,v_MinimumPrice,v_MaximumPrice]);

        v_DomainCode               = ""                       // Supplier Price List Keys/          /          /Domain
        v_PriceListCode            = ""                       //                         /          /          /Price List
        v_CurrencyCode             = ""                       //                         /          /          /Currency
        v_ProductLine              = ""                       //                         /          /          /Product Line
        v_ItemCode                 = ""                       //                         /          /          /Item
        v_UnitOfMeasure            = ""                       //                         /          /          /Unit of Measure
        v_StartDate                = ""                       //                         /          /          /Start Date
        v_rowData                  = ""                       //                         /          /          /Row Data
        v_PriceListDescription     = ""                       // Main                    /          /          /Description
        v_ProdLineDescription      = ""                       //                         /          /          /Product Line Description
        v_ItemDescription          = ""                       //                         /          /          /Item Description
        v_AmountType               = ""                       //                         /          /          /Amount Type
        v_CurrencyDescription      = ""                       //                         /          /          /(currencyDescription)
        v_UnitOfMeasureDescription = ""                       //                         /          /          /(unitOfMeasureDescription)
        v_ExpireDate               = ""                       //                         /          /          /Expiration Date
        v_IsTemporary              = ""                       //                         /          /          /Temporary
        v_ItemListPrice            = ""                       //                         /Item      /          /Item Master List Price
        v_DomainCurrency           = ""                       //                         /          /          /(domainCurrency)
        v_ThisLevelGLCost          = ""                       //                         /          /          /Total This Level GL Cost
        v_TotalGLCost              = ""                       //                         /          /          /Total GL Cost
        v_StockUOM                 = ""                       //                         /          /          /Stock UM
        v_StockUOMDescription      = ""                       //                         /          /          /(stockUOMDescription)
        v_SiteCode                 = ""                       //                         /          /          /Site
        v_SiteDescription          = ""                       //                         /          /          /(siteDescription)
        vs_DomainCode              = sysConfig.domain_code    //                         /Price List/Price List/(domainCode)
        vs_PriceListCode           = sysConfig.price_list     //                         /          /          /Price List
        vs_CurrencyCode            = sysConfig.currency_code  //                         /          /          /Currency
        vs_ProductLine             = ""                       //                         /          /          /Product Line
        vs_ItemCode                = sheetData[i][20]         //                         /          /          /Item
        vs_UnitOfMeasure           = sysConfig.um             //                         /          /          /Unit of Measure
        vs_StartDate               = "01/07/2024"             //                         /          /          /Start Date
        v_MinQuantity              = "0"                      //                         /          /          /Minimum Quantity
        v_PriceAmount              = sheetData[i][16]         //                         /          /          /Amount
        v_ListPrice                = ""                       //                         /Price List/          /List Price
        v_MinimumPrice             = ""                       //                         /          /          /Minimum Price
        v_MaximumPrice             = ""                       //                         /          /          /Maximum Price
    
        data.push([v_DomainCode,v_PriceListCode,v_CurrencyCode,v_ProductLine,v_ItemCode,v_UnitOfMeasure,v_StartDate,v_rowData,v_PriceListDescription,v_ProdLineDescription,v_ItemDescription,v_AmountType,v_CurrencyDescription,v_UnitOfMeasureDescription,v_ExpireDate,v_IsTemporary,v_ItemListPrice,v_DomainCurrency,v_ThisLevelGLCost,v_TotalGLCost,v_StockUOM,v_StockUOMDescription,v_SiteCode,v_SiteDescription,vs_DomainCode,vs_PriceListCode,vs_CurrencyCode,vs_ProductLine,vs_ItemCode,vs_UnitOfMeasure,vs_StartDate,v_MinQuantity,v_PriceAmount,v_ListPrice,v_MinimumPrice,v_MaximumPrice]);
      }
    };
  };

  /*
  var targetRange = sheet.getRange(2,48,data.length,36);
  targetRange.setValues(data);
  */
  /*
  supplier_pricelist_template_v1
  https://docs.google.com/spreadsheets/d/1XbWpYrRycjRPUZeQibW9FeVrhzz0fJra5TC5NXxdFy0/edit?gid=1932238525#gid=1932238525
  var targetRange = sheet.getRange(2,48,data.length,36);
  targetRange.setValues(data);
  var ts = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1XbWpYrRycjRPUZeQibW9FeVrhzz0fJra5TC5NXxdFy0/edit?gid=1932238525#gid=1932238525");
  */
  var ts = SpreadsheetApp.openByUrl(config.template);
  var tsheet = ts.getSheetByName(config.sheet_name);
  var targetRange = tsheet.getRange(startRow,1,data.length,data[0].length);
  targetRange.setValues(data);

  var fdata = [];
  // populate the array with the formulas.
  for (var i=0; i < data.length; i++)
  {
    fdata[i] = ['=IF(COUNTA(A' + (i+8).toString() + ':G' + (i+8).toString() + ',I' + (i+8).toString() + ':X' + (i+8).toString() + ',AH' + (i+8).toString() + ':AJ' + (i+8).toString() + ')>0,"Supplier Price List",IF(COUNTA(Y' + (i+8).toString() + ':AG' + (i+8).toString() + ')>0,"Price List",""))' ];
  }
  // set the column values.
  tsheet.getRange(startRow,config.func_col,data.length,1).setFormulas(fdata);

}

/* == Item Site Costs ================================ */
function loadTemplateISC(sysConfig,config) {
  var startRow = 10;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Worksheet");

  var sourceRange = sheet.getRange(2, 1, sheet.getLastRow(),47);
  var sheetData = sourceRange.getValues();
  var data = [];

  var numRows = sourceRange.getNumRows();
  //  Logger.log(numRows);

  for (var i=0; i < numRows; i++) {
    if (sheetData[i][0] === '' ) {
      continue;
    }
    else {
      data.push(["DBWAUS","Standard",sheetData[i][34],"11","","","","","","","","","","","","","","","","","","","","","","",]);
      data.push(["","","","","","","","","","","","","","","","DBWAUS","Standard",sheetData[i][34],"11","Material","",sheetData[i][14],"","","Yes","No",]);
    }
  };
  var targetRange = sheet.getRange(2,112,data.length,26);
  targetRange.setValues(data);




  var startRow = 10;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Visy Board Worksheet");

  var sourceRange = sheet.getRange(2, 1, sheet.getLastRow(),47);
  var sheetData = sourceRange.getValues();
  var data = [];
  var active = true;

  var numRows = sourceRange.getNumRows();
  //  Logger.log(numRows);

  for (var i=0; i < numRows; i++) {
    if (sheetData[i][0] !== '' ) {
      if (sheetData[i][21] === 'ACTIVE') {
        active = true;
      }
      else {
        active = false;
      }
    }
    if (sheetData[i][0] === '' ) {
      continue;
    }
    else {
      if (active) {
        var v_DomainCode         = "DBWAUS"         // /Item Site Cost Keys/     /Domain
        var v_CostSet            = "Standard"       // /                   /     /Cost Set
        var v_ItemCode           = sheetData[i][34] // /                   /     /Item
        var v_SiteCode           = "11"             // /                   /     /Site
        var v_rowData            = ""               // /                   /     /Row Data
        var v_ItemDescription    = ""               // /Main               /     /Description 1
        var v_UnitOfMeasure      = ""               // /                   /     /Unit of Measure
        var v_SiteDescription    = ""               // /                   /     /Description
        var v_CostSetDescription = ""               // /                   /     /Description
        var v_CostSetType        = ""               // /                   /     /Cost Set Type
        var v_CostMethod         = ""               // /                   /     /Costing Method
        var v_CostUpdate         = ""               // /                   /     /Cost Update
        var v_ThisLevelTotal     = ""               // /                   /     /This Level Total
        var v_LowerLevelTotal    = ""               // /                   /     /Lower Level Total
        var v_CostTotal          = ""               // /                   /     /Cost Total
        var v_DomainCode         = ""               // /Costs              /Costs/Domain
        var v_CostSet            = ""               // /                   /     /Cost Set
        var v_ItemCode           = ""               // /                   /     /Item Number
        var v_SiteCode           = ""               // /                   /     /Site
        var v_CostElement        = ""               // /                   /     /Element
        var v_CostCategory       = ""               // /                   /     /Category
        var v_ThisLevelCost      = ""               // /                   /     /This Level
        var v_LowerLevelCost     = ""               // /                   /     /Lower Level
        var v_TotalCost          = ""               // /                   /     /Total
        var v_IsPrimary          = ""               // /                   /     /Primary
        var v_IsAddOn            = ""               // /                   /     /Add On
        data.push([v_DomainCode,v_CostSet,v_ItemCode,v_SiteCode,v_rowData,v_ItemDescription,v_UnitOfMeasure,v_SiteDescription,v_CostSetDescription,v_CostSetType,v_CostMethod,v_CostUpdate,v_ThisLevelTotal,v_LowerLevelTotal,v_CostTotal,v_DomainCode,v_CostSet,v_ItemCode,v_SiteCode,v_CostElement,v_CostCategory,v_ThisLevelCost,v_LowerLevelCost,v_TotalCost,v_IsPrimary,v_IsAddOn]);

        v_DomainCode         = "DBWAUS"         // /Item Site Cost Keys/     /Domain
        v_CostSet            = "Standard"       // /                   /     /Cost Set
        v_ItemCode           = sheetData[i][34] // /                   /     /Item
        v_SiteCode           = "11"             // /                   /     /Site
        v_rowData            = ""               // /                   /     /Row Data
        v_ItemDescription    = ""               // /Main               /     /Description 1
        v_UnitOfMeasure      = ""               // /                   /     /Unit of Measure
        v_SiteDescription    = ""               // /                   /     /Description
        v_CostSetDescription = ""               // /                   /     /Description
        v_CostSetType        = ""               // /                   /     /Cost Set Type
        v_CostMethod         = ""               // /                   /     /Costing Method
        v_CostUpdate         = ""               // /                   /     /Cost Update
        v_ThisLevelTotal     = ""               // /                   /     /This Level Total
        v_LowerLevelTotal    = ""               // /                   /     /Lower Level Total
        v_CostTotal          = ""               // /                   /     /Cost Total
        v_DomainCode         = ""               // /Costs              /Costs/Domain
        v_CostSet            = ""               // /                   /     /Cost Set
        v_ItemCode           = ""               // /                   /     /Item Number
        v_SiteCode           = ""               // /                   /     /Site
        v_CostElement        = ""               // /                   /     /Element
        v_CostCategory       = ""               // /                   /     /Category
        v_ThisLevelCost      = ""               // /                   /     /This Level
        v_LowerLevelCost     = ""               // /                   /     /Lower Level
        v_TotalCost          = ""               // /                   /     /Total
        v_IsPrimary          = ""               // /                   /     /Primary
        v_IsAddOn            = ""               // /                   /     /Add On
        data.push([v_DomainCode,v_CostSet,v_ItemCode,v_SiteCode,v_rowData,v_ItemDescription,v_UnitOfMeasure,v_SiteDescription,v_CostSetDescription,v_CostSetType,v_CostMethod,v_CostUpdate,v_ThisLevelTotal,v_LowerLevelTotal,v_CostTotal,v_DomainCode,v_CostSet,v_ItemCode,v_SiteCode,v_CostElement,v_CostCategory,v_ThisLevelCost,v_LowerLevelCost,v_TotalCost,v_IsPrimary,v_IsAddOn]);
      }
    }
  };
  
  /*
  https://docs.google.com/spreadsheets/d/11_J6_OzIPN-IcrhZmPaAPpfQLYHl3wK8iumsFqErVag/edit?gid=1594290546#gid=1594290546
    var targetRange = sheet.getRange(2,112,data.length,26);
  targetRange.setValues(data);
  */
  
  var ts = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/11_J6_OzIPN-IcrhZmPaAPpfQLYHl3wK8iumsFqErVag/edit?gid=1594290546#gid=1594290546");
  var tsheet = ts.getSheetByName("Data");
  var targetRange = tsheet.getRange(7,1,data.length,26);
  targetRange.setValues(data);

  var fdata = [];
  // populate the array with the formulas.
  for (var i=0; i < data.length; i++)
  {
    fdata[i] = ['=IF(COUNTA(A' + (i+7).toString() + ':D' + (i+7).toString() + ',F' + (i+7).toString() + ':O' + (i+7).toString() + ')>0,"Item Site Cost",IF(COUNTA(P' + (i+7).toString() + ':Z' + (i+7).toString() + ')>0,"Costs",""))' ];
  }
  // set the column values.
  tsheet.getRange(7,5,data.length,1).setFormulas(fdata);

  




 


  
}


/* == From visy board ==================================================================================================================================================================*/

/* == Supplier Price List ============================ */
function loadTemplatea() {
  var startRow = 10;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Visy Board Worksheet");

  var sourceRange = sheet.getRange(2, 1, sheet.getLastRow(),47);
  var sheetData = sourceRange.getValues();
  var data = [];
  var active = true;

  var numRows = sourceRange.getNumRows();
  //  Logger.log(numRows);

  for (var i=0; i < numRows; i++) {
    /* breqk/end of list condition */
    if (sheetData[i][9] === '') { break; }
    if (sheetData[i][0] !== '' ) {
      if (sheetData[i][21] === 'ACTIVE') {
        active = true;
      }
      else {
        active = false;
      }
    }
    if (active) {
      if (sheetData[i][0] === '' ) {
        data.push(["","","","","","","","","","","","","","","","","","","","","","","","","DBWAUS","VIS000H","AUD","",sheetData[hrow][17],"ea","08/04/2024",sheetData[i][9],sheetData[i][14],"","",""]);
      }
      else {
        var hrow = i;
        data.push(["DBWAUS","VIS000H","AUD","",sheetData[i][17],"ea","08/04/2024","",sheetData[i][1],"",sheetData[i][18],"Price","","","30/06/2025","No","","","","","","","","","","","","","","","","","","","0","0"]);
        data.push(["","","","","","","","","","","","","","","","","","","","","","","","","DBWAUS","VIS000H","AUD","",sheetData[hrow][17],"ea","08/04/2024",sheetData[i][9],sheetData[i][14],"","",""]);
      }
    }
  };

  /*
  supplier_pricelist_template_v1
  https://docs.google.com/spreadsheets/d/1XbWpYrRycjRPUZeQibW9FeVrhzz0fJra5TC5NXxdFy0/edit?gid=1932238525#gid=1932238525
  var targetRange = sheet.getRange(2,48,data.length,36);
  targetRange.setValues(data);
  */
  var ts = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1XbWpYrRycjRPUZeQibW9FeVrhzz0fJra5TC5NXxdFy0/edit?gid=1932238525#gid=1932238525");
  var tsheet = ts.getSheetByName("Data");
  var targetRange = tsheet.getRange(8,1,data.length,36);
  targetRange.setValues(data);

  var fdata = [];
  // populate the array with the formulas.
  for (var i=0; i < data.length; i++)
  {
    fdata[i] = ['=IF(COUNTA(A' + (i+8).toString() + ':G' + (i+8).toString() + ',I' + (i+8).toString() + ':X' + (i+8).toString() + ',AH' + (i+8).toString() + ':AJ' + (i+8).toString() + ')>0,"Supplier Price List",IF(COUNTA(Y' + (i+8).toString() + ':AG' + (i+8).toString() + ')>0,"Price List",""))' ];
  }
  // set the column values.
  tsheet.getRange(8,8,data.length,1).setFormulas(fdata);
}

/* == Supplier Items ================================= */
function loadTemplate2a() {
  var startRow = 10;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Visy Board Worksheet");

  var sourceRange = sheet.getRange(2, 1, sheet.getLastRow(),47);
  var sheetData = sourceRange.getValues();
  var data = [];

  var numRows = sourceRange.getNumRows();
  //  Logger.log(numRows);

  for (var i=0; i < numRows; i++) {
    if (sheetData[i][0] === '' || sheetData[i][17] === '-') {
      continue;
    }
    else {
      data.push(["No","DBWAUS",sheetData[i][17],"VIS000",sheetData[i][0],"","","","EA","","","","","AUD","","","","","","","","","","","","","","","",]);;
    }
  };

  /*
  supplier_items_template_v1
  https://docs.google.com/spreadsheets/d/1LQwbbjTt4LmWhZcEyCmELRiP-1c_t5tQVsRuXYh24mw/edit?gid=319957131#gid=319957131
  */
  /*
  var ss = SpreadsheetApp.openById("MYSHEETKEY");
  var newSS = ss.copy("Copy of " + ss.getName());
  // Move to original folder
  var originalFolder = DriveApp.getFileById("MYSHEETKEY").getParents().next();
  var newSSFile = DriveApp.getFileById(newSS.getId());
  originalFolder.addFile(newSSFile);
  DriveApp.getRootFolder().removeFile(newSSFile);
  */
  var ts = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1LQwbbjTt4LmWhZcEyCmELRiP-1c_t5tQVsRuXYh24mw/edit?gid=319957131#gid=319957131");
  var tsheet = ts.getSheetByName("Data");
  var targetRange = tsheet.getRange(7,1,data.length,29);
  targetRange.setValues(data);

  var fdata = [];
  // populate the array with the formulas.
  for (var i=0; i < data.length; i++)
  {
    fdata[i] = ['=IF(COUNTA(B' + (i+7).toString() + ':E' + (i+7).toString() + ',G' + (i+7).toString() + ':AC' + (i+7).toString() + ')>0,"Supplier Item","")' ];
  }
  // set the column values.
  tsheet.getRange(7,6,data.length,1).setFormulas(fdata);

}

/* == Item Site Costs ================================ */
function loadTemplate3a() {
  var startRow = 10;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Visy Board Worksheet");

  var sourceRange = sheet.getRange(2, 1, sheet.getLastRow(),47);
  var sheetData = sourceRange.getValues();
  var data = [];
  var active = true;

  var numRows = sourceRange.getNumRows();
  //  Logger.log(numRows);

  for (var i=0; i < numRows; i++) {
    if (sheetData[i][0] !== '' ) {
      if (sheetData[i][21] === 'ACTIVE') {
        active = true;
      }
      else {
        active = false;
      }
    }
    if (sheetData[i][0] === '' ) {
      continue;
    }
    else {
      if (active) {
        data.push(["DBWAUS","Standard",sheetData[i][17],"11","","","","","","","","","","","","","","","","","","","","","","",]);
        data.push(["","","","","","","","","","","","","","","","DBWAUS","Standard",sheetData[i][17],"11","Material","",sheetData[i][42],"","","Yes","No",]);
      }
    }
  };

  /*
  https://docs.google.com/spreadsheets/d/11_J6_OzIPN-IcrhZmPaAPpfQLYHl3wK8iumsFqErVag/edit?gid=1594290546#gid=1594290546
    var targetRange = sheet.getRange(2,112,data.length,26);
  targetRange.setValues(data);
  */
  var ts = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/11_J6_OzIPN-IcrhZmPaAPpfQLYHl3wK8iumsFqErVag/edit?gid=1594290546#gid=1594290546");
  var tsheet = ts.getSheetByName("Data");
  var targetRange = tsheet.getRange(7,1,data.length,26);
  targetRange.setValues(data);

  var fdata = [];
  // populate the array with the formulas.
  for (var i=0; i < data.length; i++)
  {
    fdata[i] = ['=IF(COUNTA(A' + (i+7).toString() + ':D' + (i+7).toString() + ',F' + (i+7).toString() + ':O' + (i+7).toString() + ')>0,"Item Site Cost",IF(COUNTA(P' + (i+7).toString() + ':Z' + (i+7).toString() + ')>0,"Costs",""))' ];
  }
  // set the column values.
  tsheet.getRange(7,5,data.length,1).setFormulas(fdata);
}

