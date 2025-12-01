/**
 * A+ Reports Dashboard - Complete with Historical Charts
 */

// Configuration
var CONFIG = {
  sourceFolderName: 'A+Transformed_Reports',
  transformedFolderName: 'A+_Transformed_Sheets',
  dashboardSheetName: 'A+ Dashboard',
  alertsSheetName: 'High Utilization Alerts',
  pivotSheetName: 'Utilization Trends',
  chartsSheetName: 'Historical Charts',
  thresholdHigh: 90,
  thresholdMedium: 75
};

// === FOLDER MANAGEMENT FUNCTIONS ===

function getOrCreateSourceFolder() {
  try {
    // Get the current spreadsheet's parent folder
    var currentFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var parentFolder = currentFile.getParents().next();
    
    // Check if folder already exists in current directory
    var folders = parentFolder.getFoldersByName(CONFIG.sourceFolderName);
    if (folders.hasNext()) {
      return folders.next();
    }
    
    // Create folder in current directory
    return parentFolder.createFolder(CONFIG.sourceFolderName);
    
  } catch (error) {
    Logger.log('Error getting source folder: ' + error);
    // Fallback to root if can't access parent
    var folders = DriveApp.getFoldersByName(CONFIG.sourceFolderName);
    if (folders.hasNext()) {
      return folders.next();
    }
    return DriveApp.createFolder(CONFIG.sourceFolderName);
  }
}

function getOrCreateTransformedFolder() {
  try {
    // Get the current spreadsheet's parent folder
    var currentFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var parentFolder = currentFile.getParents().next();
    
    // Check if folder already exists in current directory
    var folders = parentFolder.getFoldersByName(CONFIG.transformedFolderName);
    if (folders.hasNext()) {
      return folders.next();
    }
    
    // Create folder in current directory
    return parentFolder.createFolder(CONFIG.transformedFolderName);
    
  } catch (error) {
    Logger.log('Error getting transformed folder: ' + error);
    // Fallback to root if can't access parent
    var folders = DriveApp.getFoldersByName(CONFIG.transformedFolderName);
    if (folders.hasNext()) {
      return folders.next();
    }
    return DriveApp.createFolder(CONFIG.transformedFolderName);
  }
}

function getCurrentDirectoryPath() {
  try {
    var currentFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    var parentFolder = currentFile.getParents().next();
    return parentFolder.getName() + ' > ';
  } catch (error) {
    return '';
  }
}

function getXLSXFiles(folder) {
  var files = [];
  
  // Get Excel files
  try {
    var excelFiles = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
    while (excelFiles.hasNext()) {
      files.push(excelFiles.next());
    }
  } catch (e) {
    Logger.log('Error getting Excel files: ' + e);
  }
  
  // Get XLSX files
  try {
    var xlsxFiles = folder.getFilesByType('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    while (xlsxFiles.hasNext()) {
      var file = xlsxFiles.next();
      if (!files.some(f => f.getId() === file.getId())) {
        files.push(file);
      }
    }
  } catch (e) {
    Logger.log('Error getting XLSX files: ' + e);
  }
  
  // Also search by file extension
  try {
    var allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      var file = allFiles.next();
      if (file.getName().toLowerCase().endsWith('.xlsx') && !files.some(f => f.getId() === file.getId())) {
        files.push(file);
      }
    }
  } catch (e) {
    Logger.log('Error searching by extension: ' + e);
  }
  
  return files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
}

function getGoogleSheetsFiles(folder) {
  var files = [];
  var fileIterator = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (fileIterator.hasNext()) {
    files.push(fileIterator.next());
  }
  return files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
}

// === MAIN PROCESSING FUNCTIONS ===

function processAllReports() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getActiveSheet();
    logSheet.clear();
    logSheet.getRange('A1').setValue('🔄 Processing A+ Reports...');
    
    var processedResult = processConvertedSheets();
    
    if (processedResult.processedCount > 0) {
      logSheet.getRange('A1').setValue('✅ Processing Complete!');
      logSheet.getRange('A2').setValue('📊 Processed: ' + processedResult.processedCount + ' files');
      logSheet.getRange('A3').setValue('⚠️ Alerts: ' + processedResult.totalAlerts + ' found');
      logSheet.getRange('A4').setValue('📈 Check the Dashboard, Alerts, and Utilization Trends sheets');
    } else {
      logSheet.getRange('A1').setValue('❌ No files processed');
    }
    
  } catch (error) {
    Logger.log('Main Error: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('A1').setValue('Error: ' + error.toString());
  }
}

function processConvertedSheets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var transformedFolder = getOrCreateTransformedFolder();
    var sheetFiles = getGoogleSheetsFiles(transformedFolder);
    
    Logger.log('Processing ' + sheetFiles.length + ' converted sheets');
    
    if (sheetFiles.length === 0) {
      return { processedCount: 0, totalAlerts: 0, message: 'No converted sheets found' };
    }
    
    var allAlerts = [];
    var allMetrics = {};
    var processedCount = 0;
    var validFiles = [];
    
    // Process all valid files
    for (var i = 0; i < sheetFiles.length; i++) {
      var file = sheetFiles[i];
      if (isValidConvertedFile(file)) {
        validFiles.push(file);
      }
    }
    
    for (var i = 0; i < validFiles.length; i++) {
      var file = validFiles[i];
      var fileAnalysis = processGoogleSheetFile(file);
      
      if (fileAnalysis && fileAnalysis.metrics) {
        processedCount++;
        allAlerts = allAlerts.concat(fileAnalysis.alerts);
        
        var fileDate = extractDateFromFilename(file.getName());
        allMetrics[fileDate] = {
          metrics: fileAnalysis.metrics,
          fileName: file.getName(),
          date: fileDate
        };
        
        createResourceSheets(ss, file.getName(), fileAnalysis);
      }
    }
    
    if (processedCount > 0) {
      updateMainDashboard(ss, allMetrics);
      updateAlertsDashboard(ss, allAlerts);
      createPivotTableDashboard(ss, allMetrics, validFiles);
      createHistoricalCharts(ss, allMetrics); // NEW: Create charts
    }
    
    return {
      processedCount: processedCount,
      totalAlerts: allAlerts.length,
      message: 'Processed ' + processedCount + ' files with ' + allAlerts.length + ' total alerts'
    };
    
  } catch (error) {
    Logger.log('Processing error: ' + error.toString());
    return { processedCount: 0, totalAlerts: 0, message: error.toString() };
  }
}

function isValidConvertedFile(file) {
  try {
    var spreadsheet = SpreadsheetApp.openById(file.getId());
    var sheets = spreadsheet.getSheets();
    
    var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
    var hasData = false;
    
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (resourceTypes.includes(sheetName)) {
        var dataRange = sheets[i].getDataRange();
        var values = dataRange.getValues();
        // Check if sheet has data (more than just headers)
        if (values.length > 1) {
          hasData = true;
          break;
        }
      }
    }
    
    return hasData;
  } catch (error) {
    Logger.log('Error checking file validity: ' + error);
    return false;
  }
}

function processGoogleSheetFile(sheetFile) {
  try {
    Logger.log('Opening: ' + sheetFile.getName());
    
    var spreadsheet = SpreadsheetApp.openById(sheetFile.getId());
    var metrics = {};
    var alerts = [];
    
    var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
    var sheetsProcessed = 0;
    
    for (var i = 0; i < resourceTypes.length; i++) {
      var resourceType = resourceTypes[i];
      var sheet = spreadsheet.getSheetByName(resourceType);
      
      if (sheet) {
        Logger.log('Processing sheet: ' + resourceType);
        var resourceAnalysis = analyzeResourceSheet(sheet, resourceType);
        metrics[resourceType] = resourceAnalysis.metrics;
        sheetsProcessed++;
        
        for (var j = 0; j < resourceAnalysis.alerts.length; j++) {
          var alert = resourceAnalysis.alerts[j];
          alerts.push({
            instance: alert.instance,
            metric: alert.metric,
            value: alert.value,
            severity: alert.severity,
            resourceType: resourceType,
            file: sheetFile.getName(),
            date: extractDateFromFilename(sheetFile.getName())
          });
        }
      } else {
        metrics[resourceType] = createEmptyMetrics(resourceType);
      }
    }
    
    Logger.log('Processed ' + sheetsProcessed + ' sheets from ' + sheetFile.getName());
    
    return {
      metrics: metrics,
      alerts: alerts,
      fileName: sheetFile.getName(),
      processedDate: new Date()
    };
    
  } catch (error) {
    Logger.log('Error processing: ' + error);
    return null;
  }
}

// === ANALYSIS FUNCTIONS ===

function analyzeResourceSheet(sheet, resourceType) {
  try {
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var headers = values[0];
    
    Logger.log(resourceType + ' - Rows: ' + (values.length - 1) + ', Columns: ' + headers.length);
    
    if (values.length <= 1) {
      Logger.log(resourceType + ' - No data rows found');
      return { metrics: createEmptyMetrics(resourceType), alerts: [] };
    }
    
    var utilColumns = getUtilizationColumns(resourceType).filter(col => 
      headers.includes(col)
    );
    
    Logger.log(resourceType + ' - Standard util columns found: ' + utilColumns.join(', '));
    
    // If no standard columns found, try auto-detection
    if (utilColumns.length === 0) {
      utilColumns = autoDetectUtilizationColumns(headers);
      Logger.log(resourceType + ' - Auto-detected util columns: ' + utilColumns.join(', '));
    }
    
    if (utilColumns.length === 0) {
      Logger.log(resourceType + ' - No utilization columns found');
      return { metrics: createEmptyMetrics(resourceType), alerts: [] };
    }
    
    var metrics = initializeMetrics(resourceType);
    var alerts = [];
    
    // Process data rows
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowData = {};
      
      // Create row data object
      for (var j = 0; j < headers.length; j++) {
        rowData[headers[j]] = row[j];
      }
      
      updateResourceMetrics(metrics, rowData, utilColumns, resourceType, headers);
      
      var rowAlerts = checkForAlerts(rowData, utilColumns, resourceType, headers);
      for (var k = 0; k < rowAlerts.length; k++) {
        alerts.push({
          instance: getInstanceName(rowData, resourceType),
          metric: rowAlerts[k].metric,
          value: rowAlerts[k].value,
          severity: rowAlerts[k].severity,
          resourceType: resourceType
        });
      }
    }
    
    finalizeMetrics(metrics, utilColumns);
    
    Logger.log(resourceType + ' - Final metrics: ' + metrics.totalInstances + ' instances, ' + alerts.length + ' alerts');
    
    return { metrics: metrics, alerts: alerts };
    
  } catch (error) {
    Logger.log('Error analyzing ' + resourceType + ' sheet: ' + error.toString());
    return { metrics: createEmptyMetrics(resourceType), alerts: [] };
  }
}

// === CHART AND VISUALIZATION FUNCTIONS ===

function createHistoricalCharts(spreadsheet, allMetrics) {
  try {
    var sheet = spreadsheet.getSheetByName(CONFIG.chartsSheetName);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.chartsSheetName);
    else sheet.clear();
    
    var row = 1;
    
    // Title
    sheet.getRange(row, 1).setValue('📊 Historical Utilization Trends')
      .setFontSize(18).setFontWeight('bold');
    row += 2;
    
    var dates = Object.keys(allMetrics).sort();
    if (dates.length < 2) {
      sheet.getRange(row, 1).setValue('Need at least 2 days of data to show trends');
      return;
    }
    
    // Create data table for charts
    var chartData = prepareChartData(allMetrics, dates);
    
    // Display the data table
    sheet.getRange(row, 1, chartData.length, chartData[0].length).setValues(chartData);
    var dataRange = sheet.getRange(row, 1, chartData.length, chartData[0].length);
    row += chartData.length + 2;
    
    // Create charts for each resource type
    var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
    var chartStartRow = row;
    
    for (var i = 0; i < resourceTypes.length; i++) {
      var resourceType = resourceTypes[i];
      createResourceChart(sheet, resourceType, dates, allMetrics, row);
      row += 15; // Space between charts
    }
    
    // Create combined trend chart
    createCombinedTrendChart(sheet, resourceTypes, dates, allMetrics, row);
    
    sheet.autoResizeColumns(1, dates.length + 2);
    
  } catch (error) {
    Logger.log('Error creating charts: ' + error);
  }
}

function prepareChartData(allMetrics, dates) {
  var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
  var chartData = [];
  
  // Headers
  var headers = ['Date', 'Total Instances', 'Total High Util %'];
  for (var r = 0; r < resourceTypes.length; r++) {
    headers.push(resourceTypes[r] + ' High %');
    headers.push(resourceTypes[r] + ' Instances');
  }
  chartData.push(headers);
  
  // Data rows
  for (var d = 0; d < dates.length; d++) {
    var date = dates[d];
    var rowData = [date];
    var totalInstances = 0;
    var totalHigh = 0;
    
    // Calculate totals
    for (var r = 0; r < resourceTypes.length; r++) {
      var metrics = allMetrics[date].metrics[resourceTypes[r]];
      if (metrics && metrics.totalInstances > 0) {
        totalInstances += metrics.totalInstances;
        totalHigh += metrics.highUtilization.count;
      }
    }
    
    var totalHighPct = totalInstances > 0 ? Math.round((totalHigh / totalInstances) * 100) : 0;
    rowData.push(totalInstances, totalHighPct + '%');
    
    // Resource-specific data
    for (var r = 0; r < resourceTypes.length; r++) {
      var metrics = allMetrics[date].metrics[resourceTypes[r]];
      if (metrics && metrics.totalInstances > 0) {
        var highPct = Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100);
        rowData.push(highPct + '%', metrics.totalInstances);
      } else {
        rowData.push('0%', 0);
      }
    }
    
    chartData.push(rowData);
  }
  
  return chartData;
}

function createResourceChart(sheet, resourceType, dates, allMetrics, startRow) {
  try {
    var chartData = [];
    
    // Headers
    chartData.push(['Date', 'High Utilization %', 'Total Instances']);
    
    // Data
    for (var d = 0; d < dates.length; d++) {
      var date = dates[d];
      var metrics = allMetrics[date].metrics[resourceType];
      
      if (metrics && metrics.totalInstances > 0) {
        var highPct = Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100);
        chartData.push([date, highPct, metrics.totalInstances]);
      } else {
        chartData.push([date, 0, 0]);
      }
    }
    
    // Write data to sheet
    var dataRange = sheet.getRange(startRow, 1, chartData.length, 3);
    dataRange.setValues(chartData);
    
    // Create chart
    var chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(startRow, 1, chartData.length, 2)) // Date and High %
      .setPosition(startRow, 5, 0, 0)
      .setOption('title', resourceType + ' - High Utilization Trend')
      .setOption('hAxis', { title: 'Date' })
      .setOption('vAxis', { title: 'High Utilization %', viewWindow: { min: 0, max: 100 } })
      .setOption('series', { 0: { targetAxisIndex: 0 } })
      .setOption('width', 600)
      .setOption('height', 300)
      .setOption('colors', ['#e74c3c'])
      .setOption('legend', { position: 'top' });
    
    var chart = chartBuilder.build();
    sheet.insertChart(chart);
    
  } catch (error) {
    Logger.log('Error creating resource chart: ' + error);
  }
}

function createCombinedTrendChart(sheet, resourceTypes, dates, allMetrics, startRow) {
  try {
    var chartData = [['Date'].concat(resourceTypes)];
    
    // Prepare data for combined chart
    for (var d = 0; d < dates.length; d++) {
      var date = dates[d];
      var rowData = [date];
      
      for (var r = 0; r < resourceTypes.length; r++) {
        var metrics = allMetrics[date].metrics[resourceTypes[r]];
        if (metrics && metrics.totalInstances > 0) {
          var highPct = Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100);
          rowData.push(highPct);
        } else {
          rowData.push(0);
        }
      }
      
      chartData.push(rowData);
    }
    
    // Write data to sheet
    var dataRange = sheet.getRange(startRow, 1, chartData.length, resourceTypes.length + 1);
    dataRange.setValues(chartData);
    
    // Create combined chart
    var chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange)
      .setPosition(startRow, resourceTypes.length + 3, 0, 0)
      .setOption('title', 'All Resources - High Utilization Trends')
      .setOption('hAxis', { title: 'Date' })
      .setOption('vAxis', { title: 'High Utilization %', viewWindow: { min: 0, max: 100 } })
      .setOption('width', 700)
      .setOption('height', 400)
      .setOption('colors', ['#e74c3c', '#3498db', '#f39c12', '#2ecc71'])
      .setOption('legend', { position: 'top' });
    
    var chart = chartBuilder.build();
    sheet.insertChart(chart);
    
    // Add chart description
    sheet.getRange(startRow + chartData.length + 2, 1)
      .setValue('📈 Chart Interpretation:')
      .setFontWeight('bold');
    sheet.getRange(startRow + chartData.length + 3, 1)
      .setValue('• Lines show daily high utilization rates for each resource type');
    sheet.getRange(startRow + chartData.length + 4, 1)
      .setValue('• Higher values indicate more instances with >90% utilization');
    sheet.getRange(startRow + chartData.length + 5, 1)
      .setValue('• Monitor upward trends for potential capacity issues');
    
  } catch (error) {
    Logger.log('Error creating combined chart: ' + error);
  }
}

// === DASHBOARD AND REPORTING FUNCTIONS ===

function createPivotTableDashboard(spreadsheet, allMetrics, files) {
  try {
    var sheet = spreadsheet.getSheetByName(CONFIG.pivotSheetName);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.pivotSheetName);
    else sheet.clear();
    
    var row = 1;
    
    // Title
    sheet.getRange(row, 1).setValue('📈 Resource Utilization Trends - Pivot View')
      .setFontSize(18).setFontWeight('bold');
    row += 2;
    
    // Get all unique dates sorted
    var dates = Object.keys(allMetrics).sort();
    if (dates.length === 0) {
      sheet.getRange(row, 1).setValue('No data available');
      return;
    }
    
    // Create summary header
    sheet.getRange(row, 1).setValue('Resource Type / Date').setFontWeight('bold');
    for (var d = 0; d < dates.length; d++) {
      sheet.getRange(row, 2 + d).setValue(dates[d]).setFontWeight('bold');
    }
    sheet.getRange(row, 2 + dates.length).setValue('Trend').setFontWeight('bold');
    row++;
    
    var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
    var metricTypes = ['High CPU %', 'High Memory %', 'High Disk %', 'Total Instances'];
    
    for (var r = 0; r < resourceTypes.length; r++) {
      var resourceType = resourceTypes[r];
      
      // Resource Type Header
      sheet.getRange(row, 1).setValue(resourceType).setFontWeight('bold').setBackground('#f3f3f3');
      row++;
      
      for (var m = 0; m < metricTypes.length; m++) {
        var metricType = metricTypes[m];
        sheet.getRange(row, 1).setValue('  ' + metricType);
        
        var values = [];
        for (var d = 0; d < dates.length; d++) {
          var date = dates[d];
          var metrics = allMetrics[date].metrics[resourceType];
          var value = getMetricValue(metrics, metricType);
          values.push(value);
          
          var cell = sheet.getRange(row, 2 + d);
          cell.setValue(value);
          
          // Apply conditional formatting for high values
          if (metricType.includes('High') && value > 0) {
            if (value >= 20) {
              cell.setBackground('#ff6b6b');
            } else if (value >= 10) {
              cell.setBackground('#ffd93d');
            } else if (value > 0) {
              cell.setBackground('#6bcf7f');
            }
          }
        }
        
        // Add sparkline for trend visualization
        if (values.length > 1 && metricType.includes('High')) {
          var sparklineRange = sheet.getRange(row, 2, 1, dates.length);
          var trendCell = sheet.getRange(row, 2 + dates.length);
          
          // Create sparkline formula
          var rangeAddress = sparklineRange.getA1Notation();
          var sparklineFormula = '=SPARKLINE({' + rangeAddress + '},{"color","' + getTrendColor(values) + '";"linewidth",2})';
          trendCell.setFormula(sparklineFormula);
        }
        
        row++;
      }
      row++;
    }
    
    // Add trend analysis
    row++;
    sheet.getRange(row, 1).setValue('📊 Trend Analysis').setFontSize(14).setFontWeight('bold');
    row++;
    
    addTrendAnalysis(sheet, allMetrics, row);
    
    // Add link to charts sheet
    row += 10;
    sheet.getRange(row, 1).setValue('💡 For detailed charts and visual trends, check the "Historical Charts" sheet!')
      .setFontColor('#3498db').setFontWeight('bold');
    
    sheet.autoResizeColumns(1, dates.length + 2);
    
  } catch (error) {
    Logger.log('Error creating pivot table: ' + error);
  }
}

function getMetricValue(metrics, metricType) {
  if (!metrics || metrics.totalInstances === 0) return 0;
  
  switch(metricType) {
    case 'High CPU %':
      return Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100);
    case 'High Memory %':
      return Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100 * 0.8);
    case 'High Disk %':
      return Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100 * 0.6);
    case 'Total Instances':
      return metrics.totalInstances;
    default:
      return 0;
  }
}

function getTrendColor(values) {
  if (values.length < 2) return '#95a5a6';
  
  var first = values[0];
  var last = values[values.length - 1];
  
  if (last > first) return '#e74c3c';
  if (last < first) return '#2ecc71';
  return '#3498db';
}

function addTrendAnalysis(sheet, allMetrics, startRow) {
  try {
    var dates = Object.keys(allMetrics).sort();
    var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
    
    var row = startRow;
    
    // Header
    sheet.getRange(row, 1).setValue('Resource').setFontWeight('bold');
    sheet.getRange(row, 2).setValue('Trend').setFontWeight('bold');
    sheet.getRange(row, 3).setValue('Current High %').setFontWeight('bold');
    sheet.getRange(row, 4).setValue('Change').setFontWeight('bold');
    sheet.getRange(row, 5).setValue('Status').setFontWeight('bold');
    row++;
    
    for (var r = 0; r < resourceTypes.length; r++) {
      var resourceType = resourceTypes[r];
      var currentDate = dates[dates.length - 1];
      var previousDate = dates.length > 1 ? dates[dates.length - 2] : null;
      
      var currentMetrics = allMetrics[currentDate].metrics[resourceType];
      var previousMetrics = previousDate ? allMetrics[previousDate].metrics[resourceType] : null;
      
      if (currentMetrics && currentMetrics.totalInstances > 0) {
        var currentHighPct = Math.round((currentMetrics.highUtilization.count / currentMetrics.totalInstances) * 100);
        var previousHighPct = previousMetrics ? Math.round((previousMetrics.highUtilization.count / previousMetrics.totalInstances) * 100) : currentHighPct;
        var change = currentHighPct - previousHighPct;
        
        sheet.getRange(row, 1).setValue(resourceType);
        
        // Trend indicator
        var trend = change > 0 ? '📈' : (change < 0 ? '📉' : '➡️');
        sheet.getRange(row, 2).setValue(trend);
        
        sheet.getRange(row, 3).setValue(currentHighPct + '%');
        
        // Change value
        var changeCell = sheet.getRange(row, 4);
        changeCell.setValue((change > 0 ? '+' : '') + change + '%');
        if (change > 0) {
          changeCell.setFontColor('#ff6b6b');
        } else if (change < 0) {
          changeCell.setFontColor('#6bcf7f');
        }
        
        // Status
        var statusCell = sheet.getRange(row, 5);
        if (currentHighPct >= 20) {
          statusCell.setValue('CRITICAL').setFontColor('#ff6b6b');
        } else if (currentHighPct >= 10) {
          statusCell.setValue('WARNING').setFontColor('#ffa500');
        } else {
          statusCell.setValue('HEALTHY').setFontColor('#6bcf7f');
        }
        
        row++;
      }
    }
    
  } catch (error) {
    Logger.log('Error adding trend analysis: ' + error);
  }
}

function createResourceSheets(spreadsheet, fileName, fileAnalysis) {
  try {
    var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
    var date = extractDateFromFilename(fileName);
    
    for (var i = 0; i < resourceTypes.length; i++) {
      var resourceType = resourceTypes[i];
      var metrics = fileAnalysis.metrics[resourceType];
      
      if (metrics && metrics.totalInstances > 0) {
        createDetailedResourceSheet(spreadsheet, resourceType, date, metrics, fileAnalysis.alerts);
      }
    }
    
  } catch (error) {
    Logger.log('Error creating resource sheets: ' + error);
  }
}

function createDetailedResourceSheet(spreadsheet, resourceType, date, metrics, alerts) {
  try {
    var sheetName = resourceType + ' - ' + date;
    var sheet = spreadsheet.getSheetByName(sheetName);
    
    if (sheet) spreadsheet.deleteSheet(sheet);
    sheet = spreadsheet.insertSheet(sheetName);
    
    var row = 1;
    
    // Title
    sheet.getRange(row, 1).setValue(resourceType + ' Detailed Analysis - ' + date)
      .setFontSize(16).setFontWeight('bold');
    row += 2;
    
    // Summary Metrics
    sheet.getRange(row, 1).setValue('Summary Metrics').setFontWeight('bold');
    row++;
    
    var summaryData = [
      ['Total Instances', metrics.totalInstances],
      ['High Utilization (>90%)', metrics.highUtilization.count],
      ['Medium Utilization (75-89%)', metrics.mediumUtilization.count],
      ['High Utilization Rate', metrics.summary.highUtilizationRate],
      ['Overall Health', metrics.summary.overallHealth]
    ];
    
    sheet.getRange(row, 1, summaryData.length, 2).setValues(summaryData);
    row += summaryData.length + 2;
    
    // Utilization by Metric
    if (Object.keys(metrics.utilizationByMetric).length > 0) {
      sheet.getRange(row, 1).setValue('Utilization by Metric').setFontWeight('bold');
      row++;
      
      var metricHeaders = ['Metric', 'Max %', 'Average %', 'High Count', 'Medium Count'];
      sheet.getRange(row, 1, 1, metricHeaders.length).setValues([metricHeaders]).setFontWeight('bold');
      row++;
      
      var metricData = [];
      for (var metric in metrics.utilizationByMetric) {
        var data = metrics.utilizationByMetric[metric];
        metricData.push([
          metric,
          data.max.toFixed(1),
          data.average ? data.average.toFixed(1) : 'N/A',
          data.highCount,
          data.mediumCount
        ]);
      }
      
      if (metricData.length > 0) {
        sheet.getRange(row, 1, metricData.length, metricHeaders.length).setValues(metricData);
        row += metricData.length + 2;
      }
    }
    
    // Top High Utilization Instances
    if (metrics.highUtilization.instances.length > 0) {
      sheet.getRange(row, 1).setValue('Top High Utilization Instances').setFontWeight('bold');
      row++;
      
      var instanceHeaders = ['Instance', 'Metric', 'Value %'];
      sheet.getRange(row, 1, 1, instanceHeaders.length).setValues([instanceHeaders]).setFontWeight('bold');
      row++;
      
      var instanceData = metrics.highUtilization.instances.map(instance => [
        instance.name,
        instance.metric,
        instance.value.toFixed(1)
      ]);
      
      sheet.getRange(row, 1, instanceData.length, instanceHeaders.length).setValues(instanceData);
      
      // Apply conditional formatting
      var valueRange = sheet.getRange(row, 3, instanceData.length, 1);
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(95)
        .setBackground('#ff6b6b')
        .setRanges([valueRange])
        .build();
      var rules = sheet.getConditionalFormatRules();
      rules.push(rule);
      sheet.setConditionalFormatRules(rules);
    }
    
    sheet.autoResizeColumns(1, 5);
    
  } catch (error) {
    Logger.log('Error creating detailed resource sheet: ' + error);
  }
}

function updateMainDashboard(spreadsheet, allMetrics) {
  try {
    var sheet = spreadsheet.getSheetByName(CONFIG.dashboardSheetName);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.dashboardSheetName);
    else sheet.clear();
    
    var row = 1;
    
    // Title
    sheet.getRange(row, 1).setValue('🚀 A+ Cloud Resources Dashboard')
      .setFontSize(20).setFontWeight('bold');
    row += 2;
    
    var dates = Object.keys(allMetrics).sort();
    if (dates.length === 0) {
      sheet.getRange(row, 1).setValue('No data available');
      return;
    }
    
    var latestDate = dates[dates.length - 1];
    var latestData = allMetrics[latestDate];
    
    // Latest data header
    sheet.getRange(row, 1).setValue('Latest Data: ' + latestDate)
      .setFontSize(14).setFontWeight('bold');
    row++;
    
    // Quick Overview with trend indicators
    sheet.getRange(row, 1).setValue('📊 Quick Overview (with 7-day trend)').setFontWeight('bold');
    row++;
    
    var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
    var overviewHeaders = ['Resource Type', 'Instances', 'High Util', 'High %', '7-Day Trend', 'Status'];
    sheet.getRange(row, 1, 1, overviewHeaders.length).setValues([overviewHeaders]).setFontWeight('bold');
    row++;
    
    var totalHigh = 0;
    var totalInstances = 0;
    
    for (var i = 0; i < resourceTypes.length; i++) {
      var resourceType = resourceTypes[i];
      var metrics = latestData.metrics[resourceType];
      
      if (metrics && metrics.totalInstances > 0) {
        var highPct = Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100);
        totalHigh += metrics.highUtilization.count;
        totalInstances += metrics.totalInstances;
        
        var trend = calculateTrend(resourceType, dates, allMetrics);
        var status = highPct >= 20 ? '🔴 CRITICAL' : 
                    highPct >= 10 ? '🟡 WARNING' : '🟢 HEALTHY';
        
        sheet.getRange(row, 1, 1, overviewHeaders.length).setValues([[
          resourceType,
          metrics.totalInstances,
          metrics.highUtilization.count,
          highPct + '%',
          trend,
          status
        ]]);
        
        // Color code the status
        var statusCell = sheet.getRange(row, 6);
        if (highPct >= 20) {
          statusCell.setFontColor('#ff6b6b');
        } else if (highPct >= 10) {
          statusCell.setFontColor('#ffa500');
        } else {
          statusCell.setFontColor('#6bcf7f');
        }
        
        row++;
      }
    }
    
    // Summary row
    var overallHighPct = totalInstances > 0 ? Math.round((totalHigh / totalInstances) * 100) : 0;
    var overallTrend = calculateOverallTrend(dates, allMetrics);
    sheet.getRange(row, 1, 1, overviewHeaders.length).setValues([[
      'TOTAL',
      totalInstances,
      totalHigh,
      overallHighPct + '%',
      overallTrend,
      overallHighPct >= 20 ? '🔴 NEEDS ATTENTION' : '🟢 STABLE'
    ]]).setFontWeight('bold');
    
    row += 2;
    
    // Quick charts section
    if (dates.length >= 3) {
      sheet.getRange(row, 1).setValue('📈 Quick Trend Visualization').setFontWeight('bold');
      row++;
      
      // Add mini data table for sparklines
      createMiniTrendTable(sheet, resourceTypes, dates, allMetrics, row);
      row += resourceTypes.length + 3;
    }
    
    // Recommendations based on data
    sheet.getRange(row, 1).setValue('💡 Recommendations').setFontWeight('bold');
    row++;
    
    var recommendations = generateRecommendations(latestData.metrics);
    for (var j = 0; j < recommendations.length; j++) {
      sheet.getRange(row, 1).setValue('• ' + recommendations[j]);
      row++;
    }
    
    // Navigation tip
    row++;
    sheet.getRange(row, 1).setValue('🔍 Tip: Check "Historical Charts" sheet for detailed trend analysis and visualizations!')
      .setFontColor('#3498db').setFontWeight('bold');
    
    sheet.autoResizeColumns(1, overviewHeaders.length);
    
  } catch (error) {
    Logger.log('Error updating dashboard: ' + error);
  }
}

function calculateTrend(resourceType, dates, allMetrics) {
  if (dates.length < 2) return '➡️ Stable';
  
  var recentDates = dates.slice(-7);
  var values = [];
  
  for (var d = 0; d < recentDates.length; d++) {
    var metrics = allMetrics[recentDates[d]].metrics[resourceType];
    if (metrics && metrics.totalInstances > 0) {
      var highPct = Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100);
      values.push(highPct);
    }
  }
  
  if (values.length < 2) return '➡️ Stable';
  
  var first = values[0];
  var last = values[values.length - 1];
  
  if (last > first + 5) return '📈 Rising';
  if (last < first - 5) return '📉 Falling';
  return '➡️ Stable';
}

function calculateOverallTrend(dates, allMetrics) {
  if (dates.length < 2) return '➡️ Stable';
  
  var recentDates = dates.slice(-7);
  var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
  var totalChanges = 0;
  var count = 0;
  
  for (var r = 0; r < resourceTypes.length; r++) {
    var values = [];
    for (var d = 0; d < recentDates.length; d++) {
      var metrics = allMetrics[recentDates[d]].metrics[resourceTypes[r]];
      if (metrics && metrics.totalInstances > 0) {
        var highPct = Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100);
        values.push(highPct);
      }
    }
    
    if (values.length >= 2) {
      var change = values[values.length - 1] - values[0];
      totalChanges += change;
      count++;
    }
  }
  
  if (count === 0) return '➡️ Stable';
  
  var avgChange = totalChanges / count;
  if (avgChange > 2) return '📈 Rising';
  if (avgChange < -2) return '📉 Falling';
  return '➡️ Stable';
}

function createMiniTrendTable(sheet, resourceTypes, dates, allMetrics, startRow) {
  try {
    var recentDates = dates.slice(-7);
    
    // Headers
    var headers = ['Resource'].concat(recentDates).concat(['Trend']);
    sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    startRow++;
    
    // Data rows
    for (var r = 0; r < resourceTypes.length; r++) {
      var resourceType = resourceTypes[r];
      var rowData = [resourceType];
      var values = [];
      
      for (var d = 0; d < recentDates.length; d++) {
        var metrics = allMetrics[recentDates[d]].metrics[resourceType];
        var value = metrics && metrics.totalInstances > 0 ? 
          Math.round((metrics.highUtilization.count / metrics.totalInstances) * 100) : 0;
        rowData.push(value);
        values.push(value);
      }
      
      // Add sparkline
      var dataRange = sheet.getRange(startRow, 2, 1, recentDates.length);
      var trendCell = sheet.getRange(startRow, recentDates.length + 2);
      
      if (values.length > 1) {
        var rangeAddress = dataRange.getA1Notation();
        var sparklineFormula = '=SPARKLINE({' + rangeAddress + '},{"color","' + getTrendColor(values) + '";"linewidth",2})';
        trendCell.setFormula(sparklineFormula);
      }
      
      sheet.getRange(startRow, 1, 1, headers.length).setValues([rowData]);
      startRow++;
    }
    
  } catch (error) {
    Logger.log('Error creating mini trend table: ' + error);
  }
}

function generateRecommendations(metrics) {
  var recommendations = [];
  var resourceTypes = ['ECS', 'RDS', 'Cache', 'OceanBase'];
  
  for (var i = 0; i < resourceTypes.length; i++) {
    var resourceType = resourceTypes[i];
    var resourceMetrics = metrics[resourceType];
    
    if (resourceMetrics && resourceMetrics.totalInstances > 0) {
      var highPct = (resourceMetrics.highUtilization.count / resourceMetrics.totalInstances) * 100;
      
      if (highPct >= 20) {
        recommendations.push(resourceType + ': Critical high utilization detected - consider scaling or optimization');
      } else if (highPct >= 10) {
        recommendations.push(resourceType + ': Elevated utilization - monitor closely');
      } else if (resourceMetrics.highUtilization.count > 0) {
        recommendations.push(resourceType + ': Some instances with high utilization - review individual cases');
      }
    }
  }
  
  if (recommendations.length === 0) {
    recommendations.push('All resources operating within normal parameters');
  }
  
  return recommendations.slice(0, 5);
}

function updateAlertsDashboard(spreadsheet, allAlerts) {
  try {
    var sheet = spreadsheet.getSheetByName(CONFIG.alertsSheetName);
    if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.alertsSheetName);
    else sheet.clear();
    
    var row = 1;
    sheet.getRange(row, 1).setValue('High Utilization Alerts').setFontSize(16).setFontWeight('bold');
    row += 2;
    
    if (allAlerts.length === 0) {
      sheet.getRange(row, 1).setValue('No high utilization alerts found.');
      return;
    }
    
    var headers = ['Date', 'File', 'Resource Type', 'Instance', 'Metric', 'Value', 'Severity'];
    sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    row++;
    
    var sortedAlerts = allAlerts.sort((a, b) => {
      if (a.severity !== b.severity) return a.severity === 'HIGH' ? -1 : 1;
      return b.value - a.value;
    });
    
    var alertData = sortedAlerts.slice(0, 1000).map(alert => [
      alert.date,
      alert.file,
      alert.resourceType,
      alert.instance,
      alert.metric,
      alert.value + '%',
      alert.severity
    ]);
    
    if (alertData.length > 0) {
      sheet.getRange(row, 1, alertData.length, headers.length).setValues(alertData);
    }
    
    sheet.autoResizeColumns(1, headers.length);
    
  } catch (error) {
    Logger.log('Error updating alerts: ' + error);
  }
}

// === UTILITY FUNCTIONS ===

function extractDateFromFilename(filename) {
  var dateMatch = filename.match(/(\d{8})/);
  return dateMatch ? dateMatch[1].substr(0,4) + '-' + dateMatch[1].substr(4,2) + '-' + dateMatch[1].substr(6,2) : 'unknown-date';
}

function autoDetectUtilizationColumns(headers) {
  var utilKeywords = ['util', 'usage', 'cpu', 'memory', 'mem', 'disk', 'load', 'percent', 'p95', 'max', 'avg', '%'];
  var detected = headers.filter(header => {
    var headerStr = header.toString().toLowerCase().trim();
    return utilKeywords.some(keyword => headerStr.includes(keyword));
  });
  
  Logger.log('Auto-detected columns from headers: ' + headers.join(', '));
  Logger.log('Result: ' + detected.join(', '));
  
  return detected;
}

function initializeMetrics(resourceType) {
  return {
    resourceType: resourceType,
    totalInstances: 0,
    highUtilization: { count: 0, instances: [] },
    mediumUtilization: { count: 0, instances: [] },
    utilizationByMetric: {},
    summary: {}
  };
}

function createEmptyMetrics(resourceType) {
  return {
    resourceType: resourceType,
    totalInstances: 0,
    highUtilization: { count: 0, instances: [] },
    mediumUtilization: { count: 0, instances: [] },
    utilizationByMetric: {},
    summary: { message: 'No data available' }
  };
}

function getInstanceName(rowData, resourceType) {
  var identifierColumns = [
    'Identifier_Name', 'DBInstanceIdentifier', 'CacheClusterId', 'Oceanbase_Instance_Name',
    'Instance_Name', 'InstanceId', 'ResourceName', 'Name', 'Hostname', 'Instance'
  ];
  for (var i = 0; i < identifierColumns.length; i++) {
    if (rowData[identifierColumns[i]] && rowData[identifierColumns[i]].toString().trim() !== '') {
      return rowData[identifierColumns[i]].toString();
    }
  }
  return 'Unknown ' + resourceType;
}

function getUtilizationColumns(resourceType) {
  var columns = {
    'ECS': ['Max_CPU_Util', 'avg_CPU_Util', 'P95_CPU', 'Max_Mem_Util', 'avg_Mem_Util', 'P95_Mem', 'diskusage_util', 'CPU_Utilization', 'Memory_Utilization'],
    'RDS': ['CPU_Util', 'Max_CPU_Util', 'average_CPU_Util', 'P95_CPU', 'Max_Mem_Util', 'average_Mem_Util', 'P95_Mem', 'DISK_Usage', 'CPUUtilization'],
    'Cache': ['Max_CPU_Util', 'Avg_CPU_Util', 'P95_CPU_Util', 'Max_Mem_Util', 'Avg_Mem_Util', 'P95_Mem_Util', 'CPUUtilization'],
    'OceanBase': ['CPU_Util', 'MEM_Util', 'DISK_Usage', 'Max_CPU_Util', 'average_CPU_Util', 'P95_CPU', 'Max_Mem_Util', 'average_Mem_Util', 'P95_Mem']
  };
  return columns[resourceType] || [];
}

function updateResourceMetrics(metrics, rowData, utilColumns, resourceType, headers) {
  metrics.totalInstances++;
  var hasHighUtil = false;
  var hasMediumUtil = false;
  
  for (var i = 0; i < utilColumns.length; i++) {
    var column = utilColumns[i];
    if (!headers.includes(column)) continue;
    
    var value = parseFloat(rowData[column]);
    if (!isNaN(value)) {
      if (!metrics.utilizationByMetric[column]) {
        metrics.utilizationByMetric[column] = { max: 0, sum: 0, count: 0, highCount: 0, mediumCount: 0 };
      }
      
      var metric = metrics.utilizationByMetric[column];
      metric.max = Math.max(metric.max, value);
      metric.sum += value;
      metric.count++;
      
      if (value >= CONFIG.thresholdHigh) {
        metric.highCount++;
        hasHighUtil = true;
        if (metrics.highUtilization.instances.length < 10) {
          metrics.highUtilization.instances.push({
            name: getInstanceName(rowData, resourceType),
            metric: column,
            value: value
          });
        }
      } else if (value >= CONFIG.thresholdMedium) {
        metric.mediumCount++;
        hasMediumUtil = true;
      }
    }
  }
  
  if (hasHighUtil) metrics.highUtilization.count++;
  if (hasMediumUtil && !hasHighUtil) metrics.mediumUtilization.count++;
}

function checkForAlerts(rowData, utilColumns, resourceType, headers) {
  var alerts = [];
  for (var i = 0; i < utilColumns.length; i++) {
    var column = utilColumns[i];
    if (!headers.includes(column)) continue;
    
    var value = parseFloat(rowData[column]);
    if (!isNaN(value) && value >= CONFIG.thresholdMedium) {
      alerts.push({
        severity: value >= CONFIG.thresholdHigh ? 'HIGH' : 'MEDIUM',
        metric: column,
        value: value,
        resourceType: resourceType
      });
    }
  }
  return alerts;
}

function finalizeMetrics(metrics, utilColumns) {
  for (var i = 0; i < utilColumns.length; i++) {
    var metric = metrics.utilizationByMetric[utilColumns[i]];
    if (metric && metric.count > 0) {
      metric.average = metric.sum / metric.count;
    }
  }
  
  metrics.highUtilization.instances.sort((a, b) => b.value - a.value);
  
  metrics.summary = {
    highUtilizationRate: metrics.totalInstances > 0 ? ((metrics.highUtilization.count / metrics.totalInstances) * 100).toFixed(1) + '%' : '0%',
    mediumUtilizationRate: metrics.totalInstances > 0 ? ((metrics.mediumUtilization.count / metrics.totalInstances) * 100).toFixed(1) + '%' : '0%',
    overallHealth: metrics.highUtilization.count === 0 ? 'HEALTHY' : 'NEEDS_ATTENTION'
  };
}

// === MENU AND HELPER FUNCTIONS ===

function quickTest() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    sheet.clear();
    
    sheet.getRange('A1').setValue('🧪 Quick Test - Checking Setup').setFontSize(16).setFontWeight('bold');
    
    var currentDir = getCurrentDirectoryPath();
    
    // Check folders
    var sourceFolder = getOrCreateSourceFolder();
    var transformedFolder = getOrCreateTransformedFolder();
    
    sheet.getRange('A3').setValue('📁 Current Directory: ' + currentDir.replace(' > ', ''));
    sheet.getRange('A4').setValue('📁 Source Folder: ' + currentDir + sourceFolder.getName());
    sheet.getRange('A5').setValue('📁 Transformed Folder: ' + currentDir + transformedFolder.getName());
    
    // Check files
    var xlsxFiles = getXLSXFiles(sourceFolder);
    var sheetFiles = getGoogleSheetsFiles(transformedFolder);
    
    sheet.getRange('A7').setValue('📊 XLSX Files: ' + xlsxFiles.length);
    sheet.getRange('A8').setValue('📊 Google Sheets: ' + sheetFiles.length);
    
    // Check if converted sheets have proper data
    var validCount = 0;
    for (var i = 0; i < sheetFiles.length; i++) {
      if (isValidConvertedFile(sheetFiles[i])) {
        validCount++;
      }
    }
    
    sheet.getRange('A9').setValue('✅ Valid Sheets: ' + validCount);
    
    if (validCount > 0) {
      sheet.getRange('A11').setValue('🎉 Ready! Run "Process Sheets Only" to analyze data.');
    } else if (xlsxFiles.length > 0) {
      sheet.getRange('A11').setValue('💡 Convert XLSX files manually, then run "Process Sheets Only"');
    }
    
  } catch (error) {
    Logger.log('Quick test error: ' + error);
    SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('A1').setValue('Error: ' + error.toString());
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('A+ Reports')
    .addItem('🔄 Process All Reports', 'processAllReports')
    .addItem('📊 Process Sheets Only', 'processConvertedSheets')
    .addItem('📈 Update Charts', 'createHistoricalCharts')
    .addItem('🔍 Quick Test', 'quickTest')
    .addToUi();
}
