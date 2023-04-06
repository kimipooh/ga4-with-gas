let propertyId_array = new Array();
let startDate_array = new Array();
let endDate_array = new Array();
let Name_array = new Array();

const Config_sheet = 'Configuration';

const Dimensions = [
  {name: 'country'},
//  {name:'eventName'},
];
const Metrics = [
  {name: 'activeUsers'},
  {name: 'screenPageViews'},
//  {name:'eventCount'},
];
/*
const Dimension_filters = {
  filter: {
    fieldName: 'eventName',
    stringFilter: {
      value: 'page_view'
    }
  }
}
*/
function runReport() {
  get_GA4_config();

  for(let i = 0; i < propertyId_array.length; i++){
    get_GA4(Name_array[i], propertyId_array[i], startDate_array[i], endDate_array[i], Metrics);
  }
}

function get_GA4(name, propertyId, startDate, endDate) {
  /**
   * TODO(developer): Uncomment this variable and replace with your
   *   Google Analytics 4 property ID before running the sample.
   */
    try {

    const dateRange = AnalyticsData.newDateRange();
    dateRange.startDate = startDate;
    dateRange.endDate = endDate;

    const request = AnalyticsData.newRunReportRequest();
    request.metrics = Metrics;
    request.dimensions = Dimensions;
    request.dateRanges = dateRange;
    if(typeof Dimension_filters !== 'undefined'){
      request.dimensionFilter = Dimension_filters;
    }
    const report = AnalyticsData.Properties.runReport(request,
        'properties/' + propertyId);
    if (!report.rows) {
      Logger.log('No rows returned.');
      return;
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(name);
    if(!sheet){
      let ss = spreadsheet.insertSheet();
      ss.setName(name);
      sheet = spreadsheet.getSheetByName(name);
    }else{
      sheet.clear();
    }

    // Append the headers.
    const dimensionHeaders = report.dimensionHeaders.map(
        (dimensionHeader) => {
          return dimensionHeader.name;
        });
    const metricHeaders = report.metricHeaders.map(
        (metricHeader) => {
          return metricHeader.name;
        });
    const headers = [...dimensionHeaders, ...metricHeaders];

    sheet.appendRow(headers);

    // Append the results.
    const rows = report.rows.map((row) => {
      const dimensionValues = row.dimensionValues.map(
          (dimensionValue) => {
            return dimensionValue.value;
          });
      const metricValues = row.metricValues.map(
          (metricValues) => {
            return metricValues.value;
          });
      return [...dimensionValues, ...metricValues];
    });

    sheet.getRange(2, 1, report.rows.length, headers.length)
        .setValues(rows);

    Logger.log('Done: %s', name);
  } catch (e) {
    // TODO (Developer) - Handle exception
    Logger.log('Failed with error: %s', e.error);
  }
}

function get_GA4_config() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(Config_sheet);
    if(!sheet){
      Logger.log('Cannot load Configuration sheet[%s].', Config_sheet);
      return;
    }

    const last_colum = sheet.getLastColumn();
    const data = sheet.getRange(2, 2, 5, last_colum).getValues();

    for(let cel=0; cel < last_colum-1; cel++){
      //Logger.log("%s %s", cel, last_colum);
      if(!data[0][cel]) continue;
      Name_array.push(data[0][cel].toString());
      propertyId_array.push(data[1][cel].toString());
      let sdate = Utilities.formatDate(new Date(data[2][cel]), 'JST', 'yyyy-MM-dd');
      startDate_array.push(sdate);
      let edate = Utilities.formatDate(new Date(data[3][cel]), 'JST', 'yyyy-MM-dd');
      endDate_array.push(edate);

      Logger.log('Done: %s %s %s %s', data[0][cel].toString(), data[1][cel].toString(),sdate,edate);
    }
  } catch (e) {
    // TODO (Developer) - Handle exception
    Logger.log('Failed with error: %s', e.error);
  }
}
