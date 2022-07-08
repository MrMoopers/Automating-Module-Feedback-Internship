function main(workbook: ExcelScript.Workbook, questionsText: Array<string> ) {
    // Your code here
    //let moduleId = 'CMP6002B';
    //let moduleId = 'CMP6006A';
  
  
    //questionsText = ['How much cheese do you like?', 'how good was the module?', 'comments?', 'How helpful was the lecturer?', 'comments?'];
  
    //moduleId = 'CMP6006A';
    // surveyTable = 2 or surveyTable = 3
    //let surveyTable = 1
  
    //console.log(moduleId);
    //console.log('\n\n');
    //console.log(html);
  
    //Module Specification Data
  
  
    //let sheet = workbook.getWorksheet("Sheet1");
    let sheet = workbook.getActiveWorksheet();
    sheet.addTable('A1', true);
  
  
    let table = sheet.getTables()[0];
    //let tableHeaders = sheet.getTables()[0].getHeaderRowRange().getValues();
  
    for (let i = 0; i < questionsText.length; i++) {
      table.addColumn(null, null, `${(i + 1)}. ${questionsText[i]}`);
  
    }
  
    table.getColumn('Column1').setName('When Completed');
  
  
  
  
    ////console.log(dataHeaders);
  
    //let moduleRow = -1;
    //do {
    //  moduleRow++;
    //} while (moduleRow < data.length && data[moduleRow][0] != moduleId)
  
    //let moduleName = data[moduleRow][1];
    //let moduleOrganiser = data[moduleRow][2];
  
    ////Survey Specification Data
    //let sheet_survey = workbook.getWorksheet("Sheet1");
    //let data_survey = sheet_survey.getTables()[surveyType].getRangeBetweenHeaderAndTotal().getValues();
    //let dataHeaders_survey = sheet_survey.getTables()[surveyType].getHeaderRowRange().getValues();
    ////console.log(dataHeaders_survey);
  
    //let html = '';
  }