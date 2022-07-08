function main(workbook: ExcelScript.Workbook) {
    // Your code here
    //let moduleId = 'CMP6002B';
    //let moduleId = 'CMP6006A';


    //moduleId = 'CMP6006A';
    // surveyTable = 2 or surveyTable = 3
    //let surveyTable = 1

    //console.log(moduleId);
    //console.log('\n\n');
    //console.log(html);

    //!!!!!!!!!!!!!
    //Remove: // Below

    //Module Specification Data
    let dataSheet = workbook.getWorksheet("Sheet1");
    let dataTable = dataSheet.getTables()[0].getRangeBetweenHeaderAndTotal().getValues();
    let dataHeaders = dataSheet.getTables()[0].getHeaderRowRange().getValues();

    let responses = dataTable.length;
    console.log(`responses = ${responses}`)

    workbook.addWorksheet("Report");
    let reportSheet = workbook.getWorksheet("Report");
    reportSheet.addTable('A1', true)
    let reportTable = reportSheet.getTables()[0];

    reportSheet.addTable('C1', true)
    let reportCommentsTable = reportSheet.getTables()[1];

    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic)

    //workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);

    



    console.log(dataHeaders[0].length);

    let sum = 0;
    let averageResult = 0;
    let medianResult = 0;

  //=AVERAGE(Table1[1. I found the module to be intellectually  stimulating])

  //reportSheet.getRange("C3").setValue("=6*2");
  //reportSheet.getRange("C4").setValue("=42/3");

  reportTable.addColumn(-1, null, 'Average');
  reportTable.addColumn(-1, null, 'Median');
  reportTable.addColumn(-1, null, 'Mode');
  reportTable.addColumn(-1, null, 'Strongly Disagree Count');
  reportTable.addColumn(-1, null, 'Disagree Count');
  reportTable.addColumn(-1, null, 'Neither Count');
  reportTable.addColumn(-1, null, 'Agree Count');
  reportTable.addColumn(-1, null, 'Strongly Agree Count');
  reportTable.getColumnByName('Column1').setName('Question');

  reportCommentsTable.getColumnByName('Column1').setName('Question');
  for (let k = 0; k < responses;k++)
  {
    reportCommentsTable.addColumn(-1, null, `Response${k+1}`);
  }

    let cell_address = '';
    let range_address = '';
    let questionResults:Array<number> = [];
    //Skip the time by starting at index 1.
    let commentCounter = 0;
    for (let i = 1; i < dataHeaders[0].length;i++)
    {
        //='Sheet1'!D24
      cell_address = '';
      cell_address = dataSheet.getCell(1, i).getAddress().toString();

      //console.log(cell_address);
      cell_address = cell_address.substring(7, cell_address.length - 1);
      range_address = cell_address + ':' + cell_address.substring(0, cell_address.length);

      //console.log(cell_address);

      //Sheet1!B2
      
      if (dataHeaders[0][i][dataHeaders[0][i].toString().length - 1] != '*')
      {
        reportTable.addRow(-1, 
          [
            dataHeaders[0][i],
            `=IFERROR(AVERAGE(Sheet1!${range_address}), -1)`,
            `=IFERROR(MEDIAN(Sheet1!${range_address}), -1)`,
            `=IFNA(MODE.SNGL(Sheet1!${range_address}),-1)`,
            
            `=COUNTIF(Sheet1!${range_address}, 1)`,
            `=COUNTIF(Sheet1!${range_address}, 2)`,
            `=COUNTIF(Sheet1!${range_address}, 3)`,
            `=COUNTIF(Sheet1!${range_address}, 4)`,
            `=COUNTIF(Sheet1!${range_address}, 5)`,
          ]);

          //Double check 5 is strongly agree

        reportTable.addRow(-1,
          [
            '', '', '', '', '', '', '', '', ''
          ]);
      }
      else
      {
        // reportTable.addRow(-1,
        //   [
        //     dataHeaders[0][i], ' ', '', '', '', '', '', '', ''
        //   ]);

        // reportTable.addRow(-1,
        //   [
        //     '', '', '', '', '', '', '', '', ''
        //   ]);



        let newReportCommentsTableRow:Array<string> = [];
        let emptyReportCommentsTableRow: Array<string> = [];

        newReportCommentsTableRow.push(dataHeaders[0][i].toString());
        emptyReportCommentsTableRow.push('');
        for (let x = 2; x < 2 + responses; x++) {
          newReportCommentsTableRow.push(`=Sheet1!${cell_address}${x}`);
          emptyReportCommentsTableRow.push('');
        }

        console.log(newReportCommentsTableRow);
        reportCommentsTable.addRow(-1, newReportCommentsTableRow);

        reportCommentsTable.addRow(-1, emptyReportCommentsTableRow);

        commentCounter++;
      }
    }
// - commentCounter
  for (let j = ((dataHeaders[0].length - commentCounter) * 2) - 3; j > 0;j-=2)
   {
     //console.log(j)
     reportTable.deleteRowsAt(j, 1);
   }

  for (let j = ((commentCounter) * 2) - 1; j > 0; j -= 2) {
    console.log(j)
    reportCommentsTable.deleteRowsAt(j, 1);
  }
}