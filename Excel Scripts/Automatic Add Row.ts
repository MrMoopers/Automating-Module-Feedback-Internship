function main(workbook: ExcelScript.Workbook, results: object ) {
    // Your code here
    //let moduleId = 'CMP6002B';
    //let moduleId = 'CMP6006A';


  /**results = [
    {
      "key": "ModuleId",
      "value": "CMP6002B"
    },
    {
      "key": "question1",
      "value": "1"
    },
    {
      "key": "question2",
      "value": "4"
    },
    {
      "key": "question3",
      "value": "3"
    },
    {
      "key": "question4",
      "value": "4"
    },
    {
      "key": "question5",
      "value": "4"
    },
    {
      "key": "question6",
      "value": "3"
    },
    {
      "key": "question7",
      "value": "5"
    },
    {
      "key": "question8",
      "value": "4"
    },
    {
      "key": "additionalComments9",
      "value": "1111"
    },
    {
      "key": "question10",
      "value": "3"
    },
    {
      "key": "question11",
      "value": "3"
    },
    {
      "key": "question12",
      "value": "5"
    },
    {
      "key": "additionalComments13",
      "value": "2222"
    },
    {
      "key": "question14",
      "value": "2"
    },
    {
      "key": "question15",
      "value": "2"
    },
    {
      "key": "question16",
      "value": "5"
    },
    {
      "key": "additionalComments17",
      "value": "3333"
    },
    {
      "key": "question18",
      "value": "2"
    },
    {
      "key": "question19",
      "value": "2"
    },
    {
      "key": "question20",
      "value": "4"
    },
    {
      "key": "additionalComments21",
      "value": "4444"
    },
    {
      "key": "question22",
      "value": "3"
    },
    {
      "key": "question23",
      "value": "3"
    },
    {
      "key": "question24",
      "value": "4"
    },
    {
      "key": "additionalComments25",
      "value": "55555"
    },
    {
      "key": "question26",
      "value": "4"
    },
    {
      "key": "question27",
      "value": "3"
    },
    {
      "key": "question28",
      "value": "4"
    },
    {
      "key": "additionalComments29",
      "value": "6666"
    },
    {
      "key": "question30",
      "value": "4"
    },
    {
      "key": "question31",
      "value": "3"
    },
    {
      "key": "question32",
      "value": "5"
    },
    {
      "key": "additionalComments33",
      "value": "7777"
    },
    {
      "key": "question34",
      "value": "2"
    },
    {
      "key": "question35",
      "value": "2"
    },
    {
      "key": "question36",
      "value": "1"
    },
    {
      "key": "additionalComments37",
      "value": "8888"
    },
    {
      "key": "question38",
      "value": "3"
    },
    {
      "key": "question39",
      "value": "4"
    },
    {
      "key": "question40",
      "value": "1"
    },
    {
      "key": "additionalComments41",
      "value": "9999"
    },
    {
      "key": "question42",
      "value": "2"
    },
    {
      "key": "question43",
      "value": "3"
    },
    {
      "key": "question44",
      "value": "5"
    },
    {
      "key": "additionalComments45",
      "value": "10 10 10 10"
    },
    {
      "key": "question46",
      "value": "2"
    },
    {
      "key": "question47",
      "value": "3"
    },
    {
      "key": "question48",
      "value": "4"
    },
    {
      "key": "additionalComments49",
      "value": "htrhrthtr"
    },
    {
      "key": "question50",
      "value": "4"
    },
    {
      "key": "question51",
      "value": "3"
    },
    {
      "key": "question52",
      "value": "5"
    },
    {
      "key": "additionalComments53",
      "value": "ggffgf"
    }
  ];
  **/

    //moduleId = 'CMP6006A';
    // surveyTable = 2 or surveyTable = 3
    //let surveyTable = 1

    //console.log(moduleId);
    //console.log('\n\n');
    //console.log(html);

    //Module Specification Data
    let sheet = workbook.getWorksheet("Sheet1");
    let table = sheet.getTables()[0];

    let data:Array<string> = [];

    data.push(Date().toString());
    //i skips the first two hidden values when set to 2
    for (let i = 2; i < Object.keys(results).length; i++  )
    {
        data.push(results[i]['value'])
    }
    console.log(data);
    //data.push(Date().toString());
    
    table.addRow(null, data)


    return 1;    
  

    //let data = sheet.getTables()[0].getRangeBetweenHeaderAndTotal().getValues();
    //let dataHeaders = sheet.getTables()[0].getHeaderRowRange().getValues();
    ////console.log(dataHeaders);

    //let moduleRow = -1;
    //do {
    //    moduleRow++;
    //} while (moduleRow < data.length && data[moduleRow][0] != moduleId)

    //let moduleName = data[moduleRow][1];
    //let moduleOrganiser = data[moduleRow][2];

    ////Survey Specification Data
    //let sheet_survey = workbook.getWorksheet("Sheet1");
    //let data_survey = sheet_survey.getTables()[surveyType].getRangeBetweenHeaderAndTotal().getValues();
    //let dataHeaders_survey = sheet_survey.getTables()[surveyType].getHeaderRowRange().getValues();
  //console.log(dataHeaders_survey);
}