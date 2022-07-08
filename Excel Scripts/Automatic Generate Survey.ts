function main(workbook: ExcelScript.Workbook, moduleId: string = "CMP6002B", surveyType: number = 1) {
    //Assumption: Sheets are ordered:
    // Module Criteria , Mid Module Spec. , End of Module Spec.
    // Module Criteria - Determines which staff the survey will generate questions for. Currently this is required by End Module spec., but not by mid module spec.
    // Mid Module Spec. - Determines the questions specified for the middle of the module survey. 
    // End Module Spec. - Determines the questions specified for the end of the module survey.
  
    // surveyType - number determines the type of survey created. 1 = mid module, 2 = end module.
  
    //Module Criteria
    let sheet = workbook.getWorksheets()[0]; //workbook.getWorksheet("Sheet1");
    let data = sheet.getTables()[0].getRangeBetweenHeaderAndTotal().getValues();
    let dataHeaders = sheet.getTables()[0].getHeaderRowRange().getValues();
  
    //Find the row for the selected moduleId
    let moduleRow = -1;
    do {
      moduleRow++;
    } while (moduleRow < data.length && data[moduleRow][0] != moduleId)
    let moduleName = data[moduleRow][1];
  
    //Mid / End Module Specification
    let sheet_survey = workbook.getWorksheets()[surveyType]; //workbook.getWorksheet("Sheet1");
    let data_survey = sheet_survey.getTables()[0].getRangeBetweenHeaderAndTotal().getValues();
    let dataHeaders_survey = sheet_survey.getTables()[0].getHeaderRowRange().getValues();
  
    let html = '';
    let questions:Array<string> = [];
  
    let htmlTemplateStart = `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Web form</title>
      <style>
      /* Styling the Body element i.e. Color,
      Font, Alignment */
      body {
        background-color: #EEEEEE;
    font-family: Verdana;
    text-align: center;
  }
  
  tr:nth-child(even) {
    background-color: #f2f2f2;
  }
  
  /* Styling the Form (Color, Padding, Shadow) */
  form {
    background-color: #fff;
    max-width: 500px;
    margin: 50px auto;
    padding: 30px 20px;
    box-shadow: 2px 5px 10px rgba(0, 0, 0, 0.5);
  }
  
          /* Styling form-control Class */
          .form-control {
    text-align: left;
    margin-bottom: 25px;
  }
  
          /* Styling form-control Label */
          .form-control label {
    display: block;
    margin-bottom: 10px;
  }
  
          /* Styling form-control input,
          select, textarea */
          .form-control input,
          .form-control select,
          .form-control textarea {
    border: 1px solid #777;
    border-radius: 2px;
    font-family: inherit;
    padding: 10px;
    display: block;
    width: 95%;
  }
  
          /* Styling form-control Radio
          button and Checkbox */
          .form-control input[type="radio"],
          .form-control input[type="checkbox"] {
    display: inline-block;
    width: auto;
  }
  
  /* Styling Button */
  button {
    background-color: #00aef0;
    border: 1px solid #777;
    border-radius: 2px;
    font-family: inherit;
    font-size: 21px;
    display: block;
    width: 100%;
    margin-top: 50px;
    margin-bottom: 20px;
  }
  </style></head><body><h1> [${moduleId}] ${moduleName} - ${dataHeaders_survey[0][1]}</h1><form action="https://prod-161.westeurope.logic.azure.com:443/workflows/c94642a3df8b4b259fb2a04f9fccc2a0/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=NU5qzs5sASZHrBssWMu052fCcX-3zn5uxcOPstV4NYM" method="POST"><input type="hidden" name="ModuleId" value="${moduleId}" /><input type="hidden" name="SurveyType" value="${surveyType}" />`;
  
    let htmlTemplateEnd = '<button type="submit">Submit</button></form></body></html>';
    let likertTemplateStart = '<table class="radio-table"><thead><tr><th>Question</th><th>Strongly Disagree</th><th>Disagree</th><th>Neither Agree or Disagree</th><th>Agree</th><th>Strongly Agree</th></tr></thead><tbody>';
    let likertTemplateEnd = '</tbody></table>';
    let questionHtml = '';
  
    html += htmlTemplateStart;
  
    let lastQuestionType = '';
    let i = 0;
    while (i < data_survey.length) {
      let questionType = data_survey[i][0];
      let questionText = data_survey[i][1];
  
      if (questionType == 'Likert') {
        questionHtml = `<tr><td>${questionText}</td><td><input type="radio" value="1" name="question${(i + 1)}" required/></td><td><input type="radio" value="2" name="question${(i + 1)}" required/></td><td><input type="radio" value="3" name="question${(i + 1)}" required/></td><td><input type="radio" value="4" name="question${(i + 1)}" required/></td><td><input type="radio" value="5" name="question${(i + 1)}" required/></td></tr>`;
  
        if (lastQuestionType != 'Likert') {
          //Create a new table for this new Likert.
          html += likertTemplateStart;
          html += questionHtml;
        }
        else {
          //Add question to the existing table.
          html += questionHtml;
        }
  
        questions.push(questionText.toString());
      }
      else if (questionType == 'Text') {
        //Add a textbox input.
        if (lastQuestionType == 'Likert') {
          html += likertTemplateEnd;
        }
  
        questionHtml = `<div><h2>Any Additional Coments...</h2><br><label for="html">${questionText}</label><br><textarea rows="3" cols="50" name="additionalComments${(i + 1)}" placeholder="..."></textarea></div>`;
  
        html += questionHtml;
  
        questions.push(questionText.toString() + '*');
      }
      else {
        //When 'Likert-Person' or 'Text-Person' is reached, only staff questions remain. Break from this loop and start the generation of the staff questions.
        break;
      }
  
      lastQuestionType = questionType;
      i++;
    }
  
    //Handling Staff questions:
    let numberQuestionsPerPerson = data_survey.length - i ; // - 1
    let questionNumber = 0;
    let peopleCount = 0;
  
    //Index 3 is when staff in the Module Criteria Table starts. (After ModuleId and ModuleName).
    for (let a = 3; a < dataHeaders[0].length; a++) {
      let name = data[moduleRow][a];
  
      if (name.toString().length > 0) {
        for (let b = i; b < data_survey.length; b++) {
          let questionType = data_survey[b][0];
          let questionText = data_survey[b][1];
  
          questionNumber = b + (peopleCount * numberQuestionsPerPerson)  + 1;
  
          if (questionType == 'Likert-Person') {
            questionHtml = `<tr><td>${questionText}</td><td><input type="radio" value="1" name="question${questionNumber}" required/></td><td><input type="radio" value="2" name="question${questionNumber}" required/></td><td><input type="radio" value="3" name="question${questionNumber}" required/></td><td><input type="radio" value="4" name="question${questionNumber}" required/></td><td><input type="radio" value="5" name="question${questionNumber}" required/></td></tr>`;
  
            if (lastQuestionType != 'Likert-Person') {
              //Create a new table for this new Likert-Person.
              html += `<br><br><br><label>For ${name}, please rate the following: </label><br>` + likertTemplateStart;
              html += questionHtml;
            }
            else {
              //Add question to the existing table.
              html += questionHtml;
            }
  
            questions.push(`(${name}) ${questionText.toString()}`);
          }
          else if (questionType == 'Text-Person') {
            //Add a textbox input.
            if (lastQuestionType == 'Likert-Person') {
              html += likertTemplateEnd;
            }
  
            questionHtml = `<div><h2>Any Additional Coments...</h2><br><label for="html">${questionText}</label><br><textarea rows="3" cols="50" name="additionalComments${questionNumber}" placeholder="..."></textarea></div>`;
  
            html += questionHtml;
  
            questions.push(`(${name}) ${questionText.toString()}*`);
          }
          else {
            //If there was unexpected input then give a basic error message. This does nothing in Power Automate currently. 
            console.log('Unexpected Input Error');
          }
  
          lastQuestionType = questionType.toString();
        }
  
        peopleCount++;
      }
    }
  
    html += htmlTemplateEnd;
    console.log(html);
  
    //return the html for the survey and the list of numerated questions.
    return [html, questions];
  }