

function onOpen() {

  createMenu();
     
  createHeaders();
}


function createMenu(){

  var menu = SpreadsheetApp.getUi() 
     .createMenu('QuizBanking');
   
  if (PropertiesService.getDocumentProperties().getProperty(QB_FORM_ID)) {
    menu.addItem('Reinicia el procés', 'reset');
    menu.addItem("Enviar respostes clau", 'createFlubarooResponses');
  } else {
    menu.addItem('Crear Formulari', 'uiStep1');
  }
  
  menu.addToUi();
  
}

function reset(){

  
  var documentProperties = PropertiesService.getDocumentProperties();

  documentProperties.deleteAllProperties();

  createMenu();
  //later calls onOpen()
    
}

function createHeaders() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  sheet.setColumnWidth(1, 400);

  // The size of the two-dimensional array must match the size of the range.
  var values = [
    [ "Pregunta", "Opció correcta", "Altres opcions" ]
  ];

  var range = sheet.getRange("A1:C1");
  range.setValues(values); 
  
}

function getCurrentSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function getFirstEmptyRow(sheet) {
  
  
  var column = sheet.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct][0] != "" ) {
    Logger.log("values[ct][0] : " + values[ct][0] );
    ct++;
  }
  
  ct++;
  
  Logger.log("First empty row: " + ct);
  
  return ct;
}

/**
* Listener
*/
function createForm(formName, quiz, oneAnswer, shuffle, authenticated) {
 
  var documentProperties = PropertiesService.getDocumentProperties();


  documentProperties.setProperty(QB_OPT_IS_QUIZ, quiz);
  documentProperties.setProperty(QB_OPT_ONE_ANS_PER_USER, oneAnswer);
  documentProperties.setProperty(QB_OPT_LOGIN_REQ, authenticated);
  documentProperties.setProperty(QB_OPT_SHUFFLE, shuffle);
  
  // Form configuration

  var form = FormApp.create(formName);
  form.setCollectEmail(true);
  form.setIsQuiz(quiz);
  form.setLimitOneResponsePerUser(oneAnswer);
  //form.setRequireLogin(authenticated);
  form.setShuffleQuestions(shuffle);
  
  moveForm(form);
  
 
  documentProperties.setProperty(QB_FORM_ID, form.getId());
  
 
  var sheet = getCurrentSheet();
 
  var firstEmptyRow = getFirstEmptyRow(sheet);
  
  Logger.log("firstEmptyRow: "  + firstEmptyRow);
  
  var correctAnswers = [];
    

  
  for (var row = 2; row < firstEmptyRow; row ++){
    Logger.log("Entra al bucle");
    var range = sheet.getRange(row,1,1,sheet.getMaxColumns() );
    
    var i = 0;
    var ask = range.getValues()[0][i];
     Logger.log("Ask: "  + ask);
    i++;
    
    var answers = [];
  
    var multipleChoice = form.addMultipleChoiceItem()
      .setTitle(ask)
      .showOtherOption(false)
      .setPoints(10);
    
    while( range.getValues()[0][i] != ""){

      if (i==1) correctAnswers.push(range.getValues()[0][i]);
      var choice = multipleChoice.createChoice(range.getValues()[0][i],(i == 1));
      answers.push(choice);
      i++;
    }
    
    multipleChoice.setChoices(shuffleArray(answers));

    Logger.log(ask);
    Logger.log(answers);
    
  }
  
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(QB_CORRECT_ANSWERS, JSON.stringify(correctAnswers));
  
  
  createMenu();
  
  Logger.log('Published URL: ' + form.getPublishedUrl());
  Logger.log('Editor URL: ' + form.getEditUrl());
  
  
}

function moveForm(form){
   // checks if the folder of the current course exists
  var now = new Date();
  var folderName = "QUIZ " + now.getFullYear();
  var folders = DriveApp.getFoldersByName(folderName);
  
  var folderId;
  
  while (folders.hasNext()){
    
    var folder = folders.next();
    folderId = folder.getId();
  } 
  
  if (!folderId){
     var folder = DriveApp.createFolder(folderName);
     folderId = folder.getId();
  }
  
  //move the created form to the folder
  var formFile = DriveApp.getFileById(form.getId());
  
  // Remove the file from all parent folders
  var parents = formFile.getParents();
  while (parents.hasNext()) {
    var parent = parents.next();
    parent.removeFile(formFile);
  }
  
  DriveApp.getFolderById(folderId).addFile(formFile);
  
}

function createFlubarooResponses(){

  var documentProperties = PropertiesService.getDocumentProperties();
  Logger.log(documentProperties.getProperty(QB_CORRECT_ANSWERS));
  var correctAnswers =  JSON.parse(documentProperties.getProperty(QB_CORRECT_ANSWERS));
  

   
  var form = FormApp.openById(PropertiesService.getDocumentProperties().getProperty(QB_FORM_ID));
  
  //disable restrictions temporaly
  form.setLimitOneResponsePerUser(false);
  form.setRequireLogin(false);
  form.setCollectEmail(false);
  
  var questions = form.getItems();
    
  var formResponse = form.createResponse();
    
  for (var i=0; i<questions.length; i++){
  
    var question = questions[i].asMultipleChoiceItem();
    var response = question.createResponse(correctAnswers[i]);
     
    formResponse.withItemResponse(response);
  }
    
  formResponse.submit();
  
  Logger.log(documentProperties.getProperty(QB_OPT_ONE_ANS_PER_USER));

  //restore restrictions
  form.setLimitOneResponsePerUser( documentProperties.getProperty(QB_OPT_ONE_ANS_PER_USER) === "true");
  form.setRequireLogin( documentProperties.getProperty(QB_OPT_LOGIN_REQ)  === "true");
  form.setCollectEmail(documentProperties.getProperty(QB_OPT_SHUFFLE)  === "true");
  form.setCollectEmail(true);
 

}


 

function uiStep1(){
    var html = HtmlService.createTemplateFromFile('uiStep1')
      .evaluate()
      .setWidth(350) 
      .setHeight(350)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    SpreadsheetApp.getUi()
                  .showModalDialog(html, 'Configuració del formulari');
}

