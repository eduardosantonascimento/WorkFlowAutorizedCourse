var ssID = "link"

SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados").activate();

var dados = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var lr = dados.getLastRow();

var templateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(2, 1).getValue();

//var templateText2 =  HtmlService.createHtmlOutputFromFile('mail_template').getContent(); //SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Template").getRange(14, 1).getValue();

var email_Enviado = "email_Enviado";


  function SendEmailsServidor() {
          
    function confirm_Chefia(){
      
      var dados = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      dados.getRange(i,19).setValue("AUTORIZADO"); 
    } 
       
    //Function for send email confirmation boss
    function sendChefia(){
    
      for (var i = 2; i<=lr;i++){
    
      if (dados.getRange(i,6).getValue() != null && dados.getRange(i,15).getValue() != "email_Enviado" ) { 
  
    var templateText2 =  HtmlService.createTemplateFromFile('mail_template');
         templateText2.nomeChefia = dados.getRange(i, 6).getValue(); var nomeChefia = templateText2.nomeChefia;
         templateText2.matriculaChefia = dados.getRange(i, 7).getValue(); var matrChefia = templateText2.matriculaChefia;
    var emailChefia = dados.getRange(i,8).getValue();
         templateText2.cursoServidor =  dados.getRange(i, 9).getValue(); 
         templateText2.nomeServidor = dados.getRange(i, 2).getValue();
    var anexo = dados.getRange(i, 10).getValue();
   
        // var action = CardService.newAction().setFunctionName(confirm_Chefia);
        //CardService.newTextButton().setText('AUTORIZADO').setOnClickAction(action);
        //var  messageBody1 = templateText2.replace("(nome)",nomeChefia).replace("(matrChefia)",matrChefia).replace("(cursoServidor)",cursoServidor).replace(("nomeServidor"),nomeServidor).replace(("button"),CardService.newTextButton().setText('AUTORIZADO').setOnClickAction(action));
        //var messageBody1 = templateText2;
        //Logger.log(messageBody1);
        
         templateText2.confirm_Chefia = (function confirma_Chefia(){
         var dados = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
         dados.getRange(i,19).setValue("AUTORIZADO"); 
    } )
        
        
        var htmlText = templateText2.evaluate().getContent();
        var subjectLine = "Prezado(a) " + nomeChefia + ", matrícula: " +matrChefia +" você tem um novo formulário de analise de cursos para validar!";
      
       // GmailApp.sendEmail(emailChefia, subjectLine, templateText2);
       GmailApp.sendEmail(emailChefia,subjectLine,"Seu email não suporta HTML.",{name:subjectLine, htmlBody: htmlText + anexo});
                
       dados.getRange(i, 15).setValue(email_Enviado);  // atualizacao de campo de condicional
    }
   }
  }
    
    //Main
    
    for (var i = 2; i<=lr;i++){
    
      if (dados.getRange(i,2).getValue() != null && dados.getRange(i,14).getValue() != "email_Enviado" ) { 
    
    var nomeServidor = dados.getRange(i, 2).getValue();
    var matrServidor = dados.getRange(i, 3).getValue();
    var cargoServidor = dados.getRange(i, 4).getValue();
    var emailServidor = dados.getRange(i,5).getValue();
    var emailEnviadoServidor = dados.getRange(i, 14).getValue();
    var cursoServidor =  dados.getRange(i, 9).getValue(); 
    var nomeChefia = dados.getRange(i, 6).getValue();
 
    var  messageBody = templateText.replace("(nome)",nomeServidor).replace("(matrServidor)",matrServidor).replace("(cargoServidor)",cargoServidor).replace("(emailServidor)",emailServidor).replace("(nomeChefia)",nomeChefia);
    var subjectLine = "Prezado(a) " + nomeServidor + ", matrícula: " +matrServidor +" sua análise de cursos foi preenchida e enviada para sua chefia!";
      
    GmailApp.sendEmail(emailServidor, subjectLine, messageBody);
    dados.getRange(i, 14).setValue("email_Enviado");  // atualizacao de campo de condicional
        sendChefia();
        //confirm_Chefia();
      }
    
    }
      
    
  }
