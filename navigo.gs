function onFormSubmit(e){
  
  try{
    
    var tempId = '0B1zKBLC0y6EMfjdMbXh2TGNzZmx3OFJQV01WSVI1Vi1YcVFqWnVBV3VTN2dyelBZb09MTDQ'; //Dossier temporaire 
    var ssId = "1muZfRP9lafn6UsBuLGtMqPDjwF_oWvY3X0IcHnY1i8A"; //Spreadsheet
    var ss = SpreadsheetApp.openById(ssId);
    var globalId = ss.getSheetByName('Renommage').setActiveSelection('Renommage!B3').getValue(); //Dossier global 
    var recipient = ss.getSheetByName('Renommage').setActiveSelection('Renommage!B5').getValue(); //Adresse mail support
    //throw new Error("Et bim, erreur inattendue");
    
    //Accès au justificatif
    
    var tempFolder=DriveApp.getFolderById(tempId); 
    var file = tempFolder.getFiles().next();
    
    //Récupérer la réponse du formulaire
    
    var month = e.namedValues['Mois'][0];
    var abon = e.namedValues["Type d'abonnement"][0];
    var year = file.getDateCreated().getFullYear();
    var anMois = file.getDateCreated().getMonth();
    var numMonth = month.substring(0,2);

    //Récupérer adresse mail
    
    var sheet = ss.getSheetByName('Form Responses 1');
    var mailRange1 = ss.getRangeByName('MailRange1');//B1:B
    var mails1 = mailRange1.getValues();
    var lastRow = sheet.getLastRow();
    var mail = mails1[lastRow-1].join();  
    
    //Récupérer trigramme correspondant
    
    var triRange = ss.getRangeByName('TriRange');//A2:B
    var trigrammes = triRange.getValues();  
    for(var i in trigrammes){
      if(mail == trigrammes[i][0]){
        var trigramme = trigrammes[i][1];
      }
    } 
       
    //Récupérer modèle + renommage  
    
    if(abon == 'Mensuel'){
      var renom = ss.getSheetByName('Renommage').setActiveSelection('Renommage!B1').getValue();  
      var renamed = renom.replace('TRI',trigramme).replace('Année',year).replace('Mois',numMonth);  
      file.setName(renamed);
    }else{
      var renom = ss.getSheetByName('Renommage').setActiveSelection('Renommage!B2').getValue();  
      var renamed = renom.replace('TRI',trigramme).replace('Année',year);  
      file.setName(renamed);
    }
    
    //Création onglet spreadsheet + remplissage nom, prénom
    
    var sheets = ss.getSheets()
    var newSheet = true;
    var prenom = mail.slice(0 , mail.indexOf('@')).slice(0 , mail.indexOf('.'));
    var nom = mail.slice(0 , mail.indexOf('@')).slice(mail.indexOf('.')+1);
    
    if(prenom.indexOf('-')>-1){
    
      var prenom1 = prenom.slice(0 , prenom.indexOf('-'));
      var prenom2 = prenom.slice(prenom.indexOf('-')+1);
      var Prenom1 = prenom1.charAt(0).toUpperCase() + prenom1.slice(1);
      var Prenom2 = prenom2.charAt(0).toUpperCase() + prenom2.slice(1);
      var Prenom = Prenom1+' '+Prenom2;
    
    }else{
    
      var Prenom = prenom.charAt(0).toUpperCase() + prenom.slice(1);    
    }
    
    if(nom.indexOf('-')>-1){
    
      var nom1 = nom.slice(0 , nom.indexOf('-'));
      var nom2 = nom.slice(nom.indexOf('-')+1);
      var Nom1 = nom1.charAt(0).toUpperCase() + nom1.slice(1);
      var Nom2 = nom2.charAt(0).toUpperCase() + nom2.slice(1);
      var Nom = Nom1+' '+Nom2;
    
    }else{

      var Nom = nom.charAt(0).toUpperCase() + nom.slice(1); 
    }    
    
    if(abon == 'Mensuel'){
      
      for(i in sheets){

        if(sheets[i].getName() == month){
        
          var nextRow = sheets[i].getLastRow()+1; 
          ss.setActiveSheet(sheets[i]); 
          var cell1 = sheets[i].getRange(nextRow,1);
          var cell2 = sheets[i].getRange(nextRow,2);
          cell1.setValue(Nom);
          cell2.setValue(Prenom);
          var newSheet = false;
        }
      }
      
      if(newSheet){
        var d = false;
        for(var i in sheets){
          if(sheets[i].getName().charAt(0) == '0' || sheets[i].getName().charAt(0) == '1'){
            var num = parseInt(sheets[i].getName().substring(0,2));
            var newNum = parseInt(numMonth);
            if(num>newNum){
              var pos = parseInt(i);
              var d = true;
              break;
            }
          }
        }
        if(!d){
          var pos = sheets.length;
        }
        var monthSheet = ss.insertSheet(month,pos);
        var nomCell = monthSheet.getRange(1,1);
        var prenomCell = monthSheet.getRange(1,2);
        nomCell.setValue('NOM');
        prenomCell.setValue('PRENOM');
        var nextNewRow = monthSheet.getLastRow()+1;
        var newCell1 = monthSheet.getRange(nextNewRow,1);
        var newCell2 = monthSheet.getRange(nextNewRow,2);   
        newCell1.setValue(Nom);
        newCell2.setValue(Prenom);
      } 
      
    }else{
    
      for(i in sheets){
        
        if(sheets[i].getName() == 'Annuel'){
          
          var nextRow = sheets[i].getLastRow()+1;
          ss.setActiveSheet(sheets[i]); 
          var cell1 = sheets[i].getRange(nextRow,1);
          var cell2 = sheets[i].getRange(nextRow,2);
          cell1.setValue(Nom);
          cell2.setValue(Prenom);
          var newSheet = false;
        }
      }
      
      if(newSheet){
        
        var anSheet = ss.insertSheet('Annuel',3);
        var nomCell = anSheet.getRange(1,1);
        var prenomCell = anSheet.getRange(1,2);
        nomCell.setValue('NOM');
        prenomCell.setValue('PRENOM');
        var nextNewRow = anSheet.getLastRow()+1;
        var newCell1 = anSheet.getRange(nextNewRow,1);
        var newCell2 = anSheet.getRange(nextNewRow,2);        
        newCell1.setValue(Nom);
        newCell2.setValue(Prenom);
      }
    }
    
    //Logs
    
    if(abon == 'Mensuel'){
      console.info(['Mail: '+mail],['Trigramme: '+trigramme],['Année: '+year],['Mois: '+month]);
    }else{
      console.info(['Mail: '+mail],['Trigramme: '+trigramme],['Année: '+year]);
    }
    console.info('Fichier renommé: '+renamed);  
    
    //[Dossier global] Création dossiers Année, Mois + déplacement
    
    var yearFolders = DriveApp.getFolderById(globalId).getFolders();
    var done = false;

    while(yearFolders.hasNext()){
    
      var yearFolder = yearFolders.next();
      var d = parseInt(yearFolder.getName().substring(0,4));
      
      if(abon == 'Mensuel'){
        
        if(year == d && (month=='10 Octobre'||month=='11 Novembre'||month=='12 Décembre')){
        
          var done = true;
          var done0 = false;
          
          while(yearFolder.getFolders().hasNext()){
          
            var mFolder = yearFolder.getFolders().next();
            
            if(mFolder.getName()==month){
            
              mFolder.addFile(file);
              tempFolder.removeFile(file);
              console.info('Fichier déplacé dans le dossier global (Année '+year+'-'+(year+1)+' / Mois '+month+')');
              var done0 = true;              
            }
            break;
          }
          
          if(!done0){
          
            yearFolder.createFolder(month).addFile(file);
            tempFolder.removeFile(file);
            console.info('Fichier déplacé dans le dossier global (Année '+year+'-'+(year+1)+' / Mois '+month+')');
            done0 = true;
          }
        }  

        if(year==d+1 && (month=='01 Janvier'||month=='02 Février'||month=='03 Mars'||month=='04 Avril'||month=='05 Mai'||month=='06 Juin'||month=='07 Juillet'||month=='08 Août'||month=='09 Septembre')){
          
          var done = true;
          var done0 = false;
          
          while(yearFolder.getFolders().hasNext()){
          
            var mFolder = yearFolder.getFolders().next();
            
            if(mFolder.getName()==month){
            
              mFolder.addFile(file);
              tempFolder.removeFile(file);
              console.info('Fichier déplacé dans le dossier global (Année '+(year-1)+'-'+year+' / Mois '+month+')');
              var done0 = true;              
            }
            break;
          }
          
          if(!done0){
          
            yearFolder.createFolder(month).addFile(file);
            tempFolder.removeFile(file);
            console.info('Fichier déplacé dans le dossier global (Année '+(year-1)+'-'+year+' / Mois '+month+')');
            var done0 = true;
          }
        
       } 
      }else if(year == d && anMois >= 9){
      
        yearFolder.addFile(file);
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier global (Année '+year+'-'+(year+1)+')');
        var done = true; 
        
      }else if(year == d+1 && anMois < 8){
      
        yearFolder.addFile(file);
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier global (Année '+(year-1)+'-'+year+')');
        var done = true;        
      }
      
    }
    
    if(!done){
    
      if(abon == 'Mensuel'){
      
        if(month =='10 Octobre'||month=='11 Novembre'||month=='12 Décembre'){
        
          var yFolder = DriveApp.getFolderById(globalId).createFolder(year+" - "+(year+1));
          yFolder.createFolder(month).addFile(file);
          tempFolder.removeFile(file);
          console.info('Fichier déplacé dans le dossier global (Année '+year+'-'+(year+1)+' / Mois '+month+')');
               
        }else{
        
          var yFolder = DriveApp.getFolderById(globalId).createFolder((year-1)+" - "+year);
          yFolder.createFolder(month).addFile(file);
          tempFolder.removeFile(file);
          console.info('Fichier déplacé dans le dossier global (Année '+(year-1)+'-'+year+' / Mois '+month+')');
        }
          
      }else if(anMois >= 9){  
      
        DriveApp.getFolderById(globalId).createFolder(year+'-'+(year+1)).addFile(file);
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier global (Année '+year+'-'+(year+1)+')');
      }
      
      else if(anMois < 8){
      
        DriveApp.getFolderById(globalId).createFolder((year-1)+'-'+year).addFile(file);
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier global (Année '+(year-1)+'-'+year+')');
      }
    }
    
    //[Dossier perso] Création dossier Année + déplacement
    
    /*var collabId = ss.getSheetByName('Renommage').setActiveSelection('Renommage!B4').getValue();
    var collabFolder = DriveApp.getFolderById(collabId).getFoldersByName(trigramme).next().getFoldersByName(trigramme+'-Frais').next().getFoldersByName(trigramme+'-Transports').next();  
    var lastFolders = collabFolder.getFolders();
    var done1=false;
    
    while(lastFolders.hasNext()){ 
    
      var lastFolder = lastFolders.next();    
      var d = parseInt(lastFolder.getName().substring(0,4));
      
      if(abon == 'Mensuel'){
      
        if(year==d && (month=='10 Octobre'||month=='11 Novembre'||month=='12 Décembre')){
        
          var done1 = true;
          lastFolder.addFile(file);
          tempFolder.removeFile(file);
          console.info('Fichier déplacé dans le dossier collab (Année '+year+'-'+(year+1)+')');
        }
        
        if(year==d+1 && (month=='01 Janvier'||month=='02 Février'||month=='03 Mars'||month=='04 Avril'||month=='05 Mai'||month=='06 Juin'||month=='07 Juillet'||month=='08 Août'||month=='09 Septembre')){
        
          var done1 = true; 
          lastFolder.addFile(file);
          tempFolder.removeFile(file);
          console.info('Fichier déplacé dans le dossier collab (Année '+(year-1)+'-'+year+')');
        }
        
      }else if(year == d && anMois >= 9){
      
        lastFolder.addFile(file);
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier collab (Année '+year+'-'+(year+1)+')');
        var done1 = true;  
      }
      else if(year == d+1 && anMois < 8){
      
        lastFolder.addFile(file);
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier collab (Année '+(year-1)+'-'+year+')');
        var done1 = true;        
      }
    }
    
    if(!done1){
    
      if(abon == 'Mensuel'){
      
        if(month=='10 Octobre'||month=='11 Novembre'||month=='12 Décembre'){
          collabFolder.createFolder(year+" - "+(year+1)).addFile(file);     
          tempFolder.removeFile(file);
          console.info('Fichier déplacé dans le dossier collab (Année '+year+'-'+(year+1)+')');
          
        }else{
        
        collabFolder.createFolder((year-1)+" - "+year).addFile(file);    
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier collab (Année '+(year-1)+'-'+year+')');
        }
        
      }else if(anMois >= 9){
    
        collabFolder.createFolder(year+'-'+(year+1)).addFile(file);
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier collab (Année '+year+'-'+(year+1)+')');
        
      }else if(anMois < 8){
      
        collabFolder.createFolder((year-1)+'-'+year).addFile(file);
        tempFolder.removeFile(file);
        console.info('Fichier déplacé dans le dossier collab (Année '+(year-1)+'-'+year+')');
      }
    }*/
    
  //Envoi d'un mail en cas d'erreur
  
  }catch(err){
    var tempId = '0B1zKBLC0y6EMfjdMbXh2TGNzZmx3OFJQV01WSVI1Vi1YcVFqWnVBV3VTN2dyelBZb09MTDQ';
    var tempFolder=DriveApp.getFolderById(tempId); 
    var file = tempFolder.getFiles().next();
    var mailTemplate =  HtmlService.createTemplateFromFile('mail_template'); 
    mailTemplate.params = null;
    var logoId = "0BwoJsQcPlVIZN0taY2RoYTFFWVE";
    var logoBlob = DriveApp.getFileById(logoId).getBlob().setName("logo");
    var parameters = {};
    parameters["inlineImages"] = {};
    parameters["inlineImages"]["logo"] = logoBlob;
    parameters["htmlBody"] = mailTemplate.evaluate().getContent();
    GmailApp.sendEmail(recipient,'Erreur Globale','',parameters);
    console.error(err.message);
    tempFolder.removeFile(file);    
  }
}
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [];
  menu.push({name: "Supprimer les onglets", functionName: "del"});
  ss.addMenu('Onglets',menu);
}

function del(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for(var i in sheets){
    
    var name = sheets[i].getName();
    if(name.indexOf('0')==0 || name.indexOf('1')==0 || name == 'Annuel'){
    
      ss.deleteSheet(sheets[i]);
    }
  }
}
