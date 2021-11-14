// colonne des cases à cocher (A= 0,B = 1 etc...)
var casesCol = 2;
//colonne des noms de outils 
var toolsCol = 3;
//Ligne du départ de la liste d'outils (premier outil; commence à 1 puis 2 etc )
var row = 2;
//Colonne de la liste d'outils
var col = 5;
// Noms de la feuille avec les données (Outils par tâche, à mettre entre "" ou '')
var nameDataSheet = 'Step1'; 
// Noms de la feuille avec la liste des outils nécessaires (à mettre entre "" ou '')
var nameListSheet = 'Step2'; 

// Récupére les outils nécesaires, les trie puis les ecrit dans la feuille Step2 en supprimant les doublons
function tools_needed() {
  // Récupération de la liste d'outils nécessaires
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameDataSheet);
  var data = dataSheet.getDataRange().getValues();
  var writingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameListSheet);
  var toollist = [];
  for (var i = 0; i < data.length; i++) {
    if(data[i][casesCol]){
      toollist.push(data[i][toolsCol])
    }
  } 

  // Tri et ecriture des outils dans la colonne E de la feuille Step2
  // Modifier si besoin les var row et col en fonction du fichier 

  var previousData = []; 
  toollist.sort();
  for(var i = 0; i<toollist.length; i++){
    if(previousData.includes(toollist[i])){
      Logger.log('Doublons trouvé : ' + toollist[i]);
    }else {
      previousData.push(toollist[i]);
      writingSheet.getRange(row,col).setValue(toollist[i]);
      Logger.log('Ecriture :' + toollist[i]);
      row ++;
    }
  }
}

//Reset la tool list : supprime les outils contenus dans la colonne E
function resetToolsneeded()
{
  var writingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameListSheet);
  var writingSheetData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameListSheet).getDataRange().getValues();
  for(var i = row; i< writingSheetData.length; i++){
    writingSheet.getRange(i,col).setValue("");
  }
}
