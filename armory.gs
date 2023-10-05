/** @OnlyCurrentDoc */

function ValidationVentes() {

  var spreadsheet = SpreadsheetApp.getActive();

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Stocks'), true);
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getRange('E:E').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

  //Sélection + insertion de 10 lignes dans data_ventes
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('data_ventes'), true);
  spreadsheet.getRange('3:12').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 10);

  //Décalage de 0 puis sélection
  spreadsheet.getActiveRange().offset(0, 0, 10, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A12').activate();
  spreadsheet.getCurrentCell().setFormula('=A13+1');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3:A12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A3:A12').activate();
  spreadsheet.getRange('C3').activate();
  spreadsheet.getRange('\'Fiche de Vente\'!B6:D15').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('F3').activate();
  spreadsheet.getRange('\'Fiche de Vente\'!B1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('F3:F12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('G3').activate();
  spreadsheet.getRange('\'Fiche de Vente\'!B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('G3:G12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B3').activate();

  //Fonction NOW() dans E1 pour l'horodatage
  spreadsheet.getRange('\'Fiche de Vente\'!E1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('B4').activate();
  spreadsheet.getCurrentCell().setFormula('=B3');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B4:B12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  //Nettoyage fiche de vente
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Fiche de Vente'), true);
  spreadsheet.getRange('B6:C15').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B1:D3').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Fiche de Vente'), true);
};

function ValidationVentes2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Stocks'), true);
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getRange('G:G').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('data_ventes'), true);
  spreadsheet.getRange('3:12').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 10);
  spreadsheet.getActiveRange().offset(0, 0, 10, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A12').activate();
  spreadsheet.getCurrentCell().setFormula('=A13+1');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3:A12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A3:A12').activate();
  spreadsheet.getRange('C3').activate();
  spreadsheet.getRange('\'Fiche de Vente 2\'!B6:D15').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('F3').activate();
  spreadsheet.getRange('\'Fiche de Vente 2\'!B1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('F3:F12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('G3').activate();
  spreadsheet.getRange('\'Fiche de Vente 2\'!B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('G3:G12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B3').activate();
  spreadsheet.getRange('\'Fiche de Vente 2\'!E1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('B4').activate();
  spreadsheet.getCurrentCell().setFormula('=B3');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B4:B12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

    //Nettoyage fiche de vente
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Fiche de Vente 2'), true);
  spreadsheet.getRange('B6:C15').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B1:D3').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Fiche de Vente 2'), true);
};

function ValidationVentes3() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Stocks'), true);
  spreadsheet.getRange('D:D').activate();
  spreadsheet.getRange('I:I').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('data_ventes'), true);
  spreadsheet.getRange('3:12').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 10);
  spreadsheet.getActiveRange().offset(0, 0, 10, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('A12').activate();
  spreadsheet.getCurrentCell().setFormula('=A13+1');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A3:A12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A3:A12').activate();
  spreadsheet.getRange('C3').activate();
  spreadsheet.getRange('\'Fiche de Vente 3\'!B6:D15').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('F3').activate();
  spreadsheet.getRange('\'Fiche de Vente 3\'!B1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('F3:F12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('G3').activate();
  spreadsheet.getRange('\'Fiche de Vente 3\'!B2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('G3:G12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B3').activate();
  spreadsheet.getRange('\'Fiche de Vente 3\'!E1').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('B4').activate();
  spreadsheet.getCurrentCell().setFormula('=B3');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B4:B12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

    //Nettoyage fiche de vente
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Fiche de Vente 3'), true);
  spreadsheet.getRange('B6:C15').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('B1:D3').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Fiche de Vente 3'), true);
};

function dates() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B2:D3').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('data_ventes'), true);
  spreadsheet.getRange('3:12').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 10);
  spreadsheet.getActiveRange().offset(0, 0, 10, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRange('B3').activate();
  spreadsheet.getCurrentCell().setFormula('=TODAY()');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B3:B12'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B3:B12').activate();
  spreadsheet.getRange('B3:B12').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};

function Nettoyage() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('data_ventes'), true);
  var rows = spreadsheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var rowsDeleted = 0;
  for (var i = 3; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[2] == '') { // This searches all cells in columns A (change to row[1] for columns B and so on) and deletes row if cell is empty
      spreadsheet.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
  
  // Réparation des ID suite aux suppressions
  var derniereLigne = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data_ventes").getLastRow();
  var id = 1
  for (var j = derniereLigne; j > 2; j--){
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data_ventes").getRange(j, 1).setValue(id);
    id++;
  }
};

function onEdit(t) { 
  if(SpreadsheetApp.getActiveSheet().getName() == "Salaires & Primes"){
  // Ecriture auto de date sur une colonne et ligne venant d'être modif - t: spreadsheet actuelle
    var row = t.range.getRow(); // n° de la ligne actuellement éditée
    var col = t.range.getColumn(); // n° de la colonne actuellement éditée

    if (col == 2 && row > 1 && t.range.getValue() != "" && t.source.getActiveSheet().getRange(row, 1).getValue() == "") { 
      // Si 2ème colonne et ligne > 1 contenant valeur vide et qu'une date pas déjà entrée + transfert du salaire
      var nom = t.source.getActiveSheet().getRange(row, 2).getValue();
      var salaire;
      t.source.getActiveSheet().getRange(row, 1).setValue(new Date());
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fiche Employés").getRange(1,2).setValue(nom);
      salaire = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fiche Employés").getRange(3,2).getValue();
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Salaires & Primes").getRange(row, 3).setValue(salaire);
    };

  }
  if(SpreadsheetApp.getActiveSheet().getName() == "Employés"){
    // Ecriture auto date et grade lors du recrutement de nouvel employé
    var row = t.range.getRow(); // n° de la ligne actuellement éditée
    var col = t.range.getColumn(); // n° de la colonne actuellement éditée
    if (col == 3 && row > 1 && t.range.getValue() != false) { // Si 3ème colonne et ligne > 1 contenant valeur vide
      t.source.getActiveSheet().getRange(row, 6).setValue(new Date());
      t.source.getActiveSheet().getRange(row, 8).setValue("Employé(e)");
    };
    if (col == 3 && row > 1 && t.range.getValue() != true) { // Si 3ème colonne et ligne > 1 contenant valeur vide
      var nomEmploye = t.source.getActiveSheet().getRange(row, 4).getValue();
      var prenomEmploye = t.source.getActiveSheet().getRange(row, 5).getValue();
      var identiteEmploye = prenomEmploye + " " + nomEmploye;
      var confirmation = SpreadsheetApp.getUi();
      var clic = confirmation.alert("Virer un employé","Êtes-vous sûr de vouloir virer cet employé ? Toutes ses ventes seront supprimées ainsi que ses salaires donnés.", confirmation.ButtonSet.YES_NO);
      if(clic == confirmation.Button.YES){
        // Suppression des valeurs lorsque case décochée
        for (var i = 4; i < 9; i++){
          t.source.getActiveSheet().getRange(row, i).setValue("");
        }
        effacerCompta(identiteEmploye);
      }
      if(clic == confirmation.Button.NO){
        t.source.getActiveSheet().getRange(row, 3).setValue(true);
      }
    };
  }
};

function effacerCompta(employe){

  // Suppression des ventes de l'employé
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('data_ventes'), true);
  var rows = spreadsheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var rowsDeleted = 0;

  for (var i = 2; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[5] == employe) { // This searches all cells in columns A (change to row[1] for columns B and so on) and deletes row if cell is empty
      spreadsheet.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
  
  // Réparation des ID suite aux suppressions
  var derniereLigne = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data_ventes").getLastRow();
  var id = 1
  for (var j = derniereLigne; j > 2; j--){
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data_ventes").getRange(j, 1).setValue(id);
    id++;
  }

    // Suppression des salaires de l'employé
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Salaires & Primes'), true);
  var rows = spreadsheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var rowsDeleted = 0;
  
  for (var i = 2; i <= numRows - 1; i++) {
    var row = values[i];
    if (row[1] == employe) { // This searches all cells in columns A (change to row[1] for columns B and so on) and deletes row if cell is empty
      spreadsheet.deleteRow((parseInt(i)+1) - rowsDeleted);
      rowsDeleted++;
    }
  }
}