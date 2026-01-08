var ID_ROOT_DIR = "{{ID_ROOT_DIR}}"
var ID_TEMPLATE_SS = "{{ID_TEMPLATE_SS}}"

var MONTH = ["JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"]

function doPost(e: { postData: { contents: string } }) {
  try {
    const params = JSON.parse(e.postData.contents);

    // créer le spreadsheet dans le dossier Drive
    const spreadsheet = getSpreadSheet()
    if (!spreadsheet) return

    // créer la feuille dans le sheet
    const sheet = getSheet(spreadsheet.getId())
    if (!sheet) return
    
    // L'ordre doit correspondre à tes colonnes : 
    // Sommeil, Calories, Protéines, Glucides, Lipides, Pas, Poids
    const dataArray = [
      params.sommeil,
      params.calories,
      params.proteines,
      params.glucides,
      params.lipides,
      params.pas,
      params.poids
    ];
    
    setData(spreadsheet.getId(), sheet.getName(), dataArray);
    
    return ContentService.createTextOutput("Données reçues !").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Erreur : " + err).setMimeType(ContentService.MimeType.TEXT);
  }
}

function getSpreadSheet() {
    const date = new Date()

    const root_dir = DriveApp.getFolderById(ID_ROOT_DIR)
    let name = "Health Monitoring" + date.getFullYear()

    const files = root_dir.getFilesByName(name);
    
    if (files.hasNext()) {
        return SpreadsheetApp.open(files.next());
    } else {
        // AU LIEU DE CREATE, ON FAIT UNE COPIE DU TEMPLATE
        const templateFile = DriveApp.getFileById(ID_TEMPLATE_SS);
        const newFile = templateFile.makeCopy(name, root_dir);
        
        const newSS = SpreadsheetApp.open(newFile);
        
        // Optionnel : On peut renommer ou ajuster des choses ici
        return newSS;
    }
}

function getSheet(idSpreadSheet: string) {
    const date = new Date();
    const ss = SpreadsheetApp.openById(idSpreadSheet);
    let name = MONTH[date.getMonth()];
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
        sheet = ss.insertSheet().setName(name);
        const headers = ["Date", "Sommeil (h)", "Calories (kcal)", "Prot (g)", "Glu (g)", "Lip (g)", "Pas", "Poids (kg)"];

        // Utilisation de SUBTOTAL(109, ...) au lieu de SUM
        // 109 = SOMME des cellules visibles, ignorant les autres SUBTOTAL
        // 101 = MOYENNE des cellules visibles
        const sums = [
            "TOTAL MOIS", 
            "=SUBTOTAL(101; B3:B)", // Moyenne Sommeil
            "=SUBTOTAL(109; C3:C)", // Somme Calories
            "=SUBTOTAL(109; D3:D)",
            "=SUBTOTAL(109; E3:E)",
            "=SUBTOTAL(109; F3:F)",
            "=SUBTOTAL(109; G3:G)",
            "=SUBTOTAL(101; H3:H)"  // Moyenne Poids
        ];
        
        sheet.getRange(1, 1, 2, headers.length).setValues([headers, sums]);
        
        // Style (Header & Totaux)
        sheet.getRange(1, 1, 1, headers.length).setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');
        sheet.getRange(2, 1, 1, headers.length).setBackground('#f3f3f3').setFontWeight('bold');
        sheet.setFrozenRows(2);
    }
    return sheet;
}

function setData(idSpreadSheet: string, nameSheet: string, data: any[]) {
    const ss = SpreadsheetApp.openById(idSpreadSheet);
    const sheet = ss.getSheetByName(nameSheet);
    if (!sheet) return;

    const today = new Date();
    const rowToInsert = [today.toLocaleDateString('fr-FR'), ...data];
    
    // 1. Ajouter la donnée du jour
    sheet.appendRow(rowToInsert);

    // 2. Vérifier si on est dimanche (Jour 0) pour insérer le sous-total de la semaine
    // On le fait à la fin du dimanche ou si c'est le dernier jour du mois
    if (today.getDay() === 0) {
        insertWeeklySubtotal(sheet);
    }
}

function insertWeeklySubtotal(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const lastRow = sheet.getLastRow();
    
    // On crée une ligne de séparation
    const subtotalLabel = "TOTAL SEMAINE";
    const formulas = [
        subtotalLabel,
        "=SUBTOTAL(101; INDIRECT(\"B\"&ROW()-7&\":B\"&ROW()-1))", // Moyenne sommeil sur 7 jours
        "=SUBTOTAL(109; INDIRECT(\"C\"&ROW()-7&\":C\"&ROW()-1))", // Somme calories
        "=SUBTOTAL(109; INDIRECT(\"D\"&ROW()-7&\":D\"&ROW()-1))",
        "=SUBTOTAL(109; INDIRECT(\"E\"&ROW()-7&\":E\"&ROW()-1))",
        "=SUBTOTAL(109; INDIRECT(\"F\"&ROW()-7&\":F\"&ROW()-1))",
        "=SUBTOTAL(109; INDIRECT(\"G\"&ROW()-7&\":G\"&ROW()-1))",
        "=SUBTOTAL(101; INDIRECT(\"H\"&ROW()-7&\":H\"&ROW()-1))"
    ];

    sheet.appendRow(formulas);
    
    // Style pour la ligne de semaine
    const range = sheet.getRange(sheet.getLastRow(), 1, 1, formulas.length);
    range.setBackground('#d1e0f3').setFontWeight('bold').setFontStyle('italic');
    
    // Ajouter une ligne vide pour la clarté avant la semaine suivante
    sheet.appendRow([]); 
}