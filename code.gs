// Code.gs - Côté serveur de l'application de saisie d'aléas

// Fonction doGet requise pour le déploiement web
function doGet(e) {
  try {
    Logger.log("Début de l'exécution de doGet()");
    
    // Générer le HTML du formulaire d'aléas
    var htmlOutput = HtmlService.createHtmlOutputFromFile('FormulaireAleas')
      .setTitle('Formulaire de saisie des aléas');
    
    // Configurer les options de sécurité
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    Logger.log("Fin de l'exécution de doGet() - succès");
    return htmlOutput;
    
  } catch (e) {
    Logger.log("ERREUR dans doGet : " + e.toString());
    Logger.log("Stack trace : " + e.stack);
    
    // Créer une page d'erreur
    var htmlOutput = HtmlService.createHtmlOutput(
      '<html><body>' +
      '<h1>Erreur</h1>' +
      '<p>Une erreur s\'est produite : ' + e.toString() + '</p>' +
      '<p><a href="javascript:history.back()">Retour</a></p>' +
      '</body></html>'
    );
    
    htmlOutput.setTitle('Erreur');
    return htmlOutput;
  }
}

// Fonction globale appelée lors de l'ouverture de la feuille
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Aléas')
    .addItem('Ouvrir le formulaire de saisie', 'ouvrirFormulaireAleas')
    .addToUi();
}

// Fonction qui ouvre l'application web
function ouvrirFormulaireAleas() {
  var html = HtmlService.createHtmlOutputFromFile('FormulaireAleas')
    .setWidth(800)
    .setHeight(600)
    .setTitle('Formulaire de saisie des aléas');
  SpreadsheetApp.getUi().showModalDialog(html, 'Formulaire de saisie des aléas');
}

// Fonction pour récupérer les infos d'un opérateur à partir de son SA
function getOperateurInfoBySA(sa) {
  try {
    Logger.log("Recherche de l'opérateur avec SA: " + sa);
    
    // Vérifier que le SA n'est pas vide
    if (!sa || sa.trim() === "") {
      return { success: false, message: "Veuillez saisir un numéro SA" };
    }
    
    var spreadsheet = SpreadsheetApp.openById('1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8');
    var sheet = spreadsheet.getSheetByName('Effectif');
    
    if (!sheet) {
      Logger.log("Feuille 'Effectif' introuvable");
      return { success: false, message: "Feuille 'Effectif' introuvable" };
    }
    
    // Récupérer toutes les données
    var data = sheet.getDataRange().getValues();
    
    // Trouver l'index des colonnes SA, Nom et Équipe
    var headers = data[0];
    var saIndex = -1, nomIndex = -1, equipeIndex = -1, posteIndex = -1;
    
    // Recherche case-insensitive des colonnes
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i]).toLowerCase().trim();
      if (header === "sa" || header === "numéro sa" || header === "numero sa") {
        saIndex = i;
      } else if (header === "nom" || header === "opérateur" || header === "operateur") {
        nomIndex = i;
      } else if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      } else if (header === "poste" || header === "poste habituel") {
        posteIndex = i;
      }
    }
    
    if (saIndex === -1 || nomIndex === -1) {
      Logger.log("Colonnes requises non trouvées: SA=" + saIndex + ", Nom=" + nomIndex);
      return { 
        success: false, 
        message: "Colonnes requises (SA, Nom) non trouvées dans la feuille Effectif" 
      };
    }
    
    // Normaliser le numéro SA recherché
    var saNormalise = String(sa).trim().toLowerCase();
    
    // Chercher l'opérateur correspondant au SA
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Vérifier que la cellule contient quelque chose
      if (row[saIndex] !== null && row[saIndex] !== undefined) {
        var currentSA = String(row[saIndex]).trim().toLowerCase();
        
        if (currentSA === saNormalise) {
          Logger.log("Opérateur trouvé: " + row[nomIndex]);
          return {
            success: true,
            operateur: row[nomIndex] || "",
            equipe: equipeIndex >= 0 ? (row[equipeIndex] || "") : "",
            poste: posteIndex >= 0 ? (row[posteIndex] || "") : ""
          };
        }
      }
    }
    
    Logger.log("Aucun opérateur trouvé avec le SA: " + sa);
    return { success: false, message: "Aucun opérateur trouvé avec ce numéro SA" };
    
  } catch (e) {
    Logger.log("ERREUR dans getOperateurInfoBySA: " + e.toString());
    return { success: false, message: "Erreur: " + e.toString() };
  }
}

// Fonction pour enregistrer un aléa
function saveAlea(aleaData) {
  try {
    Logger.log("Enregistrement d'un nouvel aléa: " + JSON.stringify(aleaData));
    
    var spreadsheet = SpreadsheetApp.openById('1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8');
    var sheet = spreadsheet.getSheetByName('Aleas');
    
    if (!sheet) {
      Logger.log("Feuille 'Aleas' introuvable");
      return { success: false, message: "Feuille 'Aleas' introuvable" };
    }
    
    // Convertir la durée en nombre (si fournie)
    var duree = null;
    if (aleaData.duree !== undefined && aleaData.duree !== null && aleaData.duree !== "") {
      duree = parseFloat(aleaData.duree);
      // Vérifier que la conversion a bien fonctionné
      if (isNaN(duree)) {
        Logger.log("Erreur de conversion de la durée: " + aleaData.duree);
        duree = null;
      }
    }
    
    // Traiter la date correctement
    var dateSaisie;
    try {
      // Récupérer les composants de la date (format YYYY-MM-DD)
      Logger.log("Date à convertir: " + aleaData.date);
      var dateParts = aleaData.date.split('-');
      
      if (dateParts.length !== 3) {
        Logger.log("Format de date incorrect, utilisera la date actuelle");
        dateSaisie = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else {
        // Créer une chaîne de date au format souhaité
        dateSaisie = dateParts[2] + "/" + dateParts[1] + "/" + dateParts[0]; // Format "DD/MM/YYYY"
        Logger.log("Date formatée: " + dateSaisie);
      }
    } catch (e) {
      Logger.log("Erreur lors de la conversion de la date: " + e);
      dateSaisie = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    
    // Créer le tableau de données à insérer
    var rowData = [
      aleaData.operateur,
      aleaData.sa,
      aleaData.equipe,
      aleaData.poste,
      dateSaisie, // Chaîne de date au format DD/MM/YYYY
      aleaData.typeAlea,
      aleaData.commentaire,
      duree  // Colonne H - Durée de l'aléa
    ];
    
    // Ajouter la nouvelle ligne
    var newRow = sheet.appendRow(rowData);
    
    Logger.log("Date finale enregistrée: " + dateSaisie);
    
    Logger.log("Aléa enregistré avec succès");
    return { success: true, message: "Aléa enregistré avec succès" };
    
  } catch (e) {
    Logger.log("ERREUR dans saveAlea: " + e.toString());
    return { success: false, message: "Erreur lors de l'enregistrement: " + e.toString() };
  }
}

// Fonction pour récupérer la liste des postes (pour auto-complétion)
function getListePostes() {
  try {
    var spreadsheet = SpreadsheetApp.openById('1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8');
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Historique rendement' introuvable" };
    }
    
    // Récupérer les en-têtes
    var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    var posteIndex = headers.indexOf("Poste habituel");
    
    if (posteIndex === -1) {
      return { success: false, message: "Colonne 'Poste habituel' introuvable" };
    }
    
    // Récupérer toutes les valeurs de poste
    var postesCol = sheet.getRange(3, posteIndex + 1, sheet.getLastRow() - 2, 1).getValues();
    var postes = [];
    
    // Filtrer les valeurs uniques et non vides
    postesCol.forEach(function(row) {
      var poste = row[0];
      if (poste && postes.indexOf(poste) === -1) {
        postes.push(poste);
      }
    });
    
    // Trier par ordre alphabétique
    postes.sort();
    
    return { success: true, postes: postes };
    
  } catch (e) {
    Logger.log("ERREUR dans getListePostes: " + e.toString());
    return { success: false, message: "Erreur: " + e.toString() };
  }
}

// Fonction pour tester la connexion
function testConnection() {
  try {
    return {
      success: true,
      message: "Connexion réussie",
      timestamp: new Date().toString(),
      user: Session.getEffectiveUser().getEmail()
    };
  } catch (e) {
    return {
      success: false,
      message: "Erreur lors du test de connexion: " + e.toString(),
      error: e.toString()
    };
  }
}
// Fonction pour récupérer la liste des types d'aléas
function getTypesAleas() {
  try {
    var spreadsheet = SpreadsheetApp.openById('1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8');
    var sheet = spreadsheet.getSheetByName('Type aléas');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Type aléas' introuvable" };
    }
    
    // Chercher l'index de la colonne "Aléa"
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var aleaIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).toLowerCase().trim() === "aléa" || 
          String(headers[i]).toLowerCase().trim() === "alea") {
        aleaIndex = i;
        break;
      }
    }
    
    if (aleaIndex === -1) {
      return { success: false, message: "Colonne 'Aléa' introuvable dans la feuille 'Type aléas'" };
    }
    
    // Récupérer toutes les valeurs de la colonne "Aléa"
    var aleasCol = sheet.getRange(2, aleaIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    var types = [];
    
    // Filtrer les valeurs uniques et non vides
    aleasCol.forEach(function(row) {
      var typeAlea = row[0];
      if (typeAlea && typeof typeAlea === 'string' && typeAlea.trim() !== '' && types.indexOf(typeAlea) === -1) {
        types.push(typeAlea);
      }
    });
    
    // Trier par ordre alphabétique
    types.sort();
    
    return { success: true, types: types };
    
  } catch (e) {
    Logger.log("ERREUR dans getTypesAleas: " + e.toString());
    return { success: false, message: "Erreur: " + e.toString() };
  }
}

function getAllOperateurs() {
  try {
    var spreadsheet = SpreadsheetApp.openById('1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8');
    var sheet = spreadsheet.getSheetByName('Effectif');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Effectif' introuvable" };
    }
    
    // Récupérer toutes les données
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Trouver les index des colonnes importantes
    var saIndex = -1, nomIndex = -1, equipeIndex = -1, posteIndex = -1;
    
    // Recherche case-insensitive des colonnes
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i]).toLowerCase().trim();
      if (header === "sa" || header === "numéro sa" || header === "numero sa") {
        saIndex = i;
      } else if (header === "nom" || header === "opérateur" || header === "operateur") {
        nomIndex = i;
      } else if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      } else if (header === "poste" || header === "poste habituel") {
        posteIndex = i;
      }
    }
    
    if (saIndex === -1 || nomIndex === -1) {
      return { 
        success: false, 
        message: "Colonnes requises (SA, Nom) non trouvées dans la feuille Effectif" 
      };
    }
    
    // Préparer la liste des opérateurs
    var operateurs = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Vérifier que les cellules contiennent quelque chose
      if (row[nomIndex] !== null && row[nomIndex] !== undefined && row[saIndex] !== null && row[saIndex] !== undefined) {
        var nom = String(row[nomIndex]).trim();
        var sa = String(row[saIndex]).trim();
        
        if (nom && sa) {
          operateurs.push({
            nom: nom,
            sa: sa,
            equipe: equipeIndex >= 0 ? (row[equipeIndex] || "") : "",
            poste: posteIndex >= 0 ? (row[posteIndex] || "") : ""
          });
        }
      }
    }
    
    // Trier par ordre alphabétique du nom
    operateurs.sort(function(a, b) {
      return a.nom.localeCompare(b.nom);
    });
    
    return {
      success: true,
      operateurs: operateurs
    };
    
  } catch (e) {
    Logger.log("ERREUR dans getAllOperateurs: " + e.toString());
    return { success: false, message: "Erreur: " + e.toString() };
  }
}
