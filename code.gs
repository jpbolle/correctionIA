// ====================================
// FICHIER: Code.gs - VERSION FINALE
// ====================================

// Configuration des IDs des sheets
const CONFIG = {
  GRILLES_SHEET_ID: '1ZHrq0FGeNEX3rSqSCow_D0IR7IjzPOQCTBDPPNA0jUI', // Classeur unique
  DOSSIER_CORRECTIONS_ID: '1V-OMhH5sHR9hWV-PWsUGHkL5_K6sUfUw'
};

/**
 * Fonction principale pour servir l'interface HTML
 */
function doGet() {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    return template.evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('Application de Correction')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    console.error('Erreur dans doGet:', error);
    return HtmlService.createHtmlOutput(`
      <h1>Erreur de chargement</h1>
      <p>Erreur: ${error.message}</p>
      <p>Veuillez réessayer ou contacter l'administrateur.</p>
    `);
  }
}

/**
 * Inclure des fichiers CSS/JS dans le HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 🚀 OPTIMISÉ: Récupérer élèves ET types de production en un seul appel
 */
function getElevesEtTypes() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.GRILLES_SHEET_ID);
    
    // 1. Récupérer les élèves depuis la nouvelle feuille
    const elevesSheet = spreadsheet.getSheetByName('dataEleves');
    const elevesData = elevesSheet.getRange('A2:B60').getValues(); // Colonnes A et B, lignes 2-60
    
    const eleves = [];
    for (let i = 0; i < elevesData.length; i++) {
      const email = elevesData[i][0] ? elevesData[i][0].toString().trim() : '';
      const nom = elevesData[i][1] ? elevesData[i][1].toString().trim() : '';
      
      if (nom !== '') { // Le nom doit être présent
        eleves.push({
          nom: nom,        // Colonne B = nom affiché
          email: email     // Colonne A = email en mémoire
        });
      }
    }
    
    // 2. Récupérer les types de production (noms des autres feuilles)
    const sheets = spreadsheet.getSheets();
    const typesProduction = sheets
      .map(sheet => sheet.getName())
      .filter(name => name !== 'dataEleves'); // Exclure la feuille des élèves
    
    // 3. Retourner les deux en une fois !
    return {
      success: true,
      eleves: eleves.filter((eleve, index, self) => 
        index === self.findIndex(e => e.nom === eleve.nom)
      ),
      typesProduction: typesProduction
    };
    
  } catch (error) {
    console.error('Erreur lors de la récupération des données:', error);
    return {
      success: false,
      error: error.message,
      eleves: [],
      typesProduction: []
    };
  }
}

/**
 * 🚀 OPTIMISÉ: Récupérer la liste des élèves depuis le classeur unique
 */
function getEleves() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.GRILLES_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('dataEleves');
    
    // Lire les colonnes A (email) et B (nom) des lignes 2 à 60
    const data = sheet.getRange('A2:B60').getValues();
    
    const eleves = [];
    for (let i = 0; i < data.length; i++) {
      const email = data[i][0] ? data[i][0].toString().trim() : '';
      const nom = data[i][1] ? data[i][1].toString().trim() : '';
      
      if (nom !== '') { // Le nom doit être présent
        eleves.push({
          nom: nom,        // Colonne B = nom affiché
          email: email     // Colonne A = email en mémoire
        });
      }
    }
    
    // Supprimer les doublons basés sur le nom
    return eleves.filter((eleve, index, self) => 
      index === self.findIndex(e => e.nom === eleve.nom)
    );
  } catch (error) {
    console.error('Erreur lors de la récupération des élèves:', error);
    return [];
  }
}

/**
 * Récupérer les types de production (noms des feuilles) - OPTIMISÉ
 */
function getTypesProduction() {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.GRILLES_SHEET_ID);
    const sheets = spreadsheet.getSheets();
    
    return sheets
      .map(sheet => sheet.getName())
      .filter(name => name !== 'dataEleves'); // Exclure la feuille des élèves
  } catch (error) {
    console.error('Erreur lors de la récupération des types de production:', error);
    return [];
  }
}

/**
 * Récupérer la grille de correction pour un type de production donné
 */
function getGrilleCorrection(typeProduction) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.GRILLES_SHEET_ID).getSheetByName(typeProduction);
    if (!sheet) {
      throw new Error(`Feuille "${typeProduction}" non trouvée`);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('Données insuffisantes dans la feuille');
    }
    
    const headers = data[0]; // Ligne 1: en-têtes
    const grille = {
      indicateurs: headers.slice(1, 7), // Colonnes B à G
      criteres: []
    };
    
    // Traiter chaque ligne de critères (à partir de la ligne 2)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() !== '') {
        grille.criteres.push({
          nom: data[i][0].toString().trim(),
          indicateurs: data[i].slice(1, 7), // Colonnes B à G
          points: parseFloat(data[i][7]) || 0 // Colonne H
        });
      }
    }
    
    return grille;
  } catch (error) {
    console.error('Erreur lors de la récupération de la grille:', error);
    return null;
  }
}

/**
 * 🚀 OPTIMISÉ: Récupérer l'email d'un élève depuis la feuille dataEleves
 */
function getEmailEleve(nomEleve) {
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.GRILLES_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('dataEleves');
    const data = sheet.getDataRange().getValues();
    
    // Chercher la ligne correspondant à l'élève (colonne B = nom, colonne A = email)
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === nomEleve.trim()) {
        // Retourner l'email (colonne A)
        return data[i][0] ? data[i][0].toString().trim() : null;
      }
    }
    return null;
  } catch (error) {
    console.error('Erreur lors de la récupération de l\'email:', error);
    return null;
  }
}

/**
 * Enregistrer une correction
 */
function enregistrerCorrection(nomEleve, typeProduction, corrections, pointsTotal) {
  try {
    // Créer un nouveau spreadsheet
    const nomFichier = `${nomEleve} - ${typeProduction}`;
    const nouveauSheet = SpreadsheetApp.create(nomFichier);
    
    // Copier la structure de la grille modèle
    const grilleModele = SpreadsheetApp.openById(CONFIG.GRILLES_SHEET_ID);
    const feuilleModele = grilleModele.getSheetByName(typeProduction);
    
    if (feuilleModele) {
      const dataModele = feuilleModele.getDataRange().getValues();
      const feuillePrincipale = nouveauSheet.getActiveSheet();
      feuillePrincipale.setName(typeProduction);
      
      // Copier les données de base
      if (dataModele.length > 0) {
        feuillePrincipale.getRange(1, 1, dataModele.length, dataModele[0].length).setValues(dataModele);
      }
      
      // Ajouter les corrections avec mise en forme
      ajouterCorrections(feuillePrincipale, corrections, pointsTotal);
      
      // Ajouter les bordures finales
      ajouterBordures(feuillePrincipale);
    }
    
    // Déplacer vers le dossier spécifié
    const dossier = DriveApp.getFolderById(CONFIG.DOSSIER_CORRECTIONS_ID);
    const fichier = DriveApp.getFileById(nouveauSheet.getId());
    dossier.addFile(fichier);
    DriveApp.getRootFolder().removeFile(fichier);
    
    // Partager avec l'élève (si email disponible) - avec gestion d'erreur silencieuse
    const emailEleve = getEmailEleve(nomEleve);
    if (emailEleve && emailEleve.includes('@')) {
      try {
        nouveauSheet.addViewer(emailEleve);
        console.log(`Sheet partagé avec: ${emailEleve}`);
      } catch (emailError) {
        console.warn(`Impossible de partager avec ${emailEleve}:`, emailError);
      }
    } else {
      console.log(`Aucun email valide trouvé pour ${nomEleve}, pas de partage`);
    }
    
    return {
      success: true,
      url: nouveauSheet.getUrl(),
      message: `Correction enregistrée pour ${nomEleve}`
    };
    
  } catch (error) {
    console.error('Erreur lors de l\'enregistrement:', error);
    return {
      success: false,
      message: `Erreur: ${error.message}`
    };
  }
}

/**
 * Ajouter les corrections au sheet avec mise en forme - VERSION CORRIGÉE
 */
function ajouterCorrections(sheet, corrections, pointsTotal) {
  // Calculer le total maximum possible SEULEMENT pour les critères cotés
  let pointsMaxTotal = 0;
  Object.keys(corrections).forEach(critereIndex => {
    const correction = corrections[critereIndex];
    // ✅ SOLUTION 3: Ne compter que les critères effectivement cotés
    if (correction.selected !== null && correction.pointsMax) {
      pointsMaxTotal += correction.pointsMax;
    }
  });
  
  // Ajouter les colonnes pour les corrections
  const lastCol = sheet.getLastColumn();
  const colCorrection = lastCol + 1;
  const colPoints = lastCol + 2;
  
  // En-têtes
  sheet.getRange(1, colCorrection).setValue('Correction');
  sheet.getRange(1, colPoints).setValue('Points obtenus');
  
  // Mise en forme des en-têtes
  const headerRange = sheet.getRange(1, colCorrection, 1, 2);
  headerRange.setFontWeight('bold')
            .setFontSize(12)
            .setBackground('#1a73e8')
            .setFontColor('white')
            .setHorizontalAlignment('center');
  
  // Traiter chaque correction
  let ligne = 2;
  Object.keys(corrections).forEach(critereIndex => {
    const correction = corrections[critereIndex];
    
    if (correction.selected !== null && correction.indicateur) {
      // Cellule correction (colonne J)
      const celluleCorrection = sheet.getRange(ligne, colCorrection);
      celluleCorrection.setValue(correction.indicateur);
      
      // Cellule points (colonne K)  
      const cellulePoints = sheet.getRange(ligne, colPoints);
      cellulePoints.setValue(correction.points);
      
      // Déterminer les couleurs selon le pourcentage
      let couleurFond, couleurTexte;
      const pourcentage = correction.pourcentage || 0;
      
      if (pourcentage < 50) {
        // Néant, Très insuffisant, Insuffisant = Rouge
        couleurFond = '#f44336';
        couleurTexte = 'white';
      } else {
        // Suffisant, Acquis, Parfaitement acquis = Vert
        couleurFond = '#4caf50';
        couleurTexte = 'white';
      }
      
      // Appliquer la mise en forme à la cellule correction
      celluleCorrection.setBackground(couleurFond)
                      .setFontColor(couleurTexte)
                      .setFontWeight('bold')
                      .setWrap(true)
                      .setVerticalAlignment('middle');
      
      // Mise en forme spéciale pour la colonne points
      cellulePoints.setFontSize(14)
                   .setFontWeight('bold')
                   .setBackground(couleurFond)
                   .setFontColor(couleurTexte)
                   .setHorizontalAlignment('center')
                   .setVerticalAlignment('middle');
      
      // Surligner aussi la cellule de l'indicateur choisi dans les colonnes B-G
      const ligneActuelle = ligne;
      const colonnesIndicateurs = sheet.getRange(ligneActuelle, 2, 1, 6); // Colonnes B à G
      const valeurs = colonnesIndicateurs.getValues()[0];
      
      // Trouver la colonne qui contient l'indicateur sélectionné
      for (let colIndex = 0; colIndex < valeurs.length; colIndex++) {
        if (valeurs[colIndex] && valeurs[colIndex].toString().trim() === correction.indicateur.toString().trim()) {
          const celluleIndicateur = sheet.getRange(ligneActuelle, 2 + colIndex);
          celluleIndicateur.setBackground(couleurFond)
                           .setFontColor(couleurTexte)
                           .setFontWeight('bold')
                           .setWrap(true);
          break;
        }
      }
    }
    ligne++;
  });
  
  // Ajouter le total avec le format "points obtenus / points maximum"
  const ligneTotalCorrection = ligne + 1;
  const ligneTotalPoints = ligne + 1;
  
  sheet.getRange(ligneTotalCorrection, colCorrection).setValue('TOTAL');
  sheet.getRange(ligneTotalPoints, colPoints).setValue(`${pointsTotal.toFixed(1)} / ${pointsMaxTotal}`);
  
  // Mise en forme du total
  const totalCorrectionCell = sheet.getRange(ligneTotalCorrection, colCorrection);
  const totalPointsCell = sheet.getRange(ligneTotalPoints, colPoints);
  
  totalCorrectionCell.setFontWeight('bold')
                     .setFontSize(14)
                     .setBackground('#1a73e8')
                     .setFontColor('white')
                     .setHorizontalAlignment('center');
  
  totalPointsCell.setFontWeight('bold')
                 .setFontSize(16)
                 .setBackground('#1a73e8')
                 .setFontColor('white')
                 .setHorizontalAlignment('center');
  
  // Paramétrer toutes les colonnes
  configurerColonnes(sheet);
  
  // Ajouter les commentaires si présents - VERSION CORRIGÉE
  if (corrections.commentairesGeneraux && corrections.commentairesGeneraux.trim() !== '') {
    const ligneCommentaires = ligne + 3;
    
    // Ajouter le titre des commentaires
    sheet.getRange(ligneCommentaires, 1).setValue('Commentaires généraux:');
    sheet.getRange(ligneCommentaires, 1)
         .setFontWeight('bold')
         .setFontSize(12)
         .setBackground('#fff3e0');
    
    // ⚠️ SOLUTION: Utiliser plusieurs cellules individuelles au lieu de merger
    const ligneTexteCommentaire = ligneCommentaires + 1;
    const texteCommentaire = corrections.commentairesGeneraux.toString();
    
    // Placer le texte dans la première colonne seulement
    sheet.getRange(ligneTexteCommentaire, 1)
         .setValue(texteCommentaire)
         .setWrap(true)
         .setVerticalAlignment('top')
         .setBorder(true, true, true, true, true, true)
         .setBackground('#f8f9fa');
    
    // Élargir la première colonne pour les commentaires
    sheet.setColumnWidth(1, 400);
    
    // Ajouter une bordure aux cellules adjacentes pour l'esthétique (mais sans merger)
    for (let col = 2; col <= Math.min(6, sheet.getLastColumn()); col++) {
      sheet.getRange(ligneTexteCommentaire, col)
           .setBorder(true, true, true, true, true, true)
           .setBackground('#f8f9fa');
    }
  }
  
  // ✅ MAINTENANT on peut figer la colonne A en toute sécurité
  sheet.setFrozenColumns(1);
}

/**
 * Configurer les colonnes du sheet
 */
function configurerColonnes(sheet) {
  const lastCol = sheet.getLastColumn();
  
  // Définir la largeur de toutes les colonnes à 200px
  for (let col = 1; col <= lastCol; col++) {
    sheet.setColumnWidth(col, 200);
  }
  
  // Paramétrer le retour à la ligne pour toutes les cellules
  const dataRange = sheet.getDataRange();
  dataRange.setWrap(true)
           .setVerticalAlignment('middle');
  
  // Mise en forme spéciale pour la colonne A (critères)
  const critereColumn = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1);
  critereColumn.setFontWeight('bold')
               .setVerticalAlignment('middle')
               .setWrap(true);
  
  // Mise en forme pour les en-têtes des indicateurs (ligne 1, colonnes B-G)
  const headersRange = sheet.getRange(1, 2, 1, 6);
  headersRange.setFontWeight('bold')
              .setBackground('#e3f2fd')
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle')
              .setWrap(true);
  
  // Mise en forme pour la colonne des points max (colonne H)
  const pointsMaxColumn = sheet.getRange(1, 8, sheet.getLastRow(), 1);
  pointsMaxColumn.setFontWeight('bold')
                 .setBackground('#fff3e0')
                 .setHorizontalAlignment('center')
                 .setVerticalAlignment('middle');
}

/**
 * Ajouter des bordures au tableau
 */
function ajouterBordures(sheet) {
  const dataRange = sheet.getDataRange();
  dataRange.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}

// ========================================
// FONCTIONNALITÉ IA - PRÉCORRECTION
// ========================================

/**
 * 🤖 Fonction principale de précorrection IA
 */
function precorrectionIA(typeProduction, grilleData, copieEleve) {
  try {
    console.log('🤖 Début précorrection IA pour:', typeProduction);
    
    // Vérifications
    if (!typeProduction || !grilleData || !copieEleve) {
      return {
        success: false,
        message: 'Données manquantes pour la précorrection'
      };
    }
    
    if (!copieEleve.trim()) {
      return {
        success: false,
        message: 'Le texte de la copie ne peut pas être vide'
      };
    }
    
    // Formatter la grille pour l'IA
    const grilleFormatee = formaterGrillePourIA(grilleData);
    
    // Construire les prompts
    const promptSystem = construirePromptSystem();
    const promptUser = construirePromptUser(typeProduction, grilleFormatee, copieEleve);
    
    // Appel à l'API Claude
    const reponseIA = appellerAPIClaud(promptSystem, promptUser);
    
    if (reponseIA.success) {
      return {
        success: true,
        corrections: reponseIA.data.corrections,
        commentaire_final: reponseIA.data.commentaire_final,
        message: 'Précorrection IA terminée avec succès'
      };
    } else {
      return {
        success: false,
        message: reponseIA.message || 'Erreur lors de l\'appel à l\'IA'
      };
    }
    
  } catch (error) {
    console.error('Erreur précorrection IA:', error);
    return {
      success: false,
      message: `Erreur technique: ${error.message}`
    };
  }
}

/**
 * 📋 Formater la grille de correction pour l'IA
 */
function formaterGrillePourIA(grilleData) {
  let grilleTexte = '';
  
  grilleData.criteres.forEach((critere, index) => {
    grilleTexte += `\nCRITÈRE ${index} : "${critere.nom}" (Maximum: ${critere.points} points)\n`;
    grilleTexte += 'Indicateurs disponibles :\n';
    
    const pourcentages = [0, 15, 35, 60, 80, 100];
    const niveaux = ['Néant', 'Très insuffisant', 'Insuffisant', 'Suffisant', 'Acquis', 'Parfaitement acquis'];
    
    critere.indicateurs.forEach((indicateur, indIndex) => {
      if (indicateur && indicateur.toString().trim() !== '') {
        const pourcentage = pourcentages[indIndex] || 0;
        const niveau = niveaux[indIndex] || 'Inconnu';
        const points = (critere.points * pourcentage / 100).toFixed(1);
        
        grilleTexte += `  - "${indicateur}" (${niveau} - ${pourcentage}% = ${points} pts)\n`;
      }
    });
    grilleTexte += '\n';
  });
  
  return grilleTexte;
}

/**
 * 🎯 Construire le prompt system pour Claude - VERSION CORRIGÉE
 */
function construirePromptSystem() {
  return `Tu es un assistant pédagogique expert spécialisé dans l'évaluation éducative. Tu corriges les productions d'élèves avec bienveillance et précision selon des grilles de correction structurées.

CONSIGNES IMPORTANTES :
- Sois toujours bienveillant et encourageant dans tes commentaires
- Justifie chaque choix d'indicateur de manière pédagogique
- Propose des commentaires constructifs pour aider l'élève à progresser
- Respecte scrupuleusement la grille de correction fournie
- Adapte ton vocabulaire au niveau des élèves concernés
- Choisis pour chaque critère l'indicateur le plus approprié parmi ceux proposés

ATTENTION PARTICULIÈRE À L'ORTHOGRAPHE ET À LA PONCTUATION :
- Sois BEAUCOUP PLUS SÉVÈRE sur les erreurs d'orthographe, de grammaire et de ponctuation
- Ces aspects techniques de la langue doivent être rigoureusement évalués
- Une copie avec de nombreuses fautes ne peut PAS obtenir les niveaux "Acquis" ou "Parfaitement acquis" pour les critères linguistiques
- Même 3-4 fautes d'orthographe doivent faire descendre vers "Insuffisant" ou "Suffisant" maximum
- Distingue clairement les erreurs occasionnelles des erreurs systématiques

STRUCTURE DU COMPTE RENDU CRITIQUE - CRITÈRE "RÉSUMÉ" :
- Pour le critère "résumé", vérifie impérativement que l'élève respecte la structure attendue du compte rendu critique
- Le RÉSUMÉ doit être clairement distinct et séparé de la PARTIE CRITIQUE
- La structure correcte est : 1) RÉSUMÉ objectif en premier, puis 2) PARTIE CRITIQUE avec arguments personnels
- Si l'élève mélange résumé et critique, ou ne distingue pas ces deux parties, le critère "résumé" doit être pénalisé
- Un bon résumé est objectif, neutre et présente les idées principales sans donner d'opinion personnelle

CRITÈRES À ÉVALUER ABSOLUMENT :
- Tu DOIS OBLIGATOIREMENT évaluer TOUS les critères présents dans la grille, sans exception
- N'oublie JAMAIS d'évaluer le critère "Partie critique - Qualité persuasive des arguments" s'il est présent
- Ce critère est CRUCIAL pour les productions argumentatives et doit TOUJOURS recevoir une note
- Si un critère contient "persuasif", "argumentatif", "critique" ou "qualité des arguments", il est OBLIGATOIRE de l'évaluer
- Analyse avec attention la force des arguments, leur pertinence et leur capacité de persuasion

FORMAT DE RÉPONSE EXIGÉ :
Tu dois répondre EXCLUSIVEMENT en JSON valide avec cette structure exacte :
{
  "corrections": {
    "0": {
      "indicateur": "Texte exact de l'indicateur choisi",
      "pourcentage": 60,
      "justification": "Explication bienveillante du choix"
    },
    "1": {
      "indicateur": "Texte exact de l'indicateur choisi",
      "pourcentage": 80,
      "justification": "Explication bienveillante du choix"
    }
  },
  "commentaire_final": "Commentaire général encourageant et constructif qui souligne les points forts et propose des pistes d'amélioration concrètes"
}`;
}

/**
 * 👤 Construire le prompt user pour Claude - VERSION CORRIGÉE
 */
function construirePromptUser(typeProduction, grilleFormatee, copieEleve) {
  return `GRILLE DE CORRECTION :
Type de production : ${typeProduction}

CRITÈRES À ÉVALUER :
${grilleFormatee}

COPIE DE L'ÉLÈVE À CORRIGER :
${copieEleve}

CONSIGNES SPÉCIFIQUES OBLIGATOIRES :
1. Corrige cette copie selon la grille fournie
2. Pour chaque critère, choisis l'indicateur le plus approprié et justifie ton choix avec bienveillance
3. CRITÈRE "RÉSUMÉ" : Vérifie impérativement que l'élève distingue bien le résumé (objectif, neutre) de la partie critique (arguments personnels). La structure doit être : 1) Résumé d'abord, 2) Partie critique ensuite. Pénalise si ces parties sont mélangées ou confondues.
4. CRITÈRE "PARTIE CRITIQUE - QUALITÉ PERSUASIVE" : Tu DOIS ABSOLUMENT évaluer ce critère s'il est présent. C'est OBLIGATOIRE. Analyse la force, la pertinence et la capacité de persuasion des arguments.
5. ORTHOGRAPHE/PONCTUATION : Sois BEAUCOUP PLUS SÉVÈRE. Même 3-4 fautes doivent empêcher d'atteindre "Acquis" ou "Parfaitement acquis". Pénalise davantage les erreurs techniques.
6. ÉVALUER TOUS LES CRITÈRES : N'oublie aucun critère de la grille. Chaque critère doit recevoir une évaluation.
7. Termine par un commentaire final encourageant qui souligne les points forts et propose des pistes d'amélioration concrètes

VÉRIFICATION FINALE : Assure-toi d'avoir évalué CHAQUE critère de la grille, en particulier ceux contenant "persuasif", "critique" ou "arguments".

Réponds uniquement en JSON selon le format spécifié.`;
}

/**
 * 🚀 Appeler l'API Claude d'Anthropic
 */
function appellerAPIClaud(promptSystem, promptUser) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    
    if (!apiKey) {
      return {
        success: false,
        message: 'Clé API Claude non configurée. Ajoutez CLAUDE_API_KEY dans les propriétés du script.'
      };
    }
    
    const url = 'https://api.anthropic.com/v1/messages';
    
    const payload = {
      model: 'claude-3-5-sonnet-20241022',
      max_tokens: 4000,
      system: promptSystem,
      messages: [
        {
          role: 'user',
          content: promptUser
        }
      ]
    };
    
    const options = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify(payload)
    };
    
    console.log('🚀 Envoi requête vers API Claude...');
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('📡 Réponse API Claude - Code:', responseCode);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      
      if (data.content && data.content[0] && data.content[0].text) {
        const contenuIA = data.content[0].text;
        
        try {
          // Parser la réponse JSON de Claude
          const reponseIA = JSON.parse(contenuIA);
          
          // Vérifier la structure
          if (reponseIA.corrections && reponseIA.commentaire_final) {
            return {
              success: true,
              data: reponseIA
            };
          } else {
            return {
              success: false,
              message: 'Format de réponse IA invalide'
            };
          }
          
        } catch (parseError) {
          console.error('Erreur parsing JSON IA:', parseError);
          console.log('Contenu reçu:', contenuIA);
          
          return {
            success: false,
            message: 'Réponse IA non parsable. Contenu: ' + contenuIA.substring(0, 200) + '...'
          };
        }
        
      } else {
        return {
          success: false,
          message: 'Réponse API vide ou incorrecte'
        };
      }
      
    } else {
      console.error('Erreur API Claude:', responseCode, responseText);
      
      let messageErreur = 'Erreur API Claude';
      
      try {
        const errorData = JSON.parse(responseText);
        if (errorData.error && errorData.error.message) {
          messageErreur = errorData.error.message;
        }
      } catch (e) {
        // Ignore parse error
      }
      
      return {
        success: false,
        message: `${messageErreur} (Code: ${responseCode})`
      };
    }
    
  } catch (error) {
    console.error('Erreur appel API Claude:', error);
    return {
      success: false,
      message: `Erreur technique: ${error.message}`
    };
  }
}

/**
 * 🔧 Fonction de test pour vérifier la configuration API
 */
function testerAPIClaud() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  
  if (!apiKey) {
    console.log('❌ Clé API Claude non configurée');
    return false;
  }
  
  // Prompt de test plus strict
  const promptSystemTest = `Tu es un assistant de test. Tu dois répondre EXCLUSIVEMENT en JSON valide avec cette structure exacte :
{
  "corrections": {},
  "commentaire_final": "Test de connexion réussi"
}`;
  
  const promptUserTest = `Test de connexion API. Réponds uniquement en JSON selon le format spécifié.`;
  
  console.log('🧪 Test de connexion API Claude...');
  
  const resultat = appellerAPIClaud(promptSystemTest, promptUserTest);
  
  if (resultat.success) {
    console.log('✅ API Claude fonctionnelle');
    console.log('📋 Réponse reçue:', JSON.stringify(resultat.data));
    return true;
  } else {
    console.log('❌ Erreur API Claude:', resultat.message);
    return false;
  }
}

/**
 * 🔍 Fonction de debug pour vérifier la clé API
 */
function verifierCleAPI() {
  const cle = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  console.log('Clé trouvée:', cle ? 'OUI' : 'NON');
  if (cle) {
    console.log('Début de la clé:', cle.substring(0, 10) + '...');
  }
}
