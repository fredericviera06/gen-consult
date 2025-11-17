// ============================================================
// PROJET CORO - VERSION OPTIMISÉE COMPLÈTE V 2 - 21
// ============================================================

// ========== CONFIGURATION GLOBALE ==========
const FORM_ID = '1StRiqDtHBkqTB_jZ1vttYiQ7T2CI4qjtrDVAmXZzc8w';
const SHEET_NAME_DATA = 'Data';
const SHEET_NAME_URLS = 'editResponseUrl'; // Feuille où on stocke les liens
const COL_Z = 26; // Colonne Z où on lit le type ET où on met le lien du document

// Configuration des différents types de documents
const CONFIG = {
  hospit: {
    templateId: '1_sTiRD_jC3S-smQdhuL5W9ZGL-vetGhVuq7hfE3iVk0',
    folderId: '1FbMO_uoyA1q0WTi20etPEcw8Lm3SKPJV',
    suffix: 'hospitalisation',
    fields: {
      '{{date}}': 2, '{{genre}}': 3, '{{nom}}': 4, '{{DN}}': 5,
      '{{motif}}': 6, '{{antecedents}}': 7, '{{traitement}}': 8,
      '{{FDRCV}}': 9, '{{fonctionnel}}': 10, '{{TA}}': 11,
      '{{poids}}': 12, '{{taille}}': 13, '{{EC}}': 14,
      '{{ECG}}': 15, '{{ETT}}': 16, '{{lipide}}': 17,
      '{{au total}}': 18, '{{modif ttt}}': 19, '{{suivi}}': 20,
      '{{ECO}}': 21, '{{operateur}}': 22, '{{HDM}}': 41
    }
  },
  ett: {
    templateId: '1_sTiRD_jC3S-smQdhuL5W9ZGL-vetGhVuq7hfE3iVk0',
    folderId: '1FbMO_uoyA1q0WTi20etPEcw8Lm3SKPJV',
    suffix: 'ETT',
    fields: {
      '{{date}}': 3, '{{genre}}': 4, '{{nom}}': 5, '{{DN}}': 6,
      '{{motif ETT}}': 7, '{{antecedents}}': 8, '{{traitement}}': 9,
      '{{FDRCV}}': 10, '{{fonctionnel}}': 11, '{{TA}}': 12,
      '{{poids}}': 13, '{{taille}}': 14, '{{EC}}': 15,
      '{{ECG}}': 16, '{{ETT}}': 17, '{{lipide}}': 18,
      '{{au total}}': 19, '{{modif ttt}}': 20, '{{suivi}}': 21,
      '{{ECO}}': 22, '{{operateur}}': 23
    }
  },
  coro: {
    templateId: '1nlYqN8U5GyrQ7BuFjN2cphUnOH8EMSt3xN-FJGYNRf4',
    folderId: '1j9yRuXBe5tN3AMd4QsyvxHm0jZn_lC4I',
    suffix: 'coro',
    fields: {
      '{{date}}': 2, '{{genre}}': 3, '{{nom}}': 4, '{{DN}}': 5,
      '{{antecedents}}': 7, '{{traitement}}': 8, '{{FDRCV}}': 9,
      '{{fonctionnel}}': 10, '{{indic}}': 26, '{{ETT}}': 27,
      '{{TA}}': 28, '{{EC}}': 29, '{{ECG}}': 40, '{{voie}}': 30,
      '{{dom}}': 31, '{{simpl}}': 32, '{{fermeture}}': 33,
      '{{suites}}': 34, '{{total}}': 35, '{{sortie}}': 36,
      '{{compl}}': 37, '{{code}}': 38
    }
  },
  angio: {
    templateId: '1KLoC1JhyjgySoZMErf6m7nG98XHkaWC_zG1u5o8HEvM',
    folderId: '1QT1MWkeRiOoJsBOU3TkLG_WzNhg7EsJJ',
    suffix: 'angio',
    fields: {
      '{{date}}': 2, '{{genre}}': 3, '{{nom}}': 4, '{{DN}}': 5,
      '{{antecedents}}': 7, '{{traitement}}': 8, '{{FDRCV}}': 9,
      '{{fonctionnel}}': 10, '{{motif}}': 42, '{{ETT}}': 27,
      '{{TA}}': 28, '{{EC}}': 29, '{{ECG}}': 40, '{{voie}}': 44,
      '{{fermeture}}': 45, '{{suites}}': 46, '{{sortie}}': 36,
      '{{compl}}': 37, '{{code}}': 47, '{{angio}}': 43, '{{simpl}}': 32
    }
  },
  postcoro: {
    templateId: '1Rfm4nMjBp_3eWN0xFm0etTz69s0vmqq-38eh1gb9xDA',
    folderId: '1Pwt3ixf2bes3a344X9JqOIpDtLDW2QYX',
    suffix: 'post coro',
    fields: {
      '{{date}}': 2, '{{genre}}': 3, '{{nom}}': 4, '{{DN}}': 5,
      '{{antecedents}}': 7, '{{FDRCV}}': 9, '{{indic}}': 26,
      '{{simpl}}': 32, '{{suites}}': 34, '{{angio}}': 43,
      '{{radiale}}': 51, '{{educ}}': 52, '{{ECP}}': 53,
      '{{tropo}}': 54, '{{date atc}}': 55
    }
  }
};

// ============================================================
// SECTION 1 : GÉNÉRATION DE DOCUMENTS
// ============================================================

/**
 * Fonction générique pour générer des documents depuis un template
 */
function generateDocuments(configKey, scanAll = false) {
  const config = CONFIG[configKey];
  if (!config) {
    throw new Error(`Configuration inconnue: ${configKey}`);
  }
  
  Logger.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
  Logger.log(`🚀 Début génération pour: ${configKey}`);
  Logger.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
  
  try {
    const template = DriveApp.getFileById(config.templateId);
    const folder = DriveApp.getFolderById(config.folderId);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_DATA);
    
    const lastRow = sheet.getLastRow();
    let startRow, numRows;
    
    if (scanAll) {
      startRow = 2;
      numRows = lastRow - 1;
      Logger.log(`📊 MODE COMPLET : Scan de TOUTES les ${numRows} lignes`);
    } else {
      startRow = Math.max(2, lastRow - 499);
      numRows = lastRow - startRow + 1;
      Logger.log(`📊 MODE RAPIDE : Scan des ${numRows} dernières lignes`);
    }
    
    const rows = sheet.getRange(startRow, 1, numRows, sheet.getMaxColumns()).getValues();
    
    let processed = 0;
    const colZIndex = COL_Z - 1;
    
    rows.forEach((row, index) => {
      const actualRowNumber = startRow + index;
      const cellValue = row[colZIndex] ? row[colZIndex].toString().trim() : '';
      
      if (cellValue.startsWith('http')) return;
      
      const cellValueLower = cellValue.toLowerCase();
      const shouldProcess = checkIfShouldProcess(cellValueLower, configKey);
      
      if (!shouldProcess) return;
      
      Logger.log(`\nLigne ${actualRowNumber}: "${cellValue}" → À traiter`);
      
      try {
        const fileName = `${row[2]} - ${row[4]} ${config.suffix}`;
        const copy = template.makeCopy(fileName, folder);
        const doc = DocumentApp.openById(copy.getId());
        const body = doc.getBody();
        
        Object.entries(config.fields).forEach(([token, colIndex]) => {
          const value = row[colIndex] !== undefined && row[colIndex] !== null ? row[colIndex] : '';
          body.replaceText(token, value.toString());
        });
        
        doc.saveAndClose();
        const docUrl = doc.getUrl();
        
        sheet.getRange(actualRowNumber, COL_Z).setValue(docUrl);
        
        processed++;
        Logger.log(`  ✅ Document généré: ${fileName}`);
        
      } catch (rowErr) {
        Logger.log(`  ❌ Erreur ligne ${actualRowNumber}: ${rowErr.message}`);
      }
    });
    
    Logger.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    Logger.log(`📊 RÉSULTAT: ${processed} document(s) généré(s)`);
    Logger.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    
    if (processed > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `${processed} document(s) généré(s) avec succès`, 
        '✅ Terminé', 
        5
      );
    } else {
      const scope = scanAll ? 'toute la feuille' : `les ${numRows} dernières lignes`;
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Aucune ligne à traiter dans ${scope}.`, 
        'ℹ️ Information', 
        5
      );
    }
    
    return processed;
    
  } catch (err) {
    Logger.log(`❌ ERREUR GLOBALE: ${err.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Erreur: ${err.message}`, 
      '❌ Échec', 
      10
    );
    throw err;
  }
}

/**
 * Vérifie si une ligne doit être traitée selon le contenu de la colonne Z
 */
function checkIfShouldProcess(cellValue, configKey) {
  if (!cellValue) return false;
  
  switch(configKey) {
    case 'coro':
      return cellValue.includes('coro') && !cellValue.includes('post');
    case 'postcoro':
      return cellValue.includes('post') && cellValue.includes('coro');
    case 'angio':
      return cellValue.includes('angio');
    case 'ett':
      return cellValue.includes('ett');
    case 'hospit':
      return cellValue.includes('hospit');
    default:
      return false;
  }
}

// Fonctions wrapper pour le menu
function createNewGoogleDocs() { generateDocuments('hospit'); }
function generateETT() { generateDocuments('ett'); }
function generateCoro() { generateDocuments('coro'); }
function generateAngio() { generateDocuments('angio'); }
function generatePostCoro() { generateDocuments('postcoro'); }

// ============================================================
// SECTION 2 : GÉNÉRATION PAR LOT INTELLIGENTE
// ============================================================

/**
 * Génération intelligente par lot - Traite TOUS les types automatiquement
 */
function generateAllPending() {
  const startTime = new Date();
  const ui = SpreadsheetApp.getUi();
  
  const scanModeResponse = ui.alert(
    '🔍 Mode de recherche',
    'Comment voulez-vous rechercher les documents à générer ?\n\n' +
    '• MODE RAPIDE : Scan des 500 dernières lignes (~5 sec)\n' +
    '  → Pour les nouveaux patients récents\n\n' +
    '• MODE COMPLET : Scan de TOUTES les lignes (~30-45 sec)\n' +
    '  → Pour inclure les anciens patients qui reviennent',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (scanModeResponse === ui.Button.CANCEL) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Génération annulée', 'ℹ️ Annulé', 3);
    return;
  }
  
  const scanAll = (scanModeResponse === ui.Button.NO);
  const modeName = scanAll ? 'COMPLET (toutes les lignes)' : 'RAPIDE (500 dernières)';
  
  const types = [
    { key: 'coro', name: 'Coro' },
    { key: 'angio', name: 'Angio' },
    { key: 'postcoro', name: 'Post Coro' },
    { key: 'ett', name: 'ETT' },
    { key: 'hospit', name: 'Hospit' }
  ];
  
  const results = [];
  let totalProcessed = 0;
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Analyse en cours (mode ${scanAll ? 'complet' : 'rapide'})...`, 
    '🔍 Scan des documents à générer', 
    3
  );
  
  const stats = analyzePendingDocuments(scanAll);
  
  if (stats.total === 0) {
    ui.alert(
      'Aucun document à générer', 
      `Aucune ligne ne contient "coro", "angio", "ett", "post coro" ou "hospit" en colonne Z.\n\n` +
      `Portée : ${stats.scope}`,
      ui.ButtonSet.OK
    );
    return;
  }
  
  const response = ui.alert(
    '📋 Documents à générer',
    `Mode : ${modeName}\n` +
    `Portée : ${stats.scope}\n\n` +
    `Documents trouvés :\n\n` +
    `• Coro : ${stats.coro}\n` +
    `• Angio : ${stats.angio}\n` +
    `• Post Coro : ${stats.postcoro}\n` +
    `• ETT : ${stats.ett}\n` +
    `• Hospit : ${stats.hospit}\n\n` +
    `TOTAL : ${stats.total} document(s)\n\n` +
    `Voulez-vous continuer ?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Génération annulée', 'ℹ️ Annulé', 3);
    return;
  }
  
  types.forEach(type => {
    if (stats[type.key] > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Génération ${type.name}...`, 
        `⏳ ${type.name}`, 
        2
      );
      
      try {
        const count = generateDocumentsSilent(type.key, scanAll);
        if (count > 0) {
          results.push(`✅ ${type.name}: ${count} document(s)`);
          totalProcessed += count;
        }
      } catch (err) {
        results.push(`❌ ${type.name}: Erreur - ${err.message}`);
        Logger.log(`Erreur ${type.name}: ${err.message}`);
      }
    }
  });
  
  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(1);
  
  if (totalProcessed > 0) {
    ui.alert(
      '✅ Génération terminée',
      `Mode : ${modeName}\n` +
      `${totalProcessed} document(s) généré(s) en ${duration}s\n\n` +
      results.join('\n'),
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'ℹ️ Aucun document généré',
      'Tous les documents détectés ont déjà été traités.',
      ui.ButtonSet.OK
    );
  }
}

function analyzePendingDocuments(scanAll = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_DATA);
  const lastRow = sheet.getLastRow();
  
  let startRow, numRows;
  
  if (scanAll) {
    startRow = 2;
    numRows = lastRow - 1;
  } else {
    startRow = Math.max(2, lastRow - 499);
    numRows = lastRow - startRow + 1;
  }
  
  const rows = sheet.getRange(startRow, 1, numRows, sheet.getMaxColumns()).getValues();
  const colZIndex = COL_Z - 1;
  
  const stats = {
    coro: 0, angio: 0, postcoro: 0, ett: 0, hospit: 0, total: 0,
    scope: scanAll ? 'toutes les lignes' : `${numRows} dernières lignes`
  };
  
  rows.forEach(row => {
    const cellValue = row[colZIndex] ? row[colZIndex].toString().trim() : '';
    if (cellValue.startsWith('http')) return;
    
    const cellValueLower = cellValue.toLowerCase();
    
    if (cellValueLower.includes('post') && cellValueLower.includes('coro')) {
      stats.postcoro++; stats.total++;
    } else if (cellValueLower.includes('coro')) {
      stats.coro++; stats.total++;
    } else if (cellValueLower.includes('angio')) {
      stats.angio++; stats.total++;
    } else if (cellValueLower.includes('ett')) {
      stats.ett++; stats.total++;
    } else if (cellValueLower.includes('hospit')) {
      stats.hospit++; stats.total++;
    }
  });
  
  return stats;
}

function generateDocumentsSilent(configKey, scanAll = false) {
  const config = CONFIG[configKey];
  if (!config) throw new Error(`Configuration inconnue: ${configKey}`);
  
  const template = DriveApp.getFileById(config.templateId);
  const folder = DriveApp.getFolderById(config.folderId);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_DATA);
  
  const lastRow = sheet.getLastRow();
  const startRow = scanAll ? 2 : Math.max(2, lastRow - 499);
  const numRows = scanAll ? lastRow - 1 : lastRow - startRow + 1;
  
  const rows = sheet.getRange(startRow, 1, numRows, sheet.getMaxColumns()).getValues();
  
  let processed = 0;
  const colZIndex = COL_Z - 1;
  
  rows.forEach((row, index) => {
    const actualRowNumber = startRow + index;
    const cellValue = row[colZIndex] ? row[colZIndex].toString().trim() : '';
    
    if (cellValue.startsWith('http')) return;
    
    const cellValueLower = cellValue.toLowerCase();
    if (!checkIfShouldProcess(cellValueLower, configKey)) return;
    
    try {
      const fileName = `${row[2]} - ${row[4]} ${config.suffix}`;
      const copy = template.makeCopy(fileName, folder);
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();
      
      Object.entries(config.fields).forEach(([token, colIndex]) => {
        const value = row[colIndex] !== undefined && row[colIndex] !== null ? row[colIndex] : '';
        body.replaceText(token, value.toString());
      });
      
      doc.saveAndClose();
      sheet.getRange(actualRowNumber, COL_Z).setValue(doc.getUrl());
      processed++;
      
    } catch (rowErr) {
      Logger.log(`❌ ${configKey} - Ligne ${actualRowNumber}: ${rowErr.message}`);
    }
  });
  
  return processed;
}

// ============================================================
// SECTION 3 : SYNCHRONISATION DES LIENS D'ÉDITION
// Les liens sont stockés dans la feuille 'editResponseUrl'
// et rappatriés en colonne BL via ARRAYFORMULA
// ============================================================

/**
 * Déclencheur automatique à l'envoi du formulaire
 * Ajoute le lien d'édition dans la feuille editResponseUrl
 * NOTE: Renommée en onFormSubmitSync pour éviter conflit avec le script du formulaire
 */
function onFormSubmitSync(e) {
  const maxRetries = 3;
  let attempt = 0;
  
  while (attempt < maxRetries) {
    try {
      if (!e || !e.response) {
        Logger.log('⚠️ Événement invalide, récupération manuelle de la dernière réponse');
        const form = FormApp.openById(FORM_ID);
        const responses = form.getResponses();
        if (responses.length === 0) {
          throw new Error('Aucune réponse disponible');
        }
        e = { response: responses[responses.length - 1] };
      }
      
      const formResponse = e.response;
      const timestamp = formResponse.getTimestamp();
      let editUrl = formResponse.getEditResponseUrl();
      
      if (!editUrl || editUrl === '') {
        throw new Error('URL vide - retry nécessaire');
      }
      
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_URLS);
      
      // Vérifier si cette réponse existe déjà (éviter doublons)
      const existingData = sheet.getDataRange().getValues();
      const exists = existingData.some(row => 
        row[0] && row[0].getTime && row[0].getTime() === timestamp.getTime()
      );
      
      if (!exists) {
        sheet.appendRow([timestamp, editUrl]);
        Logger.log(`✅ Lien ajouté dans editResponseUrl: ${timestamp}`);
      }
      
      return;
      
    } catch (err) {
      attempt++;
      Logger.log(`⚠️ Tentative ${attempt}/${maxRetries} échouée: ${err.message}`);
      
      if (attempt < maxRetries) {
        Utilities.sleep(2000);
      } else {
        Logger.log(`❌ Échec après ${maxRetries} tentatives`);
      }
    }
  }
}

/**
 * Synchronisation complète de tous les liens d'édition manquants
 * Remplit la feuille editResponseUrl avec tous les liens
 */
function syncAllFormEditLinks() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Synchronisation en cours...', 
      '🔄 Récupération des liens', 
      3
    );
    
    const form = FormApp.openById(FORM_ID);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_URLS);
    
    // Récupérer les timestamps existants pour éviter les doublons
    const existingData = sheet.getDataRange().getValues();
    const existingTimestamps = new Set(
      existingData.slice(1).map(row => row[0]?.getTime()).filter(Boolean)
    );
    
    // Récupérer toutes les réponses du formulaire
    const allResponses = form.getResponses();
    const newResponses = [];
    
    allResponses.forEach(response => {
      const ts = response.getTimestamp().getTime();
      if (!existingTimestamps.has(ts)) {
        try {
          const editUrl = response.getEditResponseUrl();
          if (editUrl && editUrl !== '') {
            newResponses.push([response.getTimestamp(), editUrl]);
          }
        } catch (err) {
          Logger.log(`Erreur pour réponse ${response.getTimestamp()}: ${err.message}`);
        }
      }
    });
    
    if (newResponses.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newResponses.length, 2)
        .setValues(newResponses);
      Logger.log(`✅ ${newResponses.length} nouveaux liens ajoutés`);
    }
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '✅ Synchronisation terminée',
      `${newResponses.length} lien(s) d'édition ajouté(s) dans la feuille 'editResponseUrl'.\n\n` +
      `Ces liens apparaîtront automatiquement en colonne BL grâce à votre formule ARRAYFORMULA.`,
      ui.ButtonSet.OK
    );
    
  } catch (err) {
    Logger.log(`❌ Erreur syncAllFormEditLinks: ${err.message}`);
    SpreadsheetApp.getUi().alert(
      'Erreur',
      `Erreur lors de la synchronisation: ${err.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================================
// SECTION 4 : CONFIGURATION DES DÉCLENCHEURS
// ============================================================

/**
 * Configuration du déclencheur pour onFormSubmitSync
 */
function setupFormTrigger() {
  try {
    // Supprimer les anciens déclencheurs onFormSubmitSync
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onFormSubmitSync') {
        ScriptApp.deleteTrigger(trigger);
        Logger.log('Ancien déclencheur onFormSubmitSync supprimé');
      }
    });
    
    // Créer le nouveau déclencheur
    ScriptApp.newTrigger('onFormSubmitSync')
      .forForm(FormApp.openById(FORM_ID))
      .onFormSubmit()
      .create();
    
    Logger.log('✅ Déclencheur onFormSubmitSync créé');
    
    SpreadsheetApp.getUi().alert(
      '✅ Configuration réussie',
      'Le déclencheur automatique a été configuré.\n\n' +
      'Les liens d\'édition seront maintenant ajoutés automatiquement dans la feuille "editResponseUrl" à chaque envoi du formulaire.\n\n' +
      'NOTE: Votre script de génération automatique des CR consultation (dans le formulaire) continue de fonctionner normalement.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (err) {
    Logger.log(`❌ Erreur setupFormTrigger: ${err.message}`);
    SpreadsheetApp.getUi().alert(
      'Erreur',
      `Erreur lors de la configuration: ${err.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================================
// SECTION 5 : MENU PERSONNALISÉ
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📋 Comptes rendus')
    .addItem('⚡ GÉNÉRER TOUT (intelligent)', 'generateAllPending')
    .addSeparator()
    .addItem('Compte rendu hospit', 'createNewGoogleDocs')
    .addItem('Compte rendu ETT', 'generateETT')
    .addItem('Compte rendu coro', 'generateCoro')
    .addItem('Compte rendu angio', 'generateAngio')
    .addItem('Compte rendu post coro', 'generatePostCoro')
    .addToUi();
  
  ui.createMenu('🔗 Liens formulaires')
    .addItem('🔄 Synchroniser tous les liens', 'syncAllFormEditLinks')
    .addItem('⚙️ Configurer le déclencheur automatique', 'setupFormTrigger')
    .addToUi();
}