const XLSX = require("xlsx");

/**
 * Convertit les différences en feuille Excel avec formules
 * @param {Object} comparisonResult - Résultats de la comparaison
 * @param {Object} fileAData - Données brutes du fichier A
 * @param {Object} fileBData - Données brutes du fichier B
 * @returns {Buffer} - Buffer Excel contenant le rapport
 */
const exportToExcelWithFormulas = (
  comparisonResult,
  fileAData,
  fileBData,
  fileAName,
  fileBName,
  lexiqueItems
) => {
  const workbook = XLSX.utils.book_new();

  // Feuille 1: Résumé
  const summaryData = [
    ["Réconciliation - Résumé"],
    [],
    ["Fichier Fournisseur", fileAName], // MODIFICATION: Changé de "Fichier A" à "Fichier Fournisseur"
    ["Fichier SEGUCE RDC", fileBName], // MODIFICATION: Changé de "Fichier B" à "Fichier SEGUCE RDC"
    [],
    ["Statistiques", "Fichier Fournisseur", "Fichier SEGUCE RDC", "Différence"], // MODIFICATION: Labels mis à jour
    [
      "Nombre de lignes",
      comparisonResult.summary.totalRows.fileA,
      comparisonResult.summary.totalRows.fileB,
      comparisonResult.summary.totalRows.fileA -
        comparisonResult.summary.totalRows.fileB,
    ],
    [
      "Nombre de différences trouvées",
      comparisonResult.summary.totalDifferences,
    ],
    [],
    ["Différences par colonne"],
  ];

  // Ajouter les différences par colonne
  if (comparisonResult.summary.columnDifferences) {
    comparisonResult.summary.columnDifferences.forEach((colDiff) => {
      summaryData.push([colDiff.column, colDiff.count]);
    });
  }

  // Ajouter des informations sur les doublons
  if (comparisonResult.summary.duplicates) {
    summaryData.push([]);
    summaryData.push(["Doublons de matricules"]);

    if (
      comparisonResult.summary.duplicates.fileA &&
      comparisonResult.summary.duplicates.fileA.length > 0
    ) {
      summaryData.push([
        "Doublons Fichier Fournisseur", // MODIFICATION: Label mis à jour
        comparisonResult.summary.duplicates.fileA.length,
      ]);
    }

    if (
      comparisonResult.summary.duplicates.fileB &&
      comparisonResult.summary.duplicates.fileB.length > 0
    ) {
      summaryData.push([
        "Doublons Fichier SEGUCE RDC", // MODIFICATION: Label mis à jour
        comparisonResult.summary.duplicates.fileB.length,
      ]);
    }
  }

  // Ajouter des informations sur la classification
  if (comparisonResult.summary.sequentialComparison) {
    summaryData.push([]);
    summaryData.push(["Analyse séquentielle"]);
    summaryData.push([
      "Éléments fixes vérifiés",
      comparisonResult.summary.sequentialComparison.fixedElements.totalColumns,
    ]);
    summaryData.push([
      "Erreurs dans éléments fixes",
      comparisonResult.summary.sequentialComparison.fixedElements.totalErrors,
    ]);
    summaryData.push([
      "Éléments variables vérifiés",
      comparisonResult.summary.sequentialComparison.variableElements
        .totalColumns,
    ]);
    summaryData.push([
      "Erreurs dans éléments variables",
      comparisonResult.summary.sequentialComparison.variableElements
        .totalErrors,
    ]);
  }

  summaryData.push([]);
  summaryData.push(["Date d'exportation", new Date().toLocaleString()]);

  // Créer la feuille de résumé
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(workbook, summarySheet, "Résumé");

  // Feuille 2: Détails des différences
  const detailsData = [
    [
      "ID",
      "Colonne",
      "Valeur Fichier Fournisseur",
      "Valeur Fichier SEGUCE RDC",
      "Différence",
    ], // MODIFICATION: Labels mis à jour
  ];

  // Ajouter les détails des différences
  comparisonResult.details.forEach((detail) => {
    if (detail.onlyInFileA) {
      detailsData.push([detail.id, "LIGNE COMPLÈTE", "Présent", "Absent", ""]);
    } else if (detail.onlyInFileB) {
      detailsData.push([detail.id, "LIGNE COMPLÈTE", "Absent", "Présent", ""]);
    } else if (detail.differences) {
      detail.differences.forEach((diff) => {
        detailsData.push([
          detail.id,
          diff.column,
          diff.valueA,
          diff.valueB,
          diff.difference,
        ]);
      });
    }
  });

  // Créer la feuille de détails
  const detailsSheet = XLSX.utils.aoa_to_sheet(detailsData);
  XLSX.utils.book_append_sheet(workbook, detailsSheet, "Détails");

  // Feuille 3: Données Fournisseur avec formules
  const sheetA = createSheetWithFormulas(fileAData, "Données Fournisseur"); // MODIFICATION: Nom feuille mis à jour
  XLSX.utils.book_append_sheet(workbook, sheetA, "Données Fournisseur"); // MODIFICATION: Nom feuille mis à jour

  // Feuille 4: Données SEGUCE avec formules
  const sheetB = createSheetWithFormulas(fileBData, "Données SEGUCE"); // MODIFICATION: Nom feuille mis à jour
  XLSX.utils.book_append_sheet(workbook, sheetB, "Données SEGUCE"); // MODIFICATION: Nom feuille mis à jour

  // Feuille 5: Synthèse des totaux (la même que dans l'export actuel)
  // ... (code existant pour la feuille de synthèse des totaux)

  // Feuille 6: Synthèse des doublons (nouvelle)
  if (comparisonResult.summary.duplicates) {
    const duplicatesData = [["Synthèse des doublons détectés"], []];

    if (
      comparisonResult.summary.duplicates.fileA &&
      comparisonResult.summary.duplicates.fileA.length > 0
    ) {
      duplicatesData.push(["Doublons dans le Fichier Fournisseur"]); // MODIFICATION: Label mis à jour
      duplicatesData.push(["Matricule", "Occurrences", "Lignes"]);

      comparisonResult.summary.duplicates.fileA.forEach((dup) => {
        duplicatesData.push([dup.matricule, dup.count, dup.rows.join(", ")]);
      });

      duplicatesData.push([]);
    }

    if (
      comparisonResult.summary.duplicates.fileB &&
      comparisonResult.summary.duplicates.fileB.length > 0
    ) {
      duplicatesData.push(["Doublons dans le Fichier SEGUCE RDC"]); // MODIFICATION: Label mis à jour
      duplicatesData.push(["Matricule", "Occurrences", "Lignes"]);

      comparisonResult.summary.duplicates.fileB.forEach((dup) => {
        duplicatesData.push([dup.matricule, dup.count, dup.rows.join(", ")]);
      });
    }

    const duplicatesSheet = XLSX.utils.aoa_to_sheet(duplicatesData);
    XLSX.utils.book_append_sheet(workbook, duplicatesSheet, "Doublons");
  }

  // Créer une feuille pour les écarts significatifs
  const significantDifferencesData = [
    ["Écarts significatifs"],
    [],
    [
      "Matricule",
      "Colonne",
      "Fichier Fournisseur", // MODIFICATION: Label mis à jour
      "Fichier SEGUCE RDC", // MODIFICATION: Label mis à jour
      "Différence",
      "Différence %",
    ],
  ];

  // Identifier les écarts significatifs (différence > 5% ou valeur absolue > 100)
  let significantDiffsFound = false;
  comparisonResult.details.forEach((detail) => {
    if (detail.differences) {
      detail.differences.forEach((diff) => {
        if (
          typeof diff.valueA === "number" &&
          typeof diff.valueB === "number"
        ) {
          const absDiff = Math.abs(diff.valueB - diff.valueA);
          const percentDiff =
            diff.valueA !== 0 ? (absDiff / Math.abs(diff.valueA)) * 100 : 100;

          if (percentDiff > 5 || absDiff > 100) {
            significantDiffsFound = true;
            significantDifferencesData.push([
              detail.id,
              diff.column,
              diff.valueA,
              diff.valueB,
              diff.valueB - diff.valueA,
              `${percentDiff.toFixed(2)}%`,
            ]);
          }
        }
      });
    }
  });

  if (significantDiffsFound) {
    const significantDiffsSheet = XLSX.utils.aoa_to_sheet(
      significantDifferencesData
    );
    XLSX.utils.book_append_sheet(
      workbook,
      significantDiffsSheet,
      "Écarts significatifs"
    );
  }


  // Feuille: Synthèse des écarts variables
  const variableElementsData = [
    ["Écarts des données variables"],
    [],
    [
      "Rubrique",
      "Fichier prestataire paie",
      "Fichier SEGUCE",
      "Différence",
      "Différence %",
    ],
  ];

  // Liste des éléments variables
  const variableElements = [
    "Jr. Prestés",
    "JAbsence",
    "Jrs. Maladie",
    "Congé Circ.",
    "Congé Maternité",
    "Jrs. Ferié",
    "Jrs. Congé Payés",
    "Nbre Heure Supp. 1",
    "Nbre Heure Supp. 2",
    "Nbre Heure Supp. 3",
    "Heures de nuit 25%",
    "Regule sur congé",
    "Regule",
    "Prime de bonne conduite auto",
    "Prime de production",
    "Aide sociale",
    "Prime intérim",
    "13ème mois",
    "Pagne + frais de couture",
    "Prime migration IT",
    "Aide rentrée scolaire",
    "Bonus",
    "Prime de mariage",
    "Prime Audit",
    "Panier de fin d'année",
    "Autre prime",
    "Prime Imposable Variable",
  ];

  // Extraire les écarts pour chaque élément variable
  let variableDiffsFound = false;
  variableElements.forEach((element) => {
    let totalA = 0;
    let totalB = 0;

    // Chercher parmi les détails
    comparisonResult.details.forEach((detail) => {
      if (detail.differences) {
        detail.differences.forEach((diff) => {
          if (diff.column === element) {
            if (typeof diff.valueA === "number") totalA += diff.valueA;
            if (typeof diff.valueB === "number") totalB += diff.valueB;
            variableDiffsFound = true;
          }
        });
      }
    });

    if (totalA !== 0 || totalB !== 0) {
      const difference = totalB - totalA;
      const percentDiff =
        totalA !== 0 ? (Math.abs(difference) / Math.abs(totalA)) * 100 : 100;

      variableElementsData.push([
        element,
        totalA,
        totalB,
        difference,
        `${percentDiff.toFixed(2)}%`,
      ]);
    }
  });

  if (variableDiffsFound) {
    const variableElementsSheet = XLSX.utils.aoa_to_sheet(variableElementsData);
    XLSX.utils.book_append_sheet(
      workbook,
      variableElementsSheet,
      "Écarts Variables"
    );
  }

  // Feuille: Synthèse des écarts fixes
  const fixedElementsData = [
    ["Écarts des données fixes"],
    [],
    [
      "Rubrique",
      "Fichier prestataire paie",
      "Fichier SEGUCE",
      "Différence",
      "Différence %",
    ],
  ];

  // Liste des éléments fixes
  const fixedElements = [
    "Matricule",
    "BU",
    "Pers. à Charge",
    "Enfant Légal",
    "Salaire mensuel",
    "Ancienneté mensuel",
    "Sur Salaire mensuel",
    "Taux horaire",
    "Transport mensuel",
    "Logement mensuel",
    "Prime astreinte",
    "Forfait heures supplémentaire",
    "Prime de détachement",
    "Prime Imposable Fixe",
  ];

  // Extraire les écarts pour chaque élément fixe
  let fixedDiffsFound = false;
  fixedElements.forEach((element) => {
    let totalA = 0;
    let totalB = 0;

    // Chercher parmi les détails
    comparisonResult.details.forEach((detail) => {
      if (detail.differences) {
        detail.differences.forEach((diff) => {
          if (diff.column === element) {
            if (typeof diff.valueA === "number") totalA += diff.valueA;
            if (typeof diff.valueB === "number") totalB += diff.valueB;
            fixedDiffsFound = true;
          }
        });
      }
    });

    if (totalA !== 0 || totalB !== 0) {
      const difference = totalB - totalA;
      const percentDiff =
        totalA !== 0 ? (Math.abs(difference) / Math.abs(totalA)) * 100 : 100;

      fixedElementsData.push([
        element,
        totalA,
        totalB,
        difference,
        `${percentDiff.toFixed(2)}%`,
      ]);
    }
  });

  if (fixedDiffsFound) {
    const fixedElementsSheet = XLSX.utils.aoa_to_sheet(fixedElementsData);
    XLSX.utils.book_append_sheet(workbook, fixedElementsSheet, "Écarts Fixes");
  }

  // Feuille: Synthèse complète des rubriques
  const allRubriquesData = [
    ["Synthèse des rubriques"],
    [],
    [
      "Rubrique",
      "Fichier prestataire paie",
      "Fichier SEGUCE",
      "Différence",
      "Différence %",
    ],
  ];

  // Rassembler toutes les rubriques uniques
  const allColumns = new Set();
  comparisonResult.details.forEach((detail) => {
    if (detail.differences) {
      detail.differences.forEach((diff) => {
        allColumns.add(diff.column);
      });
    }
  });

  // Calculer les totaux pour chaque rubrique
  Array.from(allColumns)
    .sort()
    .forEach((column) => {
      let totalA = 0;
      let totalB = 0;

      comparisonResult.details.forEach((detail) => {
        if (detail.differences) {
          detail.differences.forEach((diff) => {
            if (diff.column === column) {
              if (typeof diff.valueA === "number") totalA += diff.valueA;
              if (typeof diff.valueB === "number") totalB += diff.valueB;
            }
          });
        }
      });

      if (totalA !== 0 || totalB !== 0) {
        const difference = totalB - totalA;
        const percentDiff =
          totalA !== 0 ? (Math.abs(difference) / Math.abs(totalA)) * 100 : 100;

        allRubriquesData.push([
          column,
          totalA,
          totalB,
          difference,
          `${percentDiff.toFixed(2)}%`,
        ]);
      }
    });

  const allRubriquesSheet = XLSX.utils.aoa_to_sheet(allRubriquesData);
  XLSX.utils.book_append_sheet(
    workbook,
    allRubriquesSheet,
    "Synthèse rubriques"
  );

  // Feuille 7: Lexique des formules
  const lexiqueData = [
    ["Lexique des formules"],
    [],
    ["Rubrique", "Type", "Description", "Formule"],
  ];

  // Récupérer les données du lexique depuis la base de données
  // Note: Ceci sera implémenté de manière asynchrone dans la fonction app.js
  if (lexiqueItems && lexiqueItems.length > 0) {
    lexiqueItems.forEach((item) => {
      lexiqueData.push([
        item.column_name,
        item.column_type === "fixe" ? "Élément fixe" : "Élément variable",
        item.description || "-",
        item.formula || "-",
      ]);
    });
  } else {
    // Ajouter quelques formules par défaut
    lexiqueData.push([
      "Plafond Cnss",
      "Fixe",
      "Montant maximum soumis à cotisation",
      "Salaire+Ancienneté+Sur Salaire+Maladie+Congé Circ.+Ferié+Congé Maternité+Congé Payer+Heure Supplémentaire+...etc",
    ]);
    lexiqueData.push([
      "Cnss QPO",
      "Fixe",
      "CNSS Quote part ouvrier",
      "Plafond Cnss*5%",
    ]);
    lexiqueData.push([
      "Cnss QPP",
      "Fixe",
      "CNSS Quote part patronale",
      "Plafond Cnss*13%",
    ]);
    lexiqueData.push([
      "Inpp",
      "Fixe",
      "Institut National de Préparation Professionnelle",
      "Plafond Cnss*2%",
    ]);
    lexiqueData.push([
      "Onem",
      "Fixe",
      "Office National de l'Emploi",
      "Plafond Cnss*0,2%",
    ]);
  }

  const lexiqueSheet = XLSX.utils.aoa_to_sheet(lexiqueData);
  XLSX.utils.book_append_sheet(workbook, lexiqueSheet, "Lexique");

  // Convertir le classeur en buffer
  const excelBuffer = XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx",
  });

  return excelBuffer;
};

/**
 * Crée une feuille Excel à partir des données, en préservant les formules
 */
const createSheetWithFormulas = (fileData, sheetName) => {
  try {
    // Créer l'en-tête
    const headers = fileData.headers.map((h) => h.key);

    // Préparer les données
    const rows = [headers];

    fileData.data.forEach((row) => {
      const rowData = [];

      headers.forEach((header) => {
        rowData.push(row[header] !== undefined ? row[header] : "");
      });

      rows.push(rowData);
    });

    // Créer la feuille à partir des données
    const sheet = XLSX.utils.aoa_to_sheet(rows);

    // Appliquer les formules après la création de la feuille
    if (fileData.formulas) {
      Object.entries(fileData.formulas).forEach(([rowIdx, formulas]) => {
        const rowIndex = parseInt(rowIdx) + 1; // +1 pour l'en-tête

        Object.entries(formulas).forEach(([col, formula]) => {
          const colIndex = headers.indexOf(col);
          if (colIndex !== -1) {
            const cellRef = XLSX.utils.encode_cell({
              r: rowIndex,
              c: colIndex,
            });

            // Créer la cellule si elle n'existe pas
            if (!sheet[cellRef]) {
              sheet[cellRef] = { t: "n", v: 0 };
            }

            // Définir la formule
            sheet[cellRef].f = formula;
          }
        });
      });
    }

    // Ajouter des styles à la feuille
    addStylingToSheet(sheet);

    return sheet;
  } catch (error) {
    console.error(
      "Erreur lors de la création de la feuille avec formules:",
      error,
    );
    // Fallback: créer une feuille simple sans formules
    const data = [fileData.headers.map((h) => h.key)];
    fileData.data.forEach((row) => {
      const rowData = fileData.headers.map((h) => row[h.key] || "");
      data.push(rowData);
    });
    return XLSX.utils.aoa_to_sheet(data);
  }
};

/**
 * Ajoute du style à une feuille Excel
 */
const addStylingToSheet = (sheet) => {
  // Obtenir la plage de la feuille
  const range = XLSX.utils.decode_range(sheet["!ref"]);

  // Définir les largeurs de colonnes optimales
  const colWidths = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    colWidths.push({ wch: 15 }); // Largeur par défaut
  }

  // Appliquer les largeurs
  sheet["!cols"] = colWidths;

  return sheet;
};

// Exporter les fonctions
module.exports = {
  exportToExcel: exportToExcelWithFormulas,
};