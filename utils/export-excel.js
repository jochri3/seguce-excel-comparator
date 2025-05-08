const XLSX = require("xlsx");

/**
 * Convertit les différences en feuille Excel avec formules
 * @param {Object} comparisonResult - Résultats de la comparaison
 * @param {Object} fileAData - Données brutes du fichier Fournisseur
 * @param {Object} fileBData - Données brutes du fichier SEGUCE
 * @param {string} fileAName - Nom du fichier Fournisseur
 * @param {string} fileBName - Nom du fichier SEGUCE
 * @param {Array} lexiqueItems - Données du lexique (optionnel)
 * @returns {Buffer} - Buffer Excel contenant le rapport
 */
const exportToExcelWithFormulas = (
  comparisonResult,
  fileAData,
  fileBData,
  fileAName,
  fileBName,
  lexiqueItems = []
) => {
  const workbook = XLSX.utils.book_new();

  // Feuille 1: Résumé
  const summaryData = [
    ["Réconciliation - Résumé"],
    [],
    ["Fichier prestataire paie", fileAName],
    ["Fichier SEGUCE", fileBName],
    [],
    [
      "Statistiques",
      "Fichier prestataire paie",
      "Fichier SEGUCE",
      "Différence",
    ],
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
    comparisonResult.summary.columnDifferences
      .filter((colDiff) => !colDiff.column.startsWith("Col")) // Filtrer les colonnes génériques
      .forEach((colDiff) => {
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
        "Doublons Fichier prestataire paie",
        comparisonResult.summary.duplicates.fileA.length,
      ]);
    }

    if (
      comparisonResult.summary.duplicates.fileB &&
      comparisonResult.summary.duplicates.fileB.length > 0
    ) {
      summaryData.push([
        "Doublons Fichier SEGUCE",
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
      "Valeur Fichier prestataire paie",
      "Valeur Fichier SEGUCE",
      "Différence",
    ],
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

  // // Feuille 3: Données Fournisseur avec formules
  // const sheetA = createSheetWithFormulas(fileAData, "Données Fournisseur");
  // XLSX.utils.book_append_sheet(workbook, sheetA, "Données Fournisseur");

  // // Feuille 4: Données SEGUCE avec formules
  // const sheetB = createSheetWithFormulas(fileBData, "Données SEGUCE");
  // XLSX.utils.book_append_sheet(workbook, sheetB, "Données SEGUCE");

  // Feuille 5: Synthèse des doublons (nouvelle)
  if (comparisonResult.summary.duplicates) {
    const duplicatesData = [["Synthèse des doublons détectés"], []];

    if (
      comparisonResult.summary.duplicates.fileA &&
      comparisonResult.summary.duplicates.fileA.length > 0
    ) {
      duplicatesData.push(["Doublons dans le Fichier prestataire paie"]);
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
      duplicatesData.push(["Doublons dans le Fichier SEGUCE"]);
      duplicatesData.push(["Matricule", "Occurrences", "Lignes"]);

      comparisonResult.summary.duplicates.fileB.forEach((dup) => {
        duplicatesData.push([dup.matricule, dup.count, dup.rows.join(", ")]);
      });
    }

    const duplicatesSheet = XLSX.utils.aoa_to_sheet(duplicatesData);
    XLSX.utils.book_append_sheet(workbook, duplicatesSheet, "Doublons");
  }

  // Feuille 6: Écarts significatifs
  const significantDifferencesData = [
    ["Écarts significatifs"],
    [],
    [
      "Matricule",
      "Colonne",
      "Fichier prestataire paie",
      "Fichier SEGUCE",
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

  // Feuille 7: Synthèse des écarts variables
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

  // Collecter les différences par colonne pour les éléments variables
  const variableDiffs = {};
  let variableDiffsFound = false;

  comparisonResult.details.forEach((detail) => {
    if (detail.differences) {
      detail.differences.forEach((diff) => {
        const columnLower = diff.column.toLowerCase();

        // Vérifier si la colonne est dans les éléments variables
        const isVariable =
          variableElements.some((varElement) =>
            columnLower.includes(varElement.toLowerCase())
          ) ||
          !fixedElements.some((fixedElement) =>
            columnLower.includes(fixedElement.toLowerCase())
          );

        if (isVariable && !columnLower.startsWith("col")) {
          // Ignorer les colonnes génériques
          if (!variableDiffs[diff.column]) {
            variableDiffs[diff.column] = {
              totalA: 0,
              totalB: 0,
              count: 0,
            };
          }

          if (typeof diff.valueA === "number")
            variableDiffs[diff.column].totalA += diff.valueA;
          if (typeof diff.valueB === "number")
            variableDiffs[diff.column].totalB += diff.valueB;
          variableDiffs[diff.column].count++;
          variableDiffsFound = true;
        }
      });
    }
  });

  // Ajouter les différences triées au tableau
  Object.entries(variableDiffs)
    .filter(
      ([_, values]) =>
        values.count > 0 && Math.abs(values.totalA - values.totalB) > 0.001
    )
    .filter(
      ([column, _]) =>
        !fixedElements.some((fixed) =>
          column.toLowerCase().includes(fixed.toLowerCase())
        )
    )
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([column, values]) => {
      const difference = values.totalB - values.totalA;
      const percentDiff =
        values.totalA !== 0
          ? (Math.abs(difference) / Math.abs(values.totalA)) * 100
          : 100;

      variableElementsData.push([
        column,
        values.totalA,
        values.totalB,
        difference,
        `${percentDiff.toFixed(2)}%`,
      ]);
    });

  if (variableDiffsFound && variableElementsData.length > 3) {
    const variableElementsSheet = XLSX.utils.aoa_to_sheet(variableElementsData);
    XLSX.utils.book_append_sheet(
      workbook,
      variableElementsSheet,
      "Écarts Variables"
    );
  }

  // Feuille 8: Synthèse des écarts fixes
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

  // Collecter les différences par colonne pour les éléments fixes
  const fixedDiffs = {};
  let fixedDiffsFound = false;

  comparisonResult.details.forEach((detail) => {
    if (detail.differences) {
      detail.differences.forEach((diff) => {
        const columnLower = diff.column.toLowerCase();

        // Vérifier si la colonne est un élément fixe
        const isFixed = fixedElements.some((fixedElement) =>
          columnLower.includes(fixedElement.toLowerCase())
        );

        if (isFixed && !columnLower.startsWith("col")) {
          // Ignorer les colonnes génériques
          if (!fixedDiffs[diff.column]) {
            fixedDiffs[diff.column] = {
              totalA: 0,
              totalB: 0,
              count: 0,
            };
          }

          if (typeof diff.valueA === "number")
            fixedDiffs[diff.column].totalA += diff.valueA;
          if (typeof diff.valueB === "number")
            fixedDiffs[diff.column].totalB += diff.valueB;
          fixedDiffs[diff.column].count++;
          fixedDiffsFound = true;
        }
      });
    }
  });

  // Ajouter les différences triées au tableau
  Object.entries(fixedDiffs)
    .filter(
      ([_, values]) =>
        values.count > 0 && Math.abs(values.totalA - values.totalB) > 0.001
    )
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([column, values]) => {
      const difference = values.totalB - values.totalA;
      const percentDiff =
        values.totalA !== 0
          ? (Math.abs(difference) / Math.abs(values.totalA)) * 100
          : 100;

      fixedElementsData.push([
        column,
        values.totalA,
        values.totalB,
        difference,
        `${percentDiff.toFixed(2)}%`,
      ]);
    });

  if (fixedDiffsFound && fixedElementsData.length > 3) {
    const fixedElementsSheet = XLSX.utils.aoa_to_sheet(fixedElementsData);
    XLSX.utils.book_append_sheet(workbook, fixedElementsSheet, "Écarts Fixes");
  }

  // Feuille 9: Synthèse complète des rubriques
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
        if (!diff.column.startsWith("Col")) {
          // Ignorer les colonnes génériques
          allColumns.add(diff.column);
        }
      });
    }
  });

  // Calculer les totaux pour chaque rubrique
  Array.from(allColumns)
    .sort()
    .forEach((column) => {
      let totalA = 0;
      let totalB = 0;
      let count = 0;

      comparisonResult.details.forEach((detail) => {
        if (detail.differences) {
          detail.differences.forEach((diff) => {
            if (diff.column === column) {
              if (typeof diff.valueA === "number") totalA += diff.valueA;
              if (typeof diff.valueB === "number") totalB += diff.valueB;
              count++;
            }
          });
        }
      });

      if (count > 0 && Math.abs(totalA - totalB) > 0.001) {
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

  if (allRubriquesData.length > 3) {
    const allRubriquesSheet = XLSX.utils.aoa_to_sheet(allRubriquesData);
    XLSX.utils.book_append_sheet(
      workbook,
      allRubriquesSheet,
      "Synthèse rubriques"
    );
  }

  // Feuille 10: Lexique des formules
  const lexiqueData = [
    ["Lexique des formules"],
    [],
    ["Rubrique", "Type", "Description", "Formule"],
  ];

  // Récupérer les données du lexique
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

  // Feuille 11 : Pour les charges

  if (comparisonResult.summary.chargesCategories) {
    const chargesData = [
      ["Récapitulatif des charges"],
      [],
      ["", "Fichier prestataire paie", "Fichier SEGUCE", "Différence"],

      ["Charges salariales", "", "", ""],
    ];

    // Ajouter les charges salariales
    comparisonResult.summary.chargesCategories.chargesSalariales.items.forEach(
      (item) => {
        chargesData.push([
          item.column,
          item.valueA,
          item.valueB,
          item.difference,
        ]);
      }
    );

    // Ajouter le total des charges salariales
    chargesData.push([
      "Total charges salariales",
      comparisonResult.summary.chargesCategories.chargesSalariales.totalA,
      comparisonResult.summary.chargesCategories.chargesSalariales.totalB,
      comparisonResult.summary.chargesCategories.chargesSalariales.difference,
    ]);

    // Ajouter un séparateur
    chargesData.push([]);

    // Ajouter les charges patronales
    chargesData.push(["Charges patronales", "", "", ""]);
    comparisonResult.summary.chargesCategories.chargesPatronales.items.forEach(
      (item) => {
        chargesData.push([
          item.column,
          item.valueA,
          item.valueB,
          item.difference,
        ]);
      }
    );

    // Ajouter le total des charges patronales
    chargesData.push([
      "Total charges patronales",
      comparisonResult.summary.chargesCategories.chargesPatronales.totalA,
      comparisonResult.summary.chargesCategories.chargesPatronales.totalB,
      comparisonResult.summary.chargesCategories.chargesPatronales.difference,
    ]);

    // Ajouter un séparateur
    chargesData.push([]);

    // Ajouter la répartition des coûts
    chargesData.push(["Répartition des coûts liés aux salaires", "", "", ""]);
    comparisonResult.summary.chargesCategories.coutsSalaires.items.forEach(
      (item) => {
        chargesData.push([
          item.column,
          item.valueA,
          item.valueB,
          item.difference,
        ]);
      }
    );

    const chargesSheet = XLSX.utils.aoa_to_sheet(chargesData);
    XLSX.utils.book_append_sheet(workbook, chargesSheet, "Récap charges");
  }

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
      error
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
