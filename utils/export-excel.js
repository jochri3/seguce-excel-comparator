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
) => {
  const workbook = XLSX.utils.book_new();

  // Feuille 1: Résumé
  const summaryData = [
    ["Réconciliation - Résumé"],
    [],
    ["Fichier A (Fournisseur)", fileAName],
    ["Fichier B (SEGUCE RDC)", fileBName],
    [],
    ["Statistiques", "Fichier Fournisseur", "Fichier SEGUCE RDC", "Différence"],
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
        "Doublons Fichier Fournisseur",
        comparisonResult.summary.duplicates.fileA.length,
      ]);
    }

    if (
      comparisonResult.summary.duplicates.fileB &&
      comparisonResult.summary.duplicates.fileB.length > 0
    ) {
      summaryData.push([
        "Doublons Fichier SEGUCE RDC",
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
    ["ID", "Colonne", "Valeur Fichier A", "Valeur Fichier B", "Différence"],
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

  // Feuille 3: Données A avec formules
  const sheetA = createSheetWithFormulas(fileAData, "Données Fournisseur");
  XLSX.utils.book_append_sheet(workbook, sheetA, "Données Fournisseur");

  // Feuille 4: Données B avec formules
  const sheetB = createSheetWithFormulas(fileBData, "Données SEGUCE");
  XLSX.utils.book_append_sheet(workbook, sheetB, "Données SEGUCE");

  // Feuille 5: Synthèse des totaux (la même que dans l'export actuel)
  // ... (code existant pour la feuille de synthèse des totaux)

  // Feuille 6: Synthèse des doublons (nouvelle)
  if (comparisonResult.summary.duplicates) {
    const duplicatesData = [["Synthèse des doublons détectés"], []];

    if (
      comparisonResult.summary.duplicates.fileA &&
      comparisonResult.summary.duplicates.fileA.length > 0
    ) {
      duplicatesData.push(["Doublons dans le Fichier Fournisseur"]);
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
      duplicatesData.push(["Doublons dans le Fichier SEGUCE RDC"]);
      duplicatesData.push(["Matricule", "Occurrences", "Lignes"]);

      comparisonResult.summary.duplicates.fileB.forEach((dup) => {
        duplicatesData.push([dup.matricule, dup.count, dup.rows.join(", ")]);
      });
    }

    const duplicatesSheet = XLSX.utils.aoa_to_sheet(duplicatesData);
    XLSX.utils.book_append_sheet(workbook, duplicatesSheet, "Doublons");
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

    fileData.data.forEach((row, idx) => {
      const rowData = [];

      headers.forEach((header) => {
        rowData.push(row[header] || "");
      });

      rows.push(rowData);

      // Vérifier s'il y a des formules pour cette ligne
      if (fileData.formulas && fileData.formulas[idx]) {
        const formulaRow = fileData.formulas[idx];

        // Appliquer les formules
        Object.entries(formulaRow).forEach(([col, formula]) => {
          const colIndex = headers.indexOf(col);
          if (colIndex !== -1) {
            // Définir la cellule avec la formule
            const r = idx + 1; // +1 pour l'en-tête
            const c = colIndex;

            // XLSX.utils.encode_cell convertit les indices de ligne/colonne en adresse de cellule
            const cellRef = XLSX.utils.encode_cell({ r: r + 1, c: c }); // +1 car les lignes commencent à 0 dans XLSX

            // Formater la référence de cellule pour la formule
            if (!formulaRow._sheet) {
              // Si on n'a pas de référence explicite, on utilise la formule telle quelle
              rows[r + 1][c] = { f: formula };
            }
          }
        });
      }
    });

    // Créer la feuille à partir des données
    const sheet = XLSX.utils.aoa_to_sheet(rows);

    // Parcourir les formules et les appliquer
    if (fileData.formulas) {
      Object.entries(fileData.formulas).forEach(([rowIdx, formulas]) => {
        const rowIndex = parseInt(rowIdx) + 1; // +1 pour l'en-tête

        Object.entries(formulas).forEach(([col, formula]) => {
          const colIndex = headers.indexOf(col);
          if (colIndex !== -1) {
            // Récupérer la référence de cellule
            const cellRef = XLSX.utils.encode_cell({
              r: rowIndex,
              c: colIndex,
            });

            // Définir la formule
            if (!sheet[cellRef]) {
              sheet[cellRef] = {};
            }
            sheet[cellRef].f = formula;
          }
        });
      });
    }

    return sheet;
  } catch (error) {
    console.error(
      "Erreur lors de la création de la feuille avec formules:",
      error,
    );
    // En cas d'erreur, créer une feuille simple sans formules
    const data = [fileData.headers.map((h) => h.key)];
    fileData.data.forEach((row) => {
      const rowData = fileData.headers.map((h) => row[h.key] || "");
      data.push(rowData);
    });
    return XLSX.utils.aoa_to_sheet(data);
  }
};

// Exporter les fonctions
module.exports = {
  exportToExcel: exportToExcelWithFormulas,
};
