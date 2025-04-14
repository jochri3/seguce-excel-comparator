const XLSX = require("xlsx");

/**
 * Exporte les différences au format Excel
 * @param {Object} comparisonResult - Les résultats de la comparaison
 * @param {string} fileAName - Nom du fichier A
 * @param {string} fileBName - Nom du fichier B
 * @returns {Buffer} - Le contenu du fichier Excel
 */
const exportToExcel = (comparisonResult, fileAName, fileBName) => {
  // Créer un nouveau classeur
  const workbook = XLSX.utils.book_new();

  // Feuille 1: Résumé
  const summaryData = [
    ["Réconciliation - Résumé"],
    [],
    ["Fichier A (Entreprise)", fileAName],
    ["Fichier B (Client)", fileBName],
    [],
    ["Statistiques", "Fichier A", "Fichier B", "Différence"],
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
  comparisonResult.summary.columnDifferences.forEach((colDiff) => {
    summaryData.push([colDiff.column, colDiff.count]);
  });

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
    } else {
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

  // Feuille 3: Synthèse des totaux
  const getTotals = (details, isFileA = true) => {
    const totals = {
      "Cnss QPO": 0,
      IPR: 0,
      "Cnss QPP": 0,
      Inpp: 0,
      Onem: 0,
      "Total Charge Patronale": 0,
      "Coût Salarial": 0,
      "Masse Salariale": 0,
      "Net à Payer": 0,
      "Frais de Services": 0,
      "TVA 16% Frais Services": 0,
      "Total Employeur Mensuel": 0,
    };

    const keyPrefixes = [
      "cnss",
      "qpo",
      "ipr",
      "qpp",
      "inpp",
      "onem",
      "charge",
      "patronale",
      "coût",
      "cout",
      "salarial",
      "masse",
      "net",
      "payer",
      "frais",
      "service",
      "tva",
      "employeur",
      "mensuel",
    ];

    const normalizeKey = (key) => key.toLowerCase().replace(/\s+/g, "");

    // Fonction pour détecter le type de colonne
    const detectColumnType = (colName) => {
      const normalizedName = normalizeKey(colName);

      for (const key in totals) {
        if (normalizedName.includes(normalizeKey(key))) {
          return key;
        }
      }

      for (const prefix of keyPrefixes) {
        if (normalizedName.includes(prefix)) {
          if (
            normalizedName.includes("qpo") ||
            normalizedName.includes("charge salariale")
          )
            return "Cnss QPO";
          if (normalizedName.includes("ipr")) return "IPR";
          if (normalizedName.includes("qpp")) return "Cnss QPP";
          if (normalizedName.includes("inpp")) return "Inpp";
          if (normalizedName.includes("onem")) return "Onem";
          if (
            normalizedName.includes("charge patronale") ||
            normalizedName.includes("patronale")
          )
            return "Total Charge Patronale";
          if (
            (normalizedName.includes("cout") ||
              normalizedName.includes("coût")) &&
            normalizedName.includes("salarial")
          )
            return "Coût Salarial";
          if (normalizedName.includes("masse")) return "Masse Salariale";
          if (
            normalizedName.includes("net") &&
            normalizedName.includes("payer")
          )
            return "Net à Payer";
          if (
            normalizedName.includes("frais") &&
            normalizedName.includes("service")
          ) {
            if (
              normalizedName.includes("tva") ||
              normalizedName.includes("16%")
            )
              return "TVA 16% Frais Services";
            return "Frais de Services";
          }
          if (
            normalizedName.includes("total") &&
            normalizedName.includes("employeur")
          )
            return "Total Employeur Mensuel";
        }
      }

      return null;
    };

    // Parcourir les détails et extraire les valeurs totales
    details.forEach((detail) => {
      if (detail.onlyInFileA && !isFileA) return; // Ignorer si on cherche des totaux pour B
      if (detail.onlyInFileB && isFileA) return; // Ignorer si on cherche des totaux pour A

      const detailData = detail.onlyInFileA
        ? detail.rowData
        : detail.onlyInFileB
        ? detail.rowData
        : null;

      if (detailData) {
        // Cas où la ligne entière n'existe que dans un fichier
        Object.entries(detailData).forEach(([col, val]) => {
          const columnType = detectColumnType(col);
          if (columnType && typeof val === "number") {
            totals[columnType] += val;
          } else if (columnType && typeof val === "string") {
            const numVal = parseFloat(val.replace(/,/g, ""));
            if (!isNaN(numVal)) {
              totals[columnType] += numVal;
            }
          }
        });
      } else if (detail.differences) {
        // Cas des différences entre deux fichiers
        detail.differences.forEach((diff) => {
          const columnType = detectColumnType(diff.column);
          if (columnType) {
            const value = isFileA ? diff.valueA : diff.valueB;
            if (typeof value === "number") {
              totals[columnType] += value;
            } else if (typeof value === "string") {
              const numVal = parseFloat(value.replace(/,/g, ""));
              if (!isNaN(numVal)) {
                totals[columnType] += numVal;
              }
            }
          }
        });
      }
    });

    return totals;
  };

  const totalsA = getTotals(comparisonResult.details, true);
  const totalsB = getTotals(comparisonResult.details, false);

  const calculateDifference = (a, b) => {
    if (typeof a === "number" && typeof b === "number") {
      return a - b;
    }
    return 0;
  };

  const synthesisTotalsData = [
    ["Synthèse des totaux"],
    [],
    ["Rubrique", "Fichier A (Entreprise)", "Fichier B (Client)", "Différence"],
  ];

  // Ajouter les totaux pour chaque rubrique clé
  Object.keys(totalsA).forEach((key) => {
    synthesisTotalsData.push([
      key,
      totalsA[key].toFixed(2),
      totalsB[key].toFixed(2),
      calculateDifference(totalsA[key], totalsB[key]).toFixed(2),
    ]);
  });

  // Informations supplémentaires
  synthesisTotalsData.push([]);
  synthesisTotalsData.push(["Informations supplémentaires"]);
  synthesisTotalsData.push([
    "Nombre de lignes",
    comparisonResult.summary.totalRows.fileA,
    comparisonResult.summary.totalRows.fileB,
  ]);

  // Vérification des doublons de matricules
  const getMatriculeCount = (details) => {
    const matricules = new Set();
    details.forEach((detail) => {
      if (detail.id) {
        matricules.add(detail.id);
      }
    });
    return matricules.size;
  };

  const matriculeCount = getMatriculeCount(comparisonResult.details);
  const hasDuplicates =
    matriculeCount < comparisonResult.summary.totalRows.fileA ||
    matriculeCount < comparisonResult.summary.totalRows.fileB;

  synthesisTotalsData.push(["Nombre de matricules contrôlés", matriculeCount]);
  synthesisTotalsData.push([
    "Présence de doublons",
    hasDuplicates ? "OUI" : "NON",
  ]);
  synthesisTotalsData.push([
    "Nombre d'erreurs",
    comparisonResult.summary.totalDifferences,
  ]);

  // Créer la feuille de synthèse des totaux
  const synthesisTotalsSheet = XLSX.utils.aoa_to_sheet(synthesisTotalsData);
  XLSX.utils.book_append_sheet(
    workbook,
    synthesisTotalsSheet,
    "Synthèse des totaux"
  );

  // Convertir le classeur en buffer
  const excelBuffer = XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx",
  });

  return excelBuffer;
};

module.exports = {
  exportToExcel,
};
