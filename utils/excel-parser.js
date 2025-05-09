const XLSX = require("xlsx");

/**
 * Lire et parser un fichier Excel avec extraction des formules
 * @param {string} filePath - Chemin vers le fichier Excel
 * @returns {Object} - Données parsées du fichier Excel
 */
const parseExcelFile = (filePath) => {
  try {
    console.log(`Lecture du fichier: ${filePath}`);

    const workbook = XLSX.readFile(filePath, {
      cellDates: true,
      cellNF: true,
      cellFormula: true,
      raw: false,
    });

    const sheetNames = workbook.SheetNames;
    if (sheetNames.length === 0) {
      throw new Error("Le fichier Excel ne contient aucune feuille de calcul");
    }

    const worksheet = workbook.Sheets[sheetNames[0]];
    const range = XLSX.utils.decode_range(worksheet["!ref"]);

    // Extraire les données et les formules séparément
    const data = [];
    const formulas = {};
    const headers = [];

    // D'abord, détecter la ligne d'en-tête
    let headerRow = 0;
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = worksheet[cellAddress];

        if (cell && cell.t === "s" && cell.v) {
          // Chercher des mots-clés typiques d'en-têtes
          if (
            cell.v.includes("Matricule") ||
            cell.v.includes("BU") ||
            cell.v.includes("Salaire") ||
            cell.v.includes("Charge")
          ) {
            headerRow = r;
            break;
          }
        }
      }
      if (headerRow > 0) break;
    }

    // Extraire les en-têtes
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c });
      const cell = worksheet[cellAddress];
      headers.push(cell ? cell.v : `Col${c}`);
    }

    // Extraire les données et formules
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const rowData = {};
      const rowFormulas = {};
      let hasData = false;

      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = worksheet[cellAddress];
        const header = headers[c - range.s.c];

        if (cell) {
          // Extraire la valeur
          rowData[header] = cell.v;
          hasData = true;

          // Extraire la formule si elle existe
          if (cell.f) {
            rowFormulas[header] = cell.f;
          }
        }
      }

      if (hasData) {
        data.push(rowData);
        if (Object.keys(rowFormulas).length > 0) {
          formulas[r - headerRow - 1] = rowFormulas;
        }
      }
    }

    return {
      sheetName: sheetNames[0],
      headers: headers.map((h) => ({ key: h, value: h })),
      data,
      formulas,
      rawWorksheet: worksheet,
    };
  } catch (error) {
    console.error("Erreur lors du parsing du fichier Excel:", error);
    throw error;
  }
};

const getLexiconData = async () => {
  const { getDatabase } = require("../config/database");
  const db = await getDatabase();

  return new Promise((resolve, reject) => {
    db.all("SELECT * FROM lexicon_columns", [], (err, rows) => {
      if (err) {
        reject(err);
        return;
      }

      const lexicon = {
        columnsByName: {},
        fixedColumns: [],
        variableColumns: [],
      };

      rows.forEach((row) => {
        lexicon.columnsByName[row.column_name] = row;

        if (row.column_type === "fixe") {
          lexicon.fixedColumns.push(row.column_name);
        } else {
          lexicon.variableColumns.push(row.column_name);
        }
      });

      resolve(lexicon);
    });
  });
};

/**
 * Classifier les colonnes en fixes et variables
 */
const classifyColumns = (columns) => {
  // Liste des éléments fixes selon la spécification
  const fixedElements = [
    "matricule",
    "bu",
    "pers à charge",
    "enfant légal",
    "salaire mensuel",
    "ancienneté mensuel",
    "sur salaire mensuel",
    "taux horaire",
    "transport mensuel",
    "logement mensuel",
    "prime astreinte",
    "forfait heures supplémentaire",
    "prime de détachement",
    "prime imposable fixe",
  ];

  // Normaliser les noms pour la comparaison
  const normalizeForComparison = (name) => {
    return name.toLowerCase().replace(/[^a-z0-9]/g, "");
  };

  const fixedColumns = [];
  const variableColumns = [];

  columns.forEach((column) => {
    const normalizedColumn = normalizeForComparison(column);
    const isFixed = fixedElements.some((fixedElement) =>
      normalizedColumn.includes(normalizeForComparison(fixedElement)),
    );

    if (isFixed) {
      fixedColumns.push(column);
    } else {
      variableColumns.push(column);
    }
  });

  return { fixedColumns, variableColumns };
};

/**
 * Comparer les données de deux fichiers Excel
 * @param {Object} fileAData - Données du premier fichier Excel
 * @param {Object} fileBData - Données du deuxième fichier Excel
 * @returns {Object} - Résultats de la comparaison
 */
const compareExcelData = (fileAData, fileBData) => {
  const results = {
    summary: {
      totalRows: {
        fileA: fileAData.data.length,
        fileB: fileBData.data.length,
      },
      totalDifferences: 0,
      columnDifferences: [],
      columnStats: {
        commonColumns: [],
        missingInA: [],
        missingInB: [],
      },
      duplicates: {
        fileA: [],
        fileB: [],
      },
    },
    details: [],
    headers: {
      fileA: fileAData.headers,
      fileB: fileBData.headers,
    },
  };

  console.log(
    `Comparaison: Fichier A a ${fileAData.data.length} lignes, Fichier B a ${fileBData.data.length} lignes`,
  );

  // Si un des fichiers est vide, sortir
  if (fileAData.data.length === 0 || fileBData.data.length === 0) {
    console.log("Un des fichiers est vide, comparaison impossible");
    return results;
  }

  // Fonction pour normaliser le nom d'une colonne
  const normalizeColumnName = (name) =>
    String(name).toLowerCase().replace(/\s+/g, "");

  // Identifier la colonne d'identifiant unique
  let idColumnA = null;
  let idColumnB = null;

  // Trouver les colonnes d'ID potentielles
  const possibleIdColumns = [
    "matriculefictif",
    "matricule",
    "id",
    "code",
    "reference",
    "employeid",
  ];

  // Chercher dans les en-têtes du fichier A
  for (const header of fileAData.headers) {
    const normalizedName = normalizeColumnName(header.key);
    for (const idCol of possibleIdColumns) {
      if (normalizedName.includes(idCol)) {
        idColumnA = header.key;
        console.log(`Colonne d'ID trouvée dans A: "${idColumnA}"`);
        break;
      }
    }
    if (idColumnA) break;
  }

  // Chercher dans les en-têtes du fichier B
  for (const header of fileBData.headers) {
    const normalizedName = normalizeColumnName(header.key);
    for (const idCol of possibleIdColumns) {
      if (normalizedName.includes(idCol)) {
        idColumnB = header.key;
        console.log(`Colonne d'ID trouvée dans B: "${idColumnB}"`);
        break;
      }
    }
    if (idColumnB) break;
  }

  // Si aucune colonne d'ID n'est trouvée, utiliser la première colonne
  if (!idColumnA && fileAData.headers.length > 0) {
    idColumnA = fileAData.headers[0].key;
    console.log(
      `Aucune colonne d'ID trouvée dans A, utilisation de la première colonne: "${idColumnA}"`,
    );
  }

  if (!idColumnB && fileBData.headers.length > 0) {
    idColumnB = fileBData.headers[0].key;
    console.log(
      `Aucune colonne d'ID trouvée dans B, utilisation de la première colonne: "${idColumnB}"`,
    );
  }

  if (!idColumnA || !idColumnB) {
    throw new Error(
      "Impossible de trouver une colonne d'identifiant dans l'un des fichiers",
    );
  }

  // Construire la correspondance des colonnes en utilisant la normalisation
  const columnMapping = {};
  const columnsA = new Set();
  const columnsB = new Set();

  fileAData.headers.forEach((headerA) => {
    const normalizedA = normalizeColumnName(headerA.key);
    columnsA.add(headerA.key);

    fileBData.headers.forEach((headerB) => {
      const normalizedB = normalizeColumnName(headerB.key);
      if (normalizedA === normalizedB) {
        columnMapping[headerA.key] = headerB.key;
      }
    });
  });

  fileBData.headers.forEach((headerB) => {
    columnsB.add(headerB.key);
  });

  // Obtenir les colonnes communes et les colonnes manquantes
  const commonColumns = Object.keys(columnMapping).filter(
    (col) => col !== idColumnA,
  );

  const missingInA = Array.from(columnsB).filter(
    (colB) =>
      !Array.from(columnsA).some(
        (colA) => normalizeColumnName(colA) === normalizeColumnName(colB),
      ),
  );

  const missingInB = Array.from(columnsA).filter(
    (colA) =>
      !Array.from(columnsB).some(
        (colB) => normalizeColumnName(colA) === normalizeColumnName(colB),
      ),
  );

  // Mettre à jour les statistiques
  results.summary.columnStats.commonColumns = commonColumns;
  results.summary.columnStats.missingInA = missingInA;
  results.summary.columnStats.missingInB = missingInB;

  // Classifier les colonnes
  const { fixedColumns, variableColumns } = classifyColumns(commonColumns);

  // Résultats de comparaison détaillés
  results.summary.sequentialComparison = {
    fixedElements: {
      totalColumns: fixedColumns.length,
      columnsWithErrors: [],
      totalErrors: 0,
    },
    variableElements: {
      totalColumns: variableColumns.length,
      columnsWithErrors: [],
      totalErrors: 0,
    },
  };

  console.log(`Éléments fixes détectés: ${fixedColumns.length}`);
  console.log(`Éléments variables détectés: ${variableColumns.length}`);
  console.log(`${commonColumns.length} colonnes communes trouvées`);
  console.log(`${missingInA.length} colonnes présentes uniquement dans B`);
  console.log(`${missingInB.length} colonnes présentes uniquement dans A`);

  // Détection des doublons de matricules
  const detectDuplicates = (data, idColumn, fileLabel) => {
    const matriculeCounts = {};
    const duplicates = [];

    data.forEach((row, index) => {
      const matricule = String(row[idColumn]).trim();
      if (matricule) {
        if (!matriculeCounts[matricule]) {
          matriculeCounts[matricule] = 0;
        }
        matriculeCounts[matricule]++;
      }
    });

    Object.entries(matriculeCounts).forEach(([matricule, count]) => {
      if (count > 1) {
        duplicates.push({
          matricule,
          count,
          rows: data
            .map((row, index) =>
              String(row[idColumn]).trim() === matricule ? index + 1 : null,
            )
            .filter((i) => i !== null),
        });
      }
    });

    results.summary.duplicates[fileLabel] = duplicates;
  };

  detectDuplicates(fileAData.data, idColumnA, "fileA");
  detectDuplicates(fileBData.data, idColumnB, "fileB");

  // Créer un dictionnaire pour le fichier B pour faciliter la recherche
  const fileBDict = {};
  fileBData.data.forEach((row) => {
    if (row[idColumnB] !== undefined && row[idColumnB] !== null) {
      const idValue = String(row[idColumnB]).trim();
      if (!fileBDict[idValue]) {
        fileBDict[idValue] = [];
      }
      fileBDict[idValue].push(row);
    }
  });

  // Ensemble pour vérifier les matricules uniques
  const matriculesA = new Set();
  const matriculesB = new Set();

  // Compter les matricules uniques
  fileAData.data.forEach((row) => {
    if (row[idColumnA] !== undefined && row[idColumnA] !== null) {
      matriculesA.add(String(row[idColumnA]).trim());
    }
  });

  fileBData.data.forEach((row) => {
    if (row[idColumnB] !== undefined && row[idColumnB] !== null) {
      matriculesB.add(String(row[idColumnB]).trim());
    }
  });

  // Mettre à jour le nombre total de matricules contrôlés
  results.matriculeCount = matriculesA.size;

  // Vérifier les doublons
  results.hasDuplicates =
    results.summary.duplicates.fileA.length > 0 ||
    results.summary.duplicates.fileB.length > 0;

  // Fonction pour comparer une catégorie de colonnes
  const compareColumnCategory = (columns, category) => {
    columns.forEach((colA) => {
      const colB = columnMapping[colA];
      const columnErrors = [];

      fileAData.data.forEach((rowA, index) => {
        const idValue = String(rowA[idColumnA]).trim();

        if (fileBDict[idValue]) {
          const rowB = fileBDict[idValue][0]; // Pour simplifier, prendre le premier
          const valueA = parseValue(rowA[colA]);
          const valueB = parseValue(rowB[colB]);

          if (!areValuesEqual(valueA, valueB)) {
            columnErrors.push({
              matricule: idValue,
              rowIndex: index,
              valueA,
              valueB,
              difference: calculateDifference(valueA, valueB),
            });
          }
        }
      });

      if (columnErrors.length > 0) {
        results.summary.sequentialComparison[category].columnsWithErrors.push({
          column: colA,
          errorCount: columnErrors.length,
          errors: columnErrors,
        });
        results.summary.sequentialComparison[category].totalErrors +=
          columnErrors.length;
      }
    });
  };

  // Vérifier d'abord les éléments fixes
  compareColumnCategory(fixedColumns, "fixedElements");

  // Ensuite les éléments variables
  compareColumnCategory(variableColumns, "variableElements");

  // Comparer chaque ligne du fichier A avec son équivalent dans le fichier B
  fileAData.data.forEach((rowA) => {
    if (rowA[idColumnA] === undefined || rowA[idColumnA] === null) return;

    const idValue = String(rowA[idColumnA]).trim();

    // Si la ligne existe dans les deux fichiers
    if (fileBDict[idValue]) {
      fileBDict[idValue].forEach((rowB) => {
        const rowDifferences = [];

        // Comparer chaque cellule pour les colonnes communes
        commonColumns.forEach((colA) => {
          const colB = columnMapping[colA];
          const valueA = parseValue(rowA[colA]);
          const valueB = parseValue(rowB[colB]);

          // Ne pas considérer comme différence si les deux valeurs sont null/undefined
          if (
            (valueA === null && valueB === null) ||
            (valueA === undefined && valueB === undefined)
          ) {
            return;
          }

          // Comparer les valeurs
          if (!areValuesEqual(valueA, valueB)) {
            rowDifferences.push({
              column: colA,
              columnNameA: colA,
              columnNameB: colB,
              valueA,
              valueB,
              difference: calculateDifference(valueA, valueB),
              isNumeric: isNumericColumn(colA, fileAData.data),
            });
          }
        });

        // Si des différences ont été trouvées
        if (rowDifferences.length > 0) {
          results.summary.totalDifferences++;

          // Catégoriser les différences
          const categorizedDifferences = {
            fixed: [],
            variable: [],
          };

          rowDifferences.forEach((diff) => {
            if (fixedColumns.includes(diff.column)) {
              categorizedDifferences.fixed.push(diff);
            } else {
              categorizedDifferences.variable.push(diff);
            }
          });

          // Ajouter les colonnes avec différences au résumé
          rowDifferences.forEach((diff) => {
            const existingColDiff = results.summary.columnDifferences.find(
              (c) => c.column === diff.column,
            );

            if (existingColDiff) {
              existingColDiff.count++;
            } else {
              results.summary.columnDifferences.push({
                column: diff.column,
                columnNameA: diff.columnNameA,
                columnNameB: diff.columnNameB,
                count: 1,
                isNumeric: diff.isNumeric,
              });
            }
          });

          // Ajouter les détails complets
          results.details.push({
            id: idValue,
            differences: rowDifferences,
            categorizedDifferences,
            rowData: {
              matricule: idValue,
            },
          });
        }
      });
    } else {
      // Ligne présente dans A mais absente dans B
      results.summary.totalDifferences++;
      results.details.push({
        id: idValue,
        onlyInFileA: true,
        rowData: { matricule: idValue, ...rowA },
      });
    }
  });

  // Trouver les lignes qui sont dans B mais pas dans A
  matriculesB.forEach((idValue) => {
    if (!matriculesA.has(idValue)) {
      const rowB = fileBData.data.find(
        (row) => String(row[idColumnB]).trim() === idValue,
      );
      if (rowB) {
        results.summary.totalDifferences++;
        results.details.push({
          id: idValue,
          onlyInFileB: true,
          rowData: { matricule: idValue, ...rowB },
        });
      }
    }
  });

  console.log(
    `Analyse terminée: ${results.summary.totalDifferences} différences trouvées`,
  );

  return results;
};

/**
 * Parse une valeur de cellule pour s'assurer qu'elle est du bon type
 * @param {any} value - Valeur à parser
 * @returns {any} - Valeur parsée
 */
const parseValue = (value) => {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  // Si c'est déjà un nombre
  if (typeof value === "number") {
    return value;
  }

  // Si c'est une chaîne qui représente un nombre
  if (typeof value === "string") {
    // Ignorer les chaînes vides après nettoyage
    const trimmed = value.trim();
    if (trimmed === "") return null;

    try {
      // Enlever les espaces et caractères non imprimables
      let cleanValue = trimmed.replace(/\s/g, "");

      // Format européen: convertir "1.234,56" en "1234.56"
      if (
        cleanValue.includes(",") &&
        cleanValue.includes(".") &&
        cleanValue.lastIndexOf(",") > cleanValue.lastIndexOf(".")
      ) {
        cleanValue = cleanValue.replace(/\./g, "").replace(",", ".");
      }
      // Format américain: convertir "1,234.56" en "1234.56"
      else if (
        cleanValue.includes(",") &&
        cleanValue.includes(".") &&
        cleanValue.lastIndexOf(".") > cleanValue.lastIndexOf(",")
      ) {
        cleanValue = cleanValue.replace(/,/g, "");
      }
      // S'il n'y a qu'une virgule, présumer que c'est un séparateur décimal européen
      else if (cleanValue.includes(",") && !cleanValue.includes(".")) {
        cleanValue = cleanValue.replace(",", ".");
      }

      const numValue = parseFloat(cleanValue);
      if (!isNaN(numValue)) {
        return numValue;
      }
    } catch (e) {
      // En cas d'erreur, retourner la valeur d'origine
      return value;
    }
  }

  return value;
};

/**
 * Comparer deux valeurs pour déterminer si elles sont égales
 * @param {any} valueA - Première valeur
 * @param {any} valueB - Deuxième valeur
 * @returns {boolean} - True si les valeurs sont égales, false sinon
 */
const areValuesEqual = (valueA, valueB) => {
  // Si les deux valeurs sont null ou undefined, elles sont considérées égales
  if (valueA === null && valueB === null) return true;
  if (valueA === undefined && valueB === undefined) return true;

  // Si une seule valeur est null/undefined, elles sont différentes
  if (
    valueA === null ||
    valueA === undefined ||
    valueB === null ||
    valueB === undefined
  ) {
    return false;
  }

  // Gestion des dates
  if (valueA instanceof Date && valueB instanceof Date) {
    return valueA.getTime() === valueB.getTime();
  }

  // Gestion des nombres
  if (typeof valueA === "number" && typeof valueB === "number") {
    // Utiliser une tolérance plus large pour les comparaisons de nombres flottants
    // pour éviter les fausses différences dues aux erreurs d'arrondi
    return Math.abs(valueA - valueB) < 0.01;
  }

  // Conversion en chaîne pour la comparaison générale
  if (typeof valueA === "string" && typeof valueB === "string") {
    return valueA.trim() === valueB.trim();
  }

  // Si les types sont différents mais qu'un est string et l'autre number,
  // essayer de les comparer comme des nombres
  if (
    (typeof valueA === "string" && typeof valueB === "number") ||
    (typeof valueA === "number" && typeof valueB === "string")
  ) {
    const numA = typeof valueA === "string" ? parseFloat(valueA) : valueA;
    const numB = typeof valueB === "string" ? parseFloat(valueB) : valueB;

    if (!isNaN(numA) && !isNaN(numB)) {
      return Math.abs(numA - numB) < 0.01;
    }
  }

  // Conversion en chaîne pour la comparaison générale
  return String(valueA).trim() === String(valueB).trim();
};

/**
 * Calculer la différence entre deux valeurs
 * @param {any} valueA - Première valeur
 * @param {any} valueB - Deuxième valeur
 * @returns {string|number} - Différence calculée
 */
const calculateDifference = (valueA, valueB) => {
  // Si les valeurs sont numériques
  if (typeof valueA === "number" && typeof valueB === "number") {
    return valueB - valueA;
  }

  // Si une valeur est un nombre et l'autre une chaîne représentant un nombre
  if (
    (typeof valueA === "string" && typeof valueB === "number") ||
    (typeof valueA === "number" && typeof valueB === "string")
  ) {
    const numA = typeof valueA === "string" ? parseFloat(valueA) : valueA;
    const numB = typeof valueB === "string" ? parseFloat(valueB) : valueB;

    if (!isNaN(numA) && !isNaN(numB)) {
      return numB - numA;
    }
  }

  // Si les valeurs sont des dates
  if (valueA instanceof Date && valueB instanceof Date) {
    const diffInDays =
      (valueB.getTime() - valueA.getTime()) / (1000 * 3600 * 24);
    return diffInDays > 0 ? `+${diffInDays} jour(s)` : `${diffInDays} jour(s)`;
  }

  // Pour les autres types de valeurs, si elles sont égales
  if (String(valueA).trim() === String(valueB).trim()) {
    return "Identique";
  }

  // Pour les autres types de valeurs
  return "Valeurs différentes";
};

/**
 * Vérifier si une colonne contient principalement des valeurs numériques
 */
const isNumericColumn = (columnName, data) => {
  const sampleSize = Math.min(10, data.length);
  let numericCount = 0;

  for (let i = 0; i < sampleSize && i < data.length; i++) {
    const value = data[i][columnName];
    if (value !== null && value !== undefined) {
      const parsedValue = parseValue(value);
      if (typeof parsedValue === "number") {
        numericCount++;
      }
    }
  }

  return numericCount > sampleSize / 2;
};

/**
 * Extraire le mois et l'année d'un nom de fichier
 * Format attendu: NOM_DU_FICHIER_NUMERO_MOIS_ANNEE
 */
const extractDateFromFilename = (filename) => {
  const match = filename.match(/(\d{2})_(\d{4})/);
  if (match) {
    return {
      month: parseInt(match[1], 10),
      year: parseInt(match[2], 10),
    };
  }
  return null;
};

/**
 * Détecter le type de prestataire en fonction des colonnes
 */
const detectProviderType = (headers) => {
  const headerNames = headers.map((h) => h.key.toLowerCase());

  // Colonnes spécifiques au nouveau prestataire
  const newProviderColumns = [
    "prime migration it",
    "aide rentrée scolaire",
    "tva 16% frais services",
  ];

  // Vérifier la présence des colonnes spécifiques
  const hasNewProviderColumns = newProviderColumns.some((col) =>
    headerNames.includes(col),
  );

  return hasNewProviderColumns ? "nouveau" : "ancien";
};

module.exports = {
  parseExcelFile,
  compareExcelData,
  extractDateFromFilename,
  detectProviderType,
  classifyColumns,
};
