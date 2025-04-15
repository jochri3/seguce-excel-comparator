const XLSX = require("xlsx");

/**
 * Lire et parser un fichier Excel
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
    });

    const sheetNames = workbook.SheetNames;
    if (sheetNames.length === 0) {
      throw new Error("Le fichier Excel ne contient aucune feuille de calcul");
    }

    const worksheet = workbook.Sheets[sheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
      raw: false,
      dateNF: "yyyy-mm-dd",
      defval: null,
      blankrows: false,
      header: 1,
    });

    // Si les données sont vides, retourner un résultat vide
    if (jsonData.length === 0) {
      return {
        sheetName: sheetNames[0],
        headers: [],
        data: [],
        rawWorksheet: worksheet,
      };
    }

    const detectHeaderRow = (
      jsonData,
      possibleKeywords = ["Matricule", "matricule", "ID", "Id", "id", "Code"],
      maxRows = 10
    ) => {
      for (let i = 0; i < Math.min(maxRows, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row) continue;

        for (let j = 0; j < row.length; j++) {
          const cell = row[j];
          if (!cell) continue;

          if (typeof cell === "string") {
            for (const keyword of possibleKeywords) {
              if (cell.includes(keyword)) {
                return i;
              }
            }
          }
        }
      }
      // Si aucun mot-clé n'est trouvé, utiliser la première ligne comme en-tête
      return 0;
    };

    const cleanHeaders = (headers) => {
      return headers.map((header, index) => {
        if (!header || header.toString().trim() === "") {
          // Ne pas générer de noms Col_XX, utiliser un espace
          return ` ${index}`;
        }
        return header.toString().trim();
      });
    };

    const extractDataWithCleanHeaders = (jsonData, headerRowIndex) => {
      if (headerRowIndex >= jsonData.length) {
        headerRowIndex = 0; // Fallback à la première ligne
      }

      const headersRaw = jsonData[headerRowIndex] || [];
      const headers = cleanHeaders(headersRaw);

      const data = [];
      for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (
          !row ||
          row.length === 0 ||
          row.every((cell) => cell === null || cell === undefined)
        )
          continue;

        const rowData = {};
        for (let j = 0; j < headers.length; j++) {
          // Ne pas inclure les colonnes avec des en-têtes uniquement numériques ou espaces
          if (headers[j] && !headers[j].match(/^\s*\d+\s*$/)) {
            rowData[headers[j]] = j < row.length ? row[j] : null;
          }
        }

        // Ne pas inclure les lignes vides
        if (Object.keys(rowData).length > 0) {
          data.push(rowData);
        }
      }

      return {
        headers: headers
          .filter((h) => !h.match(/^\s*\d+\s*$/)) // Exclure les en-têtes numériques ou espaces
          .map((key) => ({ key, value: key })),
        data,
      };
    };

    // Détecter la ligne d'en-tête
    const headerRowIndex = detectHeaderRow(jsonData);

    // Extraire les données
    const { headers, data } = extractDataWithCleanHeaders(
      jsonData,
      headerRowIndex
    );

    console.log(
      `Données extraites: ${data.length} lignes avec ${headers.length} colonnes`
    );

    return {
      sheetName: sheetNames[0],
      headers,
      data,
      rawWorksheet: worksheet,
    };
  } catch (error) {
    console.error("Erreur lors du parsing du fichier Excel:", error);
    throw error;
  }
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
    },
    details: [],
    headers: {
      fileA: fileAData.headers,
      fileB: fileBData.headers,
    },
  };

  console.log(
    `Comparaison: Fichier A a ${fileAData.data.length} lignes, Fichier B a ${fileBData.data.length} lignes`
  );

  // Si un des fichiers est vide, sortir
  if (fileAData.data.length === 0 || fileBData.data.length === 0) {
    console.log("Un des fichiers est vide, comparaison impossible");
    return results;
  }

  // Identifier la colonne d'identifiant unique
  let idColumnA = null;
  let idColumnB = null;

  // Fonction pour normaliser le nom d'une colonne
  const normalizeColumnName = (name) =>
    String(name).toLowerCase().replace(/\s+/g, "");

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
      `Aucune colonne d'ID trouvée dans A, utilisation de la première colonne: "${idColumnA}"`
    );
  }

  if (!idColumnB && fileBData.headers.length > 0) {
    idColumnB = fileBData.headers[0].key;
    console.log(
      `Aucune colonne d'ID trouvée dans B, utilisation de la première colonne: "${idColumnB}"`
    );
  }

  if (!idColumnA || !idColumnB) {
    throw new Error(
      "Impossible de trouver une colonne d'identifiant dans l'un des fichiers"
    );
  }

  // Construire un dictionnaire de correspondance entre les colonnes des deux fichiers
  const columnMapping = {};
  fileAData.headers.forEach((headerA) => {
    const normalizedA = normalizeColumnName(headerA.key);
    fileBData.headers.forEach((headerB) => {
      const normalizedB = normalizeColumnName(headerB.key);
      if (normalizedA === normalizedB) {
        columnMapping[headerA.key] = headerB.key;
      }
    });
  });

  // Obtenir les colonnes communes pour la comparaison
  const commonColumns = Object.keys(columnMapping).filter(
    (col) => col !== idColumnA
  );

  console.log(
    `${commonColumns.length} colonnes communes trouvées pour la comparaison`
  );

  if (commonColumns.length === 0) {
    console.log(
      "Aucune colonne commune trouvée. Vérifiez les en-têtes des fichiers."
    );
    return results;
  }

  // Déterminer les colonnes numériques
  const numericColumns = [];

  // Analyser les données pour détecter les colonnes qui semblent contenir des nombres
  const detectNumericColumns = () => {
    commonColumns.forEach((colA) => {
      const colB = columnMapping[colA];

      // Vérifier si les valeurs dans cette colonne sont majoritairement numériques
      let numericCount = 0;
      const sampleSize = Math.min(fileAData.data.length, 10); // Vérifier les 10 premières lignes

      for (let i = 0; i < sampleSize; i++) {
        if (i >= fileAData.data.length) break;

        const valueA = fileAData.data[i][colA];
        if (valueA !== null && valueA !== undefined) {
          const parsedValue = parseValue(valueA);
          if (typeof parsedValue === "number") {
            numericCount++;
          }
        }
      }

      // Si plus de la moitié des valeurs sont numériques, considérer la colonne comme numérique
      if (numericCount > sampleSize / 2) {
        numericColumns.push(colA);
      }
    });
  };

  detectNumericColumns();
  console.log(`Colonnes numériques détectées: ${numericColumns.join(", ")}`);

  // Créer un dictionnaire pour le fichier B pour faciliter la recherche
  const fileBDict = {};
  fileBData.data.forEach((row) => {
    if (row[idColumnB] !== undefined && row[idColumnB] !== null) {
      const idValue = String(row[idColumnB]).trim();
      fileBDict[idValue] = row;
    }
  });

  // Ensemble pour vérifier les doublons de matricules
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
    matriculesA.size < fileAData.data.length ||
    matriculesB.size < fileBData.data.length;

  // Comparer chaque ligne du fichier A avec son équivalent dans le fichier B
  fileAData.data.forEach((rowA) => {
    if (rowA[idColumnA] === undefined || rowA[idColumnA] === null) return;

    const idValue = String(rowA[idColumnA]).trim();

    // Si la ligne existe dans les deux fichiers
    if (fileBDict[idValue]) {
      const rowB = fileBDict[idValue];
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
            isNumeric: numericColumns.includes(colA),
          });
        }
      });

      // Si des différences ont été trouvées
      if (rowDifferences.length > 0) {
        results.summary.totalDifferences++;

        // Ajouter les colonnes avec différences au résumé
        rowDifferences.forEach((diff) => {
          const existingColDiff = results.summary.columnDifferences.find(
            (c) => c.column === diff.column
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
          rowData: {
            matricule: idValue,
          },
        });
      }
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
  fileBData.data.forEach((rowB) => {
    if (rowB[idColumnB] === undefined || rowB[idColumnB] === null) return;

    const idValue = String(rowB[idColumnB]).trim();
    const rowAExists = fileAData.data.some(
      (row) =>
        row[idColumnA] !== undefined &&
        String(row[idColumnA]).trim() === idValue
    );

    if (!rowAExists) {
      results.summary.totalDifferences++;
      results.details.push({
        id: idValue,
        onlyInFileB: true,
        rowData: { matricule: idValue, ...rowB },
      });
    }
  });

  console.log(
    `Analyse terminée: ${results.summary.totalDifferences} différences trouvées`
  );

  // Ajouter les informations de matricules au résultat
  results.matriculeCount = matriculesA.size;
  results.hasDuplicates = results.hasDuplicates;

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

module.exports = {
  parseExcelFile,
  compareExcelData,
};
