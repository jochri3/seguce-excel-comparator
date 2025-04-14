const XLSX = require("xlsx");

/**
 * Lire et parser un fichier Excel
 * @param {string} filePath - Chemin vers le fichier Excel
 * @returns {Object} - Données parsées du fichier Excel
 */
const parseExcelFile = (filePath) => {
  try {
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

    const detectHeaderRow = (
      jsonData,
      keyword = "Matricule Fictif",
      maxRows = 10
    ) => {
      for (let i = 0; i < Math.min(maxRows, jsonData.length); i++) {
        const row = jsonData[i];
        if (
          row &&
          row.some((cell) => typeof cell === "string" && cell.includes(keyword))
        ) {
          return i;
        }
      }
      return -1;
    };

    const cleanHeaders = (headers) => {
      return headers.map((header, index) => {
        if (!header || header.toString().trim() === "") {
          return `Col_${index}`;
        }
        return header.toString().trim();
      });
    };

    const extractDataWithCleanHeaders = (jsonData, headerRowIndex) => {
      const headersRaw = jsonData[headerRowIndex] || [];
      const headers = cleanHeaders(headersRaw);

      const data = [];
      for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || row.length === 0) continue;

        const rowData = {};
        for (let j = 0; j < headers.length; j++) {
          rowData[headers[j]] = row[j];
        }
        data.push(rowData);
      }

      return {
        headers: headers.map((key) => ({ key, value: key })),
        data,
      };
    };

    const headerRowIndex = detectHeaderRow(jsonData);
    if (headerRowIndex === -1) {
      throw new Error(
        "Impossible de détecter une ligne d'en-tête contenant 'Matricule Fictif'"
      );
    }

    const { headers, data } = extractDataWithCleanHeaders(
      jsonData,
      headerRowIndex
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

  // Vérifier si les données sont valides
  if (fileAData.data.length === 0 || fileBData.data.length === 0) {
    throw new Error(
      "Un ou les deux fichiers ne contiennent pas de données exploitables"
    );
  }

  // Identifier la colonne d'identifiant unique (Matricule Fictif ou première colonne)
  let idColumn = "Matricule Fictif";

  // Vérifier si la colonne existe dans les deux fichiers
  if (
    !fileAData.data[0].hasOwnProperty(idColumn) ||
    !fileBData.data[0].hasOwnProperty(idColumn)
  ) {
    console.log(
      `La colonne '${idColumn}' n'est pas présente dans les deux fichiers. Recherche d'alternatives...`
    );

    // Chercher des alternatives ('Matricule', 'ID', etc.)
    const alternatives = ["Matricule", "ID", "Id", "id"];
    for (const alt of alternatives) {
      if (
        fileAData.data[0].hasOwnProperty(alt) &&
        fileBData.data[0].hasOwnProperty(alt)
      ) {
        idColumn = alt;
        console.log(`Utilisation de '${idColumn}' comme colonne d'identifiant`);
        break;
      }
    }

    // Si aucune alternative n'est trouvée, utiliser la première colonne
    if (
      !fileAData.data[0].hasOwnProperty(idColumn) ||
      !fileBData.data[0].hasOwnProperty(idColumn)
    ) {
      // Prendre la première clé de chaque objet
      const firstKeyA = Object.keys(fileAData.data[0])[0];
      const firstKeyB = Object.keys(fileBData.data[0])[0];

      if (firstKeyA === firstKeyB) {
        idColumn = firstKeyA;
        console.log(
          `Utilisation de la première colonne '${idColumn}' comme identifiant`
        );
      } else {
        throw new Error(
          "Impossible de trouver une colonne d'identifiant commune entre les fichiers"
        );
      }
    }
  }

  // Liste des colonnes à comparer (toutes les colonnes communes sauf l'identifiant)
  // Dans la fonction compareExcelData, remplacer le code qui trouve les colonnes communes par ceci:

  // Liste des colonnes à comparer
  const headerKeysA = fileAData.headers.map((h) => h.key);
  const headerKeysB = fileBData.headers.map((h) => h.key);

  // Créer une correspondance entre les noms de colonnes (insensible à la casse et aux espaces)
  const normalizeColumnName = (name) =>
    String(name).trim().toLowerCase().replace(/\s+/g, "");

  // Construire un dictionnaire de correspondance entre les colonnes des deux fichiers
  const columnMapping = {};
  headerKeysA.forEach((keyA) => {
    const normalizedA = normalizeColumnName(keyA);
    headerKeysB.forEach((keyB) => {
      const normalizedB = normalizeColumnName(keyB);
      if (
        normalizedA === normalizedB &&
        keyA !== idColumn &&
        keyB !== idColumn
      ) {
        // Associer la colonne du fichier A à celle du fichier B
        columnMapping[keyA] = keyB;
      }
    });
  });

  // Obtenir les colonnes communes pour la comparaison
  const commonColumns = Object.keys(columnMapping);

  console.log(
    `${commonColumns.length} colonnes communes trouvées pour la comparaison`
  );
  if (commonColumns.length === 0) {
    throw new Error(
      "Les fichiers n'ont aucune colonne commune pour la comparaison"
    );
  }

  // Colonnes numériques - utile pour le formatage et les comparaisons
  const numericColumns = [
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
    "Nbre Heure Supp. 4",
    "Salaire",
    "Ancienneté",
    "Sur Salaire",
    "Maladie",
    "Congé Circ.",
    "Ferié",
    "Congé Maternité",
    "Congé Payer",
    "Heure Supplémentaire",
    "Recuperation",
    "Prime Imposable Fixe",
    "Prime Imposable Variable",
    "Indemn. Logement",
    "Indemn. Transport",
    "Base Imposable",
    "Plafond Cnss",
    "Cnss QPO",
    "IPR",
    "Masse Salariale",
    "Net à Payer",
    "Cnss QPP",
    "Inpp",
    "Onem",
    "Total Charge Patronale",
    "Coût Salarial",
    "Frais de Services",
    "TVA 16% Frais Services",
    "Frais Admin.",
    "Total Employeur Mensuel",
  ];

  // Créer un dictionnaire pour le fichier B pour faciliter la recherche
  const fileBDict = {};
  fileBData.data.forEach((row) => {
    const idValue = row[idColumn];
    if (idValue) {
      fileBDict[idValue] = row;
    }
  });

  // Comparer chaque ligne du fichier A avec son équivalent dans le fichier B
  fileAData.data.forEach((rowA) => {
    const idValue = rowA[idColumn];

    // Si la ligne existe dans les deux fichiers
    if (idValue && fileBDict[idValue]) {
      const rowB = fileBDict[idValue];
      const rowDifferences = [];

      // Comparer chaque cellule pour les colonnes communes
      commonColumns.forEach((colA) => {
        const colB = columnMapping[colA]; // Obtenir la colonne correspondante dans B
        const valueA = parseValue(rowA[colA]);
        const valueB = parseValue(rowB[colB]);

        // Comparer les valeurs (avec gestion des types)
        if (!areValuesEqual(valueA, valueB)) {
          rowDifferences.push({
            column: colA, // Utiliser le nom de la colonne A pour l'affichage
            columnNameA: colA,
            columnNameB: colB,
            valueA,
            valueB,
            difference: calculateDifference(valueA, valueB),
            isNumeric:
              numericColumns.includes(colA) || numericColumns.includes(colB),
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
            nom: rowA["Nom"] || "N/A", // Ajoutez si ces colonnes existent
            prenom: rowA["Prénom"] || "N/A",
          },
        });
      }
    } else {
      // Ligne présente dans A mais absente dans B
      results.summary.totalDifferences++;
      results.details.push({
        id: idValue,
        onlyInFileA: true,
        rowData: rowA,
      });
    }
  });

  // Trouver les lignes qui sont dans B mais pas dans A
  fileBData.data.forEach((rowB) => {
    const idValue = rowB[idColumn];
    const rowAExists = fileAData.data.some((row) => row[idColumn] === idValue);

    if (idValue && !rowAExists) {
      results.summary.totalDifferences++;
      results.details.push({
        id: idValue,
        onlyInFileB: true,
        rowData: rowB,
      });
    }
  });

  console.log(
    `Analyse terminée: ${results.summary.totalDifferences} différences trouvées`
  );
  return results;
};

/**
 * Parse une valeur de cellule pour s'assurer qu'elle est du bon type
 * @param {any} value - Valeur à parser
 * @returns {any} - Valeur parsée
 */
const parseValue = (value) => {
  if (value === null || value === undefined) {
    return null;
  }

  // Si c'est déjà un nombre
  if (typeof value === "number") {
    return value;
  }

  // Si c'est une chaîne qui représente un nombre
  if (typeof value === "string") {
    // Enlever les virgules pour les nombres formatés (ex: "1,000.00")
    // Gérer également le format européen (ex: "1.000,00")
    let cleanValue = value;

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
    // Utiliser une petite tolérance pour les comparaisons de nombres flottants
    return Math.abs(valueA - valueB) < 0.001;
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
    return valueA - valueB;
  }

  // Si les valeurs sont des dates
  if (valueA instanceof Date && valueB instanceof Date) {
    const diffInDays =
      (valueA.getTime() - valueB.getTime()) / (1000 * 3600 * 24);
    return `${Math.abs(diffInDays)} jour(s)`;
  }

  // Pour les autres types de valeurs
  return "Valeurs différentes";
};

module.exports = {
  parseExcelFile,
  compareExcelData,
};
