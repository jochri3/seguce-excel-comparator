const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const excelParser = require("./utils/excel-parser");
const excelExport = require("./utils/export-excel");
const SessionManager = require("./utils/session-manager");
const { initDatabase, getDatabase } = require("./config/database");

// Configuration de l'application
const app = express();
const PORT = process.env.PORT || 3000;

// Initialiser la base de données au démarrage
initDatabase()
  .then(() => console.log("Base de données prête"))
  .catch((err) => console.error("Erreur d'initialisation DB:", err));

// Configuration des dossiers statiques et du moteur de template
app.use(express.static(path.join(__dirname, "public")));
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

// Ajouter le body parser pour lire le body des requêtes POST
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Configuration de Multer pour les uploads de fichiers
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, "uploads");
    // Vérifier si le dossier existe, sinon le créer
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    // Ajouter un timestamp pour éviter les conflits de noms de fichiers
    const uniqueSuffix = Date.now() + "-" + Math.round(Math.random() * 1e9);
    const fileExtension = path.extname(file.originalname);
    cb(null, file.fieldname + "-" + uniqueSuffix + fileExtension);
  },
});

// Filtre pour accepter uniquement les fichiers Excel
const fileFilter = (req, file, cb) => {
  const allowedFileTypes = [".xlsx", ".xls"];
  const extname = path.extname(file.originalname).toLowerCase();

  if (allowedFileTypes.includes(extname)) {
    cb(null, true);
  } else {
    cb(new Error("Seuls les fichiers Excel (.xlsx, .xls) sont acceptés"));
  }
};

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: 10 * 1024 * 1024 }, // Limite de 10 MB
});

// Routes
app.get("/", (req, res) => {
  res.render("index", { title: "SEGUGE Wages reconciliation" });
});

// Route pour traiter l'upload des deux fichiers Excel
// Route pour traiter l'upload des deux fichiers Excel
app.post(
  "/compare",
  upload.fields([
    { name: "fileA", maxCount: 1 },
    { name: "fileB", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      if (!req.files || !req.files.fileA || !req.files.fileB) {
        return res.status(400).render("error", {
          title: "Erreur",
          message: "Veuillez uploader les deux fichiers Excel",
        });
      }

      const fileAPath = req.files.fileA[0].path;
      const fileBPath = req.files.fileB[0].path;
      const fileAName = req.files.fileA[0].originalname;
      const fileBName = req.files.fileB[0].originalname;

      // Traitement des fichiers Excel
      const fileAData = await excelParser.parseExcelFile(fileAPath);
      const fileBData = await excelParser.parseExcelFile(fileBPath);

      // Extraire les informations de date du nom de fichier
      const dateInfoA = excelParser.extractDateFromFilename(fileAName);
      const dateInfoB = excelParser.extractDateFromFilename(fileBName);

      if (
        !dateInfoA ||
        !dateInfoB ||
        dateInfoA.month !== dateInfoB.month ||
        dateInfoA.year !== dateInfoB.year
      ) {
        return res.status(400).render("error", {
          title: "Erreur",
          message:
            "Les deux fichiers doivent appartenir à la même période (mois/année)",
        });
      }

      const { month, year } = dateInfoA;

      // Détecter le type de prestataire
      const providerType = excelParser.detectProviderType(fileAData.headers);

      // Obtenir ou créer la session pour ce mois
      const session = await SessionManager.getOrCreateSession(month, year);

      // Vérifier si les fichiers existent déjà
      const existingFiles = await SessionManager.getSessionHistory(
        session.session_id,
      );

      let fileAVersion = 1;
      let fileBVersion = 1;
      let isUpdate = false;

      // Vérifier si les fichiers ont déjà été uploadés
      existingFiles.forEach((file) => {
        if (file.file_type === "fileA" && file.file_name === fileAName) {
          fileAVersion = file.version + 1;
          isUpdate = true;
        }
        if (file.file_type === "fileB" && file.file_name === fileBName) {
          fileBVersion = file.version + 1;
          isUpdate = true;
        }
      });

      // Sauvegarder les nouvelles versions des fichiers
      const fileASaved = await SessionManager.saveFileVersion(
        session.session_id,
        "fileA",
        fileAName,
        fileAData.data,
        fileAData.formulas,
        "System", // À remplacer par le vrai utilisateur quand l'auth est implémentée
      );

      const fileBSaved = await SessionManager.saveFileVersion(
        session.session_id,
        "fileB",
        fileBName,
        fileBData.data,
        fileBData.formulas,
        "System",
      );

      // Réconciliation des données
      const comparisonResult = excelParser.compareExcelData(
        fileAData,
        fileBData,
      );

      // Sauvegarder les résultats de comparaison
      await SessionManager.saveComparisonResult(
        session.session_id,
        fileASaved.version, // ou Math.max(fileASaved.version, fileBSaved.version)
        comparisonResult,
      );

      // Calculer les totaux pour l'affichage
      const summary = calculateSummaryData(comparisonResult);

      // Rendre la page de comparaison avec les résultats
      req.app.locals.lastComparisonResult = comparisonResult;
      req.app.locals.fileAName = fileAName;
      req.app.locals.fileBName = fileBName;
      req.app.locals.fileAData = fileAData;
      req.app.locals.fileBData = fileBData;
      res.render("compare", {
        title: "Résultats de la réconciliation",
        fileAName,
        fileBName,
        comparisonResult,
        summary,
        session,
        isUpdate,
        providerType,
        fileAVersion: fileASaved.version,
        fileBVersion: fileBSaved.version,
      });

      // Nettoyer les fichiers uploadés après traitement
      setTimeout(() => {
        fs.unlinkSync(fileAPath);
        fs.unlinkSync(fileBPath);
      }, 5000);
    } catch (error) {
      console.error("Erreur lors de la comparaison des fichiers:", error);
      res.status(500).render("error", {
        title: "Erreur",
        message:
          "Une erreur est survenue lors de la comparaison des fichiers: " +
          error.message,
      });
    }
  },
);
// Route pour exporter les résultats en Excel
// Route pour exporter les résultats en Excel avec formules
app.get("/export-excel", (req, res) => {
  try {
    const comparisonResult = req.app.locals.lastComparisonResult;
    const fileAData = req.app.locals.fileAData;
    const fileBData = req.app.locals.fileBData;
    const fileAName = req.app.locals.fileAName;
    const fileBName = req.app.locals.fileBName;

    if (!comparisonResult) {
      return res
        .status(400)
        .send(
          "Aucun résultat de comparaison disponible. Veuillez d'abord comparer deux fichiers.",
        );
    }

    // Générer le fichier Excel
    const excelBuffer = excelExport.exportToExcel(
      comparisonResult,
      fileAData,
      fileBData,
      fileAName,
      fileBName,
    );

    // Envoyer le fichier au client
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=reconciliation_paie.xlsx",
    );
    return res.send(excelBuffer);
  } catch (error) {
    console.error("Erreur lors de l'export Excel:", error);
    res
      .status(500)
      .send("Une erreur est survenue lors de l'export: " + error.message);
  }
});

// Route pour afficher l'historique
app.get("/history", async (req, res) => {
  try {
    const months = await SessionManager.getAllSessionMonths();
    res.render("history", {
      title: "Historique des réconciliations",
      months,
    });
  } catch (error) {
    console.error("Erreur lors de la récupération de l'historique:", error);
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur lors de la récupération de l'historique",
    });
  }
});

// Route pour afficher les détails d'une session
app.get("/history/:sessionId", async (req, res) => {
  try {
    const sessionId = req.params.sessionId;
    const history = await SessionManager.getSessionHistory(sessionId);

    const groupedFiles = {
      fileA: [],
      fileB: [],
    };

    history.forEach((file) => {
      groupedFiles[file.file_type].push(file);
    });

    res.render("session-history", {
      title: `Historique - ${sessionId}`,
      sessionId,
      groupedFiles,
    });
  } catch (error) {
    console.error(
      "Erreur lors de la récupération de l'historique de session:",
      error,
    );
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur lors de la récupération de l'historique de session",
    });
  }
});

// Fonction pour calculer les données de synthèse
function calculateSummaryData(comparisonResult) {
  // Initialiser les totaux
  const totals = {
    fileA: {
      "Cnss QPO": 0,
      IPR: 0,
      "Cnss QPP": 0,
      Inpp: 0,
      Onem: 0,
      "Total Charge Patronale": 0,
      "Coût Salarial": 0,
      "Masse Salariale": 0,
      "Net à Payer": 0,
      "Net à Payer arrondi": 0,
      "Frais de Services": 0,
      "TVA 16% Frais Services": 0,
      "Total Employeur Mensuel": 0,
    },
    fileB: {
      "Cnss QPO": 0,
      IPR: 0,
      "Cnss QPP": 0,
      Inpp: 0,
      Onem: 0,
      "Total Charge Patronale": 0,
      "Coût Salarial": 0,
      "Masse Salariale": 0,
      "Net à Payer": 0,
      "Net à Payer arrondi": 0,
      "Frais de Services": 0,
      "TVA 16% Frais Services": 0,
      "Total Employeur Mensuel": 0,
    },
  };

  // Récupérer les informations sur les matricules directement du résultat
  let matriculeCount = comparisonResult.matriculeCount || 0;
  let hasDuplicates = comparisonResult.hasDuplicates || false;

  // Si ces informations ne sont pas directement disponibles, les calculer
  if (matriculeCount === 0 && comparisonResult.details.length > 0) {
    // Vérifier les doublons de matricules
    const matricules = new Set();
    comparisonResult.details.forEach((detail) => {
      if (detail.id) {
        matricules.add(detail.id);
      }
    });

    matriculeCount = matricules.size;
    hasDuplicates =
      matriculeCount < comparisonResult.summary.totalRows.fileA ||
      matriculeCount < comparisonResult.summary.totalRows.fileB;
  }

  // Détecter les colonnes qui correspondent aux totaux que nous recherchons
  const detectColumn = (columnName) => {
    if (!columnName) return null;

    const normalizedName = columnName.toLowerCase().replace(/\s+/g, "");

    if (normalizedName.includes("cnssqpo") || normalizedName.includes("qpo"))
      return "Cnss QPO";
    if (normalizedName.includes("ipr")) return "IPR";
    if (normalizedName.includes("cnssqpp") || normalizedName.includes("qpp"))
      return "Cnss QPP";
    if (normalizedName.includes("inpp")) return "Inpp";
    if (normalizedName.includes("onem")) return "Onem";
    if (
      normalizedName.includes("totalchargepatronale") ||
      normalizedName.includes("chargepatronale")
    )
      return "Total Charge Patronale";
    if (
      normalizedName.includes("coutsalarial") ||
      normalizedName.includes("coûtsalarial")
    )
      return "Coût Salarial";
    if (normalizedName.includes("massesalariale")) return "Masse Salariale";
    if (
      normalizedName.includes("netàpayer") ||
      normalizedName.includes("netapayer")
    )
      return "Net à Payer";
    if (
      normalizedName.includes("fraisdeservices") ||
      normalizedName.includes("fraisservices")
    )
      return "Frais de Services";
    if (normalizedName.includes("tva16") || normalizedName.includes("tvafrais"))
      return "TVA 16% Frais Services";
    if (normalizedName.includes("totalemployeurmensuel"))
      return "Total Employeur Mensuel";

    return null;
  };

  // Parcourir les différences pour extraire les valeurs des colonnes pertinentes
  comparisonResult.details.forEach((detail) => {
    if (detail.differences) {
      detail.differences.forEach((diff) => {
        const category = detectColumn(diff.column);
        if (category) {
          if (typeof diff.valueA === "number")
            totals.fileA[category] += diff.valueA;
          if (typeof diff.valueB === "number")
            totals.fileB[category] += diff.valueB;
        }
      });
    }
  });

  // Arrondir les totaux à 2 décimales
  for (const file in totals) {
    for (const category in totals[file]) {
      totals[file][category] = Math.round(totals[file][category] * 100) / 100;
    }
  }

  return {
    totals,
    matriculeCount,
    hasDuplicates,
    errorCount: comparisonResult.summary.totalDifferences,
  };
}

// // Routes pour la gestion du lexique
app.get("/lexique", async (req, res) => {
  try {
    const db = await getDatabase();

    db.all(
      "SELECT * FROM lexicon_columns ORDER BY column_name",
      [],
      (err, rows) => {
        if (err) {
          console.error("Erreur lors de la récupération du lexique:", err);
          return res.status(500).render("error", {
            title: "Erreur",
            message: "Erreur lors de la récupération du lexique",
          });
        }

        res.render("lexique", {
          title: "Lexique des colonnes",
          columns: rows,
          queryParams: req.query,
        });
      },
    );
  } catch (error) {
    console.error("Erreur database:", error);
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur de base de données",
    });
  }
});

// // Route pour ajouter une entrée au lexique
app.post(
  "/lexique/ajouter",
  express.urlencoded({ extended: true }),
  async (req, res) => {
    try {
      const { column_name, column_type, description, formula } = req.body;

      // Validation simple
      if (!column_name) {
        return res.status(400).render("error", {
          title: "Erreur",
          message: "Le nom de la colonne est obligatoire",
        });
      }

      const db = await getDatabase();

      db.run(
        "INSERT INTO lexicon_columns (column_name, column_type, description, formula) VALUES (?, ?, ?, ?)",
        [column_name, column_type, description, formula],
        function (err) {
          if (err) {
            console.error("Erreur lors de l'ajout au lexique:", err);
            return res.status(500).render("error", {
              title: "Erreur",
              message: "Erreur lors de l'ajout au lexique",
            });
          }

          res.redirect("/lexique?success=true");
        },
      );
    } catch (error) {
      console.error("Erreur database:", error);
      res.status(500).render("error", {
        title: "Erreur",
        message: "Erreur de base de données",
      });
    }
  },
);

// // Route pour supprimer une entrée du lexique
app.post("/lexique/supprimer/:id", async (req, res) => {
  try {
    const id = req.params.id;
    const db = await getDatabase();

    db.run("DELETE FROM lexicon_columns WHERE id = ?", [id], function (err) {
      if (err) {
        console.error("Erreur lors de la suppression:", err);
        return res.status(500).render("error", {
          title: "Erreur",
          message: "Erreur lors de la suppression",
        });
      }

      res.redirect("/lexique?deleted=true");
    });
  } catch (error) {
    console.error("Erreur database:", error);
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur de base de données",
    });
  }
});

// // Route pour l'upload du fichier lexique Excel
app.post("/lexique/upload", upload.single("lexique_file"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).render("error", {
        title: "Erreur",
        message: "Aucun fichier n'a été uploadé",
      });
    }

    // Lire le fichier Excel
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });

    // Vérifier si le format correspond à ce que nous attendons
    const headers = data[0] || [];
    const hasValidHeaders =
      headers.some(
        (h) =>
          (h && h.toLowerCase().includes("rubrique")) ||
          (h && h.toLowerCase().includes("colonne")),
      ) &&
      headers.some(
        (h) =>
          (h && h.toLowerCase().includes("définition")) ||
          (h && h.toLowerCase().includes("description")),
      ) &&
      headers.some((h) => h && h.toLowerCase().includes("formule"));

    if (!hasValidHeaders) {
      return res.status(400).render("error", {
        title: "Format invalide",
        message:
          "Le format du fichier Excel n'est pas reconnu. Assurez-vous qu'il contient les colonnes 'Nom de la rubrique', 'Définition', et 'Formule de calcul'.",
      });
    }

    // Déterminer les indices des colonnes importantes
    const colNameIndex = headers.findIndex(
      (h) =>
        (h && h.toLowerCase().includes("rubrique")) ||
        (h && h.toLowerCase().includes("colonne")),
    );
    const descriptionIndex = headers.findIndex(
      (h) =>
        (h && h.toLowerCase().includes("définition")) ||
        (h && h.toLowerCase().includes("description")),
    );
    const formulaIndex = headers.findIndex(
      (h) =>
        (h && h.toLowerCase().includes("formule")) ||
        (h && h.toLowerCase().includes("calcul")),
    );

    // Extraire les données et vérifier les conflits
    const db = await getDatabase();

    // Récupérer d'abord toutes les colonnes existantes - DÉPLACER À L'INTÉRIEUR D'UNE FONCTION ASYNC
    const existingColumnsMap = await new Promise((resolve, reject) => {
      db.all(
        "SELECT column_name, id, column_type, description, formula FROM lexicon_columns",
        [],
        (err, rows) => {
          if (err) reject(err);
          else {
            const map = {};
            rows.forEach((row) => {
              map[row.column_name] = row;
            });
            resolve(map);
          }
        },
      );
    });

    // Plus précis pour identifier les nouveaux et les conflits
    const newEntries = [];
    const conflictEntries = [];
    const existingColumnNames = Object.keys(existingColumnsMap);

    // Parcourir les données pour identifier nouveaux et conflits
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[colNameIndex]) continue; // Ignorer les lignes sans nom de colonne

      const column_name = row[colNameIndex].trim();
      const description = row[descriptionIndex]
        ? row[descriptionIndex].trim()
        : "";
      const formula = row[formulaIndex] ? row[formulaIndex].trim() : "";

      // Déterminer si c'est un élément fixe ou variable en fonction des mots-clés
      const fixedKeywords = [
        "matricule",
        "bu",
        "charge",
        "enfant",
        "salaire mensuel",
        "ancienneté",
        "taux horaire",
        "transport",
        "logement",
        "astreinte",
        "forfait",
        "prime fixe",
        "détachement",
      ];

      const isFixed = fixedKeywords.some((keyword) =>
        column_name.toLowerCase().includes(keyword.toLowerCase()),
      );

      const column_type = isFixed ? "fixe" : "variable";

      if (existingColumnNames.includes(column_name)) {
        // Comparer les valeurs pour voir s'il y a vraiment un conflit
        const existing = existingColumnsMap[column_name];
        const hasChanges =
          existing.column_type !== column_type ||
          existing.description !== description ||
          existing.formula !== formula;

        if (hasChanges) {
          conflictEntries.push({
            id: existing.id,
            column_name,
            description,
            formula,
            column_type,
            existing: {
              column_type: existing.column_type,
              description: existing.description,
              formula: existing.formula,
            },
          });
        }
      } else {
        newEntries.push({
          column_name,
          description,
          formula,
          column_type,
        });
      }
    }

    // Si des conflits existent, demander confirmation à l'utilisateur
    if (conflictEntries.length > 0) {
      // Stocker temporairement les données et rediriger vers une page de confirmation
      req.app.locals.pendingImport = {
        newEntries,
        conflictEntries,
      };

      return res.render("lexique-confirm", {
        title: "Confirmation d'import",
        newCount: newEntries.length,
        conflicts: conflictEntries,
        queryParams: req.query,
      });
    }

    // Si pas de conflits, procéder à l'import des nouvelles entrées
    const insertCount = await insertEntries(db, newEntries);

    // Nettoyer le fichier uploadé
    fs.unlinkSync(req.file.path);

    res.redirect(`/lexique?imported=${insertCount}`);
  } catch (error) {
    console.error("Erreur lors de l'import du lexique:", error);
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur lors de l'import du lexique Excel: " + error.message,
    });
  }
});

// // Fonction utilitaire pour insérer des entrées dans la base de données
async function insertEntries(db, entries) {
  let insertCount = 0;

  for (const entry of entries) {
    try {
      await new Promise((resolve, reject) => {
        // Utiliser INSERT OR IGNORE pour ignorer les erreurs de contrainte unique
        db.run(
          "INSERT OR IGNORE INTO lexicon_columns (column_name, column_type, description, formula) VALUES (?, ?, ?, ?)",
          [
            entry.column_name,
            entry.column_type,
            entry.description,
            entry.formula,
          ],
          function (err) {
            if (err) reject(err);
            else {
              if (this.changes > 0) {
                insertCount++;
              }
              resolve();
            }
          },
        );
      });
    } catch (error) {
      console.error(
        `Erreur lors de l'insertion de ${entry.column_name}:`,
        error,
      );
      // Continuer avec les entrées suivantes malgré l'erreur
    }
  }

  return insertCount;
}

// // Route pour gérer la confirmation d'import avec conflits
// // Route pour gérer la confirmation d'import avec conflits
app.post("/lexique/confirm-import", async (req, res) => {
  try {
    const { action } = req.body;
    const pendingImport = req.app.locals.pendingImport;

    if (!pendingImport) {
      return res.status(400).render("error", {
        title: "Erreur",
        message: "Aucun import en attente. Veuillez réessayer.",
      });
    }

    const db = await getDatabase();
    let updateCount = 0;
    let insertCount = 0;

    if (action === "replace_all") {
      // Mettre à jour les entrées existantes
      for (const entry of pendingImport.conflictEntries) {
        await new Promise((resolve, reject) => {
          db.run(
            "UPDATE lexicon_columns SET column_type = ?, description = ?, formula = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            [entry.column_type, entry.description, entry.formula, entry.id],
            function (err) {
              if (err) reject(err);
              else {
                if (this.changes > 0) {
                  updateCount++;
                }
                resolve();
              }
            },
          );
        });
      }

      // Ajouter les nouvelles entrées
      insertCount = await insertEntries(db, pendingImport.newEntries);
    } else if (action === "keep_existing") {
      // Ne pas toucher aux entrées existantes, ajouter uniquement les nouvelles
      insertCount = await insertEntries(db, pendingImport.newEntries);
    } else if (action === "clear_and_import") {
      // Supprimer toutes les entrées et importer les nouvelles
      await new Promise((resolve, reject) => {
        db.run("DELETE FROM lexicon_columns", [], function (err) {
          if (err) reject(err);
          else resolve();
        });
      });

      // Importer toutes les entrées
      const allEntries = pendingImport.newEntries.concat(
        pendingImport.conflictEntries.map((entry) => ({
          column_name: entry.column_name,
          column_type: entry.column_type,
          description: entry.description,
          formula: entry.formula,
        })),
      );
      insertCount = await insertEntries(db, allEntries);
    }

    // Nettoyer les données temporaires
    delete req.app.locals.pendingImport;

    res.redirect(`/lexique?imported=${insertCount}&updated=${updateCount}`);
  } catch (error) {
    console.error("Erreur lors de la confirmation d'import:", error);
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur lors de la confirmation d'import: " + error.message,
    });
  }
});

// // Route pour afficher le formulaire de modification
app.get("/lexique/editer/:id", async (req, res) => {
  try {
    const id = req.params.id;
    const db = await getDatabase();

    db.get("SELECT * FROM lexicon_columns WHERE id = ?", [id], (err, row) => {
      if (err) {
        console.error("Erreur lors de la récupération de la colonne:", err);
        return res.status(500).render("error", {
          title: "Erreur",
          message: "Erreur lors de la récupération de la colonne",
        });
      }

      if (!row) {
        return res.status(404).render("error", {
          title: "Non trouvé",
          message: "La colonne demandée n'existe pas",
        });
      }

      res.render("lexique-edit", {
        title: "Modifier une colonne",
        column: row,
        queryParams: req.query,
      });
    });
  } catch (error) {
    console.error("Erreur database:", error);
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur de base de données",
    });
  }
});

// // Route pour traiter la modification
app.post("/lexique/editer/:id", async (req, res) => {
  try {
    const id = req.params.id;
    const { column_name, column_type, description, formula } = req.body;

    // Validation
    if (!column_name) {
      return res.status(400).render("error", {
        title: "Erreur",
        message: "Le nom de la colonne est obligatoire",
      });
    }

    const db = await getDatabase();

    // Vérifier si le nouveau nom existe déjà (pour un autre ID)
    db.get(
      "SELECT id FROM lexicon_columns WHERE column_name = ? AND id != ?",
      [column_name, id],
      (err, existingRow) => {
        if (err) {
          console.error("Erreur lors de la vérification du nom:", err);
          return res.status(500).render("error", {
            title: "Erreur",
            message: "Erreur lors de la vérification du nom",
          });
        }

        if (existingRow) {
          return res.status(400).render("error", {
            title: "Nom déjà utilisé",
            message: "Ce nom de colonne est déjà utilisé par une autre entrée",
          });
        }

        // Mettre à jour l'entrée
        db.run(
          "UPDATE lexicon_columns SET column_name = ?, column_type = ?, description = ?, formula = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
          [column_name, column_type, description, formula, id],
          function (err) {
            if (err) {
              console.error("Erreur lors de la mise à jour:", err);
              return res.status(500).render("error", {
                title: "Erreur",
                message: "Erreur lors de la mise à jour",
              });
            }

            res.redirect("/lexique?updated=true");
          },
        );
      },
    );
  } catch (error) {
    console.error("Erreur database:", error);
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur de base de données",
    });
  }
});

// // Route pour exporter le lexique en CSV
app.get("/lexique/export", async (req, res) => {
  try {
    const db = await getDatabase();

    db.all(
      "SELECT column_name, column_type, description, formula FROM lexicon_columns ORDER BY column_name",
      [],
      (err, rows) => {
        if (err) {
          console.error("Erreur lors de l'export du lexique:", err);
          return res.status(500).render("error", {
            title: "Erreur",
            message: "Erreur lors de l'export du lexique",
          });
        }

        // Générer le contenu CSV
        let csvContent = "colonne,type,description,formule\n";
        rows.forEach((row) => {
          const values = [
            `"${row.column_name.replace(/"/g, '""')}"`,
            `"${row.column_type.replace(/"/g, '""')}"`,
            `"${(row.description || "").replace(/"/g, '""')}"`,
            `"${(row.formula || "").replace(/"/g, '""')}"`,
          ];
          csvContent += values.join(",") + "\n";
        });

        // Envoyer le fichier CSV
        res.setHeader("Content-Type", "text/csv");
        res.setHeader(
          "Content-Disposition",
          "attachment; filename=lexique_colonnes.csv",
        );
        res.send(csvContent);
      },
    );
  } catch (error) {
    console.error("Erreur database:", error);
    res.status(500).render("error", {
      title: "Erreur",
      message: "Erreur de base de données",
    });
  }
});

// Gestion des erreurs 404
app.use((req, res) => {
  res.status(404).render("error", {
    title: "Page non trouvée",
    message: "La page que vous recherchez n'existe pas",
  });
});

// Gestion des erreurs générales
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).render("error", {
    title: "Erreur serveur",
    message: err.message || "Une erreur s'est produite sur le serveur",
  });
});

// Démarrage du serveur
app.listen(PORT, () => {
  console.log(`Le serveur est démarré sur http://localhost:${PORT}`);
});
