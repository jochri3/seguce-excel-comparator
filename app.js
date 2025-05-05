const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const excelParser = require("./utils/excel-parser");
const excelExport = require("./utils/export-excel");
const SessionManager = require("./utils/session-manager");
const { initDatabase } = require("./config/database");

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
        session.session_id
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
        "System" // À remplacer par le vrai utilisateur quand l'auth est implémentée
      );

      const fileBSaved = await SessionManager.saveFileVersion(
        session.session_id,
        "fileB",
        fileBName,
        fileBData.data,
        fileBData.formulas,
        "System"
      );

      // Réconciliation des données
      const comparisonResult = excelParser.compareExcelData(
        fileAData,
        fileBData
      );

      // Sauvegarder les résultats de comparaison
      await SessionManager.saveComparisonResult(
        session.session_id,
        fileASaved.version, // ou Math.max(fileASaved.version, fileBSaved.version)
        comparisonResult
      );

      // Calculer les totaux pour l'affichage
      const summary = calculateSummaryData(comparisonResult);

      // Rendre la page de comparaison avec les résultats
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
  }
);
// Route pour exporter les résultats en Excel
app.get("/export-excel", (req, res) => {
  try {
    const comparisonResult = req.app.locals.lastComparisonResult;
    const fileAName = req.app.locals.fileAName;
    const fileBName = req.app.locals.fileBName;

    if (!comparisonResult) {
      return res
        .status(400)
        .send(
          "Aucun résultat de comparaison disponible. Veuillez d'abord comparer deux fichiers."
        );
    }

    // Générer le fichier Excel
    const excelBuffer = excelExport.exportToExcel(
      comparisonResult,
      fileAName,
      fileBName
    );

    // Envoyer le fichier au client
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=reconciliation_paie.xlsx"
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
      error
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
