const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const excelParser = require("./utils/excel-parser");

// Configuration de l'application
const app = express();
const PORT = process.env.PORT || 3000;

// Configuration des dossiers statiques et du moteur de template
app.use(express.static(path.join(__dirname, "public")));
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

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
  res.render("index", { title: "Réconciliation API" });
});

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

      // Réconciliation des données
      const comparisonResult = excelParser.compareExcelData(
        fileAData,
        fileBData
      );

      // Rendre la page de comparaison avec les résultats
      res.render("compare", {
        title: "Résultats de la réconciliation",
        fileAName,
        fileBName,
        comparisonResult,
      });

      // Nettoyer les fichiers uploadés après traitement (optionnel)
      setTimeout(() => {
        fs.unlinkSync(fileAPath);
        fs.unlinkSync(fileBPath);
      }, 5000); // Délai de 5 secondes avant suppression
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
