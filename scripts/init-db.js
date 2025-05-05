const { initDatabase } = require("../config/database");

initDatabase()
  .then(() => {
    console.log("Base de données initialisée avec succès");
    process.exit(0);
  })
  .catch((err) => {
    console.error("Erreur d'initialisation de la base de données:", err);
    process.exit(1);
  });
