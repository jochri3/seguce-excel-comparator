const sqlite3 = require("sqlite3").verbose();
const path = require("path");

const dbPath = path.join(__dirname, "../reconciliation.db");

const initDatabase = () => {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database(dbPath, (err) => {
      if (err) {
        console.error("Erreur de connexion à la base de données:", err);
        reject(err);
        return;
      }
      console.log("Connecté à la base de données SQLite");

      // Créer les tables si elles n'existent pas
      db.serialize(() => {
        // Table des sessions de réconciliation
        db.run(`CREATE TABLE IF NOT EXISTS sessions (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          session_id TEXT UNIQUE,
          month INTEGER,
          year INTEGER,
          file_a_name TEXT,
          file_b_name TEXT,
          provider_type TEXT,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        // Table des versions de fichiers
        db.run(`CREATE TABLE IF NOT EXISTS file_versions (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          session_id TEXT,
          file_type TEXT,
          file_name TEXT,
          version INTEGER,
          data TEXT,
          formulas TEXT,
          created_by TEXT,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          FOREIGN KEY (session_id) REFERENCES sessions(session_id)
        )`);

        // Table des résultats de comparaison
        db.run(`CREATE TABLE IF NOT EXISTS comparison_results (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id TEXT,
  version INTEGER,
  total_differences INTEGER,
  column_differences TEXT,
  details TEXT,
  totals TEXT, /* Nouvelle colonne pour stocker les totaux */
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (session_id) REFERENCES sessions(session_id)
)`);

        // Table des colonnes lexique
        db.run(`CREATE TABLE IF NOT EXISTS lexicon_columns (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          column_name TEXT UNIQUE,
          column_type TEXT,
          description TEXT,
          formula TEXT,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`);

        resolve(db);
      });
    });
  });
};

const getDatabase = () => {
  return new Promise((resolve, reject) => {
    const db = new sqlite3.Database(dbPath, (err) => {
      if (err) {
        reject(err);
      } else {
        resolve(db);
      }
    });
  });
};

module.exports = {
  initDatabase,
  getDatabase,
};
