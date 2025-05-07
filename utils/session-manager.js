const { getDatabase } = require("../config/database");
const { v4: uuidv4 } = require("uuid");

class SessionManager {
  // Créer une nouvelle session ou récupérer une existante
  static async getOrCreateSession(month, year) {
    const db = await getDatabase();

    return new Promise((resolve, reject) => {
      const sessionId = `${year}-${month.toString().padStart(2, "0")}`;

      db.get(
        "SELECT * FROM sessions WHERE session_id = ?",
        [sessionId],
        (err, row) => {
          if (err) {
            reject(err);
            return;
          }

          if (row) {
            resolve(row);
          } else {
            // Créer une nouvelle session
            db.run(
              `INSERT INTO sessions (session_id, month, year) VALUES (?, ?, ?)`,
              [sessionId, month, year],
              function (err) {
                if (err) {
                  reject(err);
                  return;
                }

                db.get(
                  "SELECT * FROM sessions WHERE id = ?",
                  [this.lastID],
                  (err, newRow) => {
                    if (err) reject(err);
                    else resolve(newRow);
                  }
                );
              }
            );
          }
        }
      );
    });
  }

  // Sauvegarder une version de fichier
  static async saveFileVersion(
    sessionId,
    fileType,
    fileName,
    data,
    formulas,
    createdBy
  ) {
    const db = await getDatabase();

    return new Promise((resolve, reject) => {
      // Obtenir la version actuelle
      db.get(
        "SELECT MAX(version) as max_version FROM file_versions WHERE session_id = ? AND file_type = ?",
        [sessionId, fileType],
        (err, row) => {
          if (err) {
            reject(err);
            return;
          }

          const version = (row.max_version || 0) + 1;

          db.run(
            `INSERT INTO file_versions (session_id, file_type, file_name, version, data, formulas, created_by) 
             VALUES (?, ?, ?, ?, ?, ?, ?)`,
            [
              sessionId,
              fileType,
              fileName,
              version,
              JSON.stringify(data),
              JSON.stringify(formulas),
              createdBy,
            ],
            function (err) {
              if (err) reject(err);
              else resolve({ id: this.lastID, version });
            }
          );
        }
      );
    });
  }

  // Obtenir l'historique d'une session
  static async getSessionHistory(sessionId) {
    const db = await getDatabase();

    return new Promise((resolve, reject) => {
      const sql = `
        SELECT fv.*, 
               CASE 
                 WHEN fv.file_type = 'fileA' THEN 'Fichier Fournisseur'
                 ELSE 'Fichier SEGUCE RDC'
               END as file_label
        FROM file_versions fv
        WHERE fv.session_id = ?
        ORDER BY fv.file_type, fv.version DESC
      `;

      db.all(sql, [sessionId], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });
  }

  // Sauvegarder les résultats de comparaison
  static async saveComparisonResult(sessionId, version, comparisonData) {
    const db = await getDatabase();

    return new Promise((resolve, reject) => {
      db.run(
        `INSERT INTO comparison_results (session_id, version, total_differences, column_differences, details) 
         VALUES (?, ?, ?, ?, ?)`,
        [
          sessionId,
          version,
          comparisonData.summary.totalDifferences,
          JSON.stringify(comparisonData.summary.columnDifferences),
          JSON.stringify(comparisonData.details),
        ],
        function (err) {
          if (err) reject(err);
          else resolve(this.lastID);
        }
      );
    });
  }

  // Obtenir tous les mois avec des sessions
  static async getAllSessionMonths() {
    const db = await getDatabase();

    return new Promise((resolve, reject) => {
      db.all(
        "SELECT DISTINCT year, month FROM sessions ORDER BY year DESC, month DESC",
        [],
        (err, rows) => {
          if (err) reject(err);
          else resolve(rows);
        }
      );
    });
  }

  static async getComparisonResults(sessionId) {
    const db = await getDatabase();

    return new Promise((resolve, reject) => {
      const sql = `
      SELECT *
      FROM comparison_results
      WHERE session_id = ?
      ORDER BY version DESC, created_at DESC
    `;

      db.all(sql, [sessionId], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });
  }
}

module.exports = SessionManager;
