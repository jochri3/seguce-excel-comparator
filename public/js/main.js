document.addEventListener("DOMContentLoaded", function () {
  // Afficher l'animation de chargement lors de la soumission du formulaire
  const form = document.querySelector("form");
  const loading = document.createElement("div");

  if (form) {
    loading.className = "loading mt-4";
    loading.innerHTML = `
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">Chargement...</span>
            </div>
            <div class="mt-3">
                <h5>Traitement des fichiers en cours...</h5>
                <p class="text-muted">Veuillez patienter pendant que nous analysons vos fichiers Excel.</p>
            </div>
        `;

    form.addEventListener("submit", function () {
      // Vérifier si les fichiers sont sélectionnés
      const fileA = document.getElementById("fileA");
      const fileB = document.getElementById("fileB");

      if (!fileA.files.length || !fileB.files.length) {
        return false;
      }

      // Ajouter l'animation de chargement
      form.appendChild(loading);
      loading.style.display = "block";

      // Désactiver le bouton de soumission
      const submitBtn = form.querySelector('button[type="submit"]');
      if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML =
          '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Traitement en cours...';
      }
    });
  }

  // Ajout d'une fonctionnalité d'export Excel si la page de comparaison est chargée
  const exportBtn = document.getElementById("exportBtn");
  if (exportBtn) {
    exportBtn.addEventListener("click", function () {
      // Tableau pour stocker les données à exporter
      const exportData = [];

      // Ajouter l'en-tête
      exportData.push([
        "Matricule",
        "Colonne",
        "Valeur Fichier A",
        "Valeur Fichier B",
        "Différence",
      ]);

      // Parcourir toutes les différences
      document.querySelectorAll(".accordion-item").forEach((item) => {
        const matricule = item.getAttribute("data-matricule");
        const diffRows = item.querySelectorAll(".diff-row");

        diffRows.forEach((row) => {
          const cells = row.querySelectorAll("td");
          if (cells.length >= 4) {
            exportData.push([
              matricule,
              cells[0].textContent.trim(),
              cells[1].textContent.trim(),
              cells[2].textContent.trim(),
              cells[3].textContent.trim(),
            ]);
          }
        });
      });

      // Créer une chaîne CSV
      const csvContent = exportData.map((row) => row.join(",")).join("\n");

      // Créer un lien de téléchargement
      const encodedUri = encodeURI("data:text/csv;charset=utf-8," + csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", "differences_apei.csv");
      document.body.appendChild(link);

      // Cliquer sur le lien et le supprimer
      link.click();
      document.body.removeChild(link);
    });
  }

  // Filtrer les résultats de comparaison
  const searchInput = document.getElementById("searchResults");
  if (searchInput) {
    searchInput.addEventListener("input", function () {
      const searchTerm = this.value.toLowerCase();
      const accordionItems = document.querySelectorAll(".accordion-item");

      accordionItems.forEach((item) => {
        const itemText = item.textContent.toLowerCase();
        if (itemText.includes(searchTerm)) {
          item.style.display = "";
        } else {
          item.style.display = "none";
        }
      });
    });
  }

  // Fonctionnalité pour filtrer par colonne
  document.querySelectorAll(".filter-btn").forEach((btn) => {
    btn.addEventListener("click", function () {
      const column = this.getAttribute("data-column");
      filterByColumn(column);
    });
  });

  // Fonction de filtrage par colonne
  function filterByColumn(columnName) {
    document.querySelectorAll(".accordion-item").forEach((item) => {
      const columns = item.getAttribute("data-columns").split(",");
      if (columns.includes(columnName)) {
        item.style.display = "";
        // Développer automatiquement les éléments qui correspondent
        const button = item.querySelector(".accordion-button");
        if (button && button.classList.contains("collapsed")) {
          button.click();
        }

        // Surligner les lignes qui correspondent à la colonne
        item.querySelectorAll(".diff-row").forEach((row) => {
          const rowColumn = row.getAttribute("data-column");
          if (rowColumn === columnName) {
            row.classList.add("table-warning");
          } else {
            row.classList.remove("table-warning");
          }
        });
      } else {
        item.style.display = "none";
      }
    });
  }

  // Effacer les filtres
  const clearFilterBtn = document.getElementById("clearFilterBtn");
  if (clearFilterBtn) {
    clearFilterBtn.addEventListener("click", function () {
      document.getElementById("searchResults").value = "";
      document.querySelectorAll(".accordion-item").forEach((item) => {
        item.style.display = "";
        // Supprimer le surlignage
        item.querySelectorAll(".diff-row").forEach((row) => {
          row.classList.remove("table-warning");
        });
      });
    });
  }

  // Amélioration de l'UX: cliquer sur une ligne de différence la met en évidence
  document.querySelectorAll(".diff-row").forEach((row) => {
    row.addEventListener("click", function () {
      // Supprimer la classe de toutes les lignes
      document.querySelectorAll(".diff-row").forEach((r) => {
        r.classList.remove("table-warning");
      });

      // Ajouter la classe à la ligne cliquée
      this.classList.add("table-warning");
    });
  });
});
