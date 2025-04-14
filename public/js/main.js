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

  // Fonctionnalité d'export Excel
  const exportExcelBtn = document.getElementById("exportExcelBtn");
  if (exportExcelBtn) {
    exportExcelBtn.addEventListener("click", function () {
      // Afficher un indicateur de chargement
      const originalText = this.innerHTML;
      this.innerHTML =
        '<span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span> Génération du fichier...';
      this.disabled = true;

      // Rediriger vers l'URL d'export
      setTimeout(() => {
        window.location.href = "/export-excel";

        // Réinitialiser le bouton après un délai
        setTimeout(() => {
          this.innerHTML = originalText;
          this.disabled = false;
        }, 2000);
      }, 500);
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
            row.classList.add("highlighted");
          } else {
            row.classList.remove("highlighted");
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
          row.classList.remove("highlighted");
        });
      });
    });
  }

  // Amélioration de l'UX: cliquer sur une ligne de différence la met en évidence
  document.querySelectorAll(".diff-row").forEach((row) => {
    row.addEventListener("click", function () {
      // Supprimer la classe de toutes les lignes
      document.querySelectorAll(".diff-row").forEach((r) => {
        r.classList.remove("highlighted");
      });

      // Ajouter la classe à la ligne cliquée
      this.classList.add("highlighted");
    });
  });

  // Recherche dans les colonnes
  const searchColumnsInput = document.getElementById("searchColumns");
  if (searchColumnsInput) {
    searchColumnsInput.addEventListener("input", function () {
      const searchTerm = this.value.toLowerCase();
      const tableRows = document.querySelectorAll("#columnsTable tbody tr");

      tableRows.forEach((row) => {
        const columnName = row
          .querySelector("td:first-child")
          .textContent.toLowerCase();
        row.style.display = columnName.includes(searchTerm) ? "" : "none";
      });
    });
  }

  // Fonctionnalité d'impression avancée
  const printBtn = document.getElementById("printBtn");
  if (printBtn) {
    printBtn.addEventListener("click", function () {
      // Préparer la page pour l'impression
      document.querySelectorAll(".accordion-collapse").forEach((collapse) => {
        // Développer tous les accordéons pour l'impression
        if (!collapse.classList.contains("show")) {
          // Trouver et cliquer sur le bouton correspondant
          const button = document.querySelector(
            `[data-bs-target="#${collapse.id}"]`
          );
          if (button) {
            button.classList.remove("collapsed");
            collapse.classList.add("show");
          }
        }
      });

      // Lancer l'impression
      window.print();

      // Restaurer l'état des accordéons après l'impression
      setTimeout(() => {
        document
          .querySelectorAll(".accordion-collapse.show")
          .forEach((collapse) => {
            // Trouver et cliquer sur le bouton correspondant
            const button = document.querySelector(
              `[data-bs-target="#${collapse.id}"]`
            );
            if (button && !button.hasAttribute("data-was-expanded")) {
              button.classList.add("collapsed");
              collapse.classList.remove("show");
            }
          });
      }, 1000);
    });
  }
});
