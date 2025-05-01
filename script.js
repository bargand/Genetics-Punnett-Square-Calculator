// DOM Elements
const traitCountInput = document.getElementById("traitCount");
const traitInput = document.getElementById("traitInput");
const punnettSquareDiv = document.getElementById("punnettSquare");
const ratioResultsDiv = document.getElementById("ratioResults");
const loader = document.getElementById("loader");
const resultsSection = document.getElementById("resultsSection");
const errorMessageDiv = document.getElementById("errorMessage");
const successMessageDiv = document.getElementById("successMessage");
const colorLegendDiv = document.getElementById("colorLegend");

// Color classes for genotypes
const COLOR_CLASSES = Array.from(
  { length: 100 },
  (_, i) => `color${i + 1}`
);
const COLOR_VALUES = [
  "#FFEBEE",
  "#E8F5E9",
  "#E3F2FD",
  "#EDE7F6",
  "#FFF8E1",
  "#FBE9E7",
  "#E0F7FA",
  "#F1F8E9",
  "#FFF3E0",
  "#E8EAF6",
  "#FCE4EC",
  "#E0F2F1",
  "#EFEBE9",
  "#EEEEEE",
  "#E1BEE7",
  "#BBDEFB",
  "#C8E6C9",
  "#FFECB3",
  "#D7CCC8",
  "#F5F5F5",
  "#FFCCCC",
  "#FFCC99",
  "#FFFF99",
  "#CCFFCC",
  "#CCE0FF",
  "#F2C2FF",
  "#D4E157",
  "#FF7043",
  "#64B5F6",
  "#F06292",
  "#9575CD",
  "#FFB74D",
  "#80DEEA",
  "#C5E1A5",
  "#FFEB3B",
  "#FF4081",
  "#FF80AB",
  "#81C784",
  "#4CAF50",
  "#FBC02D",
  "#A5D6A7",
  "#D32F2F",
  "#FFB300",
  "#8C9EFF",
  "#B39DDB",
  "#03A9F4",
  "#C2185B",
  "#8BC34A",
  "#F44336",
  "#9C27B0",
  "#BA68C8",
  "#FF5722",
  "#B2DFDB",
  "#F57C00",
  "#8E24AA",
  "#7B1FA2",
  "#01579B",
  "#C62828",
  "#F06292",
  "#29B6F6",
  "#FFB74D",
  "#81D4FA",
  "#D32F2F",
  "#6A1B9A",
  "#00E5FF",
  "#7C4DFF",
  "#FF8A65",
  "#03DAC5",
  "#90A4AE",
  "#FF80AB",
  "#FFB300",
  "#00C853",
  "#FF8F00",
  "#F44336",
  "#388E3C",
  "#8BC34A",
  "#0288D1",
  "#FFEB3B",
  "#F44336",
  "#1565C0",
  "#D32F2F",
  "#4CAF50",
  "#3F51B5",
  "#7C4DFF",
  "#90CAF9",
  "#F50057",
  "#B2EBF2",
  "#AB47BC",
  "#3E2723",
  "#FFCC80",
  "#FF9E80",
  "#00BCD4",
  "#9C27B0",
  "#F44336",
  "#FF1744",
  "#D50000",
  "#7B1FA2",
  "#03A9F4",
  "#4CAF50",
  "#color100",
  "#FFB300",
];

// Utility Functions
function showMessage(element, message) {
  element.textContent = message;
  element.style.display = "block";
  const otherElement =
    element === errorMessageDiv ? successMessageDiv : errorMessageDiv;
  otherElement.style.display = "none";
}

function hideMessages() {
  errorMessageDiv.style.display = "none";
  successMessageDiv.style.display = "none";
}

function calculateGCD(a, b) {
  return b ? calculateGCD(b, a % b) : a;
}

function normalizeGenotype(genotype) {
  let normalized = "";
  for (let i = 0; i < genotype.length; i += 2) {
    const gene = genotype.substr(i, 2);
    normalized += gene.split("").sort().join("");
  }
  return normalized;
}

function determinePhenotype(genotype) {
  let phenotype = "";
  for (let i = 0; i < genotype.length; i += 2) {
    const gene = genotype.substr(i, 2);
    phenotype += gene.match(/[A-Z]/)
      ? gene[0].toUpperCase()
      : gene[0].toLowerCase();
  }
  return phenotype;
}

function generateColorLegend(colorMap) {
  colorLegendDiv.innerHTML = "";

  Object.entries(colorMap)
    .sort((a, b) => {
      if (a[0].length !== b[0].length) return a[0].length - b[0].length;
      return a[0].localeCompare(b[0]);
    })
    .forEach(([genotype, colorClass]) => {
      const legendItem = document.createElement("div");
      legendItem.className = "legend-item";

      const colorBox = document.createElement("div");
      colorBox.className = `legend-color ${colorClass}`;

      const label = document.createElement("span");
      label.textContent = genotype;

      legendItem.appendChild(colorBox);
      legendItem.appendChild(label);
      colorLegendDiv.appendChild(legendItem);
    });
}

// Excel Download Functions
function downloadExcel() {
  const table = punnettSquareDiv.querySelector("table");
  if (!table) {
    showMessage(errorMessageDiv, "No Punnett Square to download");
    return;
  }

  // Create workbook
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.table_to_sheet(table);

  // Add colors to cells
  const colorMap = getColorMapFromTable(table);
  addColorsToWorksheet(ws, colorMap);

  // Add worksheet to workbook
  XLSX.utils.book_append_sheet(wb, ws, "Punnett Square");

  // Generate and download file
  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `Punnett_Square_${date}.xlsx`);
}

function getColorMapFromTable(table) {
  const colorMap = {};
  const cells = table.querySelectorAll('td[class^="color-"]');

  cells.forEach((cell) => {
    const address = XLSX.utils.encode_cell({
      r: cell.parentNode.rowIndex - 1,
      c: cell.cellIndex,
    });
    const colorClass = Array.from(cell.classList).find((c) =>
      c.startsWith("color-")
    );
    colorMap[address] =
      COLOR_VALUES[parseInt(colorClass.replace("color-", "")) - 1];
  });

  return colorMap;
}

function addColorsToWorksheet(ws, colorMap) {
  // Create styles for each color
  const styles = {};
  Object.entries(colorMap).forEach(([address, color]) => {
    if (!ws[address]) return;

    const styleID = `s${color.replace("#", "")}`;
    if (!styles[styleID]) {
      styles[styleID] = {
        fill: { fgColor: { rgb: color.replace("#", "") } },
        font: { bold: true },
      };
    }

    ws[address].s = styleID;
  });

  // Add styles to workbook
  if (Object.keys(styles).length > 0) {
    if (!ws["!styles"]) ws["!styles"] = {};
    Object.assign(ws["!styles"], styles);
  }
}

// Main Punnett Square Generation
function generateTable() {
  hideMessages();
  loader.style.display = "block";
  resultsSection.style.display = "none";

  setTimeout(() => {
    try {
      // Validate input
      const traitCount = parseInt(traitCountInput.value);
      if (isNaN(traitCount) || traitCount < 1 || traitCount > 7) {
        throw new Error(
          "Please enter a number between 1 and 7 for traits"
        );
      }

      const traits = traitInput.value
        .trim()
        .split(",")
        .map((t) => t.trim())
        .filter((t) => t.length > 0);

      const expectedCombinations = Math.pow(2, traitCount);
      if (traits.length !== expectedCombinations) {
        throw new Error(
          `Expected ${expectedCombinations} trait combinations for ${traitCount} trait(s)`
        );
      }

      // Generate Punnett Square
      let tableHTML = "<table><tr><th></th>";
      traits.forEach((trait) => (tableHTML += `<th>${trait}</th>`));
      tableHTML += "</tr>";

      const allGenotypes = [];
      const allPhenotypes = [];
      const colorMap = {};
      let colorIndex = 0;

      traits.forEach((rowTrait) => {
        tableHTML += `<tr><th>${rowTrait}</th>`;
        traits.forEach((colTrait) => {
          const genotype = rowTrait + colTrait;
          allGenotypes.push(genotype);
          allPhenotypes.push(determinePhenotype(genotype));

          const normalized = normalizeGenotype(genotype);
          if (!colorMap[normalized]) {
            colorMap[normalized] =
              COLOR_CLASSES[colorIndex % COLOR_CLASSES.length];
            colorIndex++;
          }

          tableHTML += `<td class="${colorMap[normalized]}">${genotype}</td>`;
        });
        tableHTML += "</tr>";
      });

      tableHTML += "</table>";
      punnettSquareDiv.innerHTML = tableHTML;

      // Calculate and display ratios
      displayGeneticRatios(allGenotypes, allPhenotypes);
      generateColorLegend(colorMap);

      resultsSection.style.display = "block";
      showMessage(
        successMessageDiv,
        "Punnett Square generated successfully!"
      );
    } catch (error) {
      showMessage(errorMessageDiv, error.message);
    } finally {
      loader.style.display = "none";
    }
  }, 300);
}

function displayGeneticRatios(genotypes, phenotypes) {
  const genotypeCounts = {};
  const phenotypeCounts = {};
  const total = genotypes.length;

  // Count genotypes and phenotypes
  genotypes.forEach((genotype, index) => {
    const normalized = normalizeGenotype(genotype);
    genotypeCounts[normalized] = (genotypeCounts[normalized] || 0) + 1;
    phenotypeCounts[phenotypes[index]] =
      (phenotypeCounts[phenotypes[index]] || 0) + 1;
  });

  // Format ratios
  const formatRatio = (counts) => {
    return Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .map(([type, count]) => {
        const gcd = calculateGCD(count, total);
        return `${type}: ${count / gcd}/${total / gcd}`;
      })
      .join(", ");
  };

  // Display results
  ratioResultsDiv.innerHTML = `
    <div class="ratio-title">Genetic Ratios</div>
    <div class="ratio-item">
      <span class="ratio-label">Genotype:</span>
      <span class="ratio-value">${formatRatio(genotypeCounts)}</span>
    </div>
    <div class="ratio-item">
      <span class="ratio-label">Phenotype:</span>
      <span class="ratio-value">${formatRatio(phenotypeCounts)}</span>
    </div>
    <div class="ratio-item">
      <span class="ratio-label">Total combinations:</span>
      <span class="ratio-value">${total}</span>
    </div>
    <button onclick="downloadExcel()" class="btn download-btn" style="margin-top: 15px;">
      Download as Excel
    </button>
  `;
}

// Initialize SheetJS library
document.addEventListener("DOMContentLoaded", function () {
  // Load SheetJS library dynamically if not already loaded
  if (typeof XLSX === "undefined") {
    const script = document.createElement("script");
    script.src =
      "https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js";
    script.onload = function () {
      console.log("SheetJS library loaded");
    };
    document.head.appendChild(script);
  }
});

document.addEventListener('DOMContentLoaded', function() {
      const dropdownHeader = document.querySelector('.dropdown-header');
      const dropdownContent = document.querySelector('.dropdown-content');
      
      dropdownHeader.addEventListener('click', function() {
          this.classList.toggle('active');
          
          if (dropdownContent.style.maxHeight) {
              dropdownContent.style.maxHeight = null;
          } else {
              dropdownContent.style.maxHeight = dropdownContent.scrollHeight + "px";
          }
      });
      
      // Optional: Auto-open if it's the user's first visit
      if (!localStorage.getItem('dropdownSeen')) {
          dropdownHeader.classList.add('active');
          dropdownContent.style.maxHeight = dropdownContent.scrollHeight + "px";
          localStorage.setItem('dropdownSeen', 'true');
      }
  });