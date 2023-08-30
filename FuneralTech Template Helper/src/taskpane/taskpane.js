
// Initialize when Office is ready
Office.onReady(function(info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("loadFieldsButton").addEventListener("click", readFile);
    document.getElementById("searchInput").addEventListener("input", filterFields);
    document.body.addEventListener("click", function() {
      const existingMenu = document.querySelector(".menu");
      if (existingMenu) existingMenu.remove();
    });
    document.getElementById("fieldList").addEventListener("click", showMenu);
  }
});

// Function to read the file and populate the field list
function readFile() {
  const fileInput = document.getElementById("fileInput");
  const file = fileInput.files[0];
  if (file) {
    const reader = new FileReader();
    reader.onload = function(e) {
      const content = e.target.result;
      if (file.name.endsWith('.csv')) {
        Papa.parse(content, {
          complete: function(results) {
            const fields = results.data.map(row => row[0]);
            populateFields(fields);
          }
        });
      }
    };
    reader.readAsText(file);
  }
}

// Function to populate the field list
function populateFields(fields) {
  const fieldList = document.getElementById("fieldList");
  fieldList.innerHTML = "";
  fields.forEach(field => {
    const listItem = document.createElement("div");
    listItem.textContent = field;
    listItem.className = "fieldItem";
    fieldList.appendChild(listItem);
  });
}

// Function to filter the field list based on search input
function filterFields() {
  const searchInput = document.getElementById("searchInput").value.toLowerCase();
  const fieldItems = document.querySelectorAll(".fieldItem");
  fieldItems.forEach(item => {
    if (item.textContent.toLowerCase().includes(searchInput)) {
      item.style.display = "block";
    } else {
      item.style.display = "none";
    }
  });
}

// Function to show the menu
function showMenu(event) {
  event.stopPropagation();
  const existingMenu = document.querySelector(".menu");
  if (existingMenu) existingMenu.remove();
  const menu = document.getElementById("menuTemplate").cloneNode(true);
  menu.style.display = "block";
  menu.className = "menu";
  const originalFieldName = event.target.closest('.fieldItem').textContent.trim();
  menu.querySelectorAll(".menuOption").forEach(option => {
    option.addEventListener("click", function(e) {
      e.stopPropagation();
      const type = e.target.getAttribute("data-type");
      insertAdvancedField(originalFieldName, type);
      menu.remove();
    });
  });
  event.target.appendChild(menu);
}

// Function to insert advanced merge fields
function insertAdvancedField(fieldName, type) {
  console.log(`insertAdvancedField called with fieldName: ${fieldName} and type: ${type}`);

  Word.run(function(context) {
    const range = context.document.getSelection();
    context.load(range);

    return context.sync().then(function() {
      let ooxml = "";

      switch (type) {
        case "fullText":
          ooxml = generateFullTextOoxml(fieldName);
          break;
        case "firstLetter":
          ooxml = generateFirstLetterOoxml(fieldName);
          break;
        case "booleanCheckbox":
          ooxml = generateBooleanCheckboxOoxml(fieldName);
          break;
        default:
          return;
      }

      range.insertOoxml(ooxml, Word.InsertLocation.end);
      return context.sync();
    });
  }).catch(function(error) {
    console.error(JSON.stringify(error));
  });
}

// Function to generate OOXML for Full Text
function generateFullTextOoxml(fieldName) {
  return `<w:fldSimple w:instr=" MERGEFIELD ${fieldName} \\* MERGEFORMAT "><w:r><w:t>${fieldName}</w:t></w:r></w:fldSimple>`;
}

// Function to generate OOXML for First Letter of Field
function generateFirstLetterOoxml(fieldName) {
  let ooxml = '<w:p>';
  for (let i = 65; i <= 90; i++) {
    const letter = String.fromCharCode(i);
    ooxml += `<w:fldSimple w:instr='IF "{ MERGEFIELD ${fieldName} }" = "${letter}" "${letter}"'><w:r><w:t>${letter}</w:t></w:r></w:fldSimple>`;
  }
  ooxml += '</w:p>';
  return ooxml;
}

// Function to generate OOXML for Boolean Checkbox
function generateBooleanCheckboxOoxml(fieldName) {
  return `<w:fldSimple w:instr='IF "{ MERGEFIELD ${fieldName} }" = " " "X" ""'><w:r><w:t>X</w:t></w:r></w:fldSimple>`;
}