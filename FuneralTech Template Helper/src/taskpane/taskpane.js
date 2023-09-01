
// Initialize when Office is ready
Office.onReady(function(info) {
  if (info.host === Office.HostType.Word) {
    document.getElementById("loadFieldsButton").addEventListener("click", readFile);
    document.getElementById("loadDefaultFieldsButton").addEventListener("click", loadDefaultFields); // New line
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
  });
}

// Function to generate OOXML for Full Text
function generateFullTextOoxml(fieldName) {
  let ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            <w:r>
              <w:fldChar w:fldCharType="begin"/>
            </w:r>
            <w:r>
              <w:instrText xml:space="preserve"> MERGEFIELD ${fieldName} </w:instrText>
            </w:r>
            <w:r>
              <w:fldChar w:fldCharType="end"/>
            </w:r>
          </w:p>
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`;
  return ooxml;
}

// Function to generate OOXML for First Letter of Field
function generateFirstLetterOoxml(fieldName) {
  let ooxml = `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body>
          <w:p>
            `;

  // Loop through the alphabet to generate the IF conditions
  for (let i = 65; i <= 90; i++) {
    const letter = String.fromCharCode(i);

    // IF field begin
    ooxml += `<w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> MERGEFIELD ${fieldName} </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:instrText xml:space="preserve"> = "${letter}*" "${letter}" "" </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  }

  ooxml += `</w:p>
          </w:body>
        </w:document>
     </pkg:xmlData>
    </pkg:part>
  </pkg:package>`;
  return ooxml;
}



function generateBooleanCheckboxOoxml(fieldName) {
  return `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
    <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
      <pkg:xmlData>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
      </pkg:xmlData>
    </pkg:part>
    <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
      <pkg:xmlData>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="begin"/></w:r>
              <w:r><w:instrText xml:space="preserve"> MERGEFIELD ${fieldName} </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
              <w:r><w:instrText xml:space="preserve"> = " " "X" "" </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="end"/></w:r>
            </w:p>
          </w:body>
        </w:document>
      </pkg:xmlData>
    </pkg:part>
  </pkg:package>`;
}

// Function to load default fields from the CSV file
function loadDefaultFields() {
  fetch('assets/field_names.csv')
    .then(response => response.text())
    .then(data => {
      const fields = data.split('\n').map(row => row.trim());
      populateFields(fields);
    })
    .catch(error => {
      console.error("Error loading default fields:", error);
    });
}