const colHeadContainer = document.getElementById("head-row-container");
const form = document.getElementById("form");
const state = {};
let totalHorizontalCells = 26;
let rows = 1000;
const dataContainer = document.getElementById("data-container");
const srNo = document.getElementById("sr-no-container");
const disEle = document.getElementById("dis-ele");
let selectedCell = null;
const funExp = document.getElementById("funExp");
const email = document.getElementById("img");
const downloadSymbol = document.getElementById("download-symbol");
const upload = document.getElementById("upload");
const uploadSymbol = document.getElementById("upload-symbol");
const createSheet = document.getElementById("create-sheet");
const sheets = document.getElementById("sheets");
let sheetCount = 1;


const defaultState = {
  fontFamily: "Arial,Sans-serif",
  fontSize: "16",
  isBold: false,
  isUnderline: false,
  isItalic: false,
  textColor: "#000000",
  backgroundColor: "#ffffff",
  horizontalAlign: "left",
};

// Creating column headers (A-Z)
for (let i = 1; i <= totalHorizontalCells; i++) {
  const cell = document.createElement("div");
  cell.innerText = String.fromCharCode(i + 64);
  cell.className = "col-head";
  colHeadContainer.appendChild(cell);
}

// Creating row numbers (1-1000)
for (let i = 1; i <= rows; i++) {
  const cell = document.createElement("div");
  cell.className = "sr-no";
  cell.innerText = i;
  srNo.appendChild(cell);
}

// Creating the grid cells
for (let i = 1; i <= rows; i++) {
  for (let j = 1; j <= totalHorizontalCells; j++) {
    const cell = document.createElement("cite");
    cell.className = "data";
    cell.contentEditable = "true";
    cell.id = `${String.fromCharCode(64 + j)}${i}`;
    cell.addEventListener("input", onChangeInnerText);
    dataContainer.appendChild(cell);
  }
}

// Updating cell visual representation
dataContainer.addEventListener("click", (e) => {
  if (!e.target.classList.contains("data")) return; // Ensure the clicked element is a data cell

  if (selectedCell) {
    selectedCell.classList.remove("active-cell"); // Remove active class from the previous cell
  }

  selectedCell = e.target;
  selectedCell.classList.add("active-cell");
  disEle.innerText = selectedCell.id;
  fillFormDataWithStateData(); // Fill the form based on the selected cell's state
});

// Fill form fields with the current state of the selected cell
function fillFormDataWithStateData() {
  if (state[selectedCell.id]) {
    const object = state[selectedCell.id];
    form["font-family"].value = object.fontFamily;
    form["font-size"].value = object.fontSize.slice(
      0,
      object.fontSize.length - 2
    ); // Remove "px"
    form["isBold"].checked = object.isBold;
    form["isItalic"].checked = object.isItalic;
    form["isUnderline"].checked = object.isUnderline;
    form["text-color"].value = object.textColor;
    form["bg-color"].value = object.backgroundColor;
    form["horizontal-align"].value = object.horizontalAlign;
  } else {
    // Reset to default values if visiting the cell for the first time
    form.reset();
    form["font-family"].value = defaultState.fontFamily;
    form["font-size"].value = defaultState.fontSize;
    form["isBold"].checked = defaultState.isBold;
    form["isItalic"].checked = defaultState.isItalic;
    form["isUnderline"].checked = defaultState.isUnderline;
    form["text-color"].value = defaultState.textColor;
    form["bg-color"].value = defaultState.backgroundColor;
    form["horizontal-align"].value = defaultState.horizontalAlign;
  }
}

// Handle form changes to apply styles to the selected cell
form.addEventListener("change", () => {
  if (!selectedCell) {
    alert("Select a cell to make changes");
    return;
  }

  const formData = {
    fontFamily: form["font-family"].value,
    fontSize: form["font-size"].value + "px",
    isBold: form["isBold"].checked,
    isItalic: form["isItalic"].checked,
    isUnderline: form["isUnderline"].checked,
    textColor: form["text-color"].value,
    bgColor: form["bg-color"].value,
    horizontalAlign: form["horizontal-align"].value,
  };

  // Save the state of the selected cell
  state[selectedCell.id] = { ...formData, innerText: selectedCell.innerText };
  applyStylesToSelectedCell(formData);
});

// Apply styles to the selected cell based on form data
function applyStylesToSelectedCell(formData) {
  selectedCell.style.fontSize = formData.fontSize;
  selectedCell.style.fontFamily = formData.fontFamily;
  selectedCell.style.fontWeight = formData.isBold ? "bold" : "normal";
  selectedCell.style.fontStyle = formData.isItalic ? "italic" : "normal";
  selectedCell.style.textDecoration = formData.isUnderline
    ? "underline"
    : "none";
  selectedCell.style.color = formData.textColor;
  selectedCell.style.backgroundColor = formData.bgColor;
  selectedCell.style.textAlign = formData.horizontalAlign;
}

// Handle formula expression input
funExp.addEventListener("keyup", (e) => {
  if (!selectedCell) {
    alert("Select a cell to make changes");
    funExp.value = "";
    return;
  }
  if (e.key === "Enter") {
    try {
      const result = new Function("return " + funExp.value)(); // Safely evaluate expressions
      selectedCell.innerText = result;
    } catch (error) {
      alert("Invalid expression");
    }
    funExp.value = ""; // Clear the input after execution
  }
});

// upon clicking email a pop up is opened with corresponding email
email.addEventListener("click", (e) => {
  window.location.href = "https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox";
});

// whenever a change is happened in the text area of cite celll we ned to update the innertext
function onChangeInnerText(e) {
  if (state[selectedCell.id]) {
    state[selectedCell.id].innerText = selectedCell.innerText;
  } else {
    state[selectedCell.id] = {
      ...defaultState,
      innerText: selectedCell.innerText,
    };
  }
}

// whener we click on upload and download symbol corresponding buttons should be triggered

downloadSymbol.addEventListener("click", () => {
  if (!selectedCell) {
    alert("select a cell");
    return;
  }

  // Simulate a click on a hidden download button (if needed)

  // Get the data from the selected cell's state
  const data = JSON.stringify(state);

  // Create a Blob with the data (set MIME type as plain text)
  let blob = new Blob([data], { type: "application/JSON" });

  // Create a download URL
  let downloadUrl = URL.createObjectURL(blob);

  // Create an anchor element dynamically
  const link = document.createElement("a");
  link.href = downloadUrl;

  // Set the filename for the download
  link.download = "abcd.json";

  // Trigger the download by simulating a click
  link.click();

  // Clean up the object URL to free memory
  URL.revokeObjectURL(downloadUrl);
});

upload.addEventListener("change", (e) => {
  let file = e.target.files[0];

  if (file.type !== "application/json") {
    alert("Please upload a valid JSON file.");
    return;
  }

  let fileReader = new FileReader();

  fileReader.onload = function (event) {
    try {
      let content = JSON.parse(event.target.result);
      fillUploadedDataIntoCell(content);
      // You can now load this content into your spreadsheet (e.g., set `state`)
    } catch (error) {
      alert("Invalid JSON file.");
      console.error(error);
    }
  };

  fileReader.readAsText(file); // Read the file as text
});

function fillUploadedDataIntoCell(content) {
  // get all the keys
  let keys = Object.keys(content);
  keys.forEach((key) => {
    const cell = document.getElementById(key);
    if (cell) {
      cell.style.fontFamily = content[key].fontFamily;
      cell.style.fontSize = content[key].fontSize.slice(0,content[key].fontSize - 2);
      cell.style.fontWeight = content[key].isBold ? "bold" : "normal";
      cell.style.fontStyle = content[key].isItalic ? "italic" : "normal";
      cell.style.textDecoration = content[key].isUnderline? "underline": "none";
      cell.style.textColor = content[key].bgColor;
      cell.style.backgroundColor = content[key].bgColor;
      cell.style.textAlign = content[key].horizontalAlign;
      cell.innerText = content[key].innerText;
    }

    state[key] = content[key];
  });
}

// multi sheet functionality
createSheet.addEventListener("click", () => {
    const newSheet = document.createElement("input");
    newSheet.type = "radio";
    newSheet.value = `sheet${sheetCount}`;
    newSheet.id = `sheet${sheetCount}`; // Updated to match label's `for` attribute
    newSheet.name = "sheets";
    newSheet.className = "sheets";
    newSheet.addEventListener("click",sheetOpening);
    
    // Create a label to display the text for the radio button
    const sheetLabel = document.createElement("label");
    sheetLabel.setAttribute("for", newSheet.id); // Link label to input
    sheetLabel.innerText = `Sheet ${sheetCount}`; // Set visible text
    
    // Append the radio button and label to the container
    sheets.appendChild(newSheet);
    sheets.appendChild(sheetLabel);
    
    sheetCount++;
});

function sheetOpening(e){
    
}

