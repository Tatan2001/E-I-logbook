let selectedPart = '';

// Load old data from localStorage if available
let allData = JSON.parse(localStorage.getItem('machineData')) || {};

// Function to display the form for a specific part
function showForm(partName) {
    selectedPart = partName;
    document.getElementById('formTitle').innerText = 'Data Entry for ' + partName;
    document.getElementById('dataForm').style.display = 'block';
}

// Function to collect form data and append to the relevant Excel sheet
function submitForm() {
    const formData = {
        OTI: document.getElementById('oti').value,
        WTI: document.getElementById('wti').value,
        Tap: document.getElementById('tap').value,
        Current: document.getElementById('current').value,
        SilicaGel: document.getElementById('silicaGel').value,
        OilLevelMOG: document.getElementById('mog').value,
        OilLevelGaugeGlass: document.getElementById('gaugeGlass').value,
        OilLeakage: document.querySelector('input[name="oilLeakage"]:checked').value,
        Remarks: document.getElementById('remarks').value
    };

    // Add the new form data to the relevant part's data array
    if (!allData[selectedPart]) {
        allData[selectedPart] = [];
    }
    allData[selectedPart].push(formData);

    // Save the updated data in localStorage
    localStorage.setItem('machineData', JSON.stringify(allData));

    // Create a new workbook
    const wb = XLSX.utils.book_new();

    // Loop through each part and create a sheet for each
    for (const part in allData) {
        const ws = XLSX.utils.json_to_sheet(allData[part]);
        XLSX.utils.book_append_sheet(wb, ws, part);
    }

    // Write the workbook to a binary string
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

    // Helper function to convert string to array buffer
    function s2ab(s) {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

    // Use FileSaver.js to save the file
    saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), "MachinePartData.xlsx");

    alert("Data updated and saved successfully in the relevant sheet!");
    document.getElementById("dataForm").reset();
}