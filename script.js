document.getElementById("uploadExcel").addEventListener("change", function (event) {
    let file = event.target.files[0];
    let reader = new FileReader();

    reader.onload = function (e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: "array" });

        let sheet = workbook.Sheets[workbook.SheetNames[0]];
        let jsonData = XLSX.utils.sheet_to_json(sheet, { raw: false });

        jsonData.forEach(row => {
            if (row["Date"] && !isNaN(row["Date"])) {
                let excelDate = parseInt(row["Date"]);
                let date = new Date((excelDate - 25569) * 86400000); // Convert to JS Date
                row["Date"] = date.toLocaleDateString("en-GB"); // Format as DD-MM-YYYY
            }
        });

        displayData(jsonData);
    };

    reader.readAsArrayBuffer(file);
});

function displayData(data) {
    let tableBody = document.querySelector("#attendanceTable tbody");
    tableBody.innerHTML = ""; // Clear previous data

    data.forEach(row => {
        addRow(row["Roll No"], row["Name"], row["Date"], "Present");
    });

    addRow("", "", "", "Present", true); // Add an empty row for new entries
}

function addRow(rollNo = "", name = "", date = "", status = "Present", isNew = false) {
    let tableBody = document.querySelector("#attendanceTable tbody");
    let tr = document.createElement("tr");

    tr.innerHTML = `
        <td><input type="text" value="${rollNo}" oninput="checkNewRow(this)"></td>
        <td><input type="text" value="${name}" oninput="checkNewRow(this)"></td>
        <td><input type="date" value="${date}" oninput="checkNewRow(this)"></td>
        <td>
            <select>
                <option value="Present" ${status === "Present" ? "selected" : ""}>Present</option>
                <option value="Absent" ${status === "Absent" ? "selected" : ""}>Absent</option>
            </select>
        </td>
    `;

    tableBody.appendChild(tr);
}

function checkNewRow(inputElement) {
    let lastRow = document.querySelector("#attendanceTable tbody tr:last-child");
    let inputs = lastRow.querySelectorAll("input");

    let rollNo = inputs[0].value.trim();
    let name = inputs[1].value.trim();
    let date = inputs[2].value.trim();

    if (rollNo !== "" && name !== "" && date !== "") {
        addRow("", "", "", "Present", true); // Add a new empty row automatically
    }
}

document.getElementById("downloadExcel").addEventListener("click", function () {
    let table = document.querySelector("#attendanceTable tbody");
    let data = [];

    table.querySelectorAll("tr").forEach(row => {
        let inputs = row.querySelectorAll("input");
        let attendanceStatus = row.querySelector("select").value;

        if (inputs[0].value.trim()) { // Only save rows with a Roll No
            data.push({
                "Roll No": inputs[0].value.trim(),
                "Name": inputs[1].value.trim(),
                "Date": inputs[2].value.trim(),
                "Status": attendanceStatus
            });
        }
    });

    let ws = XLSX.utils.json_to_sheet(data);
    let wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance");

    XLSX.writeFile(wb, "Updated_Attendance.xlsx");
});
