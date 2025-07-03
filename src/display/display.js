import { matchingData as _matchingData } from '../taskpane/taskpane.js';
let matchingData = _matchingData;

export function handleCellChange(matchingData) {
    console.log("Matching data: ", matchingData);
    console.log("length", matchingData.length);
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/display.html',
        {height: 45, width: 55},
        function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
        }
    );
    
}


// function processMessage(arg) {
//     let arr;
//     try {
//         arr = typeof arg.message === "string" ? JSON.parse(arg.message) : arg.message;
//     } catch {
//         arr = [];
//     }
//     outputAllData(arr);
// }



// const output = document.getElementById('data-output');

//     if (!matchingData || matchingData.length === 0){
//         output.innerHTML = "<p>No data to display.</p>";
//     }

//     const header = Object.keys(matchingData[0]);
//     let html = "<table border='1' style='border-collapse:collapse;'><tr>";
//     header.forEach(h => html += `<th>${h}</th>`);
//     html += "</tr>";

//     matchingData.forEach(obj => {
//         html += "<tr>";
//         header.forEach(h => html += `<td>${obj[h]}</td>`);
//         html += "</tr>";
//     });
//     html += "</table>";
//     output.innerHTML = html;