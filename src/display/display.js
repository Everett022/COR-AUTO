Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    window.onload = outputData();
  }
});

export async function handleCellChange(matchingData) {
    await Excel.run(async (context) => {
        console.log("Matching data: ", matchingData);
        Office.context.ui.displayDialogAsync(
            'https://localhost:3000/display.html',
            {height: 65, width: 55},
        );
        await context.sync();
    });
}

async function outputData(){
    const storedValue = localStorage.getItem('matchingData');
    console.log("Stored Value:", storedValue);
    if (storedValue) {
        const data = JSON.parse(storedValue);
        const output = document.getElementById('data-output');

        let html = `<table>
        <thead>
            <tr>
            <th>Item Code</th>
            <th>Job Number</th>
            <th>Quantity</th>
            <th>Start Date</th>
            </tr>
        </thead>
        <tbody>
        `;

        data.forEach(row => {
            const isDisabled = row.qty === 0 || row.date === "" || row.date == null;
            html += `<tr${isDisabled ? ' class="disabled"' : ''}>
            <td>${row.code ?? ""}</td>
            <td>${row.job ?? ""}</td>
            <td>${row.qty ?? ""}</td>
            <td>${row.fDate ?? ""}</td>
            </tr>
        `;
                });

                html += `  </tbody>
        </table>`;
        output.innerHTML = html;
    }
}