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

        let totalInventory = Number(localStorage.getItem('currentInventory'));
        if (!totalInventory || isNaN(totalInventory)) {
            totalInventory = 0;
        }

        const jobs = data
            .filter(row => row.qty > 0 && (row.date || row.fDate))
            .map(row => ({...row, dateObj: new Date(row.date || row.fDate)}));
        jobs.sort((a, b) => a.dateObj - b.dateObj);

        const coveredJobKeys = new Set();
        let inventoryLeft = totalInventory;
        for (const job of jobs) {
            if (inventoryLeft >= job.qty) {
                const key = `${job.job}__${job.dateObj.toISOString()}`;
                coveredJobKeys.add(key);
                inventoryLeft -= job.qty;
            }
        }

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
            let dateStr = row.date || row.fDate || '';
            let dateObj = dateStr ? new Date(dateStr) : null;
            const key = (row.job && dateObj) ? `${row.job}__${dateObj.toISOString()}` : '';
            const isCovered = coveredJobKeys.has(key);
            let rowClass = isDisabled ? 'disabled' : '';
            if (isCovered) rowClass += (rowClass ? ' ' : '') + 'covered';
            html += `<tr${rowClass ? ` class="${rowClass}"` : ''}>
            <td>${row.code ?? ""}</td>
            <td>${row.job ?? ""}</td>
            <td>${row.qty ?? ""}</td>
            <td>${row.fDate ?? row.date ?? ""}</td>
            </tr>
        `;
        });

        html += `  </tbody>
        </table>`;
        output.innerHTML = html;
    }
}