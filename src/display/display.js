export async function handleCellChange(matchingData) {
    await Excel.run(async (context) => {
        console.log("Matching data: ", matchingData);
        Office.context.ui.displayDialogAsync(
            'https://localhost:3000/display.html',
            {height: 45, width: 55},
        );
        await context.sync();
    });
}

window.onload = function () {
    const storedValue = localStorage.getItem('matchingData');
    console.log("Stored Value:", storedValue);
    if (storedValue) {
        const output = document.getElementById('data-output');
        output.innerHTML = "<pre>" + JSON.stringify(JSON.parse(storedValue), null, 2) + "</pre>";
    }
};

