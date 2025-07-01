Office.onReady((info) => {
    orderingWorksheet.onSelectionChanged.add(onSelectionChanged);
});

function onCellChange() {
    
}


async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        console.error(error);
    }
}

