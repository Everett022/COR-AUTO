Office.onReady((info) => {
    orderingWorksheet.onChanged.add(onCellChange);
});

async function onCellChange(event) {
    await Excel.run(async (context) => {
        await context.sync();        
        console.log("Change type of event: " + event.changeType);
        console.log("Address of event: " + event.address);
        console.log("Source of event: " + event.source);       
    }).catch(errorHandlerFunction);
}


async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        console.error(error);
    }
}

