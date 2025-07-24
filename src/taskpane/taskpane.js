import { handleCellChange } from '../display/display.js';
import { openSettings } from '../settings/settings.js';

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("generate-ordering-report").onclick = () => tryCatch(generateOrderingReport);
    document.getElementById("generate-inventory-report").onclick = () => tryCatch(generateInventoryReport);
    document.getElementById("temp-reset").onclick = () => tryCatch(resetAll);
    document.getElementById('start-date').addEventListener('input', checkDatesAndClearMessage);
    document.getElementById('end-date').addEventListener('input', checkDatesAndClearMessage);
    document.getElementById("order-filtering").addEventListener('change', filteringDropdown);
    document.getElementById("settings-button").onclick = () => tryCatch(openSettings);
  }
});

// Global Variable inits     
    let filter = [];
    let invFilter = [];
    let earlyDateMap = new Map(); 
    let startDateMap = new Map();
    let orderingWorksheet;
    let orderingTable;
    let outputJobs = new Set();

    let allData = [];
    export let matchingData = [];

async function generateOrderingReport() {
    await Excel.run(async (context) => {
        resetGenerateOrdering();
        await context.sync();

        orderingWorksheet = context.workbook.worksheets.add("Ordering");
        orderingTable = orderingWorksheet.tables.add("A1:G1", true);

        orderingTable.name = "OrderingTable";

        orderingTable.getHeaderRowRange().values = [["Case #","Demand","Current Inventory", "On Order", "Required Amount","Buy or Make", "Earliest Start Date"]];
     
        orderingTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
        orderingTable.getRange().format.autofitColumns();
        orderingTable.getRange().format.autofitRows();
        
        const startDateValue = document.getElementById('start-date').value;
        const endDateValue = document.getElementById('end-date').value;

        if(!startDateValue || !endDateValue){
            document.getElementById('message-area').textContent = "Please enter the dates";
            return;
        }else{
            document.getElementById('message-area').textContent = " ";
            dateFilter();
            await context.sync();
            importColumnData();
        }
        
        orderingWorksheet.onSelectionChanged.add(displayData);
        await context.sync();
    });
}

async function generateInventoryReport() {
    await Excel.run(async (context) => {
        resetGenerateInventory();
        await context.sync();

        const inventoryWorksheet = context.workbook.worksheets.add("Inventory At");
        const inventoryTable = inventoryWorksheet.tables.add("A1:J1", true);
        
        inventoryTable.name = "InventoryAtTable";
        inventoryTable.getHeaderRowRange().values = [["Case #", "Demand", "Qty MEB", "Qty EFW", "Total MEB + EFW", "On Order", "Start Date", "Release Date", "Qty Needed (MEB)", "Notes"]];

        inventoryTable.columns.getItemAt(2).getRange().numberFormat = [['\u20AC#,##0.00']];
        inventoryTable.getRange().format.autofitColumns();
        inventoryTable.getRange().format.autofitRows();

        const startDateValue = document.getElementById('start-date').value;
        const endDateValue = document.getElementById('end-date').value;
        

        if(!startDateValue || !endDateValue){
            document.getElementById('message-area').textContent = "Please enter the dates";
            return;
        }else{
            document.getElementById('message-area').textContent = " ";
            otherDateFilter();
            await context.sync();
            importOtherColumnData();
        }
        await context.sync();
    });
}

async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        alert(error);
        console.error(error);
    }
}

async function importColumnData() {
    await Excel.run(async (context) => {
        const inventoryReportWorksheet = context.workbook.worksheets.getItem("Inventory");
        const inventoryUsedRange = inventoryReportWorksheet.getUsedRange().load("values");

        const dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
        const dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
        
        const openPOsWorksheet = context.workbook.worksheets.getItem("Open PO's");
        const openPOsUsedRange = openPOsWorksheet.getUsedRange().load("values");
        const orderingWorksheet = context.workbook.worksheets.getItem("Ordering");

        await context.sync();

        //Dynamic fluid Placement
        const dynamicHeaders = dynamicUsedRange.values[0];
        
        const dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
        const dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
        const dynWorkIdx = dynamicHeaders.indexOf("Work Center");
        
        const dynItemCodeColumn = `${colIdxToLetter(dynItemCodeIdx)}:${colIdxToLetter(dynItemCodeIdx)}`;
        const dynItemQtyColumn = `${colIdxToLetter(dynItemQtyIdx)}:${colIdxToLetter(dynItemQtyIdx)}`;
        const dynWorkColumn = `${colIdxToLetter(dynWorkIdx)}:${colIdxToLetter(dynWorkIdx)}`;

        //Open PO's fluid Placement
        const openPOsHeaders = openPOsUsedRange.values[0];

        const openPOItemCodeIdx = openPOsHeaders.indexOf("Item Code");
        const openPOItemQtyIdx = openPOsHeaders.indexOf("Outstanding Qty");
        
        const openPOItemCodeColumn = `${colIdxToLetter(openPOItemCodeIdx)}:${colIdxToLetter(openPOItemCodeIdx)}`;
        const openPOItemQtyColumn = `${colIdxToLetter(openPOItemQtyIdx)}:${colIdxToLetter(openPOItemQtyIdx)}`;
       
        //Inventory Report Fluid Placement
        const inventoryHeaders = inventoryUsedRange.values[0];

        const invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
        const invItemQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
        
        const invRepItemCodeColumn = `${colIdxToLetter(invItemCodeIdx)}:${colIdxToLetter(invItemCodeIdx)}`;
        const invRepItemQtyColumn = `${colIdxToLetter(invItemQtyIdx)}:${colIdxToLetter(invItemQtyIdx)}`;

        //Quanity and Item Code from Dynamic, Inventory Report, and Open PO's sheets
        const dynamic = context.workbook.worksheets.getItem("Dynamic");
        const dynamicQR = dynamic.getRange(dynItemQtyColumn).getUsedRange().load("values");
        const dynamicICR = dynamic.getRange(dynItemCodeColumn).getUsedRange().load("values");
        const dynamicWork = dynamic.getRange(dynWorkColumn).getUsedRange().load("values");
        await context.sync();

        const inventoryICR = inventoryReportWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values"); 
        const inventoryQR = inventoryReportWorksheet.getRange(invRepItemQtyColumn).getUsedRange().load("values"); 

        const openPOs = context.workbook.worksheets.getItem("Open PO's");
        const openPOsICR = openPOs.getRange(openPOItemCodeColumn).getUsedRange().load("values"); 
        const openPOsQR = openPOs.getRange(openPOItemQtyColumn).getUsedRange().load("values"); 
        await context.sync();

        //Date Filtering
        const filteredICR = filter.map(item => [item.itemCode]);
        const filteredQR = filter.map(item => [item.qty]);
        
        //Sum Map Building
        const fullDynamicMap = buildSumMap(dynamicICR.values, dynamicQR.values);
        const dynamicMap = buildSumMap(filteredICR, filteredQR);
        const inventoryMap = buildSumMap(inventoryICR.values, inventoryQR.values);
        const openPOsMap = buildSumMap(openPOsICR.values, openPOsQR.values);

        const allItemCodes = new Set([
            ...dynamicMap.keys(),
            ...inventoryMap.keys(),
            ...openPOsMap.keys()
        ]);

        const result = [["Case #", "Required Amount"]]; 
        for (const code of allItemCodes) {
            const dynamicQty = dynamicMap.get(code) || 0;
            const inventoryQty = inventoryMap.get(code) || 0;
            const openPOsQty = openPOsMap.get(code) || 0;
            const toOrder = dynamicQty - inventoryQty - openPOsQty;
            if (toOrder > 0){
                    result.push([code, toOrder]);
            } 
        }

        const caseNumbers = result.map(row => [row[0]]);
        orderingWorksheet.getRange(`A1:A${caseNumbers.length}`).values = caseNumbers;

        const requiredAmounts = result.map(row => [row[1]]);
        orderingWorksheet.getRange(`E1:E${requiredAmounts.length}`).values = requiredAmounts;
        await context.sync();

        // Remove From Order
        const sell = [["Case #","Remove From Order"]];
        for (const code of allItemCodes){
            const dynamicQty = Number(fullDynamicMap.get(code)) || 0;
            const openPOsQty = Number(openPOsMap.get(code)) || 0;
            const overBuy = openPOsQty - dynamicQty;
            if (!isNaN(dynamicQty) && !isNaN(openPOsQty)) {
                if (String(code).includes("COR") && openPOsQty > dynamicQty) {
                    sell.push([code, overBuy]);
                }
            }
        }
        console.log("Sell these", sell);
        
        //Importing the Planned Start Date
        const orderingUsedRange = orderingWorksheet.getUsedRange().load("values");
        await context.sync();

        const orderingValues = orderingUsedRange.values;
        
        const startArray = [["Earliest Start Date"]];
        for (let i = 1; i < orderingValues.length; i++) {
            const itemCode = String(orderingValues[i][0]).trim();
            const start = String(earlyDateMap.get(itemCode)) || "No Start Date Established";
            const dateOnly = start.split(' ').slice(0, 4).join(' ');
            startArray.push([dateOnly]);
        }
        orderingWorksheet.getRange(`G1:G${startArray.length}`).values = startArray;

        //Importing Demand, Current Inventory, and On Order
        const caseOrder = result.map(row => row[0]);

        const demand = [["Demand"]]; 
        for (const code of caseOrder) {
            const demandQty = dynamicMap.get(code) || 0;
          if (demandQty > 0){
                demand.push([demandQty]);
          }      
        }
        
        const demandOutput = demand.map(row => [row[0]]);
        orderingWorksheet.getRange(`B1:B${demandOutput.length}`).values = demandOutput;

        const currentInventory = [["Current Inventory"]]; 
        for (const code of caseOrder.slice(1)) {
            const currentInvQty = inventoryMap.get(code) || 0;
            currentInventory.push([currentInvQty]);
        }
        
        const currentInventoryOutput = currentInventory.map(row => [row[0]]);
        orderingWorksheet.getRange(`C1:C${currentInventoryOutput.length}`).values = currentInventoryOutput;

        const onOrder = [["On Order"]]; 
        for (const code of caseOrder.slice(1)) {
            const onOrderQty = openPOsMap.get(code) || 0;
            onOrder.push([onOrderQty]);           
        }
        
        const onOrderOutput = onOrder.map(row => [row[0]]);
        orderingWorksheet.getRange(`D1:D${onOrderOutput.length}`).values = onOrderOutput;

        //Buy or Make Logic
        const orderOrMakeMap = new Map();
        for (let i = 1; i < dynamicICR.values.length; i++) {
            const code = String(dynamicICR.values[i][0]).trim();
            const work = dynamicWork.values[i] ? String(dynamicWork.values[i][0]).trim() : "";
            if (code && work) {
                if(!orderOrMakeMap.has(code)) {
                    orderOrMakeMap.set(code, new Set());
                }
                orderOrMakeMap.get(code).add(work);
            }
        }
        await context.sync();

        const orderOrMake = [["Buy or Make"]]; 
        for (const code of caseOrder.slice(1)) {
            const workCentersSet = orderOrMakeMap.get(code);
            const workCenters = workCentersSet ? Array.from(workCentersSet).join(", ") : "";
            orderOrMake.push([workCenters]);
        }
        const orderOrMakeOutput = orderOrMake.map(row => [row[0]]);
        const orderOrMakeCategory = [["Buy or Make"]];
        
        for (let i = 1; i < orderOrMakeOutput.length; i++) {
            const workCenters = orderOrMakeOutput[i][0];
            if(
                workCenters.includes("40FGAL3A") || 
                workCenters.includes("40FGAL3B") ||
                workCenters.includes("40FGAL3C") || 
                workCenters.includes("40FGSI2A") ||
                workCenters.includes("40AIFG2B")
            ) {
                orderOrMakeCategory.push(["Must Buy"]); 
            } else if (Number(requiredAmounts[i][0]) >= 300){
                orderOrMakeCategory.push(["Can Buy"]);  
            }
            else{
                orderOrMakeCategory.push(["Can Make"]);    
            }    
            await context.sync();    
        }
        orderingWorksheet.getRange(`F1:F${orderOrMakeCategory.length}`).values = orderOrMakeCategory;

        // Table Formatting
        orderingWorksheet.getRange("A:G").format.autofitColumns();
        orderingWorksheet.getRange("A:H").format.horizontalAlignment = "Center";
        orderingWorksheet.getRange("A:H").format.verticalAlignment = "Center";
        orderingWorksheet.getRange("D:D").numberFormat = [['General']];

        orderingWorksheet.getRange("A:A").format.columnWidth = 150;
        orderingWorksheet.getRange("B:B").format.columnWidth = 90;
        orderingWorksheet.getRange("C:C").format.columnWidth = 120;
        orderingWorksheet.getRange("D:D").format.columnWidth = 90;
        orderingWorksheet.getRange("E:E").format.columnWidth = 130;
        orderingWorksheet.getRange("F:F").format.columnWidth = 100;
        orderingWorksheet.getRange("B:B").format.columnWidth = 130;
        orderingWorksheet.getUsedRange().format.rowHeight = 30;

        orderingWorksheet.freezePanes.freezeRows(1); 
        
        orderingWorksheet.getRange("E1:E1").format.fill.color = "#BE5014";
        orderingWorksheet.getRange("E1:E1").format.font.color = "yellow";     

        //All border lines
        const usedRange = orderingWorksheet.getUsedRange();   
        const borders = usedRange.format.borders;
        [
            "EdgeTop",
            "EdgeBottom",
            "EdgeLeft",
            "EdgeRight",
            "InsideVertical",
            "InsideHorizontal"
        ].forEach(edge => {
            borders.getItem(edge).style = "Continuous";
            borders.getItem(edge).weight = "Thin";
            borders.getItem(edge).color = "#000000"; 
        });
        //Bold Outline Lines
        const lastRow = demandOutput.length;
        const highlight = orderingWorksheet.getRange(`E1:E${lastRow}`).format.borders;
         [
            "EdgeTop",
            "EdgeBottom",
            "EdgeLeft",
            "EdgeRight",
        ].forEach(side => {
            highlight.getItem(side).style = "Continuous";
            highlight.getItem(side).weight = "Thick";
            highlight.getItem(side).color = "#BE5014"; 
        });


        await context.sync();
    });
}

async function importOtherColumnData(event) {
    await Excel.run(async (context) => {
        const inventoryReportWorksheet = context.workbook.worksheets.getItem("Inventory");
        const inventoryUsedRange = inventoryReportWorksheet.getUsedRange().load("values");

        const dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
        const dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
        
        const openPOsWorksheet = context.workbook.worksheets.getItem("Open PO's");
        const openPOsUsedRange = openPOsWorksheet.getUsedRange().load("values");

        const inventoryWorksheet = context.workbook.worksheets.getItem("Inventory At");

        await context.sync();

        //Dynamic fluid Placement
        const dynamicHeaders = dynamicUsedRange.values[0];
        
        const dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
        const dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
        
        const dynItemCodeColumn = `${colIdxToLetter(dynItemCodeIdx)}:${colIdxToLetter(dynItemCodeIdx)}`;
        const dynItemQtyColumn = `${colIdxToLetter(dynItemQtyIdx)}:${colIdxToLetter(dynItemQtyIdx)}`;

        //Open PO's fluid Placement
        const openPOsHeaders = openPOsUsedRange.values[0];

        const openPOItemCodeIdx = openPOsHeaders.indexOf("Item Code");
        const openPOItemQtyIdx = openPOsHeaders.indexOf("Outstanding Qty");
        
        const openPOItemCodeColumn = `${colIdxToLetter(openPOItemCodeIdx)}:${colIdxToLetter(openPOItemCodeIdx)}`;
        const openPOItemQtyColumn = `${colIdxToLetter(openPOItemQtyIdx)}:${colIdxToLetter(openPOItemQtyIdx)}`;
       
        //Inventory Report Fluid Placement
        const inventoryHeaders = inventoryUsedRange.values[0];

        const invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
        const invItemQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
        const invLocationIdx = inventoryHeaders.indexOf("Location");
        
        const invRepItemCodeColumn = `${colIdxToLetter(invItemCodeIdx)}:${colIdxToLetter(invItemCodeIdx)}`;
        const invRepItemQtyColumn = `${colIdxToLetter(invItemQtyIdx)}:${colIdxToLetter(invItemQtyIdx)}`;
        const invRepLocationColumn = `${colIdxToLetter(invLocationIdx)}:${colIdxToLetter(invLocationIdx)}`;
        await context.sync();

        //Quanity and Item Code from Dynamic, Inventory Report, and Open PO's sheets
        const inventoryICR = inventoryReportWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values"); 
        const inventoryQR = inventoryReportWorksheet.getRange(invRepItemQtyColumn).getUsedRange().load("values"); 
        const inventoryLOC = inventoryReportWorksheet.getRange(invRepLocationColumn).getUsedRange().load("values");

        const openPOs = context.workbook.worksheets.getItem("Open PO's");
        const openPOsICR = openPOs.getRange(openPOItemCodeColumn).getUsedRange().load("values"); 
        const openPOsQR = openPOs.getRange(openPOItemQtyColumn).getUsedRange().load("values"); 
        await context.sync();

        //Date Filtering
        const invFilterICR = invFilter.map(item => [item.itemCode]);
        const invFilterQR = invFilter.map(item => [item.qty]);

        //Sum Map Building
        const initialEntry = buildSumMap(invFilterICR, invFilterQR);

        const result = [["Case #","Demand"]];
        for (const code of initialEntry.keys()) {
            const demandQty = initialEntry.get(code) || 0;
            if (demandQty > 0) {
                result.push([code, demandQty]);
            }
        }
        await context.sync();

        const caseNumbers = result.map(row => row[0]);
        inventoryWorksheet.getRange(`A1:A${caseNumbers.length}`).values = caseNumbers.map(val => [val]);

        const requiredAmounts = result.map(row => [row[1]]);
        console.log("Required Amounts", requiredAmounts);
        inventoryWorksheet.getRange(`B1:B${requiredAmounts.length}`).values = requiredAmounts;
        
        //Mebane-EFW Inventory Map
        const mebArray = [];
        const efwArray = [];
        for (let i = 0; i < caseNumbers.length; i++) {
            const code = String(caseNumbers[i]).trim();
            let found = false;

            for (let j = 0; j < inventoryICR.values.length; j++) {
                const invCode = String(inventoryICR.values[j][0]).trim();
                const location = String(inventoryLOC.values[j][0]).trim();
                const qty = Number(inventoryQR.values[j][0]);

                if (code === invCode) {
                    const isMeb = location.includes("MEB");
                    const isEFW = location.includes("EFW");
                    mebArray.push([code, isMeb ? qty : 0]);
                    efwArray.push([code, isEFW ? qty : 0]);
                    found = true;
                }
            }

            if (!found) {
                mebArray.push([code, 0]);
                efwArray.push([code, 0]);
            }
        }
        await context.sync();

        const mebSumMap = buildSumMap(mebArray.map(item => [item[0]]), mebArray.map(item => [item[1]]));
        const mebAmounts = Array.from(mebSumMap.entries()).map(row => [row[1]]);
        inventoryWorksheet.getRange(`C2:C${mebAmounts.length + 1}`).values = mebAmounts;

        const efwSumMap = buildSumMap(efwArray.map(item => [item[0]]), efwArray.map(item => [item[1]]));
        const efwAmounts = Array.from(efwSumMap.entries()).map(row => [row[1]]);
        inventoryWorksheet.getRange(`D2:D${efwAmounts.length + 1}`).values = efwAmounts;

        const total = mebAmounts.map((value, index) => Number(value) + efwAmounts[index][0]);
        inventoryWorksheet.getRange(`E2:E${total.length + 1}`).values = total.map(val => [val]);
        
        //Inventory On Order 
        const openPOsMap = buildSumMap(openPOsICR.values, openPOsQR.values);
        const onOrder = [["On Order"]]; 
        for (const code of caseNumbers.slice(1)) {
            const onOrderQty = openPOsMap.get(code) || 0;
            onOrder.push([onOrderQty]);           
        }
        
        const onOrderOutput = onOrder.map(row => [row[0]]);
        inventoryWorksheet.getRange(`F1:F${onOrderOutput.length}`).values = onOrderOutput;
        
        // Importing the Start and Release Date
        const startArray = [["Earliest Start Date"]];
        const releaseArray = [["Release Date"]];

        for (let i = 1; i < caseNumbers.length; i++) {
            const itemCode = String(caseNumbers[i]).trim();
            const start = String(startDateMap.get(itemCode)) || "No Start Date Established";
            const dateOnly = start.split(' ').slice(0, 4).join(' ');
            startArray.push([dateOnly]);

            let releaseDate = new Date(start);
            if (!isNaN(releaseDate)) {
                releaseDate.setDate(releaseDate.getDate() - 10);
                const adjustedRelease = releaseDate.toDateString(); 
                releaseArray.push([adjustedRelease]);
            } else {
                releaseArray.push(["Invalid Release Date"]);
            }
        }

        inventoryWorksheet.getRange(`G1:G${startArray.length}`).values = startArray;
        inventoryWorksheet.getRange(`H1:H${releaseArray.length}`).values = releaseArray;

        //Qty Needed (MEB)
        const qtyNeeded = requiredAmounts.slice(1).map((value, index) => 
            Number(value) - mebAmounts[index]
        );
        inventoryWorksheet.getRange(`I2:I${qtyNeeded.length + 1}`).values = qtyNeeded.map(val => [val]);
        await context.sync();


        //Inventory formatting
        inventoryWorksheet.getRange("A:J").format.autofitColumns();
        inventoryWorksheet.getRange("A:J").format.horizontalAlignment = "Center";
        inventoryWorksheet.getRange("A:J").format.verticalAlignment = "Center";
        inventoryWorksheet.getRange("C:C").numberFormat = [['General']];

        inventoryWorksheet.getRange("A:A").format.columnWidth = 120;
        inventoryWorksheet.getRange("B:B").format.columnWidth = 70;
        inventoryWorksheet.getRange("C:C").format.columnWidth = 70;
        inventoryWorksheet.getRange("D:D").format.columnWidth = 70;
        inventoryWorksheet.getRange("E:E").format.columnWidth = 100;
        inventoryWorksheet.getRange("F:F").format.columnWidth = 75;
        inventoryWorksheet.getRange("G:G").format.columnWidth = 90;
        inventoryWorksheet.getRange("H:H").format.columnWidth = 90;
        inventoryWorksheet.getRange("I:I").format.columnWidth = 130;
        inventoryWorksheet.getRange("J:J").format.columnWidth = 150
        inventoryWorksheet.getUsedRange().format.rowHeight = 20;

        inventoryWorksheet.freezePanes.freezeRows(1); 
        
        inventoryWorksheet.getRange("I1:I1").format.fill.color = "#BE5014";
        inventoryWorksheet.getRange("I1:I1").format.font.color = "yellow";     

        //All border lines
        const usedRange = inventoryWorksheet.getUsedRange();   
        const borders = usedRange.format.borders;
        [
            "EdgeTop",
            "EdgeBottom",
            "EdgeLeft",
            "EdgeRight",
            "InsideVertical",
            "InsideHorizontal"
        ].forEach(edge => {
            borders.getItem(edge).style = "Continuous";
            borders.getItem(edge).weight = "Thin";
            borders.getItem(edge).color = "#000000"; 
        });
        
        //Bold Outline Lines
        const lastRow = mebAmounts.length;
        const highlight = inventoryWorksheet.getRange(`I1:I${lastRow + 1}`).format.borders;
         [
            "EdgeTop",
            "EdgeBottom",
            "EdgeLeft",
            "EdgeRight",
        ].forEach(side => {
            highlight.getItem(side).style = "Continuous";
            highlight.getItem(side).weight = "Thick";
            highlight.getItem(side).color = "#BE5014"; 
        });

        await context.sync();
    });
}

async function resetAll() {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.getItemOrNullObject("Ordering").delete();
            sheets.getItemOrNullObject("Inventory At").delete();
            sheets.getItemOrNullObject("Test").delete();
            document.getElementById('start-date').value = "";
            document.getElementById('end-date').value = "";
        await context.sync();
    });
}

async function resetGenerateOrdering() {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.getItemOrNullObject("Ordering").delete();
            filter = [];
            earlyDateMap.clear();
        await context.sync();
    });
}

async function resetGenerateInventory() {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.getItemOrNullObject("Inventory At").delete();
        await context.sync();
    });
}

function colIdxToLetter(idx) {
            let letter = "";
            while (idx >= 0) {
                letter = String.fromCharCode((idx % 26) + 65) + letter;
                idx = Math.floor(idx / 26) - 1;
            }
            return letter;
}

function buildSumMap(itemCodes, qtys) {
            const map = new Map();
            for (let i = 1; i < itemCodes.length; i++) { 
                const code = itemCodes[i][0];
                const qty = Number(qtys[i][0]);
                if (code && !isNaN(qty)) {
                    map.set(code, (map.get(code) || 0) + qty);
                }
            }
            return map;
}

async function dateFilter() {
    await Excel.run(async (context) => {
        const startDateInput = document.getElementById('start-date').value;
        const endDateInput = document.getElementById('end-date').value;

        const startDate = inputDateParse(startDateInput);
        const endDate = inputDateParse(endDateInput);

        const dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
        const dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
        await context.sync();
        const dynamicHeaders = dynamicUsedRange.values[0];

        const dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
        const dynStartIdx = dynamicHeaders.indexOf("Planned Start");
        const dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
        const dynJobIdx = dynamicHeaders.indexOf("Order Number");

        const dynStartColumn = `${colIdxToLetter(dynStartIdx)}:${colIdxToLetter(dynStartIdx)}`;
        const dynItemQtyColumn = `${colIdxToLetter(dynItemQtyIdx)}:${colIdxToLetter(dynItemQtyIdx)}`;
        const dynItemCodeColumn = `${colIdxToLetter(dynItemCodeIdx)}:${colIdxToLetter(dynItemCodeIdx)}`;
        const dynJobColumn =`${colIdxToLetter(dynJobIdx)}:${colIdxToLetter(dynJobIdx)}`;

        const dynamic = context.workbook.worksheets.getItem("Dynamic");
        const dynamicICR = dynamic.getRange(dynItemCodeColumn).getUsedRange().load("values");
        const dynamicQR = dynamic.getRange(dynItemQtyColumn).getUsedRange().load("values"); 
        const plannedStart = dynamicWorksheet.getRange(dynStartColumn).getUsedRange().load("values");
        const jobNumber = dynamicWorksheet.getRange(dynJobColumn).getUsedRange().load("values");

        await context.sync();
   
        const jobLatestMap = new Map();

        for (let i = 1; i < dynamicICR.values.length; i++) {
            const itemCode = String(dynamicICR.values[i][0]).trim();
            const dateStr = plannedStart.values[i] ? String(plannedStart.values[i][0]).trim() : "";
            const date = ExcelDateToJSDate(dateStr);
            date.setHours(0,0,0,0);
            const job = String(jobNumber.values[i][0]).trim();
            const qty = Number(dynamicQR.values[i][0]);

            if (itemCode && date >= startDate && date <= endDate) {
                if (!jobLatestMap.has(job) || date > jobLatestMap.get(job).date) {
                    jobLatestMap.set(job, {itemCode, qty, date, job});
                }
            }
        }
        filter = Array.from(jobLatestMap.values());

        earlyDateMap.clear();
        for (const entry of filter) {
            const { itemCode, date } = entry;
            if (!earlyDateMap.has(itemCode) || date < earlyDateMap.get(itemCode)) {
                earlyDateMap.set(itemCode, date);
            }
        }

        filter.sort((a,b) => a.date - b.date);
        await context.sync();

        for (let i = 1; i < dynamicICR.values.length; i++){
            const itemCode = String(dynamicICR.values[i][0]).trim();
            const dateStr = plannedStart.values[i] ? String(plannedStart.values[i][0]).trim() : "";
            const date = ExcelDateToJSDate(dateStr);
            date.setHours(0,0,0,0);
            const job = String(jobNumber.values[i][0]).trim();
            const qty = Number(dynamicQR.values[i][0]);
            if (itemCode){
                allData.push({itemCode, qty, job, date});
            }
        }
    });    
}

async function otherDateFilter() {
    await Excel.run(async (context) => {
        const startDateInput = document.getElementById('start-date').value;
        const endDateInput = document.getElementById('end-date').value;

        const startDate = inputDateParse(startDateInput);
        const endDate = inputDateParse(endDateInput);

        const dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
        const dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
        await context.sync();

        const dynamicHeaders = dynamicUsedRange.values[0];

        const dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
        const dynStartIdx = dynamicHeaders.indexOf("Planned Start");
        const dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
        const dynJobIdx = dynamicHeaders.indexOf("Order Number");

        const dynStartColumn = `${colIdxToLetter(dynStartIdx)}:${colIdxToLetter(dynStartIdx)}`;
        const dynItemQtyColumn = `${colIdxToLetter(dynItemQtyIdx)}:${colIdxToLetter(dynItemQtyIdx)}`;
        const dynItemCodeColumn = `${colIdxToLetter(dynItemCodeIdx)}:${colIdxToLetter(dynItemCodeIdx)}`;
        const dynJobColumn =`${colIdxToLetter(dynJobIdx)}:${colIdxToLetter(dynJobIdx)}`;

        const dynamic = context.workbook.worksheets.getItem("Dynamic");
        const dynamicICR = dynamic.getRange(dynItemCodeColumn).getUsedRange().load("values");
        const dynamicQR = dynamic.getRange(dynItemQtyColumn).getUsedRange().load("values"); 
        const plannedStart = dynamicWorksheet.getRange(dynStartColumn).getUsedRange().load("values");
        const jobNumber = dynamicWorksheet.getRange(dynJobColumn).getUsedRange().load("values");
        await context.sync();
        const jobLatestMap = new Map();

        for (let i = 1; i < dynamicICR.values.length; i++) {
            const itemCode = String(dynamicICR.values[i][0]).trim();
            const dateStr = plannedStart.values[i] ? String(plannedStart.values[i][0]).trim() : "";
            const date = ExcelDateToJSDate(dateStr);
            date.setHours(0,0,0,0);
            const job = String(jobNumber.values[i][0]).trim();
            const qty = Number(dynamicQR.values[i][0]);

            if (itemCode && date >= startDate && date <= endDate) {
                if (!jobLatestMap.has(job) || date > jobLatestMap.get(job).date) {
                    jobLatestMap.set(job, {itemCode, qty, date, job});
                }
            }
        }
        invFilter = Array.from(jobLatestMap.values());

        startDateMap.clear();
        for (const entry of invFilter) {
            const { itemCode, date } = entry;
            if (!startDateMap.has(itemCode) || date < startDateMap.get(itemCode)) {
                startDateMap.set(itemCode, date);
            }
        }

        invFilter.sort((a,b) => a.date - b.date);
        await context.sync();
    });    
} 
function inputDateParse(str) {
    const [year, month, day] = str.split('-').map(Number);
    return new Date(year, month - 1, day);
}

function ExcelDateToJSDate(excelDate) {
    const utcDate = new Date((excelDate - 25569) * 86400 * 1000);
    return new Date(utcDate.getTime() + (utcDate.getTimezoneOffset() * 60000));
}

function checkDatesAndClearMessage() {
    const startDateValue = document.getElementById('start-date').value;
    const endDateValue = document.getElementById('end-date').value;
    if (startDateValue && endDateValue) {
        document.getElementById('message-area').textContent = "";
    }
}

async function displayData (event){
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("Ordering");
        const range = sheet.getRange(event.address);
        range.load(["columnIndex", "values", "address"]);
        await context.sync(); 
        console.log("Index Number", range.columnIndex);
        outputJobs.clear();

        if (range.columnIndex == 0){
            matchingData.length = 0; 
            const match = range.values[0];

            const allDataICR = allData.map(item => [item.itemCode]);
            const allDatajob = allData.map(item => [item.job]);
            const allDataQR = allData.map(item => [item.qty]);
            const allDatadate = allData.map(item => [item.date]);

            for (let i = 0; i < allDataICR.length; i++){
                const code = allDataICR[i][0];
                const job = allDatajob[i][0];
                const qty = allDataQR[i][0];
                const date = allDatadate[i][0]; 
                const fDate = formatDate(date);

                if (match == code) {
                    if (!outputJobs.has(job)) {
                    matchingData.push({ code, job, qty, fDate });
                    outputJobs.add(job);
                    }else {
                        const duplicateDate = earlyDateMap.get(code);
                        const idx = matchingData.findIndex(entry => entry.job === job && entry.code === code);
                        if (idx !== -1) {
                            matchingData[idx].date = duplicateDate;
                        }
                    }
                }
            }
            console.log("intial finding of Matching Data", matchingData);
            handleCellChange([...matchingData]);
            matchingData.sort((a,b) => a.date - b.date);

            localStorage.setItem("matchingData", JSON.stringify(matchingData));
        }
        else{
            console.log("Not in range");
        }
        
  });
}

async function filteringDropdown() {
    await Excel.run(async (context) => {
        const orderingWorksheet = context.workbook.worksheets.getItem("Ordering");
        const orderingTable = orderingWorksheet.tables.getItem("OrderingTable");
        const amountFilter = orderingTable.columns.getItem("Required Amount").filter;
        const buyOrMakeFilter = orderingTable.columns.getItem("Buy or Make").filter;

        switch(document.getElementById('order-filtering').value) {
            case "Intial":
                console.log("no changes made");
                amountFilter.clear();
                buyOrMakeFilter.clear();
                break;
            case "over-300":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                amountFilter.applyCustomFilter(">=300");
                break;
            case "Must-buy":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                buyOrMakeFilter.applyValuesFilter(["Must Buy"]);
                break;
            case "Can-buy":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                buyOrMakeFilter.applyValuesFilter(["Can Buy"]);
                break;
            case "Can-make":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                buyOrMakeFilter.applyValuesFilter(["Can Make"]);
                break;
            default:
                console.log("No valid filter selected");
                break;
        }
        await context.sync();
    });
}

function formatDate(dateString) {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
    }).replace(/\//g, '-');
}