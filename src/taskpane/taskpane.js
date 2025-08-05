import { handleCellChange } from '../display/display.js';
import { openSettings } from '../settings/settings.js';

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("generate-ordering-report").onclick = () => tryCatch(generateOrderingReport);
    document.getElementById("generate-inventory-report").onclick = () => tryCatch(generateInventoryReport);
    document.getElementById('start-date').addEventListener('input', checkDatesAndClearMessage);
    document.getElementById('end-date').addEventListener('input', checkDatesAndClearMessage);
    document.getElementById("order-filtering").addEventListener('change', filteringDropdown);
    document.getElementById("inventory-filtering").addEventListener('change', invFilteringDropdown);
    document.getElementById("settings-button").onclick = () => tryCatch(openSettings);
    setInterval(() => {
        test().catch(console.error);
    }, 5000);
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
            await context.sync();

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

        const inventoryReadout = context.workbook.worksheets.add("Inventory Request");
        const inventoryReadoutTable = inventoryReadout.tables.add("A1:D1", true);

        inventoryTable.name = "InventoryAtTable";
        inventoryTable.style = "TableStyleMedium10";

        inventoryReadoutTable.name = "InventoryReadoutTable";
        inventoryReadoutTable.style = "TableStyleMedium10";
        
        inventoryReadoutTable.getHeaderRowRange().values = [["Item Code", "Inventory Date", "Inventory Ref", "Inventory Qty"]];
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
            await context.sync();
            await importColumnData();
            readoutData();
        }
        await context.sync();
    });
}

async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        const messageArea = document.getElementById('message-area');
        if (messageArea) {
            messageArea.textContent = `Error: ${error.message || error}`;
            messageArea.style.color = 'red';
            setTimeout(() => {
                messageArea.textContent = '';
                messageArea.style.color = '';
            }, 5000);
        }
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
        const requiredAmounts = result.map(row => [row[1]]);
        console.log("Required Amounts", requiredAmounts);
        
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

        const efwSumMap = buildSumMap(efwArray.map(item => [item[0]]), efwArray.map(item => [item[1]]));
        const efwAmounts = Array.from(efwSumMap.entries()).map(row => [row[1]]);

        const total = mebAmounts.map((value, index) => Number(value) + efwAmounts[index][0]);
        
        //Inventory On Order 
        const openPOsMap = buildSumMap(openPOsICR.values, openPOsQR.values);
        const onOrder = [["On Order"]]; 
        for (const code of caseNumbers.slice(1)) {
            const onOrderQty = openPOsMap.get(code) || 0;
            onOrder.push([onOrderQty]);           
        }
        
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

        const qtyNeeded = requiredAmounts.slice(1).map((value, index) => 
            Number(value) - mebAmounts[index]
        );

        const filteredData = [];
        filteredData.push({
            caseNumber: "Case #",
            demand: "Demand", 
            mebQty: "Qty MEB",
            efwQty: "Qty EFW",
            totalQty: "Total MEB + EFW",
            onOrder: "On Order",
            startDate: "Start Date",
            releaseDate: "Release Date",
            qtyNeeded: "Qty Needed (MEB)",
            notes: "Notes"
        });

        for (let i = 0; i < caseNumbers.slice(1).length; i++) {
            if (qtyNeeded[i] > 0) { 
                filteredData.push({
                    caseNumber: caseNumbers[i + 1],
                    demand: requiredAmounts[i + 1][0],
                    mebQty: mebAmounts[i][0],
                    efwQty: efwAmounts[i][0],
                    totalQty: total[i],
                    onOrder: onOrder[i + 1][0],
                    startDate: startArray[i + 1][0],
                    releaseDate: releaseArray[i + 1][0],
                    qtyNeeded: qtyNeeded[i],
                    notes: ""
                });
            }
        }

        if (filteredData.length > 1) { 
            const caseNumbersFiltered = filteredData.map(row => [row.caseNumber]);
            const demandFiltered = filteredData.map(row => [row.demand]);
            const mebQtyFiltered = filteredData.map(row => [row.mebQty]);
            const efwQtyFiltered = filteredData.map(row => [row.efwQty]);
            const totalQtyFiltered = filteredData.map(row => [row.totalQty]);
            const onOrderFiltered = filteredData.map(row => [row.onOrder]);
            const startDateFiltered = filteredData.map(row => [row.startDate]);
            const releaseDateFiltered = filteredData.map(row => [row.releaseDate]);
            const qtyNeededFiltered = filteredData.map(row => [row.qtyNeeded]);
            const notesFiltered = filteredData.map(row => [row.notes]);

            inventoryWorksheet.getRange(`A1:A${caseNumbersFiltered.length}`).values = caseNumbersFiltered;
            inventoryWorksheet.getRange(`B1:B${demandFiltered.length}`).values = demandFiltered;
            inventoryWorksheet.getRange(`C1:C${mebQtyFiltered.length}`).values = mebQtyFiltered;
            inventoryWorksheet.getRange(`D1:D${efwQtyFiltered.length}`).values = efwQtyFiltered;
            inventoryWorksheet.getRange(`E1:E${totalQtyFiltered.length}`).values = totalQtyFiltered;
            inventoryWorksheet.getRange(`F1:F${onOrderFiltered.length}`).values = onOrderFiltered;
            inventoryWorksheet.getRange(`G1:G${startDateFiltered.length}`).values = startDateFiltered;
            inventoryWorksheet.getRange(`H1:H${releaseDateFiltered.length}`).values = releaseDateFiltered;
            inventoryWorksheet.getRange(`I1:I${qtyNeededFiltered.length}`).values = qtyNeededFiltered;
            inventoryWorksheet.getRange(`J1:J${notesFiltered.length}`).values = notesFiltered;
        }
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
        
        inventoryWorksheet.getRange("I1:I1").format.fill.color = "#00008B";
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
        const lastRow = filteredData.length;
        const highlight = inventoryWorksheet.getRange(`I1:I${lastRow}`).format.borders;
         [
            "EdgeTop",
            "EdgeBottom",
            "EdgeLeft",
            "EdgeRight",
        ].forEach(side => {
            highlight.getItem(side).style = "Continuous";
            highlight.getItem(side).weight = "Thick";
            highlight.getItem(side).color = "#00008B"; 
        });

        await context.sync();
    });
}

async function readoutData(){
    await Excel.run(async (context) => {
        const inventoryAtWorksheet = context.workbook.worksheets.getItem("Inventory At");
        const inventoryAtUsedRange = inventoryAtWorksheet.getUsedRange().load("values");

        const inventoryWorksheet = context.workbook.worksheets.getItem("Inventory");
        const inventoryUsedRange = inventoryWorksheet.getUsedRange().load("values");

        const inventoryRequestWorksheet = context.workbook.worksheets.getItem("Inventory Request");

        await context.sync();

        //Inventory At fluid Placement
        const inventoryAtHeaders = inventoryAtUsedRange.values[0];
        
        const invAtItemCodeIdx = inventoryAtHeaders.indexOf("Case #");
        const invAtQtyNeededIdx = inventoryAtHeaders.indexOf("Qty Needed (MEB)");
        const invAtQtyEFWIdx = inventoryAtHeaders.indexOf("Qty EFW");
        
        const invAtItemCodeColumn = `${colIdxToLetter(invAtItemCodeIdx)}:${colIdxToLetter(invAtItemCodeIdx)}`;
        const invAtQtyNeededColumn = `${colIdxToLetter(invAtQtyNeededIdx)}:${colIdxToLetter(invAtQtyNeededIdx)}`;
        const invAtQtyEFWColumn = `${colIdxToLetter(invAtQtyEFWIdx)}:${colIdxToLetter(invAtQtyEFWIdx)}`;

        //Inventory Report fluid Placement
        const inventoryHeaders = inventoryUsedRange.values[0];

        const invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
        const invDateIdx = inventoryHeaders.indexOf("Inventory Date");
        const invRefIdx = inventoryHeaders.indexOf("Inventory Ref");
        const invQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
        const invLocIdx = inventoryHeaders.indexOf("Location");
        
        const invRepItemCodeColumn = `${colIdxToLetter(invItemCodeIdx)}:${colIdxToLetter(invItemCodeIdx)}`;
        const invRepDateColumn = `${colIdxToLetter(invDateIdx)}:${colIdxToLetter(invDateIdx)}`;
        const invRepRefColumn = `${colIdxToLetter(invRefIdx)}:${colIdxToLetter(invRefIdx)}`;
        const invRepQtyColumn = `${colIdxToLetter(invQtyIdx)}:${colIdxToLetter(invQtyIdx)}`;
        const invLocColumn = `${colIdxToLetter(invLocIdx)}:${colIdxToLetter(invLocIdx)}`;

        await context.sync();

        //Get data from Inventory At sheet
        const invAtItemCodes = inventoryAtWorksheet.getRange(invAtItemCodeColumn).getUsedRange().load("values");
        const invAtQtyNeeded = inventoryAtWorksheet.getRange(invAtQtyNeededColumn).getUsedRange().load("values");
        const invAtQtyEFW = inventoryAtWorksheet.getRange(invAtQtyEFWColumn).getUsedRange().load("values");

        //Get data from Inventory sheet
        const invItemCodes = inventoryWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values");
        const invDates = inventoryWorksheet.getRange(invRepDateColumn).getUsedRange().load("values");
        const invRefs = inventoryWorksheet.getRange(invRepRefColumn).getUsedRange().load("values");
        const invQtys = inventoryWorksheet.getRange(invRepQtyColumn).getUsedRange().load("values");
        const invLoc = inventoryWorksheet.getRange(invLocColumn).getUsedRange().load("values");
        await context.sync();

        const filteredData = [];
        for (let i = 1; i < invAtItemCodes.values.length; i++) {
            const qtyNeeded = Number(invAtQtyNeeded.values[i][0]);
            const qtyEFW = Number(invAtQtyEFW.values[i][0]);
            
            if (qtyNeeded > 300 && qtyEFW > 0) {
                filteredData.push({
                    itemCode: String(invAtItemCodes.values[i][0]).trim(),
                    qtyNeeded: qtyNeeded
                });
            }
        }

        //Build inventory data map for each item code
        const inventoryDataMap = new Map();
        for (let i = 1; i < invItemCodes.values.length; i++) {
            const itemCode = String(invItemCodes.values[i][0]).trim();
            const dateStr = invDates.values[i] ? String(invDates.values[i][0]).trim() : "";
            const start = ExcelDateToJSDate(dateStr);
            start.setHours(0,0,0,0);
            const date = formatDate(start);
            const loc = invLoc.values[i] ? String(invLoc.values[i][0]).trim() : "";
            const ref = invRefs.values[i] ? String(invRefs.values[i][0]).trim() : "";
            const qty = Number(invQtys.values[i][0]);

            if (loc.includes("EFW")){
                if (itemCode && !isNaN(qty) && qty > 0) {
                    if (!inventoryDataMap.has(itemCode)) {
                        inventoryDataMap.set(itemCode, []);
                    }
                    inventoryDataMap.get(itemCode).push({
                        date: date,
                        ref: ref,
                        qty: qty
                    });
                }
            }
        }

        //Sort inventory data by date (oldest first) for each item code
        for (const [itemCode, data] of inventoryDataMap) {
            data.sort((a, b) => {
                const dateA = new Date(a.date);
                const dateB = new Date(b.date);
                return dateA - dateB;
            });
        }

        //Generate readout data
        const readoutResult = [["Item Code", "Inventory Date", "Inventory Ref", "Inventory Qty"]];
        
        for (const filteredItem of filteredData) {
            const itemCode = filteredItem.itemCode;
            const qtyNeeded = filteredItem.qtyNeeded;
            
            if (inventoryDataMap.has(itemCode)) {
                const inventoryItems = inventoryDataMap.get(itemCode);
                let totalPulled = 0;
                let palletsPulled = 0;
                
                for (const invItem of inventoryItems) {
                    if (totalPulled >= qtyNeeded) {
                        break; 
                    }
                    
                    readoutResult.push([
                        itemCode,
                        invItem.date,
                        invItem.ref,
                        invItem.qty
                    ]);
                    
                    totalPulled += invItem.qty;
                    palletsPulled++;
                }
            }
        }

        //Write data to Inventory Request table
        if (readoutResult.length > 1) {
            const itemCodes = readoutResult.map(row => [row[0]]);
            const dates = readoutResult.map(row => [row[1]]);
            const refs = readoutResult.map(row => [row[2]]);
            const qtys = readoutResult.map(row => [row[3]]);

            inventoryRequestWorksheet.getRange(`A1:A${itemCodes.length}`).values = itemCodes;
            inventoryRequestWorksheet.getRange(`B1:B${dates.length}`).values = dates;
            inventoryRequestWorksheet.getRange(`C1:C${refs.length}`).values = refs;
            inventoryRequestWorksheet.getRange(`D1:D${qtys.length}`).values = qtys;
        }
        await context.sync();

        //Formatting
        inventoryRequestWorksheet.getRange("A:D").format.autofitColumns();
        inventoryRequestWorksheet.getRange("A:D").format.horizontalAlignment = "Center";
        inventoryRequestWorksheet.getRange("A:D").format.verticalAlignment = "Center";

        inventoryRequestWorksheet.getRange("A:A").format.columnWidth = 120;
        inventoryRequestWorksheet.getRange("B:B").format.columnWidth = 100;
        inventoryRequestWorksheet.getRange("C:C").format.columnWidth = 100;
        inventoryRequestWorksheet.getRange("D:D").format.columnWidth = 100;
        inventoryRequestWorksheet.getUsedRange().format.rowHeight = 20;

        inventoryRequestWorksheet.freezePanes.freezeRows(1);

        //All border lines
        const usedRange = inventoryRequestWorksheet.getUsedRange();   
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
            sheets.getItemOrNullObject("Inventory Request").delete();
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

            const orderingTable = sheet.tables.getItem("OrderingTable");
            const tableRange = orderingTable.getDataBodyRange().load("values");
            const headers = orderingTable.getHeaderRowRange().load("values");
            await context.sync();
            const headerRow = headers.values[0];
            const codeIdx = headerRow.indexOf("Case #");
            const invIdx = headerRow.indexOf("Current Inventory");
            let currentInventory = 0;
            for (let i = 0; i < tableRange.values.length; i++) {
                const code = String(tableRange.values[i][codeIdx]).trim();
                if (code === match[0]) {
                    currentInventory = Number(tableRange.values[i][invIdx]) || 0;
                    break;
                }
            }
            localStorage.setItem('currentInventory', currentInventory);

            for (let i = 0; i < allDataICR.length; i++){
                const code = allDataICR[i][0];
                const job = allDatajob[i][0];
                const qty = allDataQR[i][0];
                const date = allDatadate[i][0]; 
                const fDate = formatDate(date);

                if (match == code) {
                    if (!outputJobs.has(job)) {
                        matchingData.push({ code, job, qty, fDate, date });
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
            matchingData.sort((a, b) => a.date - b.date);

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
        const corMinimums = {
            "CORS522_DW": 180,
            "CORCTD0033A-R2": 300,
            "COR2503503R0 RDC": 250,
            "COR2320-4731": 350,
            "COR16M006405": 350,
            "COR6064-RDC": 445,
            "CORM30402_C": 310,
            "CORM37238 FLUTED DIV": 500,
            "CORM30403_C": 250,
            "CORM39142-R1": 250,
            "CORM37912": 300,
            "CORX14107": 300,
            "COR2320-4819-R1": 300,
            "COR2320-4658-R1": 250,
            "COR17M001501": 300,
            "COR2320_5453": 300,
            "COR2320-1840-R1": 300,
            "COR2320-5573": 400,
            "COR2320_5575_R1": 349,
            "COR2320_5455": 300,
            "COR2866 RDC": 250,
            "COR6062 RDC": 250,
            "COR2320_5635": 250,
            "COR6070 RDC": 250,
            "CORM33989_H_R2": 283,
            "CORMPS13185": 250,
            "COR18M020101": 400,
            "CORX14151_B": 250,
            "COR11707_R1": 250,
            "CORM38150_R1": 250,
            "COR2320_5906": 199,
            "COR17M013701_R1": 197,
            "COR5537 PAD": 273,
            "CORM36590_ERECTED_R2": 250,
            "CORF170286A9": 340,
            "CORF170287A7-R1": 307,
            "CORF170313A6": 274,
            "COR6052": 250,
            "COR2306R0": 285,
            "CORM37798_C_R1": 185,
            "COR2320_6830_R1": 325,
            "COR19M001713": 300,
            "CORERECTOR-DW": 207,
            "COR19M001712_R1": 192,
            "CORMPS13113C_R1": 200,
            "CORL9466A4_M": 258,
            "COR2349R2_R1": 300,
            "CORMPS13182A_R1": 300,
            "CORDRF_L8916A1": 210,
            "COR19M001702_R1": 404,
            "CORL10648A2": 300,
            "CORDRF_6171_A": 257,
            "COR132790_2i": 264,
            "CORDRF_2037": 271,
            "CORDRF_2232": 300,
            "COR2251R0_R1": 300,
            "CORM38149_R1": 375,
            "COR19M012303": 375,
            "COR19M012304": 305,
            "COR20M023301": 300,
            "COR20M007911": 186,
            "CORM34795_NO INSERT HSC LID": 975,
            "CORM34795_NO INSERT HSC": 975,
            "COR20M018813": 201,
            "COR20M018814": 278,
            "CORDRF_2884_B": 350,
            "CORDRF_L9681A2": 218,
            "COR20M018815_R1": 186,
            "COR15NF011001-R1": 277,
            "COR20M026410_R1": 332,
            "COR2320_5691": 305,
            "COR16NF0805.03": 305,
            "COR2533504A_R1": 320,
            "COR16M006404_R1": 269,
            "COR18M010901": 315,
            "COR20M018812": 257,
            "COR18M025708": 325,
            "CORM37238 GLORD SLEEVE_R1": 275,
            "CORABBOTT_CAP": 138,
            "CORABBOTT_SLV": 109,
            "COR19M010308": 325,
            "COR2320_7148_R2": 263,
            "COR16M006402": 288,
            "COR18M005551_R2": 367,
            "COR21M022706": 323,
            "COR17M017101_R1": 227,
            "CORM36342": 200,
            "COR19M019402_R1": 304,
            "COR16M006403": 244,
            "COR20M026401_R1": 350,
            "COR20M026402": 363,
            "COR20M026418": 237,
            "COR18M005551_R3 NR": 350,
            "COR20M016311_R3": 216,
            "CORM34795 RSC": 385,
            "CORM38164": 450,
            "COR19M019401_R1": 300,
            "COR19M019402_R2": 300,
            "CORM34084_A": 188,
            "COR2320_6257_R2": 303,
            "COR18M000117": 202,
            "CORM21854_A": 340,
            "COR2320_5582_R1": 375,
            "COR21M012001": 417,
            "COR16M002106": 330,
            "COR16M006401_R1": 282,
            "COR16M000901_R3": 238,
            "COR17M022101_R2": 223,
            "CORM41161_B": 297,
            "COR23M012203": 261,
            "CORL14795A1": 203,
            "CORL9486A1": 279,
            "COR18M025707_R2": 265,
            "CORF210335A4": 296,
            "COR23LG0002.01": 316,
            "COR23M016803": 300,
            "CORL11558A4_M": 272,
            "CORM33618_B": 241,
            "COR17CL0236.03": 350,
            "CORCSK11804A": 282,
            "TB008-01": 500,
            "T22925-5": 800,
            "CORM37238 CAP": 154,
            "CORCSK11317B": 221,
            "CORL12394A1": 266,
            "CORMPS13215C": 231,
            "COR13442M": 330,
            "COR21LG0020.01": 305,
            "CORMPS13117J": 200,
            "CORM38149_R2": 245,
            "CORM34795_INSERT_48": 768,
            "CORL14798A1": 191,
            "COR20PF0474.02": 315,
            "COR16M002105_R1": 350,
            "COR21LG0002.01": 200,
            "CORCSK12007.02_R1": 315,
            "COR22M023719": 283,
            "CORMPS13048": 340,
            "CORMPS13150": 405,
            "COR16M006405_R1": 256,
            "COR17M027920_R2": 340,
            "CORM34795_INSERT_48_R1": 625,
            "COR23M014207": 310,
            "CORM33393_E ERECTED": 187,
            "CORM26788_C_R2": 276,
            "COR20M017216": 195,
            "COR22PF0925.07": 347,
            "CORMPS13145B": 198,
            "COR22PF0834.01_R1": 264,
            "COR24LG0016.04": 313,
            "COR24LG0016.02_R1": 334,
            "COR24LG0007.01": 261,
            "CORL9184A1_R1": 349,
            "COR22M009502_R1": 295,
            "COR22PF0924.07": 301,
            "COR17M017101_R3": 349,
            "COR22M009501": 296,
            "COR22M009503": 186,
            "CORDRF_L9680A3": 209,
            "CORMPS13150_R1": 370,
            "CORMPS13004": 465,
            "COR22M009501_R1": 292
        };

        await context.sync();

        switch(document.getElementById('order-filtering').value) {
            case "Intial":
                console.log("no changes made");
                amountFilter.clear();
                buyOrMakeFilter.clear();
                orderingTable.columns.getItem("Case #").filter.clear();
                break;
            case "cor-minimum":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                const tableRange = orderingTable.getDataBodyRange().load("values");
                await context.sync();

                const headers = orderingTable.getHeaderRowRange().load("values");
                await context.sync();

                const headerRow = headers.values[0];
                const codeIdx = headerRow.indexOf("Case #");
                const reqAmtIdx = headerRow.indexOf("Required Amount");

                const keepCodes = [];
                for (let i = 0; i < tableRange.values.length; i++) {
                    const code = String(tableRange.values[i][codeIdx]).trim();
                    const reqAmt = Number(tableRange.values[i][reqAmtIdx]);
                    const min = corMinimums[code];
                    if (min !== undefined && reqAmt >= min) {
                        keepCodes.push(code);
                    }
                }

                orderingTable.columns.getItem("Case #").filter.applyValuesFilter(keepCodes);

                break;
            case "over-300":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                orderingTable.columns.getItem("Case #").filter.clear();
                amountFilter.applyCustomFilter(">=300");
                break;
            case "Must-buy":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                orderingTable.columns.getItem("Case #").filter.clear();
                buyOrMakeFilter.applyValuesFilter(["Must Buy"]);
                break;
            case "Can-buy":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                orderingTable.columns.getItem("Case #").filter.clear();
                buyOrMakeFilter.applyValuesFilter(["Can Buy"]);
                break;
            case "Can-make":
                amountFilter.clear();
                buyOrMakeFilter.clear();
                orderingTable.columns.getItem("Case #").filter.clear();
                buyOrMakeFilter.applyValuesFilter(["Can Make"]);
                break;
            default:
                console.log("No valid filter selected");
                break;
        }
        await context.sync();
    });
}

async function invFilteringDropdown() {
    await Excel.run(async (context) => {
        const inventoryWorksheet = context.workbook.worksheets.getItem("Inventory At");
        const inventoryTable = inventoryWorksheet.tables.getItem("InventoryAtTable");
        const qtyNeededFilter = inventoryTable.columns.getItem("Qty Needed (MEB)").filter;

        switch(document.getElementById('inventory-filtering').value) {
            case "Intial case":
                console.log("no changes made");
                qtyNeededFilter.clear();
                break;
            case "over-300":
                qtyNeededFilter.clear();
                qtyNeededFilter.applyCustomFilter(">=300");
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

async function test() {
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        const targetSheets = sheets.items.filter(s =>
            ["Sheet1", "Sheet2", "Sheet3"].includes(s.name)
        );

        for (const sheet of targetSheets) {
            const ws = sheets.getItem(sheet.name);
            ws.load("name");
            await context.sync();

            const usedRange = ws.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();

            if (!usedRange.values || usedRange.values.length === 0) continue;

            const headers = usedRange.values[0].map(h => String(h).trim().toLowerCase());

            let targetName = null;
            if (
                headers.includes("item code") &&
                headers.includes("inventory qty") &&
                headers.includes("location") &&
                headers.includes("inventory date")
            ) {
                targetName = "INVENTORY";
            } else if (
                headers.includes("work center") &&
                headers.includes("planned start") &&
                headers.includes("corrugate") &&
                headers.includes("number of corrugate")
            ) {
                targetName = "DYNAMIC";
            } else if (
                headers.includes("item code") &&
                headers.includes("outstanding qty")
            ) {
                targetName = "OPEN PO'S";
            }

            if (targetName && ws.name !== targetName) {
                const nameTaken = sheets.items.some(s => s.name === targetName);
                if (!nameTaken) {
                    ws.name = targetName;
                } else {
                    let counter = 2;
                    let newName = `${targetName} (${counter})`;
                    while (sheets.items.some(s => s.name === newName)) {
                        counter++;
                        newName = `${targetName} (${counter})`;
                    }
                    ws.name = newName;
                }
                await context.sync();
            }
        }
    });
}