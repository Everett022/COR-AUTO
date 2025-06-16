/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("generate-report").onclick = () => tryCatch(generateReport);
    document.getElementById("temp-reset").onclick = () => tryCatch(resetAll);
  }
});

async function generateReport() {
    await Excel.run(async (context) => {

        const orderingWorksheet = context.workbook.worksheets.add("Ordering");
        const inventoryWorksheet = context.workbook.worksheets.add("Inventory At");
        const orderingTable = orderingWorksheet.tables.add("A1:G1", true);
        const inventoryTable = inventoryWorksheet.tables.add("A1:F1", true);

        orderingTable.name = "OrderingTable";
        inventoryTable.name = "InventoryAtTable";

        orderingTable.getHeaderRowRange().values = [["Case #","Demand","Current Inventory", "On Order", "Required Amount","Order or Make", "Planned Start Date"]];
        inventoryTable.getHeaderRowRange().values = [["Material", "Demand", "MEB", "EFW", "Release", "Planned Start Date"]];
     
        orderingTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
        orderingTable.getRange().format.autofitColumns();
        orderingTable.getRange().format.autofitRows();

        inventoryTable.columns.getItemAt(2).getRange().numberFormat = [['\u20AC#,##0.00']];
        inventoryTable.getRange().format.autofitColumns();
        inventoryTable.getRange().format.autofitRows();
        
        importColumnData();
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
        const dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");;
        
        const openPOsWorksheet = context.workbook.worksheets.getItem("Open PO's");
        const openPOsUsedRange = openPOsWorksheet.getUsedRange().load("values");

        await context.sync();

        //Dynamic fluid Placement
        const dynamicHeaders = dynamicUsedRange.values[0];
        
        const dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
        const dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
        const dynStartIdx = dynamicHeaders.indexOf("Planned Start");
        
        const dynItemCodeColumn = `${colIdxToLetter(dynItemCodeIdx)}:${colIdxToLetter(dynItemCodeIdx)}`;
        const dynItemQtyColumn = `${colIdxToLetter(dynItemQtyIdx)}:${colIdxToLetter(dynItemQtyIdx)}`;
        const dynStartColumn = `${colIdxToLetter(dynStartIdx)}:${colIdxToLetter(dynStartIdx)}`;

        //Open PO's fluid Placement
        const openPOsHeaders = openPOsUsedRange.values[0];
        const openPOItemCodeIdx = openPOsHeaders.indexOf("Item Code");
        const openPOItemQtyIdx = openPOsHeaders.indexOf("Outstanding Qty");
        
        const openPOItemCodeColumn = `${colIdxToLetter(openPOItemCodeIdx)}:${colIdxToLetter(openPOItemCodeIdx)}`;
        const openPOItemQtyColumn = `${colIdxToLetter(openPOItemQtyIdx)}:${colIdxToLetter(openPOItemQtyIdx)}`;
       
        //Inventory Report Fluid Placement
        const inventoryHeaders = inventoryUsedRange.values[0];
        console.log("Headers:", inventoryHeaders);
        
        const invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
        const invItemQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
        
        const invRepItemCodeColumn = `${colIdxToLetter(invItemCodeIdx)}:${colIdxToLetter(invItemCodeIdx)}`;
        const invRepItemQtyColumn = `${colIdxToLetter(invItemQtyIdx)}:${colIdxToLetter(invItemQtyIdx)}`;

        //Quanity and Item Code from Dynamic, Inventory Report, and Open PO's sheets
        const dynamic = context.workbook.worksheets.getItem("Dynamic");
        const dynamicICR = dynamic.getRange(dynItemCodeColumn).getUsedRange().load("values");
        const dynamicQR = dynamic.getRange(dynItemQtyColumn).getUsedRange().load("values"); 

        const inventoryICR = inventoryReportWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values"); 
        const inventoryQR = inventoryReportWorksheet.getRange(invRepItemQtyColumn).getUsedRange().load("values"); 

        const openPOs = context.workbook.worksheets.getItem("Open PO's");
        const openPOsICR = openPOs.getRange(openPOItemCodeColumn).getUsedRange().load("values"); 
        const openPOsQR = openPOs.getRange(openPOItemQtyColumn).getUsedRange().load("values"); 
        await context.sync();

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

        const dynamicMap = buildSumMap(dynamicICR.values, dynamicQR.values);
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
        
        const orderingSheet = context.workbook.worksheets.getItem("Ordering");
        const caseNumbers = result.map(row => [row[0]]);
        orderingSheet.getRange(`A1:A${caseNumbers.length}`).values = caseNumbers;
        
        const requiredAmounts = result.map(row => [row[1]]);
        orderingSheet.getRange(`E1:E${requiredAmounts.length}`).values = requiredAmounts;
        await context.sync();
        
        //Importing the Planned Start Date
        const plannedStart = dynamicWorksheet.getRange(dynStartColumn).getUsedRange().load("values"); 
        const orderingWorksheet = context.workbook.worksheets.getItem("Ordering");
        const orderingUsedRange = orderingWorksheet.getUsedRange().load("values");
        console.log("ordering used range", orderingUsedRange);
        await context.sync();

        const orderingValues = orderingUsedRange.values;
        const startMap = new Map();
        for (let i = 1; i < dynamicICR.values.length; i++) { 
            const itemCode = String(dynamicICR.values[i][0]).trim();
            const start = plannedStart.values[i] ? String(plannedStart.values[i][0]).trim() : "";
            if (itemCode) {
                startMap.set(itemCode, start);
            }
        }
        
        const startArray = [["Planned Start Date"]];
        for (let i = 1; i < orderingValues.length; i++) {
            const itemCode = String(orderingValues[i][0]).trim();
            const start = startMap.get(itemCode) || "No Start Date found";
            startArray.push([start]);
        }
        orderingWorksheet.getRange(`G1:G${startArray.length}`).values = startArray;

        //Importing Demand, Current Inventory, and On Order
        const demandSumMap = buildSumMap(orderingUsedRange.values, dynamicQR.values);
        
        const demand = [["Demand"]]; 
        for (const code of demandSumMap.keys()) {
            const dynamicQty = demandSumMap.get(code) || 0;
          if (dynamicQty > 0){
                demand.push([dynamicQty]);
          }      
        }
        
        const demandOutput = demand.map(row => [row[0]]);
        orderingSheet.getRange(`B1:B${demandOutput.length}`).values = demandOutput;

        // Table Formatting
        orderingWorksheet.getRange("G:G").numberFormat = [['m/d/yyyy h:mm']];
        orderingWorksheet.getRange("A:G").format.autofitColumns();
        orderingWorksheet.getRange("A:G").format.horizontalAlignment = "Center";
        orderingWorksheet.getRange("A:G").format.verticalAlignment = "Center";
        orderingWorksheet.freezePanes.freezeRows(1);

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

        await context.sync();
    });
}
        
async function resetAll() {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.getItemOrNullObject("Ordering").delete();
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

