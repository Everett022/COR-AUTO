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
        const inventoryWorksheet = context.workbook.worksheets.add("Inventory");
        const orderingTable = orderingWorksheet.tables.add("A1:E1", true);
        const inventoryTable = inventoryWorksheet.tables.add("A1:D1", true);

        orderingTable.name = "OrderingTable";
        inventoryTable.name = "InventoryTable";

        orderingTable.getHeaderRowRange().values = [["Item Code","Quantity", "Short Description", "Order Date","Start Date"]];
        inventoryTable.getHeaderRowRange().values = [["Product", "Stock", "Price", "Reorder Level"]];
     
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
        const inventoryReportWorksheet = context.workbook.worksheets.getItem("Inventory Report");
        const inventoryUsedRange = inventoryReportWorksheet.getUsedRange().load("values");
        const dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
        const dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");
        const openPOsWorksheet = context.workbook.worksheets.getItem("Open PO's");
        const openPOsUsedRange = openPOsWorksheet.getUsedRange().load("values");

        await context.sync();
        //Dynamic fluid Placement
        const dynamicHeaders = dynamicUsedRange.values[0];
        console.log("Headers:", dynamicHeaders);
        
        const dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
        const dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
        
        const dynItemCodeColumn = `${colIdxToLetter(dynItemCodeIdx)}:${colIdxToLetter(dynItemCodeIdx)}`;
        const dynItemQtyColumn = `${colIdxToLetter(dynItemQtyIdx)}:${colIdxToLetter(dynItemQtyIdx)}`;
        
        //Open PO's fluid Placement
        const openPOsHeaders = openPOsUsedRange.values[0];
        console.log("Headers:", openPOsHeaders);
        
        const openPOItemCodeIdx = openPOsHeaders.indexOf("Item Code");
        const openPOItemQtyIdx = openPOsHeaders.indexOf("Outstanding Qty");
        
        const openPOItemCodeColumn = `${colIdxToLetter(openPOItemCodeIdx)}:${colIdxToLetter(openPOItemCodeIdx)}`;
        const openPOItemQtyColumn = `${colIdxToLetter(openPOItemQtyIdx)}:${colIdxToLetter(openPOItemQtyIdx)}`;
        //Inventory Report Fluid Placement
        const inventoryHeaders = inventoryUsedRange.values[0];
        console.log("Headers:", inventoryHeaders);
        
        const invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
        const invItemSDIdx = inventoryHeaders.indexOf("Item Short Desc");
        const invItemQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
        
        const invRepItemCodeColumn = `${colIdxToLetter(invItemCodeIdx)}:${colIdxToLetter(invItemCodeIdx)}`;
        const invRepItemShortDescColumn = `${colIdxToLetter(invItemSDIdx)}:${colIdxToLetter(invItemSDIdx)}`;
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
                const code = itemCodes[i][0]?.trim();
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

        const result = [["Item Code", "Quantity"]]; 
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
        orderingSheet.getRange(`A1:B${result.length}`).values = result;
         
        await context.sync();
        
        //Start of the Short Description Creation Column  
        const orderingWorksheet = context.workbook.worksheets.getItem("Ordering");
        const orderingUsedRange = orderingWorksheet.getUsedRange().load("values");

        await context.sync();
        const inventoryCodeRange = inventoryReportWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values");
        const itemSD = inventoryReportWorksheet.getRange(invRepItemShortDescColumn).getUsedRange().load("values");    
        
        await context.sync();
        const descMap = new Map();
        for (let i = 1; i < inventoryCodeRange.values.length; i++) { 
            const code1 = String(inventoryCodeRange.values[i][0]).trim();
            const desc = itemSD.values[i] ? String(itemSD.values[i][0]).trim() : "";
            if (code1) {
                descMap.set(code1, desc);
            }
        }

        const orderingValues = orderingUsedRange.values;
        const descArray = [["Short Description"]];
        for (let i = 1; i < orderingValues.length; i++) {
            const code1 = String(orderingValues[i][0]).trim();
            const desc = descMap.get(code1) || "No description found";
            descArray.push([desc]);
        }

        orderingWorksheet.getRange(`C1:C${descArray.length}`).values = descArray;

        orderingWorksheet.getRange("A:E").format.autofitColumns();
        orderingWorksheet.getRange("A:E").format.horizontalAlignment = "Center";
        
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
            sheets.getItemOrNullObject("Inventory").delete();
        
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