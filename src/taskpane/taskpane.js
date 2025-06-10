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
        await context.sync();
        orderQty();
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

async function orderQty() {
    await Excel.run(async (context) => {
        const dynamic = context.workbook.worksheets.getItem("Dynamic");
        const dynamicICR = dynamic.getRange("D:D").getUsedRange().load("values");
        const dynamicQR = dynamic.getRange("E:E").getUsedRange().load("values"); 

        const inventoryReport = context.workbook.worksheets.getItem("Inventory Report");
        const inventoryICR = inventoryReport.getRange("A:A").getUsedRange().load("values"); 
        const inventoryQR = inventoryReport.getRange("E:E").getUsedRange().load("values"); 

        const openPOs = context.workbook.worksheets.getItem("Open PO's");
        const openPOsICR = openPOs.getRange("C:C").getUsedRange().load("values"); 
        const openPOsQR = openPOs.getRange("E:E").getUsedRange().load("values"); 
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