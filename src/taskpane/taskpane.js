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
//Initalize Global Variables
let orderingWorksheet, inventoryWorksheet, orderingTable, inventoryTable, inventoryReportWorksheet, inventoryUsedRange, dynamicWorksheet,dynamicUsedRange, openPOsWorksheet,
    openPOsUsedRange, dynamicHeaders, dynItemCodeIdx, dynItemQtyIdx, dynStartIdx, dynWorkIdx, dynItemCodeColumn, dynItemQtyColumn, dynStartColumn, dynWorkColumn,
    openPOsHeaders, openPOItemCodeIdx, openPOItemQtyIdx, openPOItemCodeColumn, openPOItemQtyColumn, inventoryHeaders, invItemCodeIdx, invItemQtyIdx, invRepItemCodeColumn,
    invRepItemQtyColumn, dynamic, dynamicICR, dynamicQR, dynamicWork, inventoryICR, inventoryQR, openPOs, openPOsICR, openPOsQR, dynDateColumn, plannedStart,
    orderingUsedRange;

async function generateReport() {
    await Excel.run(async (context) => {

        orderingWorksheet = context.workbook.worksheets.add("Ordering");
        inventoryWorksheet = context.workbook.worksheets.add("Inventory At");
        orderingTable = orderingWorksheet.tables.add("A1:G1", true);
        inventoryTable = inventoryWorksheet.tables.add("A1:F1", true);

        orderingTable.name = "OrderingTable";
        inventoryTable.name = "InventoryAtTable";

        orderingTable.getHeaderRowRange().values = [["Case #","Demand","Current Inventory", "On Order", "Required Amount","Buy or Make", "Earliest Start Date"]];
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
        inventoryReportWorksheet = context.workbook.worksheets.getItem("Inventory");
        inventoryUsedRange = inventoryReportWorksheet.getUsedRange().load("values");

        dynamicWorksheet = context.workbook.worksheets.getItem("Dynamic");
        dynamicUsedRange = dynamicWorksheet.getUsedRange().load("values");;
        
        openPOsWorksheet = context.workbook.worksheets.getItem("Open PO's");
        openPOsUsedRange = openPOsWorksheet.getUsedRange().load("values");

        await context.sync();

        //Dynamic fluid Placement
        dynamicHeaders = dynamicUsedRange.values[0];
        
        dynItemCodeIdx = dynamicHeaders.indexOf("Corrugate");
        dynItemQtyIdx = dynamicHeaders.indexOf("Number of Corrugate");
        dynStartIdx = dynamicHeaders.indexOf("Planned Start");
        dynWorkIdx = dynamicHeaders.indexOf("Work Center");
        
        dynItemCodeColumn = `${colIdxToLetter(dynItemCodeIdx)}:${colIdxToLetter(dynItemCodeIdx)}`;
        dynItemQtyColumn = `${colIdxToLetter(dynItemQtyIdx)}:${colIdxToLetter(dynItemQtyIdx)}`;
        dynStartColumn = `${colIdxToLetter(dynStartIdx)}:${colIdxToLetter(dynStartIdx)}`;
        dynWorkColumn = `${colIdxToLetter(dynWorkIdx)}:${colIdxToLetter(dynWorkIdx)}`;

        //Open PO's fluid Placement
        openPOsHeaders = openPOsUsedRange.values[0];
        openPOItemCodeIdx = openPOsHeaders.indexOf("Item Code");
        openPOItemQtyIdx = openPOsHeaders.indexOf("Outstanding Qty");
        
        openPOItemCodeColumn = `${colIdxToLetter(openPOItemCodeIdx)}:${colIdxToLetter(openPOItemCodeIdx)}`;
        openPOItemQtyColumn = `${colIdxToLetter(openPOItemQtyIdx)}:${colIdxToLetter(openPOItemQtyIdx)}`;
       
        //Inventory Report Fluid Placement
        inventoryHeaders = inventoryUsedRange.values[0];

        invItemCodeIdx = inventoryHeaders.indexOf("Item Code");
        invItemQtyIdx = inventoryHeaders.indexOf("Inventory Qty");
        
        invRepItemCodeColumn = `${colIdxToLetter(invItemCodeIdx)}:${colIdxToLetter(invItemCodeIdx)}`;
        invRepItemQtyColumn = `${colIdxToLetter(invItemQtyIdx)}:${colIdxToLetter(invItemQtyIdx)}`;

        //Quanity and Item Code from Dynamic, Inventory Report, and Open PO's sheets
        dynamic = context.workbook.worksheets.getItem("Dynamic");
        dynamicICR = dynamic.getRange(dynItemCodeColumn).getUsedRange().load("values");
        dynamicQR = dynamic.getRange(dynItemQtyColumn).getUsedRange().load("values"); 
        dynamicWork = dynamic.getRange(dynWorkColumn).getUsedRange().load("values");
        dyanamicDate = dynamic.getRange(dynDateColumn).getUsedRange().load("values"); 
        await context.sync();

        inventoryICR = inventoryReportWorksheet.getRange(invRepItemCodeColumn).getUsedRange().load("values"); 
        inventoryQR = inventoryReportWorksheet.getRange(invRepItemQtyColumn).getUsedRange().load("values"); 

        openPOs = context.workbook.worksheets.getItem("Open PO's");
        openPOsICR = openPOs.getRange(openPOItemCodeColumn).getUsedRange().load("values"); 
        openPOsQR = openPOs.getRange(openPOItemQtyColumn).getUsedRange().load("values"); 
        await context.sync();

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
        console.log("Case numbers and required amounts imported successfully.");

        //Importing the Planned Start Date
        plannedStart = dynamicWorksheet.getRange(dynStartColumn).getUsedRange().load("values"); 
        const orderingWorksheet = context.workbook.worksheets.getItem("Ordering");
        orderingUsedRange = orderingWorksheet.getUsedRange().load("values");
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
        
        const startArray = [["Earliest Start Date"]];
        for (let i = 1; i < orderingValues.length; i++) {
            const itemCode = String(orderingValues[i][0]).trim();
            const start = startMap.get(itemCode) || "No Start Date Established";
            startArray.push([start]);
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
        orderingSheet.getRange(`B1:B${demandOutput.length}`).values = demandOutput;

        const currentInventory = [["Current Inventory"]]; 
        for (const code of caseOrder.slice(1)) {
            const currentInvQty = inventoryMap.get(code) || 0;
            currentInventory.push([currentInvQty]);
              
        }
        
        const currentInventoryOutput = currentInventory.map(row => [row[0]]);
        orderingSheet.getRange(`C1:C${currentInventoryOutput.length}`).values = currentInventoryOutput;

        const onOrder = [["On Order"]]; 
        for (const code of caseOrder.slice(1)) {
            const onOrderQty = openPOsMap.get(code) || 0;
            onOrder.push([onOrderQty]);
              
        }
        
        const onOrderOutput = onOrder.map(row => [row[0]]);
        orderingSheet.getRange(`D1:D${onOrderOutput.length}`).values = onOrderOutput;
        
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
                orderOrMakeCategory.push(["Buy"]); 
            } else if (Number(requiredAmounts[i][0]) >= 300){
                orderOrMakeCategory.push(["Buy"]);  
            }
            else{
                orderOrMakeCategory.push(["Make"]);    
            }    
            await context.sync();    
        }
        orderingSheet.getRange(`F1:F${orderOrMakeCategory.length}`).values = orderOrMakeCategory;

        // Table Formatting
        orderingWorksheet.getRange("G:G").numberFormat = [['m/d/yyyy h:mm']];
        orderingWorksheet.getRange("A:G").format.autofitColumns();
        orderingWorksheet.getRange("A:G").format.horizontalAlignment = "Center";
        orderingWorksheet.getRange("A:G").format.verticalAlignment = "Center";
        orderingWorksheet.getRange("D:D").numberFormat = [['General']];
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

async function dateFilter(range) {
    await Excel.run(async (context) => {
        document.getElementById("myDropdown").classList.toggle("show");
        const todayDate = new Date();
        const todaysDate = new Date(todayDate);
        const todaysDateOut = todaysDate.toLocaleDateString('en-US');

        //1 Week Out from Current
        const currentDate1 = new Date();
        const nextWeekDate = currentDate1.setDate(currentDate1.getDate() + 7);
        const oneWeekOut = new Date(nextWeekDate);
        const aWeekOut = oneWeekOut.toLocaleDateString('en-US'); 
        const endDate1 = parseUSDate(aWeekOut);

        //2 Weeks Out from Current
        const currentDate2 = new Date();
        const twoWeeksDate = currentDate2.setDate(currentDate2.getDate() + 14);
        const twoWeeksOut = new Date(twoWeeksDate);
        const weekOut2 = twoWeeksOut.toLocaleDateString('en-US'); 
        const endDate2 = parseUSDate(weekOut2);

        //3 Weeks Out from Current
        const currentDate3 = new Date();
        const threeWeeksDate = currentDate3.setDate(currentDate3.getDate() + 21);
        const threeWeeksOut = new Date(threeWeeksDate);
        const weekOut3 = threeWeeksOut.toLocaleDateString('en-US'); 
        const endDate3 = parseUSDate(weekOut3);

        //4 Weeks Out from Current
        const currentDate4 = new Date();
        const fourWeeksDate = currentDate4.setDate(currentDate4.getDate() + 28);
        const fourWeeksOut = new Date(fourWeeksDate);
        const weekOut4 = fourWeeksOut.toLocaleDateString('en-US'); 
        const endDate4 = parseUSDate(weekOut4);

        //5 Weeks Out from Current
        const currentDate5 = new Date();
        const fiveWeeksDate = currentDate5.setDate(currentDate5.getDate() + 35);
        const fiveWeeksOut = new Date(fiveWeeksDate);
        const weekOut5 = fiveWeeksOut.toLocaleDateString('en-US'); 
        const endDate5 = parseUSDate(weekOut5);

        //6 Weeks Out from Current
        const currentDate6 = new Date();
        const sixWeeksDate = currentDate6.setDate(currentDate6.getDate() + 42);
        const sixWeeksOut = new Date(sixWeeksDate);
        const weekOut6 = sixWeeksOut.toLocaleDateString('en-US'); 
        const endDate6 = parseUSDate(weekOut6);

        //7 Weeks Out From Current
        const currentDate7 = new Date();
        const sevenWeeksDate = currentDate7.setDate(currentDate7.getDate() + 49);
        const sevenWeeksOut = new Date(sevenWeeksDate);
        const weekOut7 = sevenWeeksOut.toLocaleDateString('en-US'); 
        const endDate7 = parseUSDate(weekOut7);

        //8 Weeks Out from Current
        const currentDate8 = new Date();
        const eightWeeksDate = currentDate8.setDate(currentDate8.getDate() + 56);
        const eightWeeksOut = new Date(eightWeeksDate);
        const weekOut8 = eightWeeksOut.toLocaleDateString('en-US'); 
        const endDate8 = parseUSDate(weekOut8);
        
        await context.sync();
        switch(range){
            case 'one-week':
                    for (let i = 1; i < dynamicICR.values.length; i++) { 
                    const itemCode = String(dynamicICR.values[i][0]).trim();
                    const start = plannedStart.values[i] ? String(plannedStart.values[i][0]).trim() : "";
                    if (itemCode) {
                        dateMap.set(itemCode, start);
                    }
                }
                for (let i = 1; i < orderingValues.length; i++) {
                    const itemCode = String(orderingValues[i][0]).trim();
                    const start = startMap.get(itemCode) || "No Start Date Established";
                    startArray.push([start]);
                }

                break;
            case 'two-weeks':
            
                break;
            case 'three-weeks':
            
                break;
            case 'four-weeks':
            
                break;
            case 'five-weeks':
            
                break;
            case 'six-weeks':
            
                break;
            case 'seven-weeks':
            
                break;
            case 'eight-weeks':
            
                break;
        }
      await context.sync();      
    });    
}



window.onclick = function(event) {
    if (!event.target.matches('.dropbtn')) {
        var dropdowns = document.getElementsByClassName("dropdown-content");
        for (var i = 0; i < dropdowns.length; i++) {
            var openDropdown = dropdowns[i];
            if (openDropdown.classList.contains('show')) {
                openDropdown.classList.remove('show');
            } 
        }
    }
}

function parseUSDate(str) {
        const [month, day, year] = str.split('/').map(Number);
        return new Date(year, month - 1, day);
    }