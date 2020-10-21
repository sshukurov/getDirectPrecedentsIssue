/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
};

async function run() {
  try {
    await Excel.run(async (context) => {
      // Precedents are cells referenced by the formula in a cell.
      let range = context.workbook.getActiveCell();
      let directPrecedents = range.getDirectPrecedents();
      range.load("address");
      directPrecedents.areas.load("address");
      await context.sync();
    
      console.log(`Direct precedent cells of ${range.address}:`);
    
      // Use the direct precedents API to loop through precedents of the active cell. 
      for (var i = 0; i < directPrecedents.areas.items.length; i++) {
        // Highlight and console the address of each precedent cell.
        directPrecedents.areas.items[i].format.fill.color = "Yellow";
        console.log(`  ${directPrecedents.areas.items[i].address}`);
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
