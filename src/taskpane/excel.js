/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('inp').onkeyup = run
    Excel.workbook.worksheets.onSelectionChanged.add(onCellChnge)
  }
});

export async function run(event) {
  console.log(event.target.value)
  
  try {
    
    await Excel.run(async (context) => {
      const ranges = context.workbook.getActiveCell();
      ranges.values = event.target.value;
     return context.sync()

    });
  } catch (error) {
    console.error(error);
  }
}

export function onCellChnge(){
  document.getElementById('inp').value = ''
}


