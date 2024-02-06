/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;
  
  document.getElementById("protect").onclick = protect;
    document.getElementById("unprotect").onclick = unprotect;
    document.getElementById("setup").onclick = setup;
    //loadFileName();
    Office.context.document.getFilePropertiesAsync(null, (res) => {
      if (res && res.value && res.value.url) {
        document.getElementById("flph").value = res.value.url;

        var file_name = " "
        
        file_name= res.value.url.substr(res.value.url.lastIndexOf('\\') + 1)

        if(file_name.includes('/'))

           file_name = res.value.url.substr(res.value.url.lastIndexOf('/') + 1)


       document.getElementById("flnm").value =  file_name
      }
      sheetPropertiesChanged();
      //enableNotification();


    });
  }
});
async function sheetPropertiesChanged() {    
  var rangeAddress;

  await Excel.run(async context => {
    let sheet1 = context.workbook.worksheets.getItem("Sheet1");
      // Read the range address
      let cellC2Value1 = sheet1.getRange("C3").load("values");
      await context.sync();
      console.log(cellC2Value1.values[0][0].toString());
      document.getElementById("eucid").value=cellC2Value1.values[0][0].toString()
  }); 
}


Office.initialize = function () {
  // Your code here , run at the start up\\


}

/** Protecting the work sheets*/
export async function protect() {
  try {
 
    let password = "citi123";
    await Excel.run(async context => {

      await createPassword(context);

      var unprotectSheets = context.workbook.worksheets;
      unprotectSheets.load("items");
      await context.sync();

      for (var i = 0; i < unprotectSheets.items.length; i++) {
        unprotectSheets.load("protection/protected");
        await context.sync();
        unprotectSheets.items[i].protection.protect(null, password);
      }
    });
  } catch (error) {
    console.error(error);
  }
}


async function createPassword(context) {
  let sheet = context.workbook.worksheets.getItem("Sheet1");
  const randomNumber = "100" + Math.floor(1000 + Math.random() * 9000).toString();

     // Create the headers and format them to stand out.
     let headers = [
      ["Name", "value"]
    ];
    let headerRange = sheet.getRange("B2:C2");
    headerRange.values = headers;
    headerRange.format.fill.color = "#4472C4";
    headerRange.format.font.color = "white";
    headerRange.format.font.bold = true;


  let data = [
      ["EUCId", randomNumber], 
      ["password", "citi123"]
  ];

  let range = sheet.getRange("B3:C4");
  
  range.values = data;
  range.format.autofitColumns();
  range.format.font.bold= true;  
  document.getElementById("eucid").value=randomNumber.toString()
  await context.sync();

};

async function getPassword(context) {
  let sheet = context.workbook.worksheets.getItem("Sheet1");
  let cellC2Value = sheet.getRange("C4").load("values");
  await context.sync();
  return cellC2Value.values[0][0].toString();
};



/** unProtecting the work sheets*/
export async function unprotect() {
  try {
   
    await Excel.run(async context => {

      let password = await getPassword(context).catch(error => console.error(error));

      var protectSheets = context.workbook.worksheets;
      protectSheets.load("items");
      await context.sync();

      for (var i = 0; i < protectSheets.items.length; i++) {
        protectSheets.load("protection/protected");
        await context.sync();
        protectSheets.items[i].protection.unprotect(password);
      }
    });
  } catch (error) {
    console.error(error);
  }
}



/** calling the api ans adding the data to the work sheet */
export async function setup() {
  await Excel.run(async context => {

    //let user1 = await getUserName(context).catch(error => console.error(error));
    context.workbook.worksheets.getItemOrNullObject("Authorization Cover").delete();
    const sheet = context.workbook.worksheets.add("Authorization Cover");

    // Hide gridlines
    //sheet.getUsedRange().format.fill.clear();
    // Hide column headings
   
    const expensesTable = sheet.tables.add("A1:B1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    //const response = await fetch();
   //const myJson = await response.json('"{\"cards\":[{\"Property\":\"EUCID\",\"value\":\"123456\"},{\"Property\":\"ECTID\",\"value\":\"123456\"},{\"Property\":\"EUCName\",\"value\":\"myEUCNAme\"},{\"Property\":\"EUCVersion\",\"value\":\"56\"},{\"Property\":\"EUCMessage\",\"value\":\"Compilantlastperformed2050\"}]}"');

   const jsonString = '{"success":true,"cards":[{"Property":"EUCID","value":"123456"},{"Property":"ECTID","value":"123456"},{"Property":"EUCName","value":"My EUC Name"},{"Property":"EUCVersion","value":"56"},{"Property":"EUCMessage","value":"Compilant last performed 2050"}]}';
   const result = JSON.parse(jsonString);
  
    expensesTable.getHeaderRowRange().values = [["Property","value"]];
    console.log(result);
    var transactions = result["cards"];
    console.log(transactions);
    var newData = transactions.map(item => 
        [item.Property, item.value]);

     expensesTable.rows.add(null, newData);

     sheet.getUsedRange().format.autofitColumns();
     sheet.getUsedRange().format.autofitRows();

     sheet.activate()

   // sheet.visibility= Excel.SheetVisibility.hidden;
    await context.sync();
  });
}

 async function getUserName(context) {
  try {
      let tokenData = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false, forMSGraphAccess: true });
      var parts = tokenData.split(".");
      var token = JSON.parse(atob(parts[1]));
      return token.preferred_username;
  }
  catch (exception) {
    console.log(exception.message);
  }
}