window.onload = function Start() { 
   console.log('hello world'); 

   var app_1 = document.getElementById("app"); 
   app_1.innerHTML = '<b>hello world</b>'; 

   app_1.innerHTML = app_1.innerHTML + 
      '<br><input type="button" value="Add Data" onclick="loadWordData();" />'; 
} 

Office.initialize = function (reason) { 
} 

window.loadWordData = loadWordData; 

function loadWordData() { 

   Word.run(function (ctx) { 
       var myTable = new Office.TableData(); 
       myTable.headers = ["First Name", "Last Name", "Grade"]; 
       myTable.rows = [["Brittney", "Booker", "A"], ["Sanjit", "Pandit", "C"], ["Naomi", "Peacock", "B"]]; 
       Office.context.document.setSelectedDataAsync(myTable, { coercionType: Office.CoercionType.Table }); 

       return ctx.sync(); 
    }); 
} 