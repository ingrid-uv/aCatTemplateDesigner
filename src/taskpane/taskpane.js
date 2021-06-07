/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
//! import "../../assets/logo.png";

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    
    /*document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("insert-table").onclick = insertTable;*/
   
    document.getElementById("button").onclick = button;
    document.getElementById("hide/show").onclick = myFunction;
    document.getElementById("create-condition").onclick = addCondition;
    


  }
});

function button() {
  if(document.getElementById("button").innerText=="Block Beginn"){

      blockBeginn()
      document.getElementById("button").innerText="Block Ende"
      document.getElementById("button").className="btn btn-danger btn-sm form-control"


  }
    else if(document.getElementById("button").innerText=="Block Ende"){

      blockEnde()
      document.getElementById("button").innerText="Block Beginn"
      document.getElementById("button").className="btn btn-primary btn-sm form-control"
 


  }

}

function blockBeginn() {
  Word.run(function (context) {

      // TODO1: Queue commands to insert a paragraph into the document.
      var docBody = context.document.getSelection();
      docBody.insertHtml("${B:0} ", "End");
      const ctrl = docBody.insertContentControl()
      ctrl.title = "Select";
      ctrl.tag = "Select";
      ctrl.appearance = "BoundingBox";
      ctrl.color = "#589CFB";
      ctrl.parentBody.select("End")

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function blockEnde() {
  Word.run(function (context) {

      var docBody = context.document.getSelection();
      docBody.insertHtml("${B:1} ", "End");
      const ctrl = docBody.insertContentControl()
      ctrl.title = "Select";
      ctrl.tag = "Select";
      ctrl.appearance = "BoundingBox";
      ctrl.color = "#589CFB";
      ctrl.parentBody.select("End")


      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}


function myFunction() {
  var x = document.getElementById("myDIV");
  if (x.style.display === "none") {
    x.style.display = "block";
  } else {
    x.style.display = "none";
  }
}

function addCondition(){
      var action = document.getElementById("action");
      var condition = document.getElementById("condition");

      var checked = document.getElementById("check").checked;
      var feld1Input = document.getElementById("f1").value;
      var feld2Input = checked?document.getElementById("f2").value:"" ;

      var action = action.options[action.selectedIndex].value;
      var actionResult =  ":" + action;
      var conditionResult =  condition.options[condition.selectedIndex].text;
      var x = document.getElementById("alert")
      if(feld1Input.length>0 && action!="Aktion" && conditionResult!="Operator" ){
      Word.run(function (context) {

        // TODO1: Queue commands to insert a paragraph into the document.
        var docBody = context.document.getSelection();
        docBody.insertHtml("${C:"+feld1Input+":"+conditionResult+feld2Input+actionResult+"}" , "Start");
        const ctrl = docBody.insertContentControl()
        ctrl.title = "Select";
        ctrl.tag = "Select";
        ctrl.appearance = "BoundingBox";
        ctrl.color = "#589CFB";
        ctrl.parentBody.select("End")


        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    x.style.display = "none"
  }else{
    x.style.display = "block"
  }

}








