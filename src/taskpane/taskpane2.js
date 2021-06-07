var to = false;

console.log("Hi! 2")


/**
 * Nimmt die json.Data als Baumobjekt und zeigt sie auf der Taskbereichsseite des Add-Ins an.
 * 
 * Diese Funktion besteht aus drei Hauptschritten:
 * 
 * 1. Das DOM des Taskbereichs wird aktualisiert, indem die Tags "li" und "ul" eingefügt werden, um den Baum darzustellen.
 * Es wurde mit drei verschachtelten Schleifen implementiert, da json.Data drei Ebenen hat.
 * Die erste Schleife durchläuft die oberste Ebene des Baums und fügt "fieldsLable" -IDs ein.
 * Die zweite Schleife durchläuft die Mitte des Baums und fügt "blockLabel" -IDs ein.
 * Die dritte Schleife durchläuft die letzte Ebene des Baums und fügt Beschriftungen nur mit der ID-Nummer ein.
 * 
 * 2. Ruft jsTree auf, um den Baum interaktiv zu gestalten.
 * 
 * 3. Aktiviert zwei CallBack-Funktionen: onChangedSearch und onChangedTree.
 * 
 * @see https://www.jstree.com/ 
 * @param {object} data - Das Objekt, das den Baum darstellt. 
 */
function aCat(data) {
  console.log("Data loaded!");

  console.log(data.fields);
  let dataFields = data.fields;
  let ul = document.querySelector("#jstree_demo_div > ul");
  console.log(ul);

  $.each(dataFields, function(i, item) {
    var li = document.createElement("li");
    li.setAttribute("id", "fieldsLabel" + i);
    li.setAttribute("class", "F");
    li.appendChild(document.createTextNode(item.label));

    li.setAttribute("data-jstree", '{ "icon" : "../../assets/folder.png" }');
    var ul_child = document.createElement("ul");

    $.each(item.blocks, function(b, block) {
      li_child = document.createElement("li");
      li_child.setAttribute("id", "blockLabel" + b);
      li_child.setAttribute("data-jstree", '{ "icon" : "../../assets/file.png" }');
      li_child.setAttribute("class", "B");
      li_child.appendChild(document.createTextNode(block.label));

      var ul_sub_child = document.createElement("ul");
      $.each(block.fields, function(f, label) {
        li_sub_child = document.createElement("li");
        li_sub_child.setAttribute("id", f);
        li_sub_child.setAttribute("data-jstree", '{ "icon" : "../../assets/check.png" }');
        li_sub_child.setAttribute("class", "L");
        li_sub_child.appendChild(document.createTextNode(label.label));
        ul_sub_child.appendChild(li_sub_child);
      });
      li_child.appendChild(ul_sub_child);
      ul_child.appendChild(li_child);
    });
    li.appendChild(ul_child);
    ul.appendChild(li);
  });

  $("#jstree_demo_div").jstree({
    "conditionalselect": conditionalSelectFunction,
    "plugins": ["search", "conditionalselect"]
  });

  function conditionalSelectFunction(node,event) {
    var id = node.id;
    console.log("Cond sel! id = " + id)

    if( id.search( "Label" ) == -1 ) {
      console.log( "True!")
      return true;
    } else {
      console.log( "False!")
      return false;
    }
  }

  
  $("#s").keyup(onChangedSearch);

  $("#jstree_demo_div").on("changed.jstree", onChangedTree);

  Word.run(function(context) {
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.select();

    return context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

/**
 * Dies ist die CallBack-Funktion, die von jsTree aufgerufen wird, wenn der Benutzer mit dem Baum interagiert.
 * @see https://www.jstree.com/api/#/?q=.jstree%20Event&f=changed.jstree
 * @param {object} e - Parameter, die Ereignisinformationen enthalten (hier nicht verwendet).
 * @param {object} data - Enthält Informationen zum angeklickten Knoten (Element des Baums).
 */
function onChangedTree(e, data) {
  var x = document.getElementById("myDIV");
  if (x.style.display === "none") {
    Word.run(function(context) {
      var docBody = context.document.getSelection();
      docBody.insertHtml("${F:" + data.selected + "}", Word.InsertLocation.end);
      const ctrl = docBody.insertContentControl();
      ctrl.title = "Select";
      ctrl.tag = "Select";
      ctrl.appearance = "BoundingBox";
      ctrl.color = "#589CFB";
      ctrl.parentBody.select("End");

      return context.sync();
    });
  } else {
    if (!$("#check").is(":checked")) {
      document.getElementById("f1").value = data.selected;
    }
    if ($("#check").is(":checked")) {
      document.getElementById("f2").value = "F" + data.selected;
    }
  }
}
/**
 * Dies ist die CallBack-Funktion, die aufgerufen wird, wenn der Benutzer das Suchfeld ändert.
 * Diese Funktion ruft dann jsTree auf, um den Wert im Suchfeld zu finden.
 * @see https://www.jstree.com/api/#/?f=search(
 */
function onChangedSearch() {
  if (to) {
    clearTimeout(to);
  }
  to = setTimeout(function() {
    var v = $("#s").val();
    $("#jstree_demo_div")
      .jstree(true)
      .search(v);
  }, 250);
}
