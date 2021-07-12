/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */

let RESPONSEMAP = {};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      console.log("The add-in uses Word.js APIs that are not available in your version of Office.");
    }
    getResponseMap();
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function createAppButtons() {
  var e = document.getElementById("app-buttons");
  for (const response in RESPONSEMAP) {
    let btn_key = response;
    let btn_label = RESPONSEMAP[response].label;
    let btn = document.createElement("button");
    btn.innerHTML = btn_label;
    btn.classList.add("objection_button", "ms-Button", "ms-sm12");
    btn.style.marginBottom = "12px";
    btn.dataset.key = btn_key;
    btn.onclick = insertResponse;
    e.appendChild(btn);
  }
}

function insertResponse(evt) {
  Word.run(function (context) {
    var responseKey = evt.target.dataset.key;
    var text = RESPONSEMAP[responseKey].content + "\n";
    var range = context.document.getSelection();
    var inserted_range = range.insertText(text, "Start");
    inserted_range.select("End");
    return context.sync().then(function () {
      //TODO: Any action you want to take on success
      //console.log("Text added");
    });
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

/**
 * Load possible responses from server.
 *
 * @returns Object containing map of possible discovery objection responses.
 */
function getResponseMap() {
  var xhr = new XMLHttpRequest();
  xhr.open("GET", "https://restutil.jdbot.us/objection_responses", true);
  xhr.setRequestHeader("Authentication", "Basic " + btoa("tdaley:X"));
  xhr.setRequestHeader("Content-Type", "application/json");
  xhr.withCredentials = true;
  xhr.onload = function (e) {
    var response_map = JSON.parse(this.response);
    var ordered_map = Object.keys(response_map)
      .sort()
      .reduce((obj, key) => {
        obj[key] = response_map[key];
        return obj;
      }, {});
    RESPONSEMAP = ordered_map;
    createAppButtons();
  };
  xhr.send();
}
