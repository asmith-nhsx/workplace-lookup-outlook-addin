/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    run();
  }
});

function toTitleCase(str) {
  return str.replace(/\w\S*/g, function (txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    }
  );
}

export async function run() {
  // Get a reference to the current message
  var template =
    "https://nhsxnhsuk.workplace.com/work/orgsearch?filters=%7B%22name%22%3A%7B%22operator%22%3A%22contains%22%2C%22values%22%3A[%7B%22value%22%3A%22<<NAME>>%22%7D]%7D%7D";
  var item = Office.context.mailbox.item;

  // Write message property value to the task pane
  document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
  document.getElementById("item-from").innerHTML = "<b>From:</b> <br/>" + item.from.displayName;
  let listTo = document.getElementById("list-to");
  let emails = [item.from];
  emails = emails.concat(item.to);
  emails = emails.concat(item.cc);
  emails = emails.sort(function(a,b) { return a.emailAddress.localeCompare(b.emailAddress)});
  let seen = [];
  for (let add of emails) {
    let pos = add.displayName.indexOf("(DEPARTMENT OF HEALTH AND SOCIAL CARE)");
    if (pos > -1) {
      let dNames = add.displayName.substring(0, pos).split(',');
      let toName = dNames[1] + dNames[0]
      toName = toTitleCase(toName);
      if (!seen.includes(toName)) {
        let url = template.replace("<<NAME>>", toName);
        let li = document.createElement("li");
        li.innerHTML = "<a href='" + url + "' target='_workplacePopup'>" + toName + "</a>";
        listTo.appendChild(li);
        seen.push(toName);
      }
    }
  }
}
