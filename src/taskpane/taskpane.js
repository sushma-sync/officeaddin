/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // eslint-disable-next-line no-undef
    console.log("Outlook Add-in is Ready!");

    // Automatically insert signature when composing an email
    if (Office.context.mailbox.item) {
      insertSignature();
    }
  }
});

function insertSignature() {
  const signature = `
      <br/><br/>
      <strong>John Doe</strong><br/>
      <em>Product Manager | MyCompany</em><br/>
      <a href="mailto:johndoe@example.com">johndoe@example.com</a><br/>
      <a href="https://www.mycompany.com">www.mycompany.com</a>
  `;

  Office.context.mailbox.item.body.setSelectedDataAsync(
    signature,
    {
      coercionType: "html",
    },
    function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // eslint-disable-next-line no-undef
        console.error("Error inserting signature:", result.error.message);
      }
    }
  );
}
// export async function run() {
//   /**
//    * Insert your Outlook code here
//    */

//   const item = Office.context.mailbox.item;
//   let insertAt = document.getElementById("item-subject");
//   let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
//   insertAt.appendChild(label);
//   insertAt.appendChild(document.createElement("br"));
//   insertAt.appendChild(document.createTextNode(item.subject));
//   insertAt.appendChild(document.createElement("br"));
// }
