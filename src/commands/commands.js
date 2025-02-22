/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      console.log("Outlook Add-in is Ready!");
  }
});

function onMessageCompose(event) {
  const signature = `
      <br/><br/>
      <strong>John Doe</strong><br/>
      <em>Product Manager | MyCompany</em><br/>
      <a href="mailto:johndoe@example.com">johndoe@example.com</a><br/>
      <a href="https://www.mycompany.com">www.mycompany.com</a>
  `;

  Office.context.mailbox.item.body.setAsync(signature, { coercionType: "html" }, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Error inserting signature:", result.error.message);
      }
      event.completed(); // Mark event as completed
  });
}
Office.actions.associate("onMessageCompose", onMessageCompose);
/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true,
//   };

//   // Show a notification message.
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

//   // Be sure to indicate when the add-in command function is complete.
//   event.completed();
// }

// Register the function with Office.
// Office.actions.associate("action", action);
