/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
let item;
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    item = Office.context.mailbox.item;
    Office.actions.associate("onMessageSendHandler", onItemSendHandler);
    Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);
  }
});

function onItemSendHandler(event) {
  getAllRecipients()
    .then((uniqueDomains) => {
      // Display the email domains of the recipients as a comma-separated string.
      let domainsString = Array.from(uniqueDomains).join("\n");
      
      event.completed({
        allowEvent: false,
        cancelLabel: "Don't Send",
        commandId: "msgComposeOpenPaneButton",
        contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
        errorMessage: domainsString,
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
      });
    })
    .catch((error) => {
      console.error(error);
    });
}

async function getAllRecipients() {
  return new Promise((resolve, reject) => {
    let toRecipients, ccRecipients, bccRecipients;
    let uniqueDomains = new Set();

    // Verify if the mail item is an appointment or message.
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      toRecipients = item.requiredAttendees;
      ccRecipients = item.optionalAttendees;
    } else {
      toRecipients = item.to;
      ccRecipients = item.cc;
      bccRecipients = item.bcc;
    }

    // Function to add domains to the uniqueDomains set
    function addDomains(recipients) {
      recipients.forEach((recipient) => {
        let emailAddress = recipient.emailAddress;
        let domain = emailAddress.substring(emailAddress.indexOf('@') + 1);
        uniqueDomains.add(domain);
      });
    }

    // Get the recipients from the To or Required field of the item being composed.
    toRecipients.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        reject(asyncResult.error.message);
        return;
      }
      addDomains(asyncResult.value);

      // Get the recipients from the Cc or Optional field of the item being composed.
      ccRecipients.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(asyncResult.error.message);
          return;
        }
        addDomains(asyncResult.value);

        // Get the recipients from the Bcc field of the message being composed, if applicable.
        if (bccRecipients.length > 0) {
          bccRecipients.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(asyncResult.error.message);
              return;
            }
            addDomains(asyncResult.value);

            // Resolve with uniqueDomains once all recipients are processed.
            resolve(uniqueDomains);
          });
        } else {
          // Resolve with uniqueDomains if there are no Bcc recipients.
          resolve(uniqueDomains);
        }
      });
    });
  });
}
