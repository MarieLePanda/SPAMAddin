Office.initialize = function () {
};

// Function support to display a message
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function startAddin(event) {
    FowardNow();
    DeleteNow();
}

