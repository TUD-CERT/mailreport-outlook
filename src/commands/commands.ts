/* global Office, window */

// Must be run each time a new page is loaded.
Office.onReady();

function showOptions(event: Office.AddinCommands.Event) {
  //const url = new URI("options.html").absoluteTo(window.location).toString();
  const url = new URL("options.html", window.location).toString();
  const dialogOptions = { width: 20, height: 40, displayInIframe: true };
  Office.context.ui.displayDialogAsync(url, dialogOptions, () => {});

  // Be sure to indicate when the add-in command function is complete.
  //event.completed();
}

// Register the function with Office.
Office.actions.associate("showOptions", showOptions);
