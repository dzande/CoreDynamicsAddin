/* global console, document, Excel, Office */

// Run when Office is ready
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Show add-in UI if present
    const sideloadMsg = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    if (sideloadMsg) sideloadMsg.style.display = "none";
    if (appBody) appBody.style.display = "flex";

    // Attach event handlers for buttons in the taskpane
    const runBtn = document.getElementById("run");
    if (runBtn) runBtn.onclick = run;

    const dialogBtn = document.getElementById("openDialogBtn");
    if (dialogBtn) dialogBtn.onclick = openDialog;
  }
});

// ===========================
// Main Excel function
// ===========================
async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");

      // Update fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

// ===========================
// Open a full-screen dialog
// Ribbon button calls this
// ===========================
function openDialog() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/taskpane.html",
    {
      displayInIframe: false, // modeless dialog
      width: 100,             // full width
      height: 100,            // full height
      title: "Core Dynamics"  // top bar text
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Dialog failed:", asyncResult.error.message);
      } else {
        const dialog = asyncResult.value;

        // Handle close messages from the dialog
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          if (arg.message === "close") dialog.close();
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          console.log("Dialog closed.");
        });
      }
    }
  );
}

// ===========================
// Make functions global for ribbon buttons
// ===========================
window.run = run;
window.openDialog = openDialog;
