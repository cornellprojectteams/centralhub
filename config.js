// Single source of truth for the Space Status web app /exec URL.
// If the deployment URL ever changes, change it here only (used by app.js and team.html).
window.SPACE_STATUS_URL = 'https://script.google.com/macros/s/AKfycbwNbGjVcBrcsMZiOl2nXzpqZHz04nvKLm9D_aC0VJDz7Xxxf_4kLKlNSOHubPXj1X74/exec';

// Student Employee Log → Contact Info tab. The staff page reads name, NetID,
// email, “Can be a Resource for…”, and Roles (no phone / CU ID) via the Sheets
// JSONP visualization API.
window.PEOPLE_SHEET = {
  id: '1Kfffs0B0U3NZhS3YwFmRbXH3GP2pUZrCu_fylGm10CM',
  gid: '983316085',
};

// Same Student Employee Log workbook, FA25 schedule tab (weekly desk coverage).
window.SCHEDULE_SHEET = {
  id: '1Kfffs0B0U3NZhS3YwFmRbXH3GP2pUZrCu_fylGm10CM',
  gid: '1799869087',
};

// Tool & equipment inventory. The hub's inventory app (inventory.html) reads and
// writes the inventory Google Sheet through that sheet's web app
// (apps-script/08_inventory.gs).
//   apiUrl  the web app's /exec URL (Deploy > New deployment > Web app). Until it is
//           set, the hub shows the inventory as "not connected yet".
//   appUrl  the public address of inventory.html; QR labels and the poster use it.
// The Inventory menu in the sheet (Show links and setup info) prints this block filled in.
window.INVENTORY = {
  apiUrl: 'https://script.google.com/macros/s/AKfycbxrxidgCHzJL9_bgX1Z_inNzwFGoIhilh8xczHmY6JIDYZZdUhei9AziOqj0yPQO9oXMQ/exec',
  appUrl: 'https://cornellprojectteams.github.io/centralhub/inventory.html',
  sheetUrl: 'https://docs.google.com/spreadsheets/d/1Gjjqu3hX4Xcwm6w0Bp3maPg7hM0y91zAP7nnFXtwnIM/edit',
  dashboardGid: '1087637400',
  masterGid: '0',
  maintenanceGid: '2133394607',
  formUrl: 'https://forms.gle/Q2aKnKpLFVjx1My97',   // backup Google Form, in case the app is ever down
};
