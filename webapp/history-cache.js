// Runs before Office.js, which replaces window.history.pushState and
// replaceState with null (Microsoft documents this in "Referencing the
// Office JavaScript API library", with this workaround). The two methods
// are kept here and put back by office-init.js, right after Office.js.
// A file of its own, not an inline script: the page works under a
// Content-Security-Policy without 'unsafe-inline'.
window.z2ui5xlHistory = {
  pushState: window.history.pushState,
  replaceState: window.history.replaceState,
};
