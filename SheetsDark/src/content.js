const STORAGE_KEY = "enabled";
const ROOT_ATTRIBUTE = "data-sheets-dark";

function setDarkMode(enabled) {
  document.documentElement.setAttribute(ROOT_ATTRIBUTE, enabled ? "on" : "off");
}

function loadAndApplyState() {
  chrome.storage.sync.get({ [STORAGE_KEY]: true }, (result) => {
    setDarkMode(Boolean(result[STORAGE_KEY]));
  });
}

chrome.storage.onChanged.addListener((changes, areaName) => {
  if (areaName !== "sync" || !changes[STORAGE_KEY]) {
    return;
  }

  setDarkMode(Boolean(changes[STORAGE_KEY].newValue));
});

loadAndApplyState();
