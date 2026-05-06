const STORAGE_KEY = "enabled";
const checkbox = document.getElementById("enabled");

chrome.storage.sync.get({ [STORAGE_KEY]: true }, (result) => {
  checkbox.checked = Boolean(result[STORAGE_KEY]);
});

checkbox.addEventListener("change", () => {
  chrome.storage.sync.set({ [STORAGE_KEY]: checkbox.checked });
});
