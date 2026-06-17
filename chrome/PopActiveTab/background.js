async function resolveActiveTab() {
  let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) {
    [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  }
  return tab;
}

async function popActiveTabToNewWindow() {
  const tab = await resolveActiveTab();

  if (!tab?.id) {
    return;
  }

  const tabsInWindow = await chrome.tabs.query({ windowId: tab.windowId });
  if (tabsInWindow.length < 2) {
    return;
  }

  await chrome.windows.create({
    tabId: tab.id,
    focused: true,
    type: "normal",
  });
}

chrome.commands.onCommand.addListener((command) => {
  if (command === "pop-active-tab") {
    popActiveTabToNewWindow();
  }
});
