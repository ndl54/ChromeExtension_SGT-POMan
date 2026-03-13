function enableSidePanelOnActionClick() {
    chrome.sidePanel
        .setPanelBehavior({ openPanelOnActionClick: true })
        .catch((error) => console.error('Unable to enable side panel action:', error));
}

chrome.runtime.onInstalled.addListener(() => {
    enableSidePanelOnActionClick();
});

chrome.runtime.onStartup.addListener(() => {
    enableSidePanelOnActionClick();
});

enableSidePanelOnActionClick();
