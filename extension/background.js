chrome.runtime.onInstalled.addListener(() => {
  // Ajouter une option dans le clic droit sur n'importe quel lien web
  chrome.contextMenus.create({
    id: "convertLink",
    title: "Convertir ce document (PDF <-> DOCX)",
    contexts: ["link"]
  });
});

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId === "convertLink" && info.linkUrl) {
    console.log("Fichier à télécharger et convertir :", info.linkUrl);
    
    // Télécharger le document dans les téléchargements temporaires
    chrome.downloads.download({ url: info.linkUrl }, (downloadId) => {
      if (chrome.runtime.lastError) {
        console.error(chrome.runtime.lastError);
      }
    });
  }
});

// Écouter la fin du téléchargement pour lancer la conversion native
chrome.downloads.onChanged.addListener((delta) => {
  if (delta.state && delta.state.current === "complete") {
    chrome.downloads.search({ id: delta.id }, (results) => {
      if (results && results.length > 0) {
        const filePath = results[0].filename;
        const lower = filePath.toLowerCase();
        if (lower.endsWith(".pdf") || lower.endsWith(".docx")) {
          // Envoyer au programme natif
          const port = chrome.runtime.connectNative("com.convertisseur.universel");
          port.postMessage({ filePath: filePath });
          port.onMessage.addListener((response) => {
            console.log("Reponse recue du convertisseur :", response);
          });
        }
      }
    });
  }
});
