// Enhanced Background Script for ETA Invoice Exporter
class ETAEnhancedBackground {
  constructor() {
    this.activeConnections = new Map();
    this.progressData = new Map();
    this.init();
  }
  
  init() {
    console.log('ETA Enhanced Background: Service worker initialized');
    
    // Handle extension installation/update
    chrome.runtime.onInstalled.addListener((details) => {
      this.handleInstallation(details);
    });
    
    // Handle tab updates
    chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
      this.handleTabUpdate(tabId, changeInfo, tab);
    });
    
    // Handle messages from content scripts and popup
    chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
      this.handleMessage(message, sender, sendResponse);
      return true; // Keep message channel open for async responses
    });
    
    // Handle long-lived connections
    chrome.runtime.onConnect.addListener((port) => {
      this.handleConnection(port);
    });
    
    // Handle extension startup
    chrome.runtime.onStartup.addListener(() => {
      console.log('ETA Enhanced Background: Extension startup');
    });
  }
  
  handleInstallation(details) {
    console.log('ETA Enhanced Background: Installation details:', details);
    
    if (details.reason === 'install') {
      console.log('ETA Enhanced Background: First time installation');
      this.showWelcomeNotification();
    } else if (details.reason === 'update') {
      console.log('ETA Enhanced Background: Extension updated');
      this.handleUpdate(details.previousVersion);
    }
  }
  
  showWelcomeNotification() {
    // Only show if notifications are supported and enabled
    if (chrome.notifications) {
      chrome.notifications.create({
        type: 'basic',
        iconUrl: 'icons/icon48.png',
        title: 'مصدر فواتير الضرائب المصرية',
        message: 'تم تثبيت الإضافة بنجاح! يمكنك الآن تصدير جميع الفواتير تلقائياً.',
        priority: 1
      });
    }
  }
  
  handleUpdate(previousVersion) {
    console.log(`ETA Enhanced Background: Updated from version ${previousVersion}`);
    
    // Clear old cached data on major updates
    if (this.isMajorUpdate(previousVersion)) {
      chrome.storage.local.clear();
      console.log('ETA Enhanced Background: Cleared cache after major update');
    }
  }
  
  isMajorUpdate(previousVersion) {
    const current = chrome.runtime.getManifest().version;
    const prev = previousVersion || '0.0.0';
    
    const [currentMajor] = current.split('.');
    const [prevMajor] = prev.split('.');
    
    return parseInt(currentMajor) > parseInt(prevMajor);
  }
  
  handleTabUpdate(tabId, changeInfo, tab) {
    // Check if tab is ETA invoicing site
    if (changeInfo.status === 'complete' && 
        tab.url && 
        tab.url.includes('invoicing.eta.gov.eg')) {
      
      console.log('ETA Enhanced Background: ETA site loaded in tab', tabId);
      
      // Inject content script if not already present
      this.ensureContentScriptInjected(tabId);
      
      // Update badge to indicate site is supported
      this.updateBadge(tabId, 'ready');
    }
  }
  
  async ensureContentScriptInjected(tabId) {
    try {
      // Test if content script is already loaded
      const response = await chrome.tabs.sendMessage(tabId, { action: 'ping' });
      if (response && response.success) {
        console.log('ETA Enhanced Background: Content script already loaded');
        return;
      }
    } catch (error) {
      // Content script not loaded, inject it
      console.log('ETA Enhanced Background: Injecting content script');
      
      try {
        await chrome.scripting.executeScript({
          target: { tabId: tabId },
          files: ['content.js']
        });
        console.log('ETA Enhanced Background: Content script injected successfully');
      } catch (injectError) {
        console.error('ETA Enhanced Background: Failed to inject content script:', injectError);
      }
    }
  }
  
  updateBadge(tabId, status) {
    const badgeConfig = {
      ready: { text: '✓', color: '#4CAF50' },
      processing: { text: '↻', color: '#FF9800' },
      error: { text: '!', color: '#F44336' },
      complete: { text: '✓', color: '#2196F3' }
    };
    
    const config = badgeConfig[status] || { text: '', color: '#666666' };
    
    chrome.action.setBadgeText({ text: config.text, tabId });
    chrome.action.setBadgeBackgroundColor({ color: config.color, tabId });
  }
  
  handleMessage(message, sender, sendResponse) {
    console.log('ETA Enhanced Background: Received message:', message.action);
    
    switch (message.action) {
      case 'progressUpdate':
        this.handleProgressUpdate(message, sender);
        break;
        
      case 'exportStarted':
        this.handleExportStarted(message, sender);
        break;
        
      case 'exportCompleted':
        this.handleExportCompleted(message, sender);
        break;
        
      case 'exportError':
        this.handleExportError(message, sender);
        break;
        
      case 'getProgress':
        this.handleGetProgress(message, sender, sendResponse);
        break;
        
      case 'clearProgress':
        this.handleClearProgress(message, sender);
        break;
        
      default:
        console.log('ETA Enhanced Background: Unknown message action:', message.action);
    }
  }
  
  handleProgressUpdate(message, sender) {
    const tabId = sender.tab?.id;
    if (!tabId) return;
    
    // Store progress data
    this.progressData.set(tabId, {
      ...message.progress,
      timestamp: Date.now(),
      tabId: tabId
    });
    
    // Update badge to show processing
    this.updateBadge(tabId, 'processing');
    
    // Forward progress to popup if connected
    this.forwardToPopup(tabId, message);
    
    console.log('ETA Enhanced Background: Progress updated:', message.progress);
  }
  
  handleExportStarted(message, sender) {
    const tabId = sender.tab?.id;
    if (!tabId) return;
    
    console.log('ETA Enhanced Background: Export started for tab:', tabId);
    this.updateBadge(tabId, 'processing');
    
    // Store start time
    this.progressData.set(tabId, {
      startTime: Date.now(),
      status: 'processing',
      tabId: tabId
    });
  }
  
  handleExportCompleted(message, sender) {
    const tabId = sender.tab?.id;
    if (!tabId) return;
    
    console.log('ETA Enhanced Background: Export completed for tab:', tabId);
    this.updateBadge(tabId, 'complete');
    
    // Update progress data
    const existingProgress = this.progressData.get(tabId) || {};
    this.progressData.set(tabId, {
      ...existingProgress,
      status: 'completed',
      completedTime: Date.now(),
      totalRecords: message.totalRecords || 0
    });
    
    // Show completion notification
    this.showCompletionNotification(message.totalRecords || 0);
    
    // Auto-clear badge after delay
    setTimeout(() => {
      this.updateBadge(tabId, 'ready');
    }, 5000);
  }
  
  handleExportError(message, sender) {
    const tabId = sender.tab?.id;
    if (!tabId) return;
    
    console.error('ETA Enhanced Background: Export error for tab:', tabId, message.error);
    this.updateBadge(tabId, 'error');
    
    // Update progress data
    const existingProgress = this.progressData.get(tabId) || {};
    this.progressData.set(tabId, {
      ...existingProgress,
      status: 'error',
      error: message.error,
      errorTime: Date.now()
    });
    
    // Show error notification
    this.showErrorNotification(message.error);
    
    // Auto-clear badge after delay
    setTimeout(() => {
      this.updateBadge(tabId, 'ready');
    }, 10000);
  }
  
  handleGetProgress(message, sender, sendResponse) {
    const tabId = message.tabId || sender.tab?.id;
    if (!tabId) {
      sendResponse({ success: false, error: 'No tab ID provided' });
      return;
    }
    
    const progress = this.progressData.get(tabId);
    sendResponse({
      success: true,
      progress: progress || null
    });
  }
  
  handleClearProgress(message, sender) {
    const tabId = message.tabId || sender.tab?.id;
    if (!tabId) return;
    
    this.progressData.delete(tabId);
    this.updateBadge(tabId, 'ready');
    console.log('ETA Enhanced Background: Progress cleared for tab:', tabId);
  }
  
  showCompletionNotification(totalRecords) {
    if (chrome.notifications) {
      chrome.notifications.create({
        type: 'basic',
        iconUrl: 'icons/icon48.png',
        title: 'تم تصدير الفواتير بنجاح',
        message: `تم تصدير ${totalRecords} فاتورة بنجاح إلى ملف Excel.`,
        priority: 1
      });
    }
  }
  
  showErrorNotification(error) {
    if (chrome.notifications) {
      chrome.notifications.create({
        type: 'basic',
        iconUrl: 'icons/icon48.png',
        title: 'خطأ في تصدير الفواتير',
        message: `حدث خطأ أثناء التصدير: ${error}`,
        priority: 2
      });
    }
  }
  
  forwardToPopup(tabId, message) {
    // Try to forward message to popup if it's open
    const connection = this.activeConnections.get(tabId);
    if (connection) {
      try {
        connection.postMessage(message);
      } catch (error) {
        console.warn('ETA Enhanced Background: Failed to forward to popup:', error);
        this.activeConnections.delete(tabId);
      }
    }
  }
  
  handleConnection(port) {
    console.log('ETA Enhanced Background: New connection:', port.name);
    
    if (port.name === 'eta-exporter-popup') {
      // Store connection for progress forwarding
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        if (tabs[0]) {
          this.activeConnections.set(tabs[0].id, port);
          
          port.onDisconnect.addListener(() => {
            console.log('ETA Enhanced Background: Popup disconnected');
            this.activeConnections.delete(tabs[0].id);
          });
          
          // Send current progress if available
          const currentProgress = this.progressData.get(tabs[0].id);
          if (currentProgress) {
            port.postMessage({
              action: 'progressUpdate',
              progress: currentProgress
            });
          }
        }
      });
    }
  }
  
  // Cleanup old progress data periodically
  startProgressCleanup() {
    setInterval(() => {
      const now = Date.now();
      const maxAge = 30 * 60 * 1000; // 30 minutes
      
      for (const [tabId, progress] of this.progressData.entries()) {
        if (now - progress.timestamp > maxAge) {
          this.progressData.delete(tabId);
          console.log('ETA Enhanced Background: Cleaned up old progress for tab:', tabId);
        }
      }
    }, 5 * 60 * 1000); // Run every 5 minutes
  }
}

// Initialize enhanced background service
const etaEnhancedBackground = new ETAEnhancedBackground();
etaEnhancedBackground.startProgressCleanup();

console.log('ETA Enhanced Background: Service worker initialized');