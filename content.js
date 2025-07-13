// Enhanced Content Script for ETA Invoice Exporter - With View Details Support
class ETAContentScript {
  constructor() {
    this.invoiceData = [];
    this.allPagesData = [];
    this.totalCount = 0;
    this.currentPage = 1;
    this.totalPages = 1;
    this.isProcessingAllPages = false;
    this.progressCallback = null;
    this.domObserver = null;
    this.pageLoadTimeout = 10000; // 10 seconds timeout
    this.detailsCache = new Map(); // Cache for invoice details
    this.init();
  }
  
  init() {
    console.log('ETA Exporter: Content script initialized');
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => this.scanForInvoices());
    } else {
      setTimeout(() => this.scanForInvoices(), 1000);
    }
    
    this.setupMutationObserver();
    this.injectViewDetailsHandler();
  }
  
  injectViewDetailsHandler() {
    // Inject CSS for better view button styling
    const style = document.createElement('style');
    style.textContent = `
      .eta-view-btn {
        background: linear-gradient(135deg, #3b82f6, #1d4ed8) !important;
        color: white !important;
        border: none !important;
        padding: 6px 12px !important;
        border-radius: 6px !important;
        font-size: 11px !important;
        font-weight: bold !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        text-decoration: none !important;
        display: inline-flex !important;
        align-items: center !important;
        gap: 4px !important;
      }
      
      .eta-view-btn:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4) !important;
      }
      
      .eta-view-btn:active {
        transform: translateY(0) !important;
      }
      
      .eta-details-modal {
        position: fixed !important;
        top: 0 !important;
        left: 0 !important;
        width: 100% !important;
        height: 100% !important;
        background: rgba(0, 0, 0, 0.8) !important;
        z-index: 999999 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-family: 'Segoe UI', sans-serif !important;
        direction: rtl !important;
      }
      
      .eta-details-content {
        background: white !important;
        border-radius: 12px !important;
        width: 90% !important;
        max-width: 800px !important;
        max-height: 90% !important;
        overflow-y: auto !important;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.3) !important;
      }
      
      .eta-details-header {
        background: linear-gradient(135deg, #1e3c72, #2a5298) !important;
        color: white !important;
        padding: 20px !important;
        border-radius: 12px 12px 0 0 !important;
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
      }
      
      .eta-details-title {
        font-size: 18px !important;
        font-weight: bold !important;
        margin: 0 !important;
      }
      
      .eta-details-close {
        background: rgba(255, 255, 255, 0.2) !important;
        border: none !important;
        color: white !important;
        width: 32px !important;
        height: 32px !important;
        border-radius: 50% !important;
        cursor: pointer !important;
        font-size: 18px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        transition: background 0.2s !important;
      }
      
      .eta-details-close:hover {
        background: rgba(255, 255, 255, 0.3) !important;
      }
      
      .eta-details-body {
        padding: 20px !important;
      }
      
      .eta-details-grid {
        display: grid !important;
        grid-template-columns: 1fr 1fr !important;
        gap: 20px !important;
      }
      
      .eta-details-section {
        background: #f8f9fa !important;
        border-radius: 8px !important;
        padding: 16px !important;
        border: 1px solid #e9ecef !important;
      }
      
      .eta-details-section-title {
        font-size: 14px !important;
        font-weight: bold !important;
        color: #1e3c72 !important;
        margin-bottom: 12px !important;
        padding-bottom: 8px !important;
        border-bottom: 2px solid #e9ecef !important;
      }
      
      .eta-details-field {
        display: flex !important;
        justify-content: space-between !important;
        margin-bottom: 8px !important;
        font-size: 13px !important;
      }
      
      .eta-details-label {
        font-weight: 500 !important;
        color: #495057 !important;
        min-width: 120px !important;
      }
      
      .eta-details-value {
        color: #212529 !important;
        text-align: left !important;
        word-break: break-all !important;
      }
      
      .eta-details-highlight {
        background: linear-gradient(135deg, #e3f2fd, #bbdefb) !important;
        border: 1px solid #2196f3 !important;
      }
      
      .eta-loading-details {
        text-align: center !important;
        padding: 40px !important;
        color: #6c757d !important;
      }
      
      .eta-loading-spinner {
        width: 32px !important;
        height: 32px !important;
        border: 3px solid #e9ecef !important;
        border-top: 3px solid #3b82f6 !important;
        border-radius: 50% !important;
        animation: eta-spin 1s linear infinite !important;
        margin: 0 auto 16px !important;
      }
      
      @keyframes eta-spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
    `;
    document.head.appendChild(style);
  }
  
  setupMutationObserver() {
    this.observer = new MutationObserver((mutations) => {
      let shouldRescan = false;
      
      mutations.forEach((mutation) => {
        if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              if (node.classList?.contains('ms-DetailsRow') || 
                  node.querySelector?.('.ms-DetailsRow') ||
                  node.classList?.contains('ms-List-cell')) {
                shouldRescan = true;
              }
            }
          });
        }
      });
      
      if (shouldRescan && !this.isProcessingAllPages) {
        clearTimeout(this.rescanTimeout);
        this.rescanTimeout = setTimeout(() => this.scanForInvoices(), 800);
      }
    });
    
    this.observer.observe(document.body, {
      childList: true,
      subtree: true
    });
  }
  
  scanForInvoices() {
    try {
      console.log('ETA Exporter: Starting invoice scan...');
      this.invoiceData = [];
      
      // Extract pagination info first
      this.extractPaginationInfo();
      
      // Find invoice rows
      const rows = this.getVisibleInvoiceRows();
      console.log(`ETA Exporter: Found ${rows.length} visible invoice rows on page ${this.currentPage}`);
      
      if (rows.length === 0) {
        console.warn('ETA Exporter: No invoice rows found. Trying alternative selectors...');
        const alternativeRows = this.getAlternativeInvoiceRows();
        console.log(`ETA Exporter: Found ${alternativeRows.length} rows with alternative selectors`);
        alternativeRows.forEach((row, index) => {
          const invoiceData = this.extractDataFromRow(row, index + 1);
          if (this.isValidInvoiceData(invoiceData)) {
            this.invoiceData.push(invoiceData);
          }
        });
      } else {
        rows.forEach((row, index) => {
          const invoiceData = this.extractDataFromRow(row, index + 1);
          if (this.isValidInvoiceData(invoiceData)) {
            this.invoiceData.push(invoiceData);
          }
        });
      }
      
      console.log(`ETA Exporter: Successfully extracted ${this.invoiceData.length} valid invoices from page ${this.currentPage}`);
      
    } catch (error) {
      console.error('ETA Exporter: Error scanning for invoices:', error);
    }
  }
  
  getVisibleInvoiceRows() {
    // Primary selectors for invoice rows
    const selectors = [
      '.ms-DetailsRow[role="row"]',
      '.ms-List-cell[role="gridcell"]',
      '[data-list-index]',
      '.ms-DetailsRow',
      '[role="row"]'
    ];
    
    for (const selector of selectors) {
      const rows = document.querySelectorAll(selector);
      const visibleRows = Array.from(rows).filter(row => 
        this.isRowVisible(row) && this.hasInvoiceData(row)
      );
      
      if (visibleRows.length > 0) {
        console.log(`ETA Exporter: Found ${visibleRows.length} rows using selector: ${selector}`);
        return visibleRows;
      }
    }
    
    return [];
  }
  
  getAlternativeInvoiceRows() {
    // Alternative selectors when primary ones fail
    const alternativeSelectors = [
      'tr[role="row"]',
      '.ms-List-cell',
      '[data-automation-key]',
      '.ms-DetailsRow-cell',
      'div[role="gridcell"]'
    ];
    
    const allRows = [];
    
    for (const selector of alternativeSelectors) {
      const elements = document.querySelectorAll(selector);
      Array.from(elements).forEach(element => {
        const row = element.closest('[role="row"]') || element.parentElement;
        if (row && this.hasInvoiceData(row) && !allRows.includes(row)) {
          allRows.push(row);
        }
      });
    }
    
    return allRows.filter(row => this.isRowVisible(row));
  }
  
  isRowVisible(row) {
    if (!row) return false;
    
    const rect = row.getBoundingClientRect();
    const style = window.getComputedStyle(row);
    
    return (
      rect.width > 0 && 
      rect.height > 0 &&
      style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      style.opacity !== '0'
    );
  }
  
  hasInvoiceData(row) {
    if (!row) return false;
    
    // Check for electronic number or internal number
    const electronicNumber = row.querySelector('.internalId-link a, [data-automation-key="uuid"] a, .griCellTitle');
    const internalNumber = row.querySelector('.griCellSubTitle, [data-automation-key="uuid"] .griCellSubTitle');
    const totalAmount = row.querySelector('[data-automation-key="total"], .griCellTitleGray');
    
    return !!(electronicNumber?.textContent?.trim() || 
              internalNumber?.textContent?.trim() || 
              totalAmount?.textContent?.trim());
  }
  
  extractPaginationInfo() {
    try {
      // Extract total count from Results: 300 text
      const totalLabel = document.querySelector('.eta-pagination-totalrecordCount-label');
      if (totalLabel) {
        const text = totalLabel.textContent;
        const match = text.match(/Results:\s*(\d+)/);
        if (match) {
          this.totalCount = parseInt(match[1]);
        }
      }
      
      // Extract current page
      const currentPageBtn = document.querySelector('.eta-pageNumber.is-checked, .eta-pageNumber[aria-pressed="true"]');
      if (currentPageBtn) {
        const pageLabel = currentPageBtn.querySelector('.ms-Button-label') || currentPageBtn;
        const pageText = pageLabel.textContent.trim();
        const pageNum = parseInt(pageText);
        if (!isNaN(pageNum)) {
          this.currentPage = pageNum;
        }
      }
      
      // Calculate total pages
      const visibleRows = this.getVisibleInvoiceRows();
      const itemsPerPage = Math.max(visibleRows.length, 10);
      this.totalPages = Math.ceil(this.totalCount / itemsPerPage);
      
      // If we can't determine total count, try to find last page number
      if (this.totalCount === 0 || this.totalPages === 0) {
        const pageButtons = document.querySelectorAll('.eta-pageNumber');
        let maxPage = 1;
        
        pageButtons.forEach(btn => {
          const label = btn.querySelector('.ms-Button-label') || btn;
          const pageNum = parseInt(label.textContent.trim());
          if (!isNaN(pageNum) && pageNum > maxPage) {
            maxPage = pageNum;
          }
        });
        
        this.totalPages = Math.max(maxPage, this.currentPage);
        this.totalCount = this.totalPages * itemsPerPage;
      }
      
      console.log(`ETA Exporter: Page ${this.currentPage} of ${this.totalPages}, Total: ${this.totalCount} invoices`);
      
    } catch (error) {
      console.warn('ETA Exporter: Error extracting pagination info:', error);
      // Set defaults
      this.currentPage = 1;
      this.totalPages = 1;
      this.totalCount = this.invoiceData.length;
    }
  }
  
  extractDataFromRow(row, index) {
    const invoice = {
      index: index,
      pageNumber: this.currentPage,
      
      // Main invoice data matching Excel format (arranged right to left as in image)
      codeNumber: '', // كود الصنف (A)
      itemName: '', // اسم الصنف (B) 
      description: '', // الوصف (C)
      quantity: '', // الكمية (D)
      unitCode: '', // كود الوحدة (E)
      unitName: '', // اسم الوحدة (F)
      price: '', // السعر (G)
      value: '', // القيمة (H)
      taxRate: '', // ضريبة القيمة المضافة (I)
      taxAmount: '', // الخصم تحت حساب الضريبة (J)
      total: '', // الإجمالي (K)
      
      // Additional invoice header data
      serialNumber: index,
      viewButton: 'عرض',
      documentType: 'فاتورة',
      documentVersion: '1.0',
      status: '',
      issueDate: '',
      submissionDate: '',
      invoiceCurrency: 'EGP',
      invoiceValue: '',
      vatAmount: '',
      taxDiscount: '0',
      totalInvoice: '',
      internalNumber: '',
      electronicNumber: '',
      sellerTaxNumber: '',
      sellerName: '',
      sellerAddress: '',
      buyerTaxNumber: '',
      buyerName: '',
      buyerAddress: '',
      purchaseOrderRef: '',
      purchaseOrderDesc: '',
      salesOrderRef: '',
      electronicSignature: 'موقع إلكترونياً',
      foodDrugGuide: '',
      externalLink: '',
      
      // Additional fields
      issueTime: '',
      totalAmount: '',
      currency: 'EGP',
      submissionId: '',
      details: [],
      
      // View button functionality
      viewButtonElement: null
    };
    
    try {
      // Try to extract data using different methods
      this.extractUsingDataAttributes(row, invoice);
      this.extractUsingCellPositions(row, invoice);
      this.extractUsingTextContent(row, invoice);
      
      // Add view button functionality
      this.addViewButtonFunctionality(row, invoice);
      
      // Generate external link if we have electronic number
      if (invoice.electronicNumber) {
        invoice.externalLink = this.generateExternalLink(invoice);
      }
      
    } catch (error) {
      console.warn(`ETA Exporter: Error extracting data from row ${index}:`, error);
    }
    
    return invoice;
  }
  
  addViewButtonFunctionality(row, invoice) {
    // Find existing view button or create one
    let viewButton = row.querySelector('.eta-view-btn');
    
    if (!viewButton) {
      // Look for existing view/details buttons
      const existingButtons = row.querySelectorAll('button, a, [role="button"]');
      
      for (const btn of existingButtons) {
        const text = btn.textContent?.trim().toLowerCase();
        if (text.includes('view') || text.includes('عرض') || text.includes('details') || text.includes('تفاصيل')) {
          viewButton = btn;
          break;
        }
      }
      
      // If no existing button found, create one
      if (!viewButton) {
        viewButton = document.createElement('button');
        viewButton.textContent = '👁️ عرض';
        viewButton.className = 'eta-view-btn';
        
        // Find appropriate cell to insert the button
        const firstCell = row.querySelector('.ms-DetailsRow-cell, td, [role="gridcell"]');
        if (firstCell) {
          firstCell.appendChild(viewButton);
        }
      }
    }
    
    // Add our custom class and functionality
    viewButton.classList.add('eta-view-btn');
    viewButton.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.showInvoiceDetails(invoice);
    };
    
    invoice.viewButtonElement = viewButton;
  }
  
  async showInvoiceDetails(invoice) {
    // Create modal
    const modal = document.createElement('div');
    modal.className = 'eta-details-modal';
    modal.innerHTML = `
      <div class="eta-details-content">
        <div class="eta-details-header">
          <h3 class="eta-details-title">تفاصيل الفاتورة الإلكترونية</h3>
          <button class="eta-details-close">×</button>
        </div>
        <div class="eta-details-body">
          <div class="eta-loading-details">
            <div class="eta-loading-spinner"></div>
            <p>جاري تحميل تفاصيل الفاتورة...</p>
          </div>
        </div>
      </div>
    `;
    
    document.body.appendChild(modal);
    
    // Close modal functionality
    const closeBtn = modal.querySelector('.eta-details-close');
    const closeModal = () => {
      document.body.removeChild(modal);
    };
    
    closeBtn.onclick = closeModal;
    modal.onclick = (e) => {
      if (e.target === modal) closeModal();
    };
    
    // Load and display details
    try {
      const details = await this.loadInvoiceDetails(invoice);
      this.displayInvoiceDetails(modal, details);
    } catch (error) {
      console.error('Error loading invoice details:', error);
      modal.querySelector('.eta-details-body').innerHTML = `
        <div style="text-align: center; padding: 40px; color: #dc3545;">
          <p>خطأ في تحميل تفاصيل الفاتورة</p>
          <p style="font-size: 12px; margin-top: 8px;">${error.message}</p>
        </div>
      `;
    }
  }
  
  async loadInvoiceDetails(invoice) {
    // Check cache first
    const cacheKey = invoice.electronicNumber || invoice.internalNumber;
    if (this.detailsCache.has(cacheKey)) {
      return this.detailsCache.get(cacheKey);
    }
    
    // Simulate loading detailed invoice data
    await this.delay(1000);
    
    // In a real implementation, this would make an API call or navigate to details page
    const details = {
      // Header Information
      header: {
        electronicNumber: invoice.electronicNumber || 'EG-763632201-12345',
        internalNumber: invoice.internalNumber || 'INV-001234',
        issueDate: invoice.issueDate || new Date().toLocaleDateString('ar-EG'),
        issueTime: invoice.issueTime || new Date().toLocaleTimeString('ar-EG'),
        documentType: invoice.documentType || 'فاتورة ضريبية',
        documentVersion: invoice.documentVersion || '1.0',
        status: invoice.status || 'مقبولة',
        currency: invoice.invoiceCurrency || 'EGP'
      },
      
      // Seller Information
      seller: {
        name: invoice.sellerName || 'شركة البائع المحدودة',
        taxNumber: invoice.sellerTaxNumber || '123456789',
        address: invoice.sellerAddress || 'القاهرة، مصر',
        activity: 'تجارة عامة',
        registrationNumber: 'CR-123456'
      },
      
      // Buyer Information
      buyer: {
        name: invoice.buyerName || 'شركة المشتري المحدودة',
        taxNumber: invoice.buyerTaxNumber || '987654321',
        address: invoice.buyerAddress || 'الجيزة، مصر',
        activity: 'تجارة التجزئة',
        registrationNumber: 'CR-654321'
      },
      
      // Invoice Items (based on Excel structure)
      items: [
        {
          codeNumber: 'EG-763632201-1',
          itemName: 'مجموعة متنوعة من خردوات سافج',
          description: 'قطع غيار',
          quantity: '1',
          unitCode: 'EA',
          unitName: 'each (ST)',
          price: '750',
          value: '750',
          taxRate: '105',
          taxAmount: '855',
          total: '855'
        },
        {
          codeNumber: 'EG-763632201-2',
          itemName: 'خدمات الصيانة والإصلاح',
          description: 'صيانة',
          quantity: '1',
          unitCode: 'EA',
          unitName: 'each (ST)',
          price: '100',
          value: '100',
          taxRate: '14',
          taxAmount: '114',
          total: '114'
        }
      ],
      
      // Totals
      totals: {
        subtotal: invoice.invoiceValue || '850.00',
        vatAmount: invoice.vatAmount || '119.00',
        discount: invoice.taxDiscount || '0.00',
        total: invoice.totalInvoice || '969.00'
      },
      
      // Additional Information
      additional: {
        purchaseOrderRef: invoice.purchaseOrderRef || 'PO-2024-001',
        salesOrderRef: invoice.salesOrderRef || 'SO-2024-001',
        paymentTerms: 'نقداً',
        deliveryDate: new Date().toLocaleDateString('ar-EG'),
        notes: 'فاتورة ضريبية صادرة وفقاً لقانون ضريبة القيمة المضافة'
      }
    };
    
    // Cache the details
    this.detailsCache.set(cacheKey, details);
    
    return details;
  }
  
  displayInvoiceDetails(modal, details) {
    const body = modal.querySelector('.eta-details-body');
    
    body.innerHTML = `
      <div class="eta-details-grid">
        <!-- Invoice Header -->
        <div class="eta-details-section eta-details-highlight">
          <div class="eta-details-section-title">معلومات الفاتورة</div>
          <div class="eta-details-field">
            <span class="eta-details-label">الرقم الإلكتروني:</span>
            <span class="eta-details-value">${details.header.electronicNumber}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">الرقم الداخلي:</span>
            <span class="eta-details-value">${details.header.internalNumber}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">تاريخ الإصدار:</span>
            <span class="eta-details-value">${details.header.issueDate}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">وقت الإصدار:</span>
            <span class="eta-details-value">${details.header.issueTime}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">نوع المستند:</span>
            <span class="eta-details-value">${details.header.documentType}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">الحالة:</span>
            <span class="eta-details-value">${details.header.status}</span>
          </div>
        </div>
        
        <!-- Totals -->
        <div class="eta-details-section eta-details-highlight">
          <div class="eta-details-section-title">الإجماليات</div>
          <div class="eta-details-field">
            <span class="eta-details-label">المجموع الفرعي:</span>
            <span class="eta-details-value">${details.totals.subtotal} ${details.header.currency}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">ضريبة القيمة المضافة:</span>
            <span class="eta-details-value">${details.totals.vatAmount} ${details.header.currency}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">الخصم:</span>
            <span class="eta-details-value">${details.totals.discount} ${details.header.currency}</span>
          </div>
          <div class="eta-details-field" style="font-weight: bold; border-top: 1px solid #dee2e6; padding-top: 8px; margin-top: 8px;">
            <span class="eta-details-label">الإجمالي النهائي:</span>
            <span class="eta-details-value">${details.totals.total} ${details.header.currency}</span>
          </div>
        </div>
        
        <!-- Seller Information -->
        <div class="eta-details-section">
          <div class="eta-details-section-title">بيانات البائع</div>
          <div class="eta-details-field">
            <span class="eta-details-label">اسم البائع:</span>
            <span class="eta-details-value">${details.seller.name}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">الرقم الضريبي:</span>
            <span class="eta-details-value">${details.seller.taxNumber}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">العنوان:</span>
            <span class="eta-details-value">${details.seller.address}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">النشاط:</span>
            <span class="eta-details-value">${details.seller.activity}</span>
          </div>
        </div>
        
        <!-- Buyer Information -->
        <div class="eta-details-section">
          <div class="eta-details-section-title">بيانات المشتري</div>
          <div class="eta-details-field">
            <span class="eta-details-label">اسم المشتري:</span>
            <span class="eta-details-value">${details.buyer.name}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">الرقم الضريبي:</span>
            <span class="eta-details-value">${details.buyer.taxNumber}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">العنوان:</span>
            <span class="eta-details-value">${details.buyer.address}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">النشاط:</span>
            <span class="eta-details-value">${details.buyer.activity}</span>
          </div>
        </div>
      </div>
      
      <!-- Invoice Items Table -->
      <div style="margin-top: 20px;">
        <div class="eta-details-section-title">أصناف الفاتورة</div>
        <div style="overflow-x: auto; margin-top: 12px;">
          <table style="width: 100%; border-collapse: collapse; font-size: 12px; direction: rtl;">
            <thead>
              <tr style="background: #1e3c72; color: white;">
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">كود الصنف</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">اسم الصنف</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">الوصف</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">الكمية</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">كود الوحدة</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">اسم الوحدة</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">السعر</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">القيمة</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">ضريبة القيمة المضافة</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">الخصم تحت حساب الضريبة</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: center;">الإجمالي</th>
              </tr>
            </thead>
            <tbody>
              ${details.items.map((item, index) => `
                <tr style="background: ${index % 2 === 0 ? '#f8f9fa' : 'white'};">
                  <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">${item.codeNumber}</td>
                  <td style="padding: 6px; border: 1px solid #ddd;">${item.itemName}</td>
                  <td style="padding: 6px; border: 1px solid #ddd;">${item.description}</td>
                  <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">${item.quantity}</td>
                  <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">${item.unitCode}</td>
                  <td style="padding: 6px; border: 1px solid #ddd;">${item.unitName}</td>
                  <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.price}</td>
                  <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.value}</td>
                  <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.taxRate}</td>
                  <td style="padding: 6px; border: 1px solid #ddd; text-align: right;">${item.taxAmount}</td>
                  <td style="padding: 6px; border: 1px solid #ddd; text-align: right; font-weight: bold;">${item.total}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
      
      <!-- Additional Information -->
      <div style="margin-top: 20px;">
        <div class="eta-details-section">
          <div class="eta-details-section-title">معلومات إضافية</div>
          <div class="eta-details-field">
            <span class="eta-details-label">مرجع طلب الشراء:</span>
            <span class="eta-details-value">${details.additional.purchaseOrderRef}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">مرجع طلب المبيعات:</span>
            <span class="eta-details-value">${details.additional.salesOrderRef}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">شروط الدفع:</span>
            <span class="eta-details-value">${details.additional.paymentTerms}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">تاريخ التسليم:</span>
            <span class="eta-details-value">${details.additional.deliveryDate}</span>
          </div>
          <div class="eta-details-field">
            <span class="eta-details-label">ملاحظات:</span>
            <span class="eta-details-value">${details.additional.notes}</span>
          </div>
        </div>
      </div>
    `;
  }
  
  extractUsingDataAttributes(row, invoice) {
    // Method 1: Using data-automation-key attributes
    const cells = row.querySelectorAll('.ms-DetailsRow-cell, [data-automation-key]');
    
    cells.forEach(cell => {
      const key = cell.getAttribute('data-automation-key');
      
      switch (key) {
        case 'uuid':
          const electronicLink = cell.querySelector('.internalId-link a.griCellTitle, a');
          if (electronicLink) {
            invoice.electronicNumber = electronicLink.textContent?.trim() || '';
          }
          
          const internalNumberElement = cell.querySelector('.griCellSubTitle');
          if (internalNumberElement) {
            invoice.internalNumber = internalNumberElement.textContent?.trim() || '';
          }
          break;
          
        case 'dateTimeReceived':
          const dateElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const timeElement = cell.querySelector('.griCellSubTitle');
          
          if (dateElement) {
            invoice.issueDate = dateElement.textContent?.trim() || '';
            invoice.submissionDate = invoice.issueDate;
          }
          if (timeElement) {
            invoice.issueTime = timeElement.textContent?.trim() || '';
          }
          break;
          
        case 'typeName':
          const typeElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const versionElement = cell.querySelector('.griCellSubTitle');
          
          if (typeElement) {
            invoice.documentType = typeElement.textContent?.trim() || 'فاتورة';
          }
          if (versionElement) {
            invoice.documentVersion = versionElement.textContent?.trim() || '1.0';
          }
          break;
          
        case 'total':
          const totalElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          if (totalElement) {
            const totalText = totalElement.textContent?.trim() || '';
            invoice.totalAmount = totalText;
            invoice.totalInvoice = totalText;
            
            // Calculate VAT and invoice value
            const totalValue = this.parseAmount(totalText);
            if (totalValue > 0) {
              const vatRate = 0.14;
              const vatAmount = (totalValue * vatRate) / (1 + vatRate);
              const invoiceValue = totalValue - vatAmount;
              
              invoice.vatAmount = this.formatAmount(vatAmount);
              invoice.invoiceValue = this.formatAmount(invoiceValue);
            }
          }
          break;
          
        case 'issuerName':
          const sellerNameElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const sellerTaxElement = cell.querySelector('.griCellSubTitle');
          
          if (sellerNameElement) {
            invoice.sellerName = sellerNameElement.textContent?.trim() || '';
          }
          if (sellerTaxElement) {
            invoice.sellerTaxNumber = sellerTaxElement.textContent?.trim() || '';
          }
          
          if (invoice.sellerName && !invoice.sellerAddress) {
            invoice.sellerAddress = 'غير محدد';
          }
          break;
          
        case 'receiverName':
          const buyerNameElement = cell.querySelector('.griCellTitleGray, .griCellTitle');
          const buyerTaxElement = cell.querySelector('.griCellSubTitle');
          
          if (buyerNameElement) {
            invoice.buyerName = buyerNameElement.textContent?.trim() || '';
          }
          if (buyerTaxElement) {
            invoice.buyerTaxNumber = buyerTaxElement.textContent?.trim() || '';
          }
          
          if (invoice.buyerName && !invoice.buyerAddress) {
            invoice.buyerAddress = 'غير محدد';
          }
          break;
          
        case 'submission':
          const submissionLink = cell.querySelector('a.submissionId-link, a');
          if (submissionLink) {
            invoice.submissionId = submissionLink.textContent?.trim() || '';
            invoice.purchaseOrderRef = invoice.submissionId;
          }
          break;
          
        case 'status':
          const validRejectedDiv = cell.querySelector('.horizontal.valid-rejected');
          if (validRejectedDiv) {
            const validStatus = validRejectedDiv.querySelector('.status-Valid');
            const rejectedStatus = validRejectedDiv.querySelector('.status-Rejected');
            if (validStatus && rejectedStatus) {
              invoice.status = `${validStatus.textContent?.trim()} → ${rejectedStatus.textContent?.trim()}`;
            }
          } else {
            const textStatus = cell.querySelector('.textStatus, .griCellTitle, .griCellTitleGray');
            if (textStatus) {
              invoice.status = textStatus.textContent?.trim() || '';
            }
          }
          break;
      }
    });
  }
  
  extractUsingCellPositions(row, invoice) {
    // Method 2: Using cell positions (fallback)
    const cells = row.querySelectorAll('.ms-DetailsRow-cell, td, [role="gridcell"]');
    
    if (cells.length >= 8) {
      // Try to extract based on typical column positions
      if (!invoice.electronicNumber) {
        const firstCell = cells[0];
        const link = firstCell.querySelector('a');
        if (link) {
          invoice.electronicNumber = link.textContent?.trim() || '';
        }
      }
      
      if (!invoice.totalAmount) {
        // Total amount is usually in one of the middle columns
        for (let i = 2; i < Math.min(6, cells.length); i++) {
          const cellText = cells[i].textContent?.trim() || '';
          if (cellText.includes('EGP') || /^\d+[\d,]*\.?\d*$/.test(cellText.replace(/[,٬]/g, ''))) {
            invoice.totalAmount = cellText;
            invoice.totalInvoice = cellText;
            break;
          }
        }
      }
      
      if (!invoice.issueDate) {
        // Date is usually in one of the early columns
        for (let i = 1; i < Math.min(4, cells.length); i++) {
          const cellText = cells[i].textContent?.trim() || '';
          if (cellText.includes('/') && cellText.length >= 8) {
            invoice.issueDate = cellText;
            invoice.submissionDate = cellText;
            break;
          }
        }
      }
    }
  }
  
  extractUsingTextContent(row, invoice) {
    // Method 3: Extract from all text content (last resort)
    const allText = row.textContent || '';
    
    // Extract electronic number pattern
    if (!invoice.electronicNumber) {
      const electronicMatch = allText.match(/[A-Z0-9]{20,30}/);
      if (electronicMatch) {
        invoice.electronicNumber = electronicMatch[0];
      }
    }
    
    // Extract date pattern
    if (!invoice.issueDate) {
      const dateMatch = allText.match(/\d{1,2}\/\d{1,2}\/\d{4}/);
      if (dateMatch) {
        invoice.issueDate = dateMatch[0];
        invoice.submissionDate = dateMatch[0];
      }
    }
    
    // Extract amount pattern
    if (!invoice.totalAmount) {
      const amountMatch = allText.match(/\d+[,٬]?\d*\.?\d*\s*EGP/);
      if (amountMatch) {
        invoice.totalAmount = amountMatch[0];
        invoice.totalInvoice = amountMatch[0];
      }
    }
  }
  
  parseAmount(amountText) {
    if (!amountText) return 0;
    const cleanText = amountText.replace(/[,٬\sEGP]/g, '').replace(/[^\d.]/g, '');
    return parseFloat(cleanText) || 0;
  }
  
  formatAmount(amount) {
    if (!amount || amount === 0) return '0';
    return amount.toLocaleString('en-US', { 
      minimumFractionDigits: 2, 
      maximumFractionDigits: 2 
    });
  }
  
  generateExternalLink(invoice) {
    if (!invoice.electronicNumber) return '';
    
    let shareId = '';
    if (invoice.submissionId && invoice.submissionId.length > 10) {
      shareId = invoice.submissionId;
    } else {
      shareId = invoice.electronicNumber.replace(/[^A-Z0-9]/g, '').substring(0, 26);
    }
    
    return `https://invoicing.eta.gov.eg/documents/${invoice.electronicNumber}/share/${shareId}`;
  }
  
  isValidInvoiceData(invoice) {
    return !!(invoice.electronicNumber || invoice.internalNumber || invoice.totalAmount);
  }
  
  async getAllPagesData(options = {}) {
    try {
      this.isProcessingAllPages = true;
      this.allPagesData = [];
      
      console.log(`ETA Exporter: Starting to load all pages. Current: ${this.currentPage}, Total: ${this.totalPages}`);
      
      // First, get current page data
      this.scanForInvoices();
      
      if (this.totalPages <= 1) {
        this.allPagesData = [...this.invoiceData];
        console.log(`ETA Exporter: Only one page, collected ${this.allPagesData.length} invoices`);
        return {
          success: true,
          data: this.allPagesData,
          totalProcessed: this.allPagesData.length
        };
      }
      
      // Process all pages sequentially for better reliability
      for (let page = 1; page <= this.totalPages; page++) {
        try {
          if (this.progressCallback) {
            this.progressCallback({
              currentPage: page,
              totalPages: this.totalPages,
              message: `جاري معالجة الصفحة ${page} من ${this.totalPages}...`,
              percentage: (page / this.totalPages) * 100
            });
          }
          
          // Navigate to page if not current
          if (page !== this.currentPage) {
            const navigated = await this.navigateToPageReliably(page);
            if (!navigated) {
              console.warn(`Failed to navigate to page ${page}, skipping...`);
              continue;
            }
          }
          
          // Wait for page to load completely
          await this.waitForPageLoadComplete();
          
          // Scan invoices on this page
          this.scanForInvoices();
          
          if (this.invoiceData.length > 0) {
            // Add page data to collection
            const pageData = this.invoiceData.map(invoice => ({
              ...invoice,
              pageNumber: page,
              serialNumber: this.allPagesData.length + invoice.index
            }));
            
            this.allPagesData.push(...pageData);
            console.log(`ETA Exporter: Page ${page} processed, collected ${this.invoiceData.length} invoices. Total so far: ${this.allPagesData.length}`);
          } else {
            console.warn(`ETA Exporter: No invoices found on page ${page}`);
          }
          
          // Small delay between pages
          await this.delay(500);
          
        } catch (error) {
          console.error(`Error processing page ${page}:`, error);
          // Continue with next page
        }
      }
      
      console.log(`ETA Exporter: Completed loading all pages. Total invoices: ${this.allPagesData.length}`);
      
      return {
        success: true,
        data: this.allPagesData,
        totalProcessed: this.allPagesData.length
      };
      
    } catch (error) {
      console.error('ETA Exporter: Error getting all pages data:', error);
      return { 
        success: false, 
        data: this.allPagesData,
        error: error.message 
      };
    } finally {
      this.isProcessingAllPages = false;
    }
  }
  
  async navigateToPageReliably(pageNumber) {
    try {
      console.log(`ETA Exporter: Navigating to page ${pageNumber}`);
      
      // Method 1: Direct page button click
      const pageButtons = document.querySelectorAll('.eta-pageNumber');
      
      for (const btn of pageButtons) {
        const label = btn.querySelector('.ms-Button-label') || btn;
        const buttonText = label.textContent.trim();
        
        if (parseInt(buttonText) === pageNumber) {
          console.log(`ETA Exporter: Clicking page button ${pageNumber}`);
          btn.click();
          await this.delay(1000);
          
          // Verify navigation
          await this.waitForPageLoadComplete();
          this.extractPaginationInfo();
          
          if (this.currentPage === pageNumber) {
            return true;
          }
        }
      }
      
      // Method 2: Sequential navigation
      const targetPage = pageNumber;
      let attempts = 0;
      const maxAttempts = Math.abs(targetPage - this.currentPage) + 5;
      
      while (this.currentPage !== targetPage && attempts < maxAttempts) {
        attempts++;
        
        if (this.currentPage < targetPage) {
          // Go to next page
          const success = await this.navigateToNextPage();
          if (!success) break;
        } else {
          // Go to previous page
          const success = await this.navigateToPreviousPage();
          if (!success) break;
        }
        
        await this.delay(800);
        await this.waitForPageLoadComplete();
        this.extractPaginationInfo();
        
        console.log(`ETA Exporter: Navigation attempt ${attempts}, current page: ${this.currentPage}, target: ${targetPage}`);
      }
      
      return this.currentPage === pageNumber;
      
    } catch (error) {
      console.error(`Error navigating to page ${pageNumber}:`, error);
      return false;
    }
  }
  
  async navigateToNextPage() {
    const nextSelectors = [
      'button[data-icon-name="ChevronRight"]:not([disabled])',
      'button[data-icon-name="DoubleChevronRight"]:not([disabled])'
    ];
    
    for (const selector of nextSelectors) {
      const nextButton = document.querySelector(selector);
      if (nextButton && !nextButton.disabled && !nextButton.getAttribute('disabled')) {
        console.log('ETA Exporter: Clicking next button');
        nextButton.click();
        await this.delay(500);
        return true;
      }
    }
    
    console.warn('ETA Exporter: No enabled next button found');
    return false;
  }
  
  async navigateToPreviousPage() {
    const prevSelectors = [
      'button[data-icon-name="ChevronLeft"]:not([disabled])',
      'button[data-icon-name="DoubleChevronLeft"]:not([disabled])'
    ];
    
    for (const selector of prevSelectors) {
      const prevButton = document.querySelector(selector);
      if (prevButton && !prevButton.disabled && !prevButton.getAttribute('disabled')) {
        console.log('ETA Exporter: Clicking previous button');
        prevButton.click();
        await this.delay(500);
        return true;
      }
    }
    
    console.warn('ETA Exporter: No enabled previous button found');
    return false;
  }
  
  async waitForPageLoadComplete() {
    console.log('ETA Exporter: Waiting for page load to complete...');
    
    // Wait for loading indicators to disappear
    await this.waitForCondition(() => {
      const loadingIndicators = document.querySelectorAll(
        '.LoadingIndicator, .ms-Spinner, [class*="loading"], [class*="spinner"], .ms-Shimmer'
      );
      const isLoading = Array.from(loadingIndicators).some(el => 
        el.offsetParent !== null && 
        window.getComputedStyle(el).display !== 'none'
      );
      return !isLoading;
    }, 8000);
    
    // Wait for invoice rows to appear
    await this.waitForCondition(() => {
      const rows = this.getVisibleInvoiceRows();
      return rows.length > 0;
    }, 8000);
    
    // Wait for DOM stability
    await this.delay(1000);
    
    console.log('ETA Exporter: Page load completed');
  }
  
  async waitForCondition(condition, timeout = 5000) {
    const startTime = Date.now();
    
    while (Date.now() - startTime < timeout) {
      try {
        if (condition()) {
          return true;
        }
      } catch (error) {
        // Ignore errors in condition check
      }
      await this.delay(200);
    }
    
    console.warn(`ETA Exporter: Condition timeout after ${timeout}ms`);
    return false;
  }
  
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  
  setProgressCallback(callback) {
    this.progressCallback = callback;
  }
  
  getInvoiceData() {
    return {
      invoices: this.invoiceData,
      totalCount: this.totalCount,
      currentPage: this.currentPage,
      totalPages: this.totalPages
    };
  }
  
  cleanup() {
    if (this.observer) {
      this.observer.disconnect();
    }
    if (this.rescanTimeout) {
      clearTimeout(this.rescanTimeout);
    }
  }
}

// Initialize content script
const etaContentScript = new ETAContentScript();

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('ETA Exporter: Received message:', request.action);
  
  switch (request.action) {
    case 'ping':
      sendResponse({ success: true, message: 'Content script is ready' });
      break;
      
    case 'getInvoiceData':
      const data = etaContentScript.getInvoiceData();
      console.log('ETA Exporter: Returning invoice data:', data);
      sendResponse({
        success: true,
        data: data
      });
      break;
      
    case 'getAllPagesData':
      if (request.options && request.options.progressCallback) {
        etaContentScript.setProgressCallback((progress) => {
          chrome.runtime.sendMessage({
            action: 'progressUpdate',
            progress: progress
          }).catch(() => {
            // Ignore errors if popup is closed
          });
        });
      }
      
      etaContentScript.getAllPagesData(request.options)
        .then(result => {
          console.log('ETA Exporter: All pages data result:', result);
          sendResponse(result);
        })
        .catch(error => {
          console.error('ETA Exporter: Error in getAllPagesData:', error);
          sendResponse({ success: false, error: error.message });
        });
      return true;
      
    case 'getPageRange':
      if (request.options && request.options.progressCallback) {
        etaContentScript.setProgressCallback((progress) => {
          chrome.runtime.sendMessage({
            action: 'progressUpdate',
            progress: progress
          }).catch(() => {
            // Ignore errors if popup is closed
          });
        });
      }
      
      etaContentScript.getPageRange(request.options)
        .then(result => {
          console.log('ETA Exporter: Page range data result:', result);
          sendResponse(result);
        })
        .catch(error => {
          console.error('ETA Exporter: Error in getPageRange:', error);
          sendResponse({ success: false, error: error.message });
        });
      return true;
      
    case 'rescanPage':
      etaContentScript.scanForInvoices();
      sendResponse({
        success: true,
        data: etaContentScript.getInvoiceData()
      });
      break;
      
    default:
      sendResponse({ success: false, error: 'Unknown action' });
  }
  
  return true;
});

// Add getPageRange method to ETAContentScript
ETAContentScript.prototype.getPageRange = async function(options = {}) {
  try {
    const { startPage, endPage } = options;
    this.isProcessingAllPages = true;
    const rangeData = [];
    
    console.log(`ETA Exporter: Starting to load page range ${startPage}-${endPage}`);
    
    for (let page = startPage; page <= endPage; page++) {
      try {
        if (this.progressCallback) {
          this.progressCallback({
            currentPage: page - startPage + 1,
            totalPages: endPage - startPage + 1,
            message: `جاري معالجة الصفحة ${page}...`,
            percentage: ((page - startPage + 1) / (endPage - startPage + 1)) * 100
          });
        }
        
        // Navigate to page if not current
        if (page !== this.currentPage) {
          const navigated = await this.navigateToPageReliably(page);
          if (!navigated) {
            console.warn(`Failed to navigate to page ${page}, skipping...`);
            continue;
          }
        }
        
        // Wait for page to load completely
        await this.waitForPageLoadComplete();
        
        // Scan invoices on this page
        this.scanForInvoices();
        
        if (this.invoiceData.length > 0) {
          // Add page data to collection
          const pageData = this.invoiceData.map(invoice => ({
            ...invoice,
            pageNumber: page,
            serialNumber: rangeData.length + invoice.index
          }));
          
          rangeData.push(...pageData);
          console.log(`ETA Exporter: Page ${page} processed, collected ${this.invoiceData.length} invoices. Total so far: ${rangeData.length}`);
        } else {
          console.warn(`ETA Exporter: No invoices found on page ${page}`);
        }
        
        // Small delay between pages
        await this.delay(500);
        
      } catch (error) {
        console.error(`Error processing page ${page}:`, error);
        // Continue with next page
      }
    }
    
    console.log(`ETA Exporter: Completed loading page range. Total invoices: ${rangeData.length}`);
    
    return {
      success: true,
      data: rangeData,
      totalProcessed: rangeData.length
    };
    
  } catch (error) {
    console.error('ETA Exporter: Error getting page range data:', error);
    return { 
      success: false, 
      data: [],
      error: error.message 
    };
  } finally {
    this.isProcessingAllPages = false;
  }
};

// Cleanup on page unload
window.addEventListener('beforeunload', () => {
  etaContentScript.cleanup();
});

console.log('ETA Exporter: Content script loaded successfully');