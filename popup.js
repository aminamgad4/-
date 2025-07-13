class ETAInvoiceExporter {
  constructor() {
    this.invoiceData = [];
    this.totalCount = 0;
    this.currentPage = 1;
    this.totalPages = 1;
    this.isProcessing = false;
    this.isCancelled = false;
    this.processedInvoices = 0;
    this.startTime = null;
    
    this.elements = {};
    this.initializeElements();
    this.attachEventListeners();
    this.initializeApp();
  }
  
  initializeElements() {
    // Stats elements
    this.elements.currentPageCount = document.getElementById('currentPageCount');
    this.elements.totalInvoicesCount = document.getElementById('totalInvoicesCount');
    this.elements.currentPageNumber = document.getElementById('currentPageNumber');
    this.elements.totalPagesCount = document.getElementById('totalPagesCount');
    
    // Mode selection
    this.elements.modeOptions = document.querySelectorAll('.mode-option');
    this.elements.rangeInputs = document.getElementById('range-inputs');
    this.elements.startPage = document.getElementById('startPage');
    this.elements.endPage = document.getElementById('endPage');
    
    // Progress elements
    this.elements.progressContainer = document.getElementById('progressContainer');
    this.elements.progressBar = document.getElementById('progressBar');
    this.elements.progressText = document.getElementById('progressText');
    this.elements.processedPages = document.getElementById('processedPages');
    this.elements.processedInvoices = document.getElementById('processedInvoices');
    this.elements.estimatedTime = document.getElementById('estimatedTime');
    this.elements.cancelBtn = document.getElementById('cancelBtn');
    
    // Buttons
    this.elements.excelBtn = document.getElementById('excelBtn');
    this.elements.jsonBtn = document.getElementById('jsonBtn');
    this.elements.closeBtn = document.getElementById('closeBtn');
    
    // Status
    this.elements.status = document.getElementById('status');
    
    // Checkboxes
    this.elements.selectAll = document.getElementById('option-select-all');
    this.elements.fieldCheckboxes = document.querySelectorAll('input[type="checkbox"][id^="option-"]:not(#option-select-all)');
  }
  
  attachEventListeners() {
    // Mode selection
    this.elements.modeOptions.forEach(option => {
      option.addEventListener('click', () => this.selectMode(option));
    });
    
    // Range mode inputs
    document.querySelector('input[value="range"]').addEventListener('change', (e) => {
      this.elements.rangeInputs.style.display = e.target.checked ? 'block' : 'none';
    });
    
    // Select all checkbox
    this.elements.selectAll.addEventListener('change', (e) => {
      this.toggleAllCheckboxes(e.target.checked);
    });
    
    // Cancel button
    this.elements.cancelBtn.addEventListener('click', () => this.cancelProcess());
    
    // Export buttons
    this.elements.excelBtn.addEventListener('click', () => this.startExport('excel'));
    this.elements.jsonBtn.addEventListener('click', () => this.startExport('json'));
    this.elements.closeBtn.addEventListener('click', () => window.close());
  }
  
  selectMode(option) {
    // Remove selected class from all options
    this.elements.modeOptions.forEach(opt => opt.classList.remove('selected'));
    
    // Add selected class to clicked option
    option.classList.add('selected');
    
    // Check the radio button
    const radio = option.querySelector('input[type="radio"]');
    if (radio) {
      radio.checked = true;
    }
    
    // Show/hide range inputs
    const isRange = option.dataset.mode === 'range';
    this.elements.rangeInputs.style.display = isRange ? 'block' : 'none';
  }
  
  toggleAllCheckboxes(checked) {
    this.elements.fieldCheckboxes.forEach(checkbox => {
      checkbox.checked = checked;
    });
  }
  
  async initializeApp() {
    try {
      this.showStatus('جاري فحص الصفحة الحالية...', 'loading');
      
      // Check if we're on the correct page
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      
      if (!tab.url.includes('invoicing.eta.gov.eg')) {
        throw new Error('يرجى الانتقال إلى بوابة الفواتير الإلكترونية المصرية');
      }
      
      // Ensure content script is loaded
      await this.ensureContentScriptLoaded(tab.id);
      
      // Load initial data
      await this.loadCurrentPageData();
      this.updateStatsDisplay();
      this.showStatus('جاهز للتصدير', 'success');
      
    } catch (error) {
      this.showStatus('خطأ في تحميل البيانات: ' + error.message, 'error');
      this.disableButtons();
    }
  }
  
  async ensureContentScriptLoaded(tabId) {
    try {
      // Try to ping the content script
      await chrome.tabs.sendMessage(tabId, { action: 'ping' });
    } catch (error) {
      // Content script not loaded, inject it
      console.log('Content script not found, injecting...');
      
      try {
        await chrome.scripting.executeScript({
          target: { tabId: tabId },
          files: ['content.js']
        });
        
        // Wait a bit for the script to initialize
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Try to ping again
        await chrome.tabs.sendMessage(tabId, { action: 'ping' });
        console.log('Content script successfully injected and ready');
      } catch (injectError) {
        throw new Error('فشل في تحميل المكونات المطلوبة. يرجى إعادة تحميل الصفحة.');
      }
    }
  }
  
  async loadCurrentPageData() {
    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      const response = await this.sendMessageWithRetry(tab.id, { action: 'getInvoiceData' });
      
      if (!response || !response.success) {
        throw new Error('فشل في الحصول على بيانات الفواتير');
      }
      
      this.invoiceData = response.data.invoices || [];
      this.totalCount = response.data.totalCount || this.invoiceData.length;
      this.currentPage = response.data.currentPage || 1;
      this.totalPages = response.data.totalPages || 1;
      
    } catch (error) {
      console.error('Load error:', error);
      throw error;
    }
  }
  
  async sendMessageWithRetry(tabId, message, maxRetries = 3) {
    for (let i = 0; i < maxRetries; i++) {
      try {
        const response = await chrome.tabs.sendMessage(tabId, message);
        return response;
      } catch (error) {
        console.log(`Message attempt ${i + 1} failed:`, error);
        
        if (i === maxRetries - 1) {
          // Last attempt failed
          if (error.message.includes('Could not establish connection')) {
            throw new Error('فشل في الاتصال مع الصفحة. يرجى إعادة تحميل الصفحة والمحاولة مرة أخرى.');
          }
          throw error;
        }
        
        // Wait before retry
        await new Promise(resolve => setTimeout(resolve, 500 * (i + 1)));
        
        // Try to ensure content script is loaded again
        try {
          await this.ensureContentScriptLoaded(tabId);
        } catch (ensureError) {
          console.warn('Failed to ensure content script:', ensureError);
        }
      }
    }
  }
  
  updateStatsDisplay() {
    this.elements.currentPageCount.textContent = this.invoiceData.length;
    this.elements.totalInvoicesCount.textContent = this.totalCount.toLocaleString('ar-EG');
    this.elements.currentPageNumber.textContent = this.currentPage;
    this.elements.totalPagesCount.textContent = this.totalPages;
    
    // Update range inputs max values
    this.elements.startPage.max = this.totalPages;
    this.elements.endPage.max = this.totalPages;
    this.elements.endPage.value = this.totalPages;
  }
  
  getSelectedMode() {
    const checkedRadio = document.querySelector('input[name="downloadMode"]:checked');
    return checkedRadio ? checkedRadio.value : 'current';
  }
  
  getSelectedFields() {
    const fields = {};
    this.elements.fieldCheckboxes.forEach(checkbox => {
      const fieldName = checkbox.id.replace('option-', '').replace(/-/g, '_');
      fields[fieldName] = checkbox.checked;
    });
    return fields;
  }
  
  validateSelections() {
    const selectedFields = this.getSelectedFields();
    const hasSelectedField = Object.values(selectedFields).some(selected => selected);
    
    if (!hasSelectedField) {
      throw new Error('يرجى اختيار حقل واحد على الأقل للتصدير');
    }
    
    const mode = this.getSelectedMode();
    if (mode === 'range') {
      const startPage = parseInt(this.elements.startPage.value);
      const endPage = parseInt(this.elements.endPage.value);
      
      if (!startPage || !endPage || startPage < 1 || endPage < 1 || startPage > endPage) {
        throw new Error('يرجى إدخال نطاق صحيح للصفحات');
      }
      
      if (endPage > this.totalPages) {
        throw new Error(`نطاق الصفحات يتجاوز العدد الإجمالي (${this.totalPages})`);
      }
    }
    
    return true;
  }
  
  async startExport(format) {
    if (this.isProcessing) {
      this.showStatus('جاري المعالجة... يرجى الانتظار', 'loading');
      return;
    }
    
    try {
      this.validateSelections();
      
      this.isProcessing = true;
      this.isCancelled = false;
      this.processedInvoices = 0;
      this.startTime = Date.now();
      
      this.disableButtons();
      this.showProgress();
      
      const mode = this.getSelectedMode();
      const selectedFields = this.getSelectedFields();
      
      let exportData = [];
      
      if (mode === 'current') {
        exportData = await this.exportCurrentPage();
      } else if (mode === 'all') {
        exportData = await this.exportAllPages();
      } else if (mode === 'range') {
        const startPage = parseInt(this.elements.startPage.value);
        const endPage = parseInt(this.elements.endPage.value);
        exportData = await this.exportPageRange(startPage, endPage);
      }
      
      if (!this.isCancelled && exportData.length > 0) {
        await this.generateExportFile(exportData, format, selectedFields);
        this.showStatus(`تم تصدير ${exportData.length} فاتورة بنجاح!`, 'success');
      } else if (this.isCancelled) {
        this.showStatus('تم إلغاء العملية', 'error');
      } else {
        this.showStatus('لم يتم العثور على بيانات للتصدير', 'error');
      }
      
    } catch (error) {
      this.showStatus('خطأ في التصدير: ' + error.message, 'error');
    } finally {
      this.isProcessing = false;
      this.enableButtons();
      this.hideProgress();
    }
  }
  
  async exportCurrentPage() {
    this.updateProgress(1, 1, 'تصدير الصفحة الحالية...');
    await this.delay(500); // Simulate processing
    return [...this.invoiceData];
  }
  
  async exportAllPages() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    
    const allData = await this.sendMessageWithRetry(tab.id, { 
      action: 'getAllPagesData',
      options: { progressCallback: true }
    });
    
    if (!allData || !allData.success) {
      throw new Error('فشل في تحميل جميع الصفحات: ' + (allData?.error || 'خطأ غير معروف'));
    }
    
    return allData.data;
  }
  
  async exportPageRange(startPage, endPage) {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    
    const rangeData = await this.sendMessageWithRetry(tab.id, { 
      action: 'getPageRange',
      options: { 
        startPage: startPage,
        endPage: endPage,
        progressCallback: true 
      }
    });
    
    if (!rangeData || !rangeData.success) {
      throw new Error('فشل في تحميل النطاق المحدد: ' + (rangeData?.error || 'خطأ غير معروف'));
    }
    
    return rangeData.data;
  }
  
  updateProgress(current, total, message) {
    const percentage = (current / total) * 100;
    
    this.elements.progressBar.style.width = `${Math.min(percentage, 100)}%`;
    this.elements.progressText.textContent = message;
    this.elements.processedPages.textContent = current;
    
    this.updateProgressDetails();
  }
  
  updateProgressDetails() {
    this.elements.processedInvoices.textContent = this.processedInvoices.toLocaleString('ar-EG');
    
    // Calculate estimated time remaining
    if (this.startTime && this.processedInvoices > 0) {
      const elapsed = Date.now() - this.startTime;
      const rate = this.processedInvoices / elapsed;
      const remaining = this.totalCount - this.processedInvoices;
      const estimatedMs = remaining / rate;
      
      if (estimatedMs > 0 && estimatedMs < Infinity) {
        const minutes = Math.floor(estimatedMs / 60000);
        const seconds = Math.floor((estimatedMs % 60000) / 1000);
        this.elements.estimatedTime.textContent = `${minutes}:${seconds.toString().padStart(2, '0')}`;
      } else {
        this.elements.estimatedTime.textContent = '-';
      }
    }
  }
  
  cancelProcess() {
    this.isCancelled = true;
    this.showStatus('جاري إلغاء العملية...', 'loading');
  }
  
  showProgress() {
    this.elements.progressContainer.style.display = 'block';
    this.elements.status.style.display = 'none';
  }
  
  hideProgress() {
    this.elements.progressContainer.style.display = 'none';
  }
  
  async generateExportFile(data, format, selectedFields) {
    this.updateProgress(100, 100, 'جاري إنشاء الملف...');
    
    if (format === 'excel') {
      this.generateExcelFile(data, selectedFields);
    } else if (format === 'json') {
      this.generateJSONFile(data, selectedFields);
    }
  }
  
  generateExcelFile(data, selectedFields) {
    const wb = XLSX.utils.book_new();
    
    // Create headers based on selected fields
    const headers = [];
    const fieldMap = {
      serial_number: 'تسلسل',
      details_button: 'عرض',
      document_type: 'نوع المستند',
      document_version: 'نسخة المستند',
      status: 'الحالة',
      issue_date: 'تاريخ الإصدار',
      submission_date: 'تاريخ التقديم',
      invoice_currency: 'عملة الفاتورة',
      invoice_value: 'قيمة الفاتورة',
      vat_amount: 'ضريبة القيمة المضافة',
      tax_discount: 'الخصم تحت حساب الضريبة',
      total_invoice: 'إجمالي الفاتورة',
      internal_number: 'الرقم الداخلي',
      electronic_number: 'الرقم الإلكتروني',
      seller_tax_number: 'الرقم الضريبي للبائع',
      seller_name: 'اسم البائع',
      seller_address: 'عنوان البائع',
      buyer_tax_number: 'الرقم الضريبي للمشتري',
      buyer_name: 'اسم المشتري',
      buyer_address: 'عنوان المشتري',
      purchase_order_ref: 'مرجع طلب الشراء',
      purchase_order_desc: 'وصف طلب الشراء',
      sales_order_ref: 'مرجع طلب المبيعات',
      electronic_signature: 'التوقيع الإلكتروني',
      food_drug_guide: 'دليل الغذاء والدواء',
      external_link: 'الرابط الخارجي'
    };
    
    // Build headers array based on selected fields
    Object.keys(selectedFields).forEach(field => {
      if (selectedFields[field] && fieldMap[field]) {
        headers.push(fieldMap[field]);
      }
    });
    
    // Create data rows
    const rows = [headers];
    
    data.forEach((invoice, index) => {
      const row = [];
      
      Object.keys(selectedFields).forEach(field => {
        if (selectedFields[field]) {
          let value = '';
          
          switch (field) {
            case 'serial_number':
              value = index + 1;
              break;
            case 'details_button':
              value = 'عرض';
              break;
            case 'document_type':
              value = invoice.documentType || 'فاتورة';
              break;
            case 'document_version':
              value = invoice.documentVersion || '1.0';
              break;
            case 'status':
              value = invoice.status || '';
              break;
            case 'issue_date':
              value = invoice.issueDate || '';
              break;
            case 'submission_date':
              value = invoice.submissionDate || '';
              break;
            case 'invoice_currency':
              value = invoice.invoiceCurrency || 'EGP';
              break;
            case 'invoice_value':
              value = invoice.invoiceValue || '';
              break;
            case 'vat_amount':
              value = invoice.vatAmount || '';
              break;
            case 'tax_discount':
              value = invoice.taxDiscount || '0.00';
              break;
            case 'total_invoice':
              value = invoice.totalInvoice || '';
              break;
            case 'internal_number':
              value = invoice.internalNumber || '';
              break;
            case 'electronic_number':
              value = invoice.electronicNumber || '';
              break;
            case 'seller_tax_number':
              value = invoice.sellerTaxNumber || '';
              break;
            case 'seller_name':
              value = invoice.sellerName || '';
              break;
            case 'seller_address':
              value = invoice.sellerAddress || '';
              break;
            case 'buyer_tax_number':
              value = invoice.buyerTaxNumber || '';
              break;
            case 'buyer_name':
              value = invoice.buyerName || '';
              break;
            case 'buyer_address':
              value = invoice.buyerAddress || '';
              break;
            case 'purchase_order_ref':
              value = invoice.purchaseOrderRef || '';
              break;
            case 'purchase_order_desc':
              value = invoice.purchaseOrderDesc || '';
              break;
            case 'sales_order_ref':
              value = invoice.salesOrderRef || '';
              break;
            case 'electronic_signature':
              value = invoice.electronicSignature || 'موقع إلكترونياً';
              break;
            case 'food_drug_guide':
              value = invoice.foodDrugGuide || '';
              break;
            case 'external_link':
              value = invoice.externalLink || '';
              break;
          }
          
          row.push(value);
        }
      });
      
      rows.push(row);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    
    // Format the worksheet
    this.formatWorksheet(ws, headers.length, data.length);
    
    XLSX.utils.book_append_sheet(wb, ws, 'فواتير مصلحة الضرائب');
    
    // Add statistics sheet
    this.addStatisticsSheet(wb, data);
    
    // Generate filename
    const timestamp = new Date().toISOString().split('T')[0];
    const mode = this.getSelectedMode();
    let modeText = '';
    
    if (mode === 'current') {
      modeText = `Page${this.currentPage}`;
    } else if (mode === 'all') {
      modeText = 'AllPages';
    } else if (mode === 'range') {
      const startPage = this.elements.startPage.value;
      const endPage = this.elements.endPage.value;
      modeText = `Pages${startPage}-${endPage}`;
    }
    
    const filename = `ETA_Invoices_${modeText}_${timestamp}.xlsx`;
    
    XLSX.writeFile(wb, filename);
  }
  
  formatWorksheet(ws, headerCount, dataCount) {
    // Set column widths
    const colWidths = Array(headerCount).fill({ wch: 20 });
    ws['!cols'] = colWidths;
    
    // Style header row
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let col = 0; col < headerCount; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      
      ws[cellAddress].s = {
        font: { bold: true, color: { rgb: "FFFFFF" }, size: 12 },
        fill: { fgColor: { rgb: "1E3C72" } },
        alignment: { horizontal: "center", vertical: "center" },
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        }
      };
    }
    
    // Style data rows
    for (let row = 1; row <= dataCount; row++) {
      const isEvenRow = row % 2 === 0;
      const fillColor = isEvenRow ? "F8F9FA" : "FFFFFF";
      
      for (let col = 0; col < headerCount; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;
        
        ws[cellAddress].s = {
          fill: { fgColor: { rgb: fillColor } },
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "thin", color: { rgb: "E0E0E0" } },
            bottom: { style: "thin", color: { rgb: "E0E0E0" } },
            left: { style: "thin", color: { rgb: "E0E0E0" } },
            right: { style: "thin", color: { rgb: "E0E0E0" } }
          }
        };
      }
    }
  }
  
  addStatisticsSheet(wb, data) {
    const stats = this.calculateStatistics(data);
    
    const statsData = [
      ['إحصائيات الفواتير المصدرة', ''],
      ['', ''],
      ['تاريخ التصدير', new Date().toLocaleString('ar-EG')],
      ['نمط التصدير', this.getModeDescription()],
      ['عدد الفواتير المصدرة', data.length],
      ['إجمالي قيمة الفواتير', this.formatCurrency(stats.totalValue) + ' EGP'],
      ['إجمالي ضريبة القيمة المضافة', this.formatCurrency(stats.totalVAT) + ' EGP'],
      ['متوسط قيمة الفاتورة', this.formatCurrency(stats.averageValue) + ' EGP'],
      ['', ''],
      ['إحصائيات حسب الحالة', ''],
      ...Object.entries(stats.statusCounts).map(([status, count]) => [status, count]),
      ['', ''],
      ['معلومات إضافية', ''],
      ['النطاق الزمني', stats.dateRange],
      ['عدد البائعين المختلفين', stats.uniqueSellers],
      ['عدد المشترين المختلفين', stats.uniqueBuyers]
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(statsData);
    ws['!cols'] = [{ wch: 30 }, { wch: 20 }];
    
    XLSX.utils.book_append_sheet(wb, ws, 'الإحصائيات');
  }
  
  getModeDescription() {
    const mode = this.getSelectedMode();
    if (mode === 'current') {
      return `الصفحة الحالية (${this.currentPage})`;
    } else if (mode === 'all') {
      return `جميع الصفحات (1-${this.totalPages})`;
    } else if (mode === 'range') {
      const startPage = this.elements.startPage.value;
      const endPage = this.elements.endPage.value;
      return `نطاق الصفحات (${startPage}-${endPage})`;
    }
    return 'غير محدد';
  }
  
  calculateStatistics(data) {
    const stats = {
      totalValue: 0,
      totalVAT: 0,
      averageValue: 0,
      statusCounts: {},
      dateRange: '',
      uniqueSellers: 0,
      uniqueBuyers: 0
    };
    
    const sellers = new Set();
    const buyers = new Set();
    const dates = [];
    
    data.forEach(invoice => {
      // Calculate totals
      const value = parseFloat(invoice.invoiceValue?.replace(/[,٬]/g, '') || 0);
      const vat = parseFloat(invoice.vatAmount?.replace(/[,٬]/g, '') || 0);
      
      stats.totalValue += value;
      stats.totalVAT += vat;
      
      // Count status
      const status = invoice.status || 'غير محدد';
      stats.statusCounts[status] = (stats.statusCounts[status] || 0) + 1;
      
      // Track unique sellers and buyers
      if (invoice.sellerName) sellers.add(invoice.sellerName);
      if (invoice.buyerName) buyers.add(invoice.buyerName);
      
      // Track dates
      if (invoice.issueDate) dates.push(new Date(invoice.issueDate));
    });
    
    stats.averageValue = data.length > 0 ? stats.totalValue / data.length : 0;
    stats.uniqueSellers = sellers.size;
    stats.uniqueBuyers = buyers.size;
    
    // Calculate date range
    if (dates.length > 0) {
      const minDate = new Date(Math.min(...dates));
      const maxDate = new Date(Math.max(...dates));
      stats.dateRange = `${minDate.toLocaleDateString('ar-EG')} - ${maxDate.toLocaleDateString('ar-EG')}`;
    }
    
    return stats;
  }
  
  formatCurrency(value) {
    if (!value || value === 0) return '0';
    return parseFloat(value).toLocaleString('ar-EG', { 
      minimumFractionDigits: 2, 
      maximumFractionDigits: 2 
    });
  }
  
  generateJSONFile(data, selectedFields) {
    const exportData = {
      metadata: {
        exportDate: new Date().toISOString(),
        exportMode: this.getModeDescription(),
        totalRecords: data.length,
        selectedFields: Object.keys(selectedFields).filter(key => selectedFields[key]),
        source: 'ETA Invoice Exporter Enhanced',
        statistics: this.calculateStatistics(data)
      },
      invoices: data
    };
    
    const blob = new Blob([JSON.stringify(exportData, null, 2)], { 
      type: 'application/json;charset=utf-8' 
    });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    
    const timestamp = new Date().toISOString().split('T')[0];
    const mode = this.getSelectedMode();
    let modeText = '';
    
    if (mode === 'current') {
      modeText = `Page${this.currentPage}`;
    } else if (mode === 'all') {
      modeText = 'AllPages';
    } else if (mode === 'range') {
      const startPage = this.elements.startPage.value;
      const endPage = this.elements.endPage.value;
      modeText = `Pages${startPage}-${endPage}`;
    }
    
    a.download = `ETA_Invoices_${modeText}_${timestamp}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }
  
  showStatus(message, type = '') {
    this.elements.status.textContent = message;
    this.elements.status.className = `status ${type}`;
    this.elements.status.style.display = 'flex';
    
    if (type === 'loading') {
      this.elements.status.innerHTML = `
        <span class="loading-spinner"></span>
        ${message}
      `;
    }
    
    if (type === 'success' || type === 'error') {
      setTimeout(() => {
        if (!this.isProcessing) {
          this.elements.status.style.display = 'none';
        }
      }, 5000);
    }
  }
  
  disableButtons() {
    this.elements.excelBtn.disabled = true;
    this.elements.jsonBtn.disabled = true;
  }
  
  enableButtons() {
    this.elements.excelBtn.disabled = false;
    this.elements.jsonBtn.disabled = false;
  }
  
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// Initialize the enhanced exporter when the DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  new ETAInvoiceExporter();
});

// Listen for progress updates from content script
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'progressUpdate') {
    // Handle progress updates if needed
    console.log('Progress update:', message.progress);
  }
});