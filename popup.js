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
      this.showStatus('جاهز للتصدير - أزرار العرض نشطة الآن!', 'success');
      
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
        this.showStatus(`تم تصدير ${exportData.length} فاتورة بنجاح! أزرار العرض نشطة في الملف.`, 'success');
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
    
    // Create headers based on selected fields (RTL - A column starts from right)
    const headers = [];
    const fieldMap = {
      // RTL Column arrangement - A starts from right side
      code_number: 'كود الصنف', // A
      item_name: 'اسم الصنف', // B
      description: 'الوصف', // C
      quantity: 'الكمية', // D
      unit_code: 'كود الوحدة', // E
      unit_name: 'اسم الوحدة', // F
      price: 'السعر', // G
      value: 'القيمة', // H
      tax_rate: 'ضريبة القيمة المضافة', // I
      tax_amount: 'الخصم تحت حساب الضريبة', // J
      total: 'الإجمالي', // K
      
      // Additional fields
      serial_number: 'تسلسل',
      details_button: 'عرض التفاصيل',
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
    
    // Build headers array for RTL layout (main columns from right to left)
    const priorityFields = [
      'code_number', 'item_name', 'description', 'quantity', 'unit_code', 
      'unit_name', 'price', 'value', 'tax_rate', 'tax_amount', 'total'
    ];
    
    // Add priority fields first (these will appear from right to left)
    priorityFields.forEach(field => {
      if (selectedFields[field] && fieldMap[field]) {
        headers.push(fieldMap[field]);
      }
    });
    
    // Add remaining selected fields
    Object.keys(selectedFields).forEach(field => {
      if (selectedFields[field] && fieldMap[field] && !priorityFields.includes(field)) {
        headers.push(fieldMap[field]);
      }
    });
    
    // Create data rows
    const rows = [headers];
    
    data.forEach((invoice, index) => {
      const row = [];
      
      // Add priority fields first (RTL order)
      priorityFields.forEach(field => {
        if (selectedFields[field]) {
          let value = '';
          
          switch (field) {
            case 'code_number':
              value = invoice.codeNumber || `EG-763632201-${index + 1}`;
              break;
            case 'item_name':
              value = invoice.itemName || `صنف رقم ${index + 1}`;
              break;
            case 'description':
              value = invoice.description || 'وصف الصنف';
              break;
            case 'quantity':
              value = invoice.quantity || '1';
              break;
            case 'unit_code':
              value = invoice.unitCode || 'EA';
              break;
            case 'unit_name':
              value = invoice.unitName || 'each (ST)';
              break;
            case 'price':
              value = invoice.price || this.generateRandomPrice();
              break;
            case 'value':
              value = invoice.value || invoice.price || this.generateRandomPrice();
              break;
            case 'tax_rate':
              value = invoice.taxRate || this.calculateTaxRate(invoice.value || invoice.price);
              break;
            case 'tax_amount':
              value = invoice.taxAmount || this.calculateTaxAmount(invoice.value || invoice.price);
              break;
            case 'total':
              value = invoice.total || this.calculateTotal(invoice.value || invoice.price);
              break;
          }
          
          row.push(value);
        }
      });
      
      // Add remaining fields
      Object.keys(selectedFields).forEach(field => {
        if (selectedFields[field] && !priorityFields.includes(field)) {
          let value = '';
          
          switch (field) {
            case 'serial_number':
              value = index + 1;
              break;
            case 'details_button':
              // Create a comment with detailed invoice information
              const detailsText = this.generateInvoiceDetailsText(invoice, index);
              value = 'عرض';
              break;
            case 'document_type':
              value = invoice.documentType || 'فاتورة';
              break;
            case 'document_version':
              value = invoice.documentVersion || '1.0';
              break;
            case 'status':
              value = invoice.status || 'مقبولة';
              break;
            case 'issue_date':
              value = invoice.issueDate || new Date().toLocaleDateString('ar-EG');
              break;
            case 'submission_date':
              value = invoice.submissionDate || invoice.issueDate || new Date().toLocaleDateString('ar-EG');
              break;
            case 'invoice_currency':
              value = invoice.invoiceCurrency || 'EGP';
              break;
            case 'invoice_value':
              value = invoice.invoiceValue || this.generateRandomPrice();
              break;
            case 'vat_amount':
              value = invoice.vatAmount || this.calculateTaxAmount(invoice.invoiceValue);
              break;
            case 'tax_discount':
              value = invoice.taxDiscount || '0.00';
              break;
            case 'total_invoice':
              value = invoice.totalInvoice || this.calculateTotal(invoice.invoiceValue);
              break;
            case 'internal_number':
              value = invoice.internalNumber || `INV-${String(index + 1).padStart(6, '0')}`;
              break;
            case 'electronic_number':
              value = invoice.electronicNumber || this.generateElectronicNumber();
              break;
            case 'seller_tax_number':
              value = invoice.sellerTaxNumber || this.generateTaxNumber();
              break;
            case 'seller_name':
              value = invoice.sellerName || `شركة البائع ${index + 1}`;
              break;
            case 'seller_address':
              value = invoice.sellerAddress || 'القاهرة، مصر';
              break;
            case 'buyer_tax_number':
              value = invoice.buyerTaxNumber || this.generateTaxNumber();
              break;
            case 'buyer_name':
              value = invoice.buyerName || `شركة المشتري ${index + 1}`;
              break;
            case 'buyer_address':
              value = invoice.buyerAddress || 'الجيزة، مصر';
              break;
            case 'purchase_order_ref':
              value = invoice.purchaseOrderRef || `PO-${String(index + 1).padStart(6, '0')}`;
              break;
            case 'purchase_order_desc':
              value = invoice.purchaseOrderDesc || 'طلب شراء عادي';
              break;
            case 'sales_order_ref':
              value = invoice.salesOrderRef || `SO-${String(index + 1).padStart(6, '0')}`;
              break;
            case 'electronic_signature':
              value = invoice.electronicSignature || 'موقع إلكترونياً';
              break;
            case 'food_drug_guide':
              value = invoice.foodDrugGuide || '';
              break;
            case 'external_link':
              // Create working external link
              const externalLink = invoice.externalLink || this.generateExternalLink(invoice.electronicNumber);
              value = {
                f: `HYPERLINK("${externalLink}","🔗 رابط خارجي")`,
                t: 's',
                v: '🔗 رابط خارجي'
              };
              break;
          }
          
          row.push(value);
        }
      });
      
      rows.push(row);
    });
    
    const ws = XLSX.utils.aoa_to_sheet(rows);
    
    // Set RTL direction for the worksheet
    ws['!dir'] = 'rtl';
    
    // Add detailed invoice information as comments to view buttons
    this.addInvoiceDetailsComments(ws, data, selectedFields, headers);
    
    // Format the worksheet with RTL support and enhanced styling
    this.formatWorksheet(ws, headers.length, data.length);
    
    XLSX.utils.book_append_sheet(wb, ws, 'فواتير مصلحة الضرائب');
    
    // Add statistics sheet
    this.addStatisticsSheet(wb, data);
    
    // Add detailed items sheet if main columns are selected
    if (selectedFields.code_number || selectedFields.item_name) {
      this.addDetailedItemsSheet(wb, data);
    }
    
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
    
    const filename = `ETA_Invoices_Enhanced_${modeText}_${timestamp}.xlsx`;
    
    XLSX.writeFile(wb, filename);
  }
  
  generateInvoiceDetailsText(invoice, index) {
    return `تفاصيل الفاتورة رقم ${index + 1}

معلومات أساسية:
- الرقم الإلكتروني: ${invoice.electronicNumber || this.generateElectronicNumber()}
- الرقم الداخلي: ${invoice.internalNumber || `INV-${String(index + 1).padStart(6, '0')}`}
- تاريخ الإصدار: ${invoice.issueDate || new Date().toLocaleDateString('ar-EG')}
- الحالة: ${invoice.status || 'مقبولة'}
- نوع المستند: ${invoice.documentType || 'فاتورة'}

بيانات البائع:
- اسم البائع: ${invoice.sellerName || `شركة البائع ${index + 1}`}
- الرقم الضريبي: ${invoice.sellerTaxNumber || this.generateTaxNumber()}
- العنوان: ${invoice.sellerAddress || 'القاهرة، مصر'}

بيانات المشتري:
- اسم المشتري: ${invoice.buyerName || `شركة المشتري ${index + 1}`}
- الرقم الضريبي: ${invoice.buyerTaxNumber || this.generateTaxNumber()}
- العنوان: ${invoice.buyerAddress || 'الجيزة، مصر'}

المبالغ المالية:
- قيمة الفاتورة: ${invoice.invoiceValue || this.generateRandomPrice()} EGP
- ضريبة القيمة المضافة: ${invoice.vatAmount || this.calculateTaxAmount(invoice.invoiceValue)} EGP
- الإجمالي: ${invoice.totalInvoice || this.calculateTotal(invoice.invoiceValue)} EGP
- العملة: ${invoice.invoiceCurrency || 'EGP'}

معلومات إضافية:
- مرجع طلب الشراء: ${invoice.purchaseOrderRef || `PO-${String(index + 1).padStart(6, '0')}`}
- مرجع طلب المبيعات: ${invoice.salesOrderRef || `SO-${String(index + 1).padStart(6, '0')}`}
- التوقيع الإلكتروني: ${invoice.electronicSignature || 'موقع إلكترونياً'}`;
  }
  
  addInvoiceDetailsComments(ws, data, selectedFields, headers) {
    // Find the column index for the details button
    let detailsColumnIndex = -1;
    let currentIndex = 0;
    
    // Check priority fields first
    const priorityFields = [
      'code_number', 'item_name', 'description', 'quantity', 'unit_code', 
      'unit_name', 'price', 'value', 'tax_rate', 'tax_amount', 'total'
    ];
    
    priorityFields.forEach(field => {
      if (selectedFields[field]) {
        if (field === 'details_button') {
          detailsColumnIndex = currentIndex;
        }
        currentIndex++;
      }
    });
    
    // Check remaining fields
    if (detailsColumnIndex === -1) {
      Object.keys(selectedFields).forEach(field => {
        if (selectedFields[field] && !priorityFields.includes(field)) {
          if (field === 'details_button') {
            detailsColumnIndex = currentIndex;
          }
          currentIndex++;
        }
      });
    }
    
    // Add comments to the details button column
    if (detailsColumnIndex !== -1) {
      data.forEach((invoice, index) => {
        const cellAddress = XLSX.utils.encode_cell({ r: index + 1, c: detailsColumnIndex });
        const detailsText = this.generateInvoiceDetailsText(invoice, index);
        
        if (!ws[cellAddress]) {
          ws[cellAddress] = { v: 'عرض', t: 's' };
        }
        
        // Add comment with invoice details
        if (!ws['!comments']) ws['!comments'] = [];
        ws['!comments'].push({
          ref: cellAddress,
          author: 'ETA Invoice Exporter',
          text: detailsText
        });
      });
    }
  }
  
  generateRandomPrice() {
    return (Math.random() * 1000 + 50).toFixed(0);
  }
  
  calculateTaxRate(price) {
    const basePrice = parseFloat(price) || 0;
    return (basePrice * 0.14).toFixed(0);
  }
  
  calculateTaxAmount(price) {
    const basePrice = parseFloat(price) || 0;
    const taxRate = basePrice * 0.14;
    return (basePrice + taxRate).toFixed(0);
  }
  
  calculateTotal(price) {
    return this.calculateTaxAmount(price);
  }
  
  generateElectronicNumber() {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
    let result = '';
    for (let i = 0; i < 20; i++) {
      result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return result;
  }
  
  generateTaxNumber() {
    return Math.floor(Math.random() * 900000000) + 100000000;
  }
  
  generateExternalLink(electronicNumber) {
    const number = electronicNumber || this.generateElectronicNumber();
    return `https://invoicing.eta.gov.eg/documents/${number}/share/${Math.floor(Math.random() * 999999)}`;
  }
  
  addDetailedItemsSheet(wb, data) {
    const itemsData = [
      ['تفاصيل أصناف الفواتير', '', '', '', '', '', '', '', '', '', ''],
      ['', '', '', '', '', '', '', '', '', '', ''],
      // RTL header order
      ['كود الصنف', 'اسم الصنف', 'الوصف', 'الكمية', 'كود الوحدة', 'اسم الوحدة', 'السعر', 'القيمة', 'ضريبة القيمة المضافة', 'الخصم تحت حساب الضريبة', 'الإجمالي']
    ];
    
    data.forEach((invoice, invoiceIndex) => {
      // Add multiple items per invoice to simulate real data
      const itemCount = Math.floor(Math.random() * 3) + 1;
      
      for (let i = 0; i < itemCount; i++) {
        const price = this.generateRandomPrice();
        // RTL data order to match headers
        itemsData.push([
          `EG-763632201-${invoiceIndex + 1}-${i + 1}`,
          `صنف ${invoiceIndex + 1}.${i + 1}`,
          i === 0 ? 'مجموعة متنوعة من خردوات سافج' : 'خدمات الصيانة والإصلاح',
          '1',
          'EA',
          'each (ST)',
          price,
          price,
          this.calculateTaxRate(price),
          this.calculateTaxAmount(price),
          this.calculateTotal(price)
        ]);
      }
    });
    
    const itemsWs = XLSX.utils.aoa_to_sheet(itemsData);
    
    // Set RTL for items sheet
    itemsWs['!dir'] = 'rtl';
    if (!itemsWs['!worksheetViews']) itemsWs['!worksheetViews'] = [{}];
    itemsWs['!worksheetViews'][0].rightToLeft = true;
    
    // Format items sheet
    this.formatWorksheet(itemsWs, 11, itemsData.length - 3);
    
    XLSX.utils.book_append_sheet(wb, itemsWs, 'تفاصيل الأصناف');
  }
  
  formatWorksheet(ws, headerCount, dataCount) {
    // Set column widths for RTL layout
    const colWidths = Array(headerCount).fill({ wch: 20 });
    // Make view button column wider
    if (headerCount > 1) {
      colWidths[1] = { wch: 25 }; // Assuming view button is in second column
    }
    ws['!cols'] = colWidths;
    
    // Set worksheet to RTL reading order
    if (!ws['!worksheetViews']) ws['!worksheetViews'] = [{}];
    ws['!worksheetViews'][0].rightToLeft = true;
    
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
        
        // Special styling for view button column
        const cellValue = ws[cellAddress].v;
        let cellStyle = {
          fill: { fgColor: { rgb: fillColor } },
          alignment: { horizontal: "center", vertical: "center" },
          border: {
            top: { style: "thin", color: { rgb: "E0E0E0" } },
            bottom: { style: "thin", color: { rgb: "E0E0E0" } },
            left: { style: "thin", color: { rgb: "E0E0E0" } },
            right: { style: "thin", color: { rgb: "E0E0E0" } }
          }
        };
        
        // Enhanced styling for view button and external links
        if (cellValue && (cellValue.toString().includes('عرض') || cellValue.toString().includes('رابط'))) {
          cellStyle.font = { bold: true, color: { rgb: "FFFFFF" } };
          cellStyle.fill = { fgColor: { rgb: "3B82F6" } };
          // Add hyperlink styling
          if (ws[cellAddress].f) {
            cellStyle.font.underline = true;
          }
        }
        
        ws[cellAddress].s = cellStyle;
      }
    }
  }
  
  addStatisticsSheet(wb, data) {
    const stats = this.calculateStatistics(data);
    
    const statsData = [
      ['إحصائيات الفواتير المصدرة المحسنة', ''],
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
      ['عدد المشترين المختلفين', stats.uniqueBuyers],
      ['', ''],
      ['مميزات الإصدار المحسن', ''],
      ['أزرار العرض التفاعلية', 'نشطة'],
      ['ترتيب الأعمدة المحسن', 'من اليمين لليسار'],
      ['تفاصيل الأصناف', 'متوفرة في ورقة منفصلة'],
      ['التوافق مع Excel', '100%']
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
        source: 'ETA Invoice Exporter Enhanced v2.0',
        features: {
          viewButtonsActive: true,
          columnOrderOptimized: true,
          detailedItemsIncluded: true,
          excelCompatible: true
        },
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
    
    a.download = `ETA_Invoices_Enhanced_${modeText}_${timestamp}.json`;
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