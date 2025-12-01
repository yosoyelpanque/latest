document.addEventListener('DOMContentLoaded', () => {
    
    let state = {
        inventoryEditMode: false,
        loggedIn: false, currentUser: null, inventory: [], additionalItems: [],
        resguardantes: [], activeResguardante: null, locations: {}, areas: [], areaNames: {},
        lastAutosave: null, sessionStartTime: null, additionalPhotos: {}, locationPhotos: {},
        notes: {}, photos: {}, theme: 'light',
        inventoryFinished: false,
        areaDirectory: {},
        closedAreas: {},
        completedAreas: {}, 
        persistentAreas: [],
        serialNumberCache: new Set(),
        cameraStream: null,
        readOnlyMode: false,
        activityLog: [],
        institutionalReportCheckboxes: {},
        actionCheckboxes: {
            labels: {},
            notes: {},
            additional: {},
            mismatched: {},
            personal: {}
        },
        reportCheckboxes: {
            notes: {},
            mismatched: {}
        },
        mapLayout: { 'page1': {} }, 
        currentLayoutPage: 'page1',
        layoutPageNames: { 'page1': 'Página 1' },
        layoutImages: {},
        layoutPageColors: { 'page1': '#ffffff' }, 
        layoutItemColors: {} 
    };
    
    let logoClickCount = 0;
    // Variables temporales para gestión de ubicaciones (Creación y Edición)
    let tempUserLocations = []; 
    let tempEditUserLocations = []; // NUEVO: Para el modal de edición

    function generateUUID() {
        if (crypto && crypto.randomUUID) {
            return crypto.randomUUID();
        }
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
            var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
            return v.toString(16);
        });
    }
    // --- CÓDIGO NUEVO DE SEGURIDAD ---
    function escapeHTML(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }
     const verifiers = {
        '41290': 'BENÍTEZ HERNÁNDEZ MARIO',
        '41292': 'ESCAMILLA VILLEGAS BRYAN ANTONY',
        '41282': 'LÓPEZ QUINTANA ALDO',
        '41287': 'MARIN ESPINOSA MIGUEL',
        '41289': 'SANCHEZ ARELLANES RICARDO',
        '41293': 'EDSON OSNAR TORRES JIMENEZ',
        '15990': 'CHÁVEZ SÁNCHEZ ALFONSO',
        '17326': 'DOMÍNGUEZ VAZQUEZ FRANCISCO JAVIER',
        '11885': 'ESTRADA HERNÁNDEZ ROBERTO',
        '19328': 'LÓPEZ ESTRADA LEOPOLDO',
        '44925': 'MENDOZA SOLARES JOSE JUAN',
        '16990': 'PÉREZ RODRÍGUEZ DANIEL',
        '16000': 'PÉREZ YAÑEZ JUAN JOSE',
        '17812': 'RODRÍGUEZ RAMÍREZ RENE',
        '44095': 'LOPEZ JIMENEZ ALAN GABRIEL',
        '2875': 'VIZCAINO ROJAS ALVARO'
    };
    
    const photoDB = {
        db: null,
        init: function() {
            return new Promise((resolve, reject) => {
                const request = indexedDB.open('InventarioProPhotosDB', 2); 
                request.onupgradeneeded = (event) => {
                    const db = event.target.result;
                    if (!db.objectStoreNames.contains('photos')) db.createObjectStore('photos');
                    if (!db.objectStoreNames.contains('layoutImages')) db.createObjectStore('layoutImages');
                };
                request.onsuccess = (event) => { this.db = event.target.result; resolve(); };
                request.onerror = (event) => { console.error('Error con IndexedDB:', event.target.errorCode); reject(event.target.errorCode); };
            });
        },
        setItem: function(storeName, key, value) {
            return new Promise((resolve, reject) => {
                if (!this.db) return reject('DB not initialized');
                const transaction = this.db.transaction([storeName], 'readwrite');
                const store = transaction.objectStore(storeName);
                const request = store.put(value, key);
                request.onsuccess = () => resolve();
                request.onerror = (event) => reject(event.target.error);
            });
        },
        getItem: function(storeName, key) {
            return new Promise((resolve, reject) => {
                if (!this.db) return reject('DB not initialized');
                const transaction = this.db.transaction([storeName], 'readonly');
                const store = transaction.objectStore(storeName);
                const request = store.get(key);
                request.onsuccess = () => resolve(request.result);
                request.onerror = (event) => reject(event.target.error);
            });
        },
        deleteItem: function(storeName, key) {
            return new Promise((resolve, reject) => {
                if (!this.db) return reject('DB not initialized');
                const transaction = this.db.transaction([storeName], 'readwrite');
                const store = transaction.objectStore(storeName);
                const request = store.delete(key);
                request.onsuccess = () => resolve();
                request.onerror = (event) => reject(event.target.error);
            });
        },
        getAllItems: function(storeName) {
            return new Promise((resolve, reject) => {
                if (!this.db) return reject('DB not initialized');
                const transaction = this.db.transaction([storeName], 'readonly');
                const store = transaction.objectStore(storeName);
                const keysRequest = store.getAllKeys();
                const valuesRequest = store.getAll();

                Promise.all([
                    new Promise((res, rej) => { keysRequest.onsuccess = () => res(keysRequest.result); keysRequest.onerror = (e) => rej(e.target.error); }),
                    new Promise((res, rej) => { valuesRequest.onsuccess = () => res(valuesRequest.result); valuesRequest.onerror = (e) => rej(e.target.error); })
                ]).then(([keys, values]) => {
                    const result = keys.map((key, index) => ({ key, value: values[index] }));
                    resolve(result);
                }).catch(reject);
            });
        }
    };

    const elements = {
        loginPage: document.getElementById('login-page'), mainApp: document.getElementById('main-app'),
        employeeNumberInput: document.getElementById('employee-number-input'),
        employeeLoginBtn: document.getElementById('employee-login-btn'),
        clearSessionLink: document.getElementById('clear-session-link'),
        currentUserDisplay: document.getElementById('current-user-name'),
        fileInput: document.getElementById('file-input'),
        uploadBtn: document.getElementById('upload-btn'), logoutBtn: document.getElementById('logout-btn'),
        dashboard: {
            headerAndDashboard: document.getElementById('header-and-dashboard'),
            toggleBtn: document.getElementById('dashboard-toggle-btn'),
            dailyProgressCard: document.getElementById('daily-progress-card'),
            progressTooltip: document.getElementById('progress-tooltip'),
        },
        totalItemsEl: document.getElementById('total-items'), locatedItemsEl: document.getElementById('located-items'),
        pendingItemsEl: document.getElementById('pending-items'), dailyProgressEl: document.getElementById('daily-progress'),
        workingAreasCountEl: document.getElementById('working-areas-count'),
        additionalItemsCountEl: document.getElementById('additional-items-count'),
        tabsContainer: document.getElementById('tabs-container'), tabContents: document.querySelectorAll('.tab-content'),
        mainContentArea: document.getElementById('main-content-area'),
        logo: {
            container: document.getElementById('logo-container'),
            img: document.getElementById('logo-img'),
            title: document.querySelector('#main-app header div:nth-child(1) > div:nth-child(2) > h2')
        },
        activeUserBanner: {
            banner: document.getElementById('active-user-banner'),
            name: document.getElementById('active-user-banner-name'),
            area: document.getElementById('active-user-banner-area'), 
            deactivateBtn: document.getElementById('deactivate-user-btn')
        },
        userForm: {
            name: document.getElementById('user-name'), locationSelect: document.getElementById('user-location-select'),
            locationManual: document.getElementById('user-location-manual'), areaSelect: document.getElementById('user-area-select'),
            createBtn: document.getElementById('create-user-btn'), list: document.getElementById('registered-users-list'),
            // Elementos de creación multi-ubicación
            addLocationBtn: document.getElementById('add-location-btn'),
            locationsList: document.getElementById('new-user-locations-list')
        },
        // ACTUALIZADO: Elementos del modal de edición de usuario con soporte multi-ubicación
        editUserModal: {
            modal: document.getElementById('edit-user-modal'),
            name: document.getElementById('edit-user-name'),
            locationType: document.getElementById('edit-user-location-type'),
            locationManual: document.getElementById('edit-user-location-manual'),
            addLocationBtn: document.getElementById('edit-add-location-btn'),
            locationsList: document.getElementById('edit-user-locations-list'),
            areaSelect: document.getElementById('edit-user-area'),
            saveBtn: document.getElementById('edit-user-save-btn'),
            cancelBtn: document.getElementById('edit-user-cancel-btn')
        },
        inventory: {
            tableBody: document.getElementById('inventory-table-body'),
            searchInput: document.getElementById('search-input'), qrScanBtn: document.getElementById('qr-scan-btn'),
            clearSearchBtn: document.getElementById('clear-search-btn'), ubicadoBtn: document.getElementById('ubicado-btn'),
            reEtiquetarBtn: document.getElementById('re-etiquetar-btn'),
            desubicarBtn: document.getElementById('desubicar-btn'),
            addNoteBtn: document.getElementById('add-note-btn'),
            prevPageBtn: document.getElementById('prev-page-btn'),
            nextPageBtn: document.getElementById('next-page-btn'), pageInfo: document.getElementById('page-info'),
            statusFilter: document.getElementById('status-filter'), areaFilter: document.getElementById('area-filter-inventory'),
            bookTypeFilter: document.getElementById('book-type-filter'),
            selectAllCheckbox: document.getElementById('select-all-checkbox')
        },
        adicionales: {
            form: document.getElementById('adicional-form'),
            addBtn: document.getElementById('add-adicional-btn'), list: document.getElementById('adicionales-list'),
            areaFilter: document.getElementById('ad-area-filter'),
            userFilter: document.getElementById('ad-user-filter'),
            printResguardoBtn: document.getElementById('print-adicionales-resguardo-btn'),
            total: document.getElementById('additional-items-total')
        },
        reports: {
            areaProgressContainer: document.getElementById('area-progress-container'),
            stats: document.getElementById('general-stats'), 
            areaFilter: document.getElementById('report-area-filter'),
            userFilter: document.getElementById('report-user-filter'),
            reportButtons: document.querySelectorAll('.report-btn'),
            exportLabelsXlsxBtn: document.getElementById('export-labels-xlsx-btn'),
            exportXlsxBtn: document.getElementById('export-xlsx-btn'),
            reportViewModal: {
                modal: document.getElementById('report-view-modal'),
                title: document.getElementById('report-view-title'),
                closeBtn: document.getElementById('report-view-close-btn'),
                closeFooterBtn: document.getElementById('report-view-close-footer-btn'),
                content: document.getElementById('report-view-content'),
                tableHead: document.getElementById('report-view-table-head'),
                tableBody: document.getElementById('report-view-table-body')
            }
        },
        settings: {
            themes: document.querySelectorAll('[data-theme]'), autosaveInterval: document.getElementById('autosave-interval'),
            loadedListsContainer: document.getElementById('loaded-lists-container'),
            exportSessionBtn: document.getElementById('export-session-btn'),
            importSessionBtn: document.getElementById('import-session-btn'),
            importFileInput: document.getElementById('import-file-input'),
            finalizeInventoryBtn: document.getElementById('finalize-inventory-btn'),
            summaryAuthor: document.getElementById('summary-author'),
            summaryAreaResponsible: document.getElementById('summary-area-responsible'),
            summaryLocation: document.getElementById('summary-location'),
            directoryContainer: document.getElementById('directory-container'),
            directoryCount: document.getElementById('directory-count'),
            aboutHeader: document.getElementById('about-header'),
            aboutContent: document.getElementById('about-content'),
            importPhotosBtn: document.getElementById('import-photos-btn'),
            importPhotosInput: document.getElementById('import-photos-input'),
            restorePhotosBtn: document.getElementById('restore-photos-from-backup-btn'),
            restorePhotosInput: document.getElementById('restore-photos-input')
        },
        loadingOverlay: {
            overlay: document.getElementById('loading-overlay'),
            spinner: document.getElementById('loading-spinner'),
            text: document.getElementById('loading-text')
        },
        importProgress: {
            modal: document.getElementById('import-progress-modal'),
            text: document.getElementById('import-progress-text'),
            bar: document.getElementById('import-progress-bar')
        },
        confirmationModal: document.getElementById('confirmation-modal'), modalTitle: document.getElementById('modal-title'),
        modalText: document.getElementById('modal-text'), modalConfirmBtn: document.getElementById('modal-confirm'),
        modalCancelBtn: document.getElementById('modal-cancel'), toastContainer: document.getElementById('toast-container'),
        
        addAdicionalesConfirm: {
            modal: document.getElementById('add-adicionales-confirm-modal'),
            yesBtn: document.getElementById('add-adicionales-yes'),
            noBtn: document.getElementById('add-adicionales-no')
        },

        notesModal: document.getElementById('notes-modal'), noteTextarea: document.getElementById('note-textarea'),
        noteSaveBtn: document.getElementById('note-save-btn'), noteCancelBtn: document.getElementById('note-cancel-btn'),
        itemDetailsModal: {
            modal: document.getElementById('item-details-modal'),
            title: document.getElementById('item-details-title'),
            content: document.getElementById('item-details-content'),
            closeBtn: document.getElementById('item-details-close-btn')
        },
        qrDisplayModal: {
            modal: document.getElementById('qr-display-modal'),
            title: document.getElementById('qr-display-title'),
            container: document.getElementById('qr-code-display'),
            closeBtn: document.getElementById('qr-display-close-btn')
        },
        transferPhotoModal: {
            modal: document.getElementById('transfer-photo-modal'),
            title: document.getElementById('transfer-photo-title'),
            text: document.getElementById('transfer-photo-text'),
            preview: document.getElementById('transfer-photo-preview'),
            search: document.getElementById('transfer-photo-search'),
            select: document.getElementById('transfer-photo-select'),
            skipBtn: document.getElementById('transfer-photo-skip-btn'),
            cancelBtn: document.getElementById('transfer-photo-cancel-btn'),
            confirmBtn: document.getElementById('transfer-photo-confirm-btn')
        },
        formatoEntradaModal: {
            modal: document.getElementById('formato-entrada-modal'),
            siBtn: document.getElementById('formato-entrada-si'),
            noBtn: document.getElementById('formato-entrada-no')
        },
        editAdicionalModal: {
            modal: document.getElementById('edit-adicional-modal'),
            form: document.getElementById('edit-adicional-form'),
            saveBtn: document.getElementById('edit-adicional-save-btn'),
            cancelBtn: document.getElementById('edit-adicional-cancel-btn')
        },
        photo: {
            modal: document.getElementById('photo-modal'),
            title: document.getElementById('photo-modal-title'),
            input: document.getElementById('photo-input'),
            message: document.getElementById('photo-message'),
            closeBtn: document.getElementById('photo-close-btn'),
            viewContainer: document.getElementById('photo-view-container'),
            uploadContainer: document.getElementById('photo-upload-container'),
            img: document.getElementById('item-photo-img'),
            deleteBtn: document.getElementById('delete-photo-btn'),
            useCameraBtn: document.getElementById('use-camera-btn'),
            cameraViewContainer: document.getElementById('camera-view-container'),
            cameraStream: document.getElementById('camera-stream'),
            photoCanvas: document.getElementById('photo-canvas'),
            captureBtn: document.getElementById('capture-photo-btn'),
            switchToUploadBtn: document.getElementById('switch-to-upload-btn'),
            cameraSelect: document.getElementById('photo-camera-select')
        },
        qrScannerModal: document.getElementById('qr-scanner-modal'),
        qrReader: document.getElementById('qr-reader'), 
        qrScannerCloseBtn: document.getElementById('qr-scanner-close-btn'),
        qrCameraSelect: document.getElementById('qr-camera-select'),
        areaClosure: {
            modal: document.getElementById('area-closure-modal'),
            title: document.getElementById('area-closure-title'),
            responsibleInput: document.getElementById('area-closure-responsible'),
            locationInput: document.getElementById('area-closure-location'),
            confirmBtn: document.getElementById('area-closure-confirm-btn'),
            cancelBtn: document.getElementById('area-closure-cancel-btn')
        },
        reassignModal: {
            modal: document.getElementById('reassign-modal'),
            title: document.getElementById('reassign-title'),
            text: document.getElementById('reassign-text'),
            areaSelect: document.getElementById('reassign-area-select'),
            confirmBtn: document.getElementById('reassign-confirm-btn'),
            keepBtn: document.getElementById('reassign-keep-btn'),
            deleteAllBtn: document.getElementById('reassign-delete-all-btn'),
            cancelBtn: document.getElementById('reassign-cancel-btn'),
        },
        readOnlyOverlay: document.getElementById('read-only-mode-overlay'),
        log: {
            modal: document.getElementById('log-modal'),
            content: document.getElementById('log-content'),
            showBtn: document.getElementById('show-log-btn'),
            closeBtn: document.getElementById('log-close-btn')
        },
        detailView: {
            modal: document.getElementById('item-detail-view-modal'),
            title: document.getElementById('detail-view-title'),
            closeBtn: document.getElementById('detail-view-close-btn'),
            photoContainer: document.getElementById('detail-view-photo-container'),
            photo: document.getElementById('detail-view-photo'),
            noPhoto: document.getElementById('detail-view-no-photo'),
            clave: document.getElementById('detail-view-clave'),
            descripcion: document.getElementById('detail-view-descripcion'),
            marca: document.getElementById('detail-view-marca'),
            modelo: document.getElementById('detail-view-modelo'),
            serie: document.getElementById('detail-view-serie'),
            usuario: document.getElementById('detail-view-usuario'),
            ubicacionEspecifica: document.getElementById('detail-view-ubicacion-especifica'),
            area: document.getElementById('detail-view-area'),
            areaWarning: document.getElementById('detail-view-area-warning'),
            ubicarBtn: document.getElementById('detail-view-ubicar-btn'),
            reetiquetarBtn: document.getElementById('detail-view-reetiquetar-btn'),
            notaBtn: document.getElementById('detail-view-nota-btn'),
            fotoBtn: document.getElementById('detail-view-foto-btn')
        },
        userDetailView: {
            modal: document.getElementById('user-detail-view-modal'),
            title: document.getElementById('user-detail-view-title'),
            closeBtn: document.getElementById('user-detail-view-close-btn'),
            closeFooterBtn: document.getElementById('user-detail-view-close-footer-btn'),
            photoContainer: document.getElementById('user-detail-view-photo-container'),
            photo: document.getElementById('user-detail-view-photo'),
            noPhoto: document.getElementById('user-detail-view-no-photo'),
            name: document.getElementById('user-detail-view-name'),
            area: document.getElementById('user-detail-view-area'),
            location: document.getElementById('user-detail-view-location')
        },
        adicionalDetailView: {
            modal: document.getElementById('adicional-detail-view-modal'),
            title: document.getElementById('adicional-detail-view-title'),
            closeBtn: document.getElementById('adicional-detail-view-close-btn'),
            closeFooterBtn: document.getElementById('adicional-detail-view-close-footer-btn'),
            photoContainer: document.getElementById('adicional-detail-view-photo-container'),
            photo: document.getElementById('adicional-detail-view-photo'),
            noPhoto: document.getElementById('adicional-detail-view-no-photo'),
            descripcion: document.getElementById('adicional-detail-view-descripcion'),
            clave: document.getElementById('adicional-detail-view-clave'),
            claveAsignada: document.getElementById('adicional-detail-view-claveAsignada'),
            marca: document.getElementById('adicional-detail-view-marca'),
            modelo: document.getElementById('adicional-detail-view-modelo'),
            serie: document.getElementById('adicional-detail-view-serie'),
            area: document.getElementById('adicional-detail-view-area'),
            usuario: document.getElementById('adicional-detail-view-usuario'),
            ubicacionEspecifica: document.getElementById('adicional-detail-view-ubicacion-especifica'), // Esta es la línea crítica
            tipo: document.getElementById('adicional-detail-view-tipo')
        },
        preprintModal: {
            modal: document.getElementById('preprint-edit-modal'),
            title: document.getElementById('preprint-title'),
            fieldsContainer: document.getElementById('preprint-fields'),
            dateInput: document.getElementById('preprint-date'),
            confirmBtn: document.getElementById('preprint-confirm-btn'),
            cancelBtn: document.getElementById('preprint-cancel-btn')
        },
        layoutEditor: { 
            modal: document.getElementById('layout-editor-modal'),
            openBtn: document.getElementById('open-layout-editor-btn'),
            closeBtn: document.getElementById('layout-close-btn'),
            saveBtn: document.getElementById('layout-save-btn'),
            printBtn: document.getElementById('layout-print-btn'),
            sidebar: document.getElementById('layout-sidebar-locations'),
            toolsSidebar: document.getElementById('layout-tools-sidebar'),
            canvas: document.getElementById('layout-canvas'),
            canvasWrapper: document.getElementById('layout-canvas-wrapper'),
            pagePrev: document.getElementById('layout-page-prev'),
            pageNext: document.getElementById('layout-page-next'),
            pageAdd: document.getElementById('layout-page-add'),
            pageReset: document.getElementById('layout-page-reset'),
            pageRemove: document.getElementById('layout-page-remove'),
            pageName: document.getElementById('layout-page-name'),
            addImageBtn: document.getElementById('layout-add-image-btn'),
            imageInput: document.getElementById('layout-image-input')
        }, 
        // ... dentro de const elements = { ...
        batchModal: {
            modal: document.getElementById('batch-print-modal'),
            closeBtn: document.getElementById('batch-close-btn'),
            cancelBtn: document.getElementById('batch-cancel-btn'),
            generateBtn: document.getElementById('batch-generate-btn'),
            dateInput: document.getElementById('batch-date'),
            entregaInput: document.getElementById('batch-entrega'),
            cargoInput: document.getElementById('batch-cargo-entrega'),
            includeAdditionals: document.getElementById('batch-include-additionals'),
            areaNameDisplay: document.getElementById('batch-area-name'),
            usersListContainer: document.getElementById('batch-users-list'),
            selectAllBtn: document.getElementById('batch-select-all'),
            deselectAllBtn: document.getElementById('batch-deselect-all'),
            countDisplay: document.getElementById('batch-selected-count')
        },
        printContainer: document.getElementById('print-view-container'),
        printTemplates: {
            sessionSummary: document.getElementById('print-session-summary'),
            areaClosure: document.getElementById('print-area-closure'),
            resguardo: document.getElementById('print-resguardo'),
            simplePending: document.getElementById('print-simple-pending'),
            tasksReport: document.getElementById('print-tasks-report'),
            layout: document.getElementById('print-layout-view')
        }
    };
    let currentPage = 1;
    const itemsPerPage = 50; 
    let filteredItems = [];
    let html5QrCode;

    function preparePrint(activeTemplateId, options = {}) {
        const { date } = options;
        const dateToPrint = date || new Date().toLocaleDateString('es-MX'); 
        
        document.querySelectorAll('.print-page').forEach(page => {
            page.classList.remove('active');
        });

        const activeTemplate = document.getElementById(activeTemplateId);
        if (activeTemplate) {
            const dateElement = activeTemplate.querySelector('.print-header-date');
            if (dateElement) {
                if (dateElement.id.includes('date')) { 
                     dateElement.textContent = `Fecha: ${dateToPrint}`;
                } else { 
                    dateElement.textContent = dateToPrint;
                }
            }

            activeTemplate.classList.add('active');
            
            if (activeTemplateId === 'print-layout-view') {
                document.querySelectorAll('.print-page.layout-clone').forEach(clone => {
                    const cloneDateEl = clone.querySelector('.print-header-date');
                    if (cloneDateEl) cloneDateEl.textContent = `Fecha: ${dateToPrint}`;
                    clone.classList.add('active');
                });
            }
            
            window.print();
        } else {
            showToast('Error: No se encontró la plantilla de impresión.', 'error');
        }
    }
    
    // ACTUALIZADO: Soporte robusto para contar ubicaciones múltiples y nombres manuales
    function recalculateLocationCounts() {
state.locations = {};

state.resguardantes.forEach(user => {
    // Determinar qué ubicaciones procesar
    let locsToProcess = [];
    if (user.locations && Array.isArray(user.locations) && user.locations.length > 0) {
        locsToProcess = user.locations;
    } else if (user.locationWithId) {
        locsToProcess = [user.locationWithId];
    }

    locsToProcess.forEach(locFull => {
        if (!locFull) return;

        // MEJORA: Regex más robusta.
        // ^(.*?): Captura cualquier cosa al inicio (non-greedy)
        // \s*: Permite CERO o MÁS espacios antes del número (tolera "OFICINA1" o "OFICINA   1")
        // (\d+)$: Captura el número al final
        const baseMatch = locFull.match(/^(.*?)\s*(\d+)$/);
        
        let base;
        if (baseMatch) {
            // baseMatch[1] es el texto limpio antes del número
            base = baseMatch[1].trim().toUpperCase(); 
        } else {
            // Si no tiene número, usamos todo el texto limpio
            base = locFull.trim().toUpperCase();
        }

        // Incrementamos el contador para esa base normalizada
        state.locations[base] = (state.locations[base] || 0) + 1;
    });
});

console.log('Contadores de ubicación recalculados (Normalizados):', state.locations);
// No necesitamos logActivity aquí para no saturar el log, es una operación interna
}

    function logActivity(action, details = '') {
        const timestamp = new Date().toLocaleString('es-MX');
        const logEntry = `[${timestamp}] ${action}: ${details}`;
        
        state.activityLog.push(logEntry);

        // OPTIMIZACIÓN: Prevenir que el log llene la memoria localStorage
        // Si supera 500 entradas, borramos las más antiguas
        if (state.activityLog.length > 500) {
            // Mantiene solo los últimos 500
            state.activityLog = state.activityLog.slice(-500); 
        }
    }

    function handleModalNavigation(modalElement) {
        const focusableElements = modalElement.querySelectorAll('button, [href], input, select, textarea');
        if (focusableElements.length === 0) return;

        const firstElement = focusableElements[0];
        const lastElement = focusableElements[focusableElements.length - 1];
        
        firstElement.focus();

        const keydownHandler = (e) => {
            if (e.key === 'Tab') {
                if (e.shiftKey && document.activeElement === firstElement) {
                    lastElement.focus();
                    e.preventDefault();
                } else if (!e.shiftKey && document.activeElement === lastElement) {
                    firstElement.focus();
                    e.preventDefault();
                }
            } else if (e.key === 'Enter') {
                if (document.activeElement.tagName === 'TEXTAREA') return;

                const confirmBtn = modalElement.querySelector('#modal-confirm, #note-save-btn, #edit-adicional-save-btn, #edit-user-save-btn, #preprint-confirm-btn');
                // Evitar conflicto con botones de añadir ubicación
                if (confirmBtn && document.activeElement !== confirmBtn && document.activeElement.id !== 'add-location-btn' && document.activeElement.id !== 'edit-add-location-btn') {
                    e.preventDefault();
                    confirmBtn.click();;
                }
            } else if (e.key === 'Escape') {
                const cancelBtn = modalElement.querySelector('#modal-cancel, #note-cancel-btn, #photo-close-btn, #edit-adicional-cancel-btn, #edit-user-cancel-btn, #log-close-btn, #preprint-cancel-btn, #layout-close-btn');
                if (cancelBtn) cancelBtn.click();
            }
        };

        modalElement.addEventListener('keydown', keydownHandler);
        
        return () => modalElement.removeEventListener('keydown', keydownHandler);
    }
    
    function debounce(func, delay) {
        let timeout;
        return function(...args) {
            const context = this;
            clearTimeout(timeout);
            timeout = setTimeout(() => func.apply(context, args), delay);
        };
    }

    function showToast(message, type = 'info') {
        const toast = document.createElement('div');
        const bgColor = type === 'error' ? 'bg-red-500' : (type === 'warning' ? 'bg-yellow-500' : 'bg-green-500');
        toast.className = `toast-notification show rounded-lg p-4 text-white shadow-lg transition-all duration-300 transform translate-y-2 opacity-0 ${bgColor}`;
        toast.textContent = message;
        elements.toastContainer.appendChild(toast);
        setTimeout(() => { toast.classList.remove('translate-y-2', 'opacity-0'); }, 10);
        setTimeout(() => {
            toast.classList.add('translate-y-2', 'opacity-0');
            toast.addEventListener('transitionend', () => toast.remove());
        }, 3000);
    }
    
    function showUndoToast(message, onUndo) {
        const toast = document.createElement('div');
        let timeoutId;

        const closeToast = () => {
            toast.classList.add('opacity-0');
            toast.addEventListener('transitionend', () => toast.remove());
            clearTimeout(timeoutId);
        };

        toast.className = 'toast-notification flex items-center justify-between show rounded-lg p-4 text-white shadow-lg transition-all duration-300 transform opacity-0 bg-slate-700';
        toast.innerHTML = `<span>${message}</span>`;

        const undoButton = document.createElement('button');
        undoButton.className = 'ml-4 font-bold underline';
        undoButton.textContent = 'Deshacer';
        undoButton.onclick = () => {
            onUndo();
            closeToast();
        };
        
        toast.appendChild(undoButton);
        elements.toastContainer.appendChild(toast);

        setTimeout(() => { toast.classList.remove('opacity-0'); }, 10);
        timeoutId = setTimeout(closeToast, 5000);
    }
    function updateSerialNumberCache() {
        state.serialNumberCache.clear();
        state.inventory.forEach(item => {
            if (item.SERIE) state.serialNumberCache.add(String(item.SERIE).trim().toLowerCase());
            if (item['CLAVE UNICA']) state.serialNumberCache.add(String(item['CLAVE UNICA']).trim().toLowerCase());
        });
        state.additionalItems.forEach(item => {
            if (item.serie) state.serialNumberCache.add(String(item.serie).trim().toLowerCase());
            if (item.clave) state.serialNumberCache.add(String(item.clave).trim().toLowerCase());
        });
    }
    
    function showConfirmationModal(title, text, onConfirm, options = {}) {
        const { confirmText = 'Confirmar', cancelText = 'Cancelar', onCancel = () => {} } = options;
        elements.modalCancelBtn.style.display = '';
        elements.modalTitle.textContent = title;
        elements.modalText.textContent = text;
        elements.modalConfirmBtn.textContent = confirmText;
        elements.modalCancelBtn.textContent = cancelText;
        elements.confirmationModal.classList.add('show');
        
        const cleanup = handleModalNavigation(elements.confirmationModal);

        const confirmHandler = () => {
            onConfirm();
            closeModal();
        };

        const cancelHandler = () => {
            onCancel();
            closeModal();
        };
        
        const closeModal = () => {
            elements.confirmationModal.classList.remove('show');
            elements.modalConfirmBtn.removeEventListener('click', confirmHandler);
            elements.modalCancelBtn.removeEventListener('click', cancelHandler);
            cleanup();
        };

        elements.modalConfirmBtn.addEventListener('click', confirmHandler, { once: true });
        elements.modalCancelBtn.addEventListener('click', cancelHandler, { once: true });
    }
    function loadState() {
        try {
            const storedState = localStorage.getItem('inventarioProState');
            if (storedState) {
                const loaded = JSON.parse(storedState);
                const defaultState = { 
                    locationPhotos: {}, 
                    activityLog: [], 
                    institutionalReportCheckboxes: {},
                    actionCheckboxes: { labels: {}, notes: {}, additional: {}, mismatched: {}, personal: {} },
                    reportCheckboxes: { notes: {}, mismatched: {} },
                    completedAreas: {},
                    mapLayout: { 'page1': {} },
                    currentLayoutPage: 'page1',
                    layoutPageNames: { 'page1': 'Página 1' },
                    layoutImages: {},
                    layoutPageColors: { 'page1': '#ffffff' },
                    layoutItemColors: {}
                }; 
                state = { ...defaultState, ...state, ...loaded };
                
                if (!state.mapLayout || !state.mapLayout.page1) {
                     if (Object.keys(state.mapLayout || {}).length > 0 && !state.mapLayout.page1) {
                        const oldLayout = { ...state.mapLayout };
                        state.mapLayout = { 'page1': oldLayout };
                        state.currentLayoutPage = 'page1';
                        state.layoutPageNames = { 'page1': 'Página 1' };
                    }
                }
                updateSerialNumberCache();
                return true;
            }
        } catch (e) { 
            console.error('Error al cargar el estado:', e);
            localStorage.removeItem('inventarioProState');
        }
        return false;
    }

    function deleteDB(dbName) {
        return new Promise((resolve, reject) => {
            const request = indexedDB.deleteDatabase(dbName);
            request.onsuccess = () => resolve();
            request.onerror = (event) => reject(event.target.error);
            request.onblocked = () => {
                console.warn('La eliminación de IndexedDB fue bloqueada.');
                resolve(); 
            };
        });
    }
// CONTINÚA EN LA PARTE 2
// ... VIENE DE LA PARTE 1

    // CÓDIGO CORREGIDO
    async function resetInventoryState() {
        const currentUser = state.currentUser;
        const theme = state.theme;

        // 1. Mostrar loading para feedback visual
        elements.loadingOverlay.text.textContent = 'Reiniciando sistema...';
        elements.loadingOverlay.overlay.classList.add('show');

        // 2. Definir estado limpio
        state = {
            loggedIn: true, currentUser, inventory: [], additionalItems: [],
            resguardantes: [], activeResguardante: null, locations: {}, areas: [], areaNames: {},
            lastAutosave: null, sessionStartTime: new Date().toISOString(), additionalPhotos: {}, locationPhotos: {},
            notes: {}, photos: {}, theme,
            inventoryFinished: false,
            areaDirectory: {},
            closedAreas: {},
            completedAreas: {},
            persistentAreas: [],
            serialNumberCache: new Set(),
            cameraStream: null,
            readOnlyMode: false,
            activityLog: [],
            institutionalReportCheckboxes: {},
            actionCheckboxes: { labels: {}, notes: {}, additional: {}, mismatched: {}, personal: {} },
            reportCheckboxes: { notes: {}, mismatched: {} },
            mapLayout: { 'page1': {} },
            currentLayoutPage: 'page1',
            layoutPageNames: { 'page1': 'Página 1' },
            layoutImages: {},
            layoutPageColors: { 'page1': '#ffffff' },
            layoutItemColors: {}
        };

        try {
            // 3. Cerrar conexión explícitamente
            if (photoDB.db) {
                photoDB.db.close();
                photoDB.db = null;
            }

            // 4. Intentar borrar con manejo de bloqueo
            await new Promise((resolve, reject) => {
                const req = indexedDB.deleteDatabase('InventarioProPhotosDB');
                req.onsuccess = () => resolve();
                req.onerror = (e) => reject(e);
                req.onblocked = () => {
                    console.warn("Borrado de DB bloqueado. Intentando forzar recarga...");
                    // Si está bloqueado, no podemos hacer mucho más que resolver y dejar que init() intente recrear
                    resolve(); 
                };
            });

            // 5. Reinicializar
            await photoDB.init();
            
            showToast('Se ha iniciado un nuevo inventario.', 'info');
            logActivity('Sesión reiniciada', `Nuevo inventario iniciado por ${currentUser.name}.`);
            saveState();
            showMainApp();

        } catch (error) {
            console.error("Error crítico al reiniciar DB:", error);
            showToast('Error al limpiar base de datos. Se recomienda recargar la página (F5).', 'error');
            // Aún así intentamos guardar el estado limpio en localStorage
            saveState();
            showMainApp();
        } finally {
            elements.loadingOverlay.overlay.classList.remove('show');
        }
    }

    function saveState() {
        if (state.readOnlyMode) return;

        try {
            const stateToSave = { ...state };
            delete stateToSave.serialNumberCache;
            delete stateToSave.cameraStream;
            localStorage.setItem('inventarioProState', JSON.stringify(stateToSave));
        } catch (e) {
            console.error('Error Crítico al guardar el estado:', e);
            state.readOnlyMode = true;
            checkReadOnlyMode(); 
            showConfirmationModal(
                '¡ALERTA! Almacenamiento Lleno',
                'No se puede guardar más progreso porque el almacenamiento del navegador está lleno. La aplicación se ha puesto en "Modo de Sólo Lectura" para prevenir pérdida de datos. Por favor, exporte su sesión actual desde la pestaña de Ajustes y comience una nueva sesión.',
                () => {},
                { confirmText: 'Entendido', cancelText: '' }
            );
            if(elements.modalCancelBtn) elements.modalCancelBtn.style.display = 'none';
            if (autosaveIntervalId) clearInterval(autosaveIntervalId);
        }
    }

    let autosaveIntervalId;
    function startAutosave() {
        const interval = (parseInt(elements.settings.autosaveInterval.value) || 30) * 1000;
        if (autosaveIntervalId) clearInterval(autosaveIntervalId);
        autosaveIntervalId = setInterval(() => { 
            if (!state.readOnlyMode) {
                saveState(); 
                showToast('Progreso guardado automáticamente.');
                logActivity('Autoguardado', 'El progreso de la sesión se guardó automáticamente.');
            }
        }, interval);
    }

    function checkReadOnlyMode() {
        if (state.readOnlyMode) {
            elements.readOnlyOverlay.classList.remove('hidden');
            document.querySelectorAll(`
                #upload-btn, #file-input, #create-user-btn, .edit-user-btn, 
                .delete-user-btn, .activate-user-btn, #ubicado-btn, #re-etiquetar-btn, 
                #add-note-btn, .inventory-item-checkbox, #select-all-checkbox, #add-adicional-btn, 
                .edit-adicional-btn, .delete-adicional-btn, #note-save-btn, #delete-photo-btn, 
                #photo-input, #use-camera-btn, #capture-photo-btn, .delete-list-btn, 
                #finalize-inventory-btn, #import-session-btn, #import-file-input, 
                #summary-area-responsible, #summary-location, #generate-summary-btn,
                #user-name, #user-location-select, #user-location-manual, #user-area-select,
                #adicional-form input, #adicional-form button, #edit-adicional-form input,
                .save-new-clave-btn, .new-clave-input, .report-btn,
                #open-layout-editor-btn, #layout-save-btn, #layout-page-add, #layout-page-remove, #layout-page-name,
                #layout-add-image-btn, #layout-image-input, #layout-page-reset,
                #add-location-btn, #edit-add-location-btn
            `).forEach(el => {
                el.disabled = true;
                el.style.cursor = 'not-allowed';
                if (el.tagName === 'BUTTON' || el.tagName === 'LABEL') {
                    el.style.opacity = '0.6';
                }
            });
            elements.noteTextarea.readOnly = true;
        } else {
            elements.readOnlyOverlay.classList.add('hidden');
        }
    }


    function renderDashboard() {
        const totalItems = state.inventory.length;
        const locatedItems = state.inventory.filter(item => item.UBICADO === 'SI').length;
        const todayStr = new Date().toISOString().slice(0, 10);
        
        const dailyInventoryProgress = state.inventory.filter(item => item.fechaUbicado && item.fechaUbicado.startsWith(todayStr)).length;
        const dailyAdditionalProgress = state.additionalItems.filter(item => item.fechaRegistro && item.fechaRegistro.startsWith(todayStr)).length;
        const dailyTotal = dailyInventoryProgress + dailyAdditionalProgress;

        elements.totalItemsEl.textContent = totalItems;
        elements.locatedItemsEl.textContent = locatedItems;
        elements.pendingItemsEl.textContent = totalItems - locatedItems;
        elements.dailyProgressEl.textContent = dailyTotal;
        elements.workingAreasCountEl.textContent = new Set(state.inventory.map(item => item.areaOriginal)).size;
        elements.additionalItemsCountEl.textContent = state.additionalItems.length;
    }

    function renderAreaProgress() {
        const container = elements.reports.areaProgressContainer;
        if (!container) return;

        const selectedArea = elements.reports.areaFilter.value;
        const selectedUser = elements.reports.userFilter.value;

        let itemsToStats = state.inventory;

        if (selectedArea !== 'all') {
            itemsToStats = itemsToStats.filter(i => i.areaOriginal === selectedArea);
        }
        if (selectedUser !== 'all') {
            itemsToStats = itemsToStats.filter(i => i['NOMBRE DE USUARIO'] === selectedUser);
        }

        container.innerHTML = '';
        const areas = [...new Set(itemsToStats.map(i => i.areaOriginal))].sort();

        if (areas.length === 0) {
            container.innerHTML = '<p class="text-sm text-gray-500 dark:text-slate-400">No hay áreas para los filtros seleccionados.</p>';
            return;
        }

        let progressHtml = '';
        areas.forEach(area => {
            const areaItems = itemsToStats.filter(i => i.areaOriginal === area);
            const total = areaItems.length;
            if (total === 0) return;
            
            const located = areaItems.filter(i => i.UBICADO === 'SI').length;
            const percent = Math.round((located / total) * 100);
            const areaName = state.areaNames[area] || `Área ${area}`;
            
            const barColor = percent === 100 ? 'bg-green-500' : 'bg-blue-600';

            progressHtml += `
                <div>
                    <div class="flex justify-between mb-1">
                        <span class="text-sm font-medium text-gray-700 dark:text-slate-300">${areaName}</span>
                        <span class="text-sm font-medium text-gray-700 dark:text-slate-300">${located} / ${total} (${percent}%)</span>
                    </div>
                    <div class="progress-bar-container">
                        <div class="${barColor} h-2.5 rounded-full" style="width: ${percent}%"></div>
                    </div>
                </div>
            `;
        });
        container.innerHTML = progressHtml;
    }

    function updateActiveUserBanner() {
        const { banner, name, area } = elements.activeUserBanner;
        const selectDesktop = document.getElementById('active-user-location-select');
        const selectMobile = document.getElementById('active-user-location-select-mobile');
        
        const tabsToShowOn = ['users', 'inventory', 'adicionales'];
        const currentTab = document.querySelector('.tab-btn.active')?.dataset.tab;

        if (state.activeResguardante && tabsToShowOn.includes(currentTab)) {
            // 1. Mostrar datos básicos
            name.textContent = state.activeResguardante.name;
            const areaName = state.areaNames[state.activeResguardante.area] || `Área ${state.activeResguardante.area}`;
            area.textContent = areaName;
            
            // 2. Llenar los selectores de ubicación
            const locations = state.activeResguardante.locations && state.activeResguardante.locations.length > 0 
                              ? state.activeResguardante.locations 
                              : [state.activeResguardante.locationWithId || 'Ubicación Única'];

            const optionsHtml = locations.map(loc => `<option value="${loc}">${loc}</option>`).join('');
            
            if(selectDesktop) {
                selectDesktop.innerHTML = optionsHtml;
                // Sincronizar cambios entre móvil y escritorio
                selectDesktop.onchange = (e) => { if(selectMobile) selectMobile.value = e.target.value; };
            }
            
            if(selectMobile) {
                selectMobile.innerHTML = optionsHtml;
                selectMobile.onchange = (e) => { if(selectDesktop) selectDesktop.value = e.target.value; };
            }

            banner.classList.remove('hidden');
        } else {
            banner.classList.add('hidden');
        }
    }

    function changeTab(tabName) {
        elements.tabContents.forEach(tab => tab.classList.remove('active'));
        document.getElementById(`${tabName}-tab`).classList.add('active');
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.tab === tabName));
        
        const contentArea = elements.mainContentArea;
        contentArea.className = 'p-6 rounded-xl shadow-md glass-effect';
        contentArea.classList.add(`bg-tab-${tabName}`);

        logActivity('Navegación', `Se cambió a la pestaña: ${tabName}.`);
        updateActiveUserBanner();

        if (tabName === 'inventory') {
            currentPage = 1;
            filterAndRenderInventory();
            setTimeout(() => elements.inventory.searchInput.focus(), 100);
        }
        if (tabName === 'users') { renderUserList(); }
        if (tabName === 'reports') {
            renderAreaProgress();
            renderReportStats();
            populateReportFilters();
        }
        if (tabName === 'settings') {
            renderLoadedLists();
            renderDirectory();
        }
        if (tabName === 'adicionales') {
            populateAdicionalesFilters();
            renderAdicionalesList();
            setTimeout(() => document.getElementById('ad-clave').focus(), 100);
        }
    }

    function updateTheme(theme) {
        document.body.classList.toggle('dark-mode', theme === 'dark');
        state.theme = theme;
        logActivity('Ajustes', `Tema cambiado a ${theme}.`);
    }

    function processFile(file) {
        if (state.readOnlyMode) return showToast('Modo de solo lectura: no se pueden cargar nuevos archivos.', 'warning');
        const fileName = file.name;

        const proceedWithUpload = () => {
            elements.loadingOverlay.overlay.classList.add('show');
            elements.dashboard.headerAndDashboard.classList.add('hidden');
            
            // Resetear texto de carga
            elements.loadingOverlay.text.textContent = 'Leyendo archivo...';

            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = e.target.result;
                    const workbook = XLSX.read(data, { type: 'binary' });
                    const sheetName = workbook.SheetNames[0];
                    const sheet = workbook.Sheets[sheetName];
                    const tipoLibro = sheet['B7']?.v || sheet['L7']?.v || 'Sin Tipo';
                    
                    // Aquí llamamos a la función que ahora será asíncrona/chunked
                    // Nota: Ya no usamos 'finally' aquí, addItemsFromFile se encargará de cerrar el overlay
                    addItemsFromFile(sheet, tipoLibro, fileName);
                } catch (error) {
                    console.error("Error processing file: ", error);
                    showToast('Error al procesar el archivo. Asegúrate de que el formato es correcto.', 'error');
                    elements.loadingOverlay.overlay.classList.remove('show');
                }
            };
            reader.onerror = () => {
                elements.loadingOverlay.overlay.classList.remove('show');
                showToast('Error al leer el archivo.', 'error');
            };
            reader.readAsBinaryString(file);
        };

        const isFileAlreadyLoaded = state.inventory.some(item => item.fileName === fileName);
        
        if (isFileAlreadyLoaded) {
            showConfirmationModal(
                'Archivo Duplicado',
                `El archivo "${fileName}" ya fue cargado. ¿Deseas reemplazar los datos existentes de este archivo con el nuevo?`,
                () => {
                    const itemsFromThisFile = state.inventory.filter(item => item.fileName === fileName).length;
                    logActivity('Archivo reemplazado', `Archivo "${fileName}" con ${itemsFromThisFile} bienes fue reemplazado.`);
                    state.inventory = state.inventory.filter(item => item.fileName !== fileName);
                    proceedWithUpload();
                }
            );
        } else {
            proceedWithUpload();
        }
    }
    // --- FUNCIÓN MEJORADA v2: Detecta Texto y Números de Serie de Excel ---
    function findReportDateSmart(sheet) {
        if (!sheet['!ref']) return 'S/F';
        
        const range = XLSX.utils.decode_range(sheet['!ref']);
        // Buscamos en las primeras 10 filas y 30 columnas para asegurar
        const maxRow = Math.min(range.e.r, 10); 
        const maxCol = Math.min(range.e.c, 30); 

        // Regex flexible: Busca dd/mm/aaaa o dd-mm-aaaa
        const dateRegex = /(\d{2})[\/\-](\d{2})[\/\-](\d{4})/;

        for (let R = 0; R <= maxRow; ++R) {
            for (let C = 0; C <= maxCol; ++C) {
                const cellRef = XLSX.utils.encode_cell({c: C, r: R});
                const cell = sheet[cellRef];

                if (!cell) continue;

                // CASO 1: Excel guarda la fecha como NÚMERO (ej: 45192)
                // 43000 es aprox año 2017, 60000 es futuro lejano. Filtramos ruido.
                if (cell.t === 'n' && cell.v > 43000 && cell.v < 60000) {
                    try {
                        const dateObj = XLSX.SSF.parse_date_code(cell.v);
                        if (dateObj && dateObj.d && dateObj.m && dateObj.y) {
                            const day = String(dateObj.d).padStart(2, '0');
                            const month = String(dateObj.m).padStart(2, '0');
                            return `${day}/${month}/${dateObj.y}`;
                        }
                    } catch (e) { console.error('Error convirtiendo fecha Excel', e); }
                }

                // CASO 2: La fecha es TEXTO (ej: "Fecha: 23/09/2025")
                if (cell.v) {
                    const val = String(cell.v);
                    const match = val.match(dateRegex);
                    if (match) {
                        // match[0] es la fecha exacta encontrada
                        return match[0]; 
                    }
                }
            }
        }
        return 'S/F'; // Si falla todo
    }
    function extractResponsibleInfo(sheet) {
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
        const contentRows = data.filter(row => row.some(cell => cell !== null && String(cell).trim() !== ''));

        if (contentRows.length >= 2) {
            const nameRow = contentRows[contentRows.length - 2];
            const titleRow = contentRows[contentRows.length - 1];
            
            const name = nameRow.find(cell => cell !== null && String(cell).trim() !== '');
            const title = titleRow.find(cell => cell !== null && String(cell).trim() !== '');

            if (name && title && isNaN(name) && isNaN(title) && String(name).length > 3 && String(title).length > 3) {
                return { name: String(name).trim(), title: String(title).trim() };
            }
        }

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            for (let j = 0; j < row.length; j++) {
                if (String(row[j]).trim().toLowerCase() === 'responsable') {
                    if (i + 3 < data.length) {
                        const name = data[i + 2] ? String(data[i + 2][j] || '').trim() : null;
                        const title = data[i + 3] ? String(data[i + 3][j] || '').trim() : null;
                        if (name && title) return { name, title };
                    }
                }
            }
        }
        
        return null;
    }
    function addItemsFromFile(sheet, tipoLibro, fileName) {
        // 1. Extracción preliminar de metadatos (rápido)
        const areaString = sheet['A10']?.v || 'Sin Área';
        const area = areaString.match(/AREA\s(\d+)/)?.[1] || 'Sin Área';
        const printDate = findReportDateSmart(sheet);
        const listId = Date.now();
        
        if (area && !state.areaNames[area]) {
            state.areaNames[area] = areaString;
        }
        
        const responsible = extractResponsibleInfo(sheet);
        if (area && !state.areaDirectory[area]) {
            if (responsible) {
                state.areaDirectory[area] = {
                    fullName: areaString,
                    name: responsible.name,
                    title: responsible.title,
                };
            }
        }

        // 2. Convertir hoja a JSON crudo (esto puede tardar un poco, pero es nativo de la librería)
        const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 11 });
        const totalRows = rawData.length;
        const claveUnicaRegex = /^(?:\d{5,6}|0\.\d+)$/;
        
        let processedRows = 0;
        const chunkSize = 500; // Procesar 500 filas por bloque
        const newItemsBatch = [];

        // 3. Función recursiva para procesar por lotes sin bloquear la UI
        function processChunk() {
            const end = Math.min(processedRows + chunkSize, totalRows);
            
            for (let i = processedRows; i < end; i++) {
                const row = rawData[i];
                const clave = String(row[0] || '').trim();
                
                if (claveUnicaRegex.test(clave)) {
                    newItemsBatch.push({
                        'CLAVE UNICA': clave, 
                        'DESCRIPCION': String(row[1] || ''), 
                        'OFICIO': row[2] || '', 
                        'TIPO': row[3] || '',
                        'MARCA': row[4] || '', 
                        'MODELO': row[5] || '', 
                        'SERIE': row[6] || '', 
                        'FECHA DE INICIO': row[7] || '',
                        'REMISIÓN': row[8] || '', 
                        'FECHA DE REMISIÓN': row[9] || '', 
                        'FACTURA': row[10] || '', 
                        'FECHA DE FACTURA': row[11] || '', 
                        'AÑO': row[12] || '',
                        'NOMBRE DE USUARIO': '', 
                        'UBICADO': 'NO', 
                        'IMPRIMIR ETIQUETA': 'NO',
                        'listadoOriginal': tipoLibro, 
                        'areaOriginal': area,
                        'listId': listId, 
                        'fileName': fileName, 
                        'printDate': printDate
                    });
                }
            }

            processedRows = end;

            // Actualizar UI de progreso
            if (elements.loadingOverlay.text) {
                const percent = Math.round((processedRows / totalRows) * 100);
                elements.loadingOverlay.text.textContent = `Procesando: ${percent}% (${processedRows}/${totalRows} bienes)...`;
            }

            if (processedRows < totalRows) {
                // Si faltan filas, programar el siguiente bloque para dejar respirar a la UI
                setTimeout(processChunk, 0);
            } else {
                // 4. Finalizar proceso
                finalizeImport();
            }
        }

        function finalizeImport() {
            state.inventory = state.inventory.concat(newItemsBatch);
            state.inventoryFinished = false; 
            
            logActivity('Archivo cargado', `Archivo "${fileName}" con ${newItemsBatch.length} bienes para el área ${area}. Tipo: ${tipoLibro}.`);

            const responsibleName = responsible?.name || 'No detectado';
            const toastMessage = `Área ${area}: Se cargaron ${newItemsBatch.length} bienes. Responsable: ${responsibleName}.`;
            
            // Guardar y renderizar
            saveState();
            renderDashboard();
            populateAreaSelects();
            populateReportFilters();
            populateBookTypeFilter();
            currentPage = 1;
            filterAndRenderInventory();
            renderLoadedLists();
            renderDirectory();
            updateSerialNumberCache();

            // Ocultar overlay
            elements.loadingOverlay.overlay.classList.remove('show');
            showToast(toastMessage, 'success');
        }

        // Iniciar el procesamiento
        processChunk();
    }
// CONTINÚA EN LA PARTE 3
// ... VIENE DE LA PARTE 2

    function filterAndRenderInventory() {
        const searchTerm = elements.inventory.searchInput.value.trim().toLowerCase();
        const statusFilter = elements.inventory.statusFilter.value;
        const areaFilter = elements.inventory.areaFilter.value;
        const bookTypeFilter = elements.inventory.bookTypeFilter.value;

        filteredItems = state.inventory.filter(item =>
            (!searchTerm || [item['CLAVE UNICA'], item['DESCRIPCION'], item['MARCA'], item['MODELO'], item['SERIE']].some(f => String(f||'').toLowerCase().includes(searchTerm))) &&
            (statusFilter === 'all' || item.UBICADO === statusFilter) &&
            (areaFilter === 'all' || item.areaOriginal === areaFilter) &&
            (bookTypeFilter === 'all' || item.listadoOriginal === bookTypeFilter)
        );
        
        renderInventoryTable();

        if (searchTerm && filteredItems.length === 1 && String(filteredItems[0]['CLAVE UNICA']).toLowerCase() === searchTerm) {
            showItemDetailView(filteredItems[0]['CLAVE UNICA']);
        }

        const additionalResultsContainer = document.getElementById('additional-search-results-container');
        const additionalResultsList = document.getElementById('additional-search-results-list');

        if (!searchTerm) {
            additionalResultsContainer.classList.add('hidden');
            return;
        }

        const additionalMatches = state.additionalItems.filter(item =>
            (item.clave && String(item.clave).toLowerCase().includes(searchTerm)) ||
            (item.descripcion && item.descripcion.toLowerCase().includes(searchTerm)) ||
            (item.marca && item.marca.toLowerCase().includes(searchTerm)) ||
            (item.modelo && item.modelo.toLowerCase().includes(searchTerm)) ||
            (item.serie && String(item.serie).toLowerCase().includes(searchTerm)) ||
            (item.claveAsignada && String(item.claveAsignada).toLowerCase().includes(searchTerm))
        );

        if (additionalMatches.length > 0) {
            additionalResultsList.innerHTML = additionalMatches.map(item => {
                const isPersonal = item.personal === 'Si';
                const itemClass = isPersonal ? 'personal-item' : 'additional-item';
                return `
                    <div class="flex items-center justify-between p-3 rounded-lg shadow-sm border-l-4 ${itemClass}">
                        <div>
                            <p class="font-semibold">${item.descripcion}</p>
                            <p class="text-sm opacity-80">Clave: ${item.clave || 'N/A'}, Serie: ${item.serie || 'N/A'}, Clave Asignada: ${item.claveAsignada || 'N/A'}</p>
                            <p class="text-xs opacity-70 mt-1">Asignado a: <strong>${item.usuario}</strong></p>
                        </div>
                        <i class="fa-solid fa-star text-purple-400" title="Bien Adicional"></i>
                    </div>
                `;
            }).join('');
            additionalResultsContainer.classList.remove('hidden');
        } else {
            additionalResultsContainer.classList.add('hidden');
        }
    }

    function highlightText(text, searchTerm) {
        if (!searchTerm.trim() || !text) {
            return text;
        }
        const escapedTerm = searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(`(${escapedTerm})`, 'gi');
        return String(text).replace(regex, `<mark class="bg-yellow-300 rounded-sm px-1">$1</mark>`);
    }
    
    function createInventoryRowElement(item) {
        const searchTerm = elements.inventory.searchInput.value.trim();
        const clave = item['CLAVE UNICA'] || '';
        const descripcion = item['DESCRIPCION'] || '';
        const marca = item['MARCA'] || '';
        const modelo = item['MODELO'] || '';
        const serie = item['SERIE'] || '';
        const usuario = item['NOMBRE DE USUARIO'] || '';

        const row = document.createElement('tr');
        let rowClasses = 'hover:bg-gray-50 dark:hover:bg-gray-700 transition-colors';
        
        // Si NO estamos editando, el cursor es pointer para abrir detalles
        if (!state.inventoryEditMode) rowClasses += ' cursor-pointer';
        
        if (state.notes[clave]) rowClasses += ' has-note';
        if (item.UBICADO === 'SI') rowClasses += ' item-located';
        row.className = rowClasses;
        row.dataset.clave = clave;
        
        const mismatchTag = item.areaIncorrecta ? `<span class="mismatched-area-tag" title="Ubicado en el área de otro listado">⚠️</span>` : '';
        
        const userData = state.resguardantes.find(u => u.name === usuario);
        
        let locationDisplay = '';
        if (item.ubicacionEspecifica) {
            locationDisplay = `📍 Encontrado en: ${item.ubicacionEspecifica}`;
        } else if (userData) {
             locationDisplay = userData.locations ? userData.locations.join(', ') : userData.locationWithId;
        }
        
        const userDetails = userData 
            ? `${userData.name}\nÁrea: ${userData.area}\n${locationDisplay}` 
            : usuario;
        
        const truncate = (str, len) => (str && String(str).length > len ? String(str).substring(0, len) + '...' : str || '');

        // --- LÓGICA DE EDICIÓN ---
        const isEdit = state.inventoryEditMode;
        const editClass = isEdit ? 'inventory-editable-cell' : '';
        const contentEditableAttr = isEdit ? 'contenteditable="true"' : '';
        
        // Helper para renderizar celda (Editable o Normal) - VERSIÓN SEGURA
        const renderCell = (field, value, truncateLen = 0) => {
            // Sanitizamos el valor SIEMPRE primero para evitar XSS
            const safeValue = escapeHTML(value);
            
            if (isEdit) {
                // En modo edición usamos safeValue
                return `<div class="${editClass} w-full h-full min-h-[24px]" ${contentEditableAttr} data-field="${field}">${safeValue}</div>`;
            } else {
                // Modo normal: truncado sobre el valor seguro
                const text = truncateLen > 0 ? truncate(safeValue, truncateLen) : safeValue;
                return highlightText(text, escapeHTML(searchTerm));
            }
        };

        row.innerHTML = `
            <td class="px-2 py-2"><input type="checkbox" class="inventory-item-checkbox rounded"></td>
            <td class="px-2 py-2 text-sm font-mono" title="${escapeHTML(clave)}">${highlightText(escapeHTML(clave), searchTerm)}</td>

            <td class="px-2 py-2 text-sm" title="${escapeHTML(descripcion)}">
                ${renderCell('DESCRIPCION', descripcion, 30)}
                ${!isEdit ? mismatchTag : ''} 
            </td>
            <td class="px-2 py-2 text-sm" title="${escapeHTML(marca)}">
                ${renderCell('MARCA', marca)}
            </td>
            <td class="px-2 py-2 text-sm" title="${escapeHTML(modelo)}">
                ${renderCell('MODELO', modelo)}
            </td>
            <td class="px-2 py-2 text-sm" title="${escapeHTML(serie)}">
                ${renderCell('SERIE', serie)}
            </td>
            <td class="px-2 py-2 text-sm" title="${escapeHTML(userDetails)}">
                ${highlightText(escapeHTML(usuario), searchTerm)}
            </td>
            <td class="px-2 py-2 text-sm text-center">${item['UBICADO'] === 'SI' ? '<span class="text-green-600 font-bold">SI</span>' : '<span class="text-red-400">NO</span>'}</td>
            <td class="px-2 py-2 text-sm text-center">${item['IMPRIMIR ETIQUETA'] === 'SI' ? '<span class="text-orange-500 font-bold">SI</span>' : 'NO'}</td>
            <td class="px-2 py-2 text-center">
                <div class="flex items-center justify-center space-x-3">
                    <i class="fa-solid fa-note-sticky text-xl ${state.notes[clave] ? 'text-yellow-500' : 'text-gray-400'} note-icon cursor-pointer hover:scale-110 transition-transform" title="Añadir/Ver Nota"></i>
                    <i class="fa-solid fa-camera text-xl ${state.photos[clave] ? 'text-indigo-500' : 'text-gray-400'} camera-icon cursor-pointer hover:scale-110 transition-transform" title="Añadir/Ver Foto"></i>
                    <i class="fa-solid fa-circle-info text-xl text-gray-400 hover:text-blue-500 md:hidden view-details-btn cursor-pointer" title="Ver Detalles"></i>
                    <i class="fa-solid fa-qrcode text-xl text-gray-400 hover:text-indigo-500 view-qr-btn cursor-pointer" title="Ver Código QR"></i>
                </div>
            </td>`;
        
        return row;
    }

    function renderInventoryTable() {
        const { tableBody, pageInfo, prevPageBtn, nextPageBtn } = elements.inventory;
        const fragment = document.createDocumentFragment();

        const totalPages = Math.ceil(filteredItems.length / itemsPerPage) || 1;
        if (currentPage > totalPages) currentPage = totalPages;
        
        const start = (currentPage - 1) * itemsPerPage;
        const end = start + itemsPerPage;
        const itemsToRender = filteredItems.slice(start, end);

        if (itemsToRender.length === 0) {
            const emptyRow = document.createElement('tr');
            const cell = document.createElement('td');
            cell.colSpan = 12; 
            cell.className = 'text-center py-4 text-gray-500';
            cell.textContent = 'No se encontraron bienes con los filtros actuales.';
            emptyRow.appendChild(cell);
            fragment.appendChild(emptyRow);
        } else {
            itemsToRender.forEach(item => {
                const rowElement = createInventoryRowElement(item);
                fragment.appendChild(rowElement);
            });
        }
        
        tableBody.innerHTML = '';
        tableBody.appendChild(fragment);

        pageInfo.textContent = `Página ${currentPage} de ${totalPages}`;
        prevPageBtn.disabled = currentPage === 1;
        nextPageBtn.disabled = currentPage >= totalPages;
    }
    function handleInventoryActions(action) {
        if (state.readOnlyMode) return showToast('Modo de solo lectura: no se pueden realizar acciones.', 'warning');
        const selectedClaves = Array.from(document.querySelectorAll('.inventory-item-checkbox:checked')).map(cb => cb.closest('tr').dataset.clave);
        if (selectedClaves.length === 0) return showToast('Seleccione al menos un bien.', 'error');
        
        if (action === 'desubicar') {
            showConfirmationModal('Des-ubicar Bienes', `¿Estás seguro de que quieres marcar ${selectedClaves.length} bien(es) como NO ubicados? Esto eliminará la asignación de usuario.`, () => {
                selectedClaves.forEach(clave => {
                    const item = state.inventory.find(i => i['CLAVE UNICA'] === clave);
                    if (item) {
                        item.UBICADO = 'NO';
                        item['NOMBRE DE USUARIO'] = '';
                        item['IMPRIMIR ETIQUETA'] = 'NO'; 
                        item.fechaUbicado = null;
                        item.areaIncorrecta = false;
                        logActivity('Bien des-ubicado', `Clave: ${clave}`);
                        checkAreaCompletion(item.areaOriginal); 
                    }
                });
                showToast(`${selectedClaves.length} bien(es) marcado(s) como NO ubicado(s).`);
                filterAndRenderInventory(); renderDashboard(); saveState();
            });
            return; 
        }

        if (!state.activeResguardante) {
            return showToast('Debe activar un usuario para poder ubicar o re-etiquetar bienes.', 'error');
        }
        const activeUser = state.activeResguardante;
        const { searchInput } = elements.inventory;

        selectedClaves.forEach(clave => {
            const item = state.inventory.find(i => i['CLAVE UNICA'] === clave);
            if (!item) return;

            const isAssignedToOther = item.UBICADO === 'SI' && item['NOMBRE DE USUARIO'] && item['NOMBRE DE USUARIO'] !== activeUser.name;
            
            const processItem = () => {
                assignItem(item, activeUser); 
                
                if (action === 're-etiquetar') {
                    item['IMPRIMIR ETIQUETA'] = 'SI';
                    logActivity('Bien marcado para re-etiquetar', `Clave: ${clave}, Usuario: ${activeUser.name}`);
                } else if (action === 'ubicar') {
                    if (item['IMPRIMIR ETIQUETA'] === 'SI') {
                        item['IMPRIMIR ETIQUETA'] = 'NO';
                        logActivity('Marca de re-etiquetar quitada al ubicar', `Clave: ${clave}, Usuario: ${activeUser.name}`);
                    } else {
                        logActivity('Bien ubicado', `Clave: ${clave}, Usuario: ${activeUser.name}`);
                    }
                }
            };

            if (isAssignedToOther) {
                showConfirmationModal('Reasignar Bien', `El bien ${clave} ya está asignado a ${item['NOMBRE DE USUARIO']}. ¿Deseas reasignarlo a ${activeUser.name}?`, () => {
                    logActivity('Bien reasignado', `Clave: ${clave} de ${item['NOMBRE DE USUARIO']} a ${activeUser.name}`);
                    processItem();
                    showToast(`Bien ${clave} reasignado a ${activeUser.name}.`);
                    filterAndRenderInventory(); renderDashboard(); saveState();
                });
            } else {
                processItem(); 
            }
        });
        const requiresConfirmation = selectedClaves.some(clave => {
             const item = state.inventory.find(i => i['CLAVE UNICA'] === clave);
             return item && item.UBICADO === 'SI' && item['NOMBRE DE USUARIO'] && item['NOMBRE DE USUARIO'] !== activeUser.name;
        });

        if (!requiresConfirmation) {
             const message = action === 'ubicar' ? `Se ubicaron ${selectedClaves.length} bienes.` : `Se marcaron ${selectedClaves.length} bienes para re-etiquetar y fueron ubicados.`;
             showToast(message);
             searchInput.value = '';
             searchInput.focus();
             filterAndRenderInventory(); renderDashboard(); saveState();
        } else {
            showToast(`Algunos bienes requerían confirmación para reasignar.`);
            document.querySelectorAll('.inventory-item-checkbox:checked').forEach(cb => cb.checked = false);
        }
    }
    function assignItem(item, user) {
        // 1. Obtener la ubicación precisa del selector del banner
        const selectDesktop = document.getElementById('active-user-location-select');
        const preciseLocation = selectDesktop ? selectDesktop.value : (user.locationWithId || 'N/A');

        // 2. Asignar datos estándar
        item.UBICADO = 'SI';
        item['NOMBRE DE USUARIO'] = user.name;
        item.fechaUbicado = new Date().toISOString();
        item.areaIncorrecta = item.areaOriginal !== user.area;
        
        // 3. GUARDAR LA UBICACIÓN PRECISA (El Anclaje)
        item.ubicacionEspecifica = preciseLocation; 
        
        logActivity('Bien Ubicado', `Clave: ${item['CLAVE UNICA']} anclado a: ${preciseLocation}`);

        checkAreaCompletion(item.areaOriginal);
        checkInventoryCompletion();
    }

    function checkInventoryCompletion() {
        if (state.inventoryFinished || state.inventory.length === 0) return;

        const allLocated = state.inventory.every(item => item.UBICADO === 'SI');
        if (allLocated) {
            state.inventoryFinished = true;
            logActivity('Inventario completado', 'Todos los bienes han sido ubicados.');
            showConfirmationModal(
                '¡Inventario Completado!',
                '¡Felicidades! Has ubicado todos los bienes. ¿Deseas generar el Resumen de Sesión y el Plan de Acción?',
                () => { 
                    showPreprintModal('session_summary');
                }
            );
            saveState();
        }
    }
    function checkAreaCompletion(areaId) {
        if (!areaId || state.closedAreas[areaId]) {
            return; 
        }

        const areaItems = state.inventory.filter(item => item.areaOriginal === areaId);
        const isAreaComplete = areaItems.length > 0 && areaItems.every(item => item.UBICADO === 'SI');
        const wasPreviouslyComplete = !!state.completedAreas[areaId];

        if (isAreaComplete && !wasPreviouslyComplete) {
            state.completedAreas[areaId] = true; 
            logActivity('Área completada', `Todos los bienes del área ${areaId} han sido ubicados.`);
            showToast(`¡Área ${state.areaNames[areaId] || areaId} completada! Puedes generar el Acta de Cierre desde la pestaña de Ajustes.`);
            saveState(); 
            renderLoadedLists(); 
        } else if (!isAreaComplete && wasPreviouslyComplete) {
            delete state.completedAreas[areaId];
            logActivity('Área ya no completada', `El área ${areaId} ahora tiene bienes pendientes.`);
            saveState();
            renderLoadedLists(); 
        }
    }
    
    // ACTUALIZADO: Renderizado de usuarios con lógica visual de multi-ubicación
    function renderUserList() {
        const list = elements.userForm.list;
        const searchInput = document.getElementById('user-search-input');
        const userCountBadge = document.getElementById('user-count-badge');
        
        const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';

        const filteredUsers = state.resguardantes.filter(user => {
            if (!searchTerm) return true;
            // Búsqueda mejorada: incluye arrays de ubicaciones
            const locationsString = (user.locations || [user.locationWithId]).join(' ').toLowerCase();
            return (
                user.name.toLowerCase().includes(searchTerm) ||
                locationsString.includes(searchTerm) ||
                String(user.area).toLowerCase().includes(searchTerm)
            );
        });
        
        if (userCountBadge) {
            userCountBadge.textContent = `${filteredUsers.length} de ${state.resguardantes.length} Total`;
        }

        list.innerHTML = filteredUsers.length === 0 
            ? `<p class="text-gray-500">No se encontraron usuarios.</p>` 
            : '';
            
        filteredUsers.forEach((user) => {
            const originalIndex = state.resguardantes.findIndex(u => u.id === user.id);
            const isActive = state.activeResguardante?.id === user.id;
            const item = document.createElement('div');
            item.className = `flex items-center justify-between p-2 rounded-lg shadow-sm transition-colors cursor-pointer ${isActive ? 'active-user border-l-4 border-green-500' : 'non-active-user'}`;
            item.dataset.userId = user.id;
            
            // Determinar color del icono de cámara si hay foto en la ubicación principal
            const hasLocationPhoto = state.locationPhotos && state.locationPhotos[user.locationWithId];
            const photoIconColor = hasLocationPhoto ? 'text-indigo-500' : 'text-gray-400';

            // Renderizado inteligente de ubicaciones
            let locationText = '';
            if (user.locations && user.locations.length > 0) {
                // Mostrar las primeras 2 ubicaciones y "+X más" si hay muchas
                const displayLocs = user.locations.slice(0, 2).join(', ');
                const remaining = user.locations.length - 2;
                locationText = remaining > 0 ? `${displayLocs} (+${remaining})` : displayLocs;
            } else {
                locationText = user.locationWithId || 'Sin ubicación';
            }

            item.innerHTML = `
                <div class="flex-grow user-info-clickable" data-user-id="${user.id}">
                   <p class="font-semibold">${escapeHTML(user.name)}</p>
                   <p class="text-sm text-gray-500 dark:text-gray-400">
                       <span class="font-medium text-xs bg-gray-200 dark:bg-slate-700 px-1 rounded text-gray-600 dark:text-gray-300 mr-1">Área ${escapeHTML(user.area)}</span>
                       ${escapeHTML(locationText)}
                   </p>
                </div>
                <div class="space-x-2 flex items-center">
                    <i class="fa-solid fa-camera text-xl ${photoIconColor} cursor-pointer location-photo-btn" data-location-id="${user.locationWithId}" title="Gestionar foto de la ubicación principal"></i>
                    <button data-index="${originalIndex}" class="activate-user-btn px-3 py-1 rounded-lg text-xs font-bold transition-colors ${isActive ? 'text-white bg-green-600' : 'text-gray-700 bg-gray-200 hover:bg-gray-300'}">${isActive ? 'Activo' : 'Activar'}</button>
                    <button data-index="${originalIndex}" class="edit-user-btn px-3 py-1 rounded-lg text-xs font-bold text-white bg-blue-500 hover:bg-blue-600">Editar</button>
                    <button data-index="${originalIndex}" class="delete-user-btn px-3 py-1 rounded-lg text-xs font-bold text-white bg-red-500 hover:bg-red-600">Eliminar</button>
                </div>`;
            list.appendChild(item);
        });
    }
    
    function renderAdicionalesList() {
        const listEl = elements.adicionales.list;
        const filterUser = elements.adicionales.userFilter.value;
        const filterArea = elements.adicionales.areaFilter.value;
        
        let filtered = state.additionalItems;

        if (filterArea && filterArea !== 'all') {
            const usersInArea = state.resguardantes
                .filter(user => user.area === filterArea)
                .map(user => user.name);
            filtered = filtered.filter(item => usersInArea.includes(item.usuario));
        }

        if (filterUser && filterUser !== 'all') {
            filtered = filtered.filter(item => item.usuario === filterUser);
        }
        
        elements.adicionales.total.textContent = `${filtered.length} de ${state.additionalItems.length} Total`;

        if (filtered.length === 0) {
            listEl.innerHTML = '<p class="text-gray-500">No hay bienes adicionales con los filtros seleccionados.</p>';
            return;
        }

        listEl.innerHTML = filtered.map((item, index) => {
            const isPersonal = item.personal === 'Si';
            const itemClass = isPersonal ? 'personal-item' : 'additional-item';
            
            let personalTag = '';
            if (isPersonal) {
                if (item.tieneFormatoEntrada === true) {
                    personalTag = `<span class="font-bold text-xs ml-2" title="Tiene formato de entrada"><i class="fa-solid fa-file-circle-check text-green-600"></i> (Personal)</span>`;
                } else if (item.tieneFormatoEntrada === false) {
                    personalTag = `<span class="font-bold text-xs ml-2" title="No tiene formato de entrada"><i class="fa-solid fa-file-circle-exclamation text-amber-600"></i> (Personal)</span>`;
                } else {
                    personalTag = `<span class="font-bold text-xs ml-2">(Personal)</span>`;
                }
            }

            const hasPhoto = state.additionalPhotos[item.id];

            return `<div data-id="${item.id}" class="adicional-item-clickable flex items-center justify-between p-3 rounded-lg shadow-sm border-l-4 ${itemClass} cursor-pointer">
                <div class="flex items-center" data-id="${item.id}">
                    <span class="font-bold text-lg mr-3">${index + 1}.</span>
                    <div>
                        <p class="font-semibold">${item.descripcion}${personalTag}</p>
                        <p class="text-sm opacity-80">Clave: ${item.clave || 'N/A'}, Marca: ${item.marca || 'N/A'}, Serie: ${item.serie || 'N/A'}</p>
                        <p class="text-sm opacity-70">Usuario: ${item.usuario}</p>
                    </div>
                </div>
                <div class="space-x-2">
                    <button data-id="${item.id}" class="adicional-photo-btn action-btn ${hasPhoto ? 'text-indigo-500' : ''}"><i class="fa-solid fa-camera"></i></button>
                    <button data-id="${item.id}" class="edit-adicional-btn action-btn"><i class="fa-solid fa-pencil"></i></button>
                    <button data-id="${item.id}" class="delete-adicional-btn action-btn"><i class="fa-solid fa-trash-can"></i></button>
                </div>
            </div>`
        }).join('');
    }

    function showEditAdicionalModal(id) {
        if (state.readOnlyMode) return;
        const item = state.additionalItems.find(i => i.id === id);
        if (!item) return;

        const { modal, form, saveBtn } = elements.editAdicionalModal;
        form.elements['clave'].value = item.clave || '';
        form.elements['descripcion'].value = item.descripcion || '';
        form.elements['marca'].value = item.marca || '';
        form.elements['modelo'].value = item.modelo || '';
        form.elements['serie'].value = item.serie || '';
        form.elements['area'].value = item.area || '';
        form.elements['personal'].value = item.personal || 'No';
        
        saveBtn.dataset.id = id;
        modal.classList.add('show');
    }

    // Variable global para rastrear la URL de la foto activa y evitar fugas de memoria
    let activeDetailPhotoUrl = null;

    function updateDetailViewPhoto(clave) {
        const { detailView } = elements;
        
        // 1. Limpieza de memoria previa (CRÍTICO)
        if (activeDetailPhotoUrl) {
            URL.revokeObjectURL(activeDetailPhotoUrl);
            activeDetailPhotoUrl = null;
        }
        
        // 2. Resetear UI
        detailView.photo.classList.add('hidden');
        detailView.noPhoto.classList.remove('hidden');
        detailView.photo.src = ''; 

        if (state.photos[clave]) {
            photoDB.getItem('photos', `inventory-${clave}`).then(imageBlob => {
                if (imageBlob) {
                    // Crear nueva URL y guardarla en la variable de rastreo
                    activeDetailPhotoUrl = URL.createObjectURL(imageBlob);
                    
                    detailView.photo.src = activeDetailPhotoUrl;
                    detailView.photo.classList.remove('hidden');
                    detailView.noPhoto.classList.add('hidden');
                }
            }).catch(() => {
                detailView.photo.classList.add('hidden');
                detailView.noPhoto.classList.remove('hidden');
            });
        }
    }

    let activeModalPhotoUrl = null;

    function showPhotoModal(type, id) {
        const { modal, title, input, deleteBtn, viewContainer, uploadContainer, cameraViewContainer, img } = elements.photo;
        
        // 1. Limpieza de memoria previa
        if (activeModalPhotoUrl) {
            URL.revokeObjectURL(activeModalPhotoUrl);
            activeModalPhotoUrl = null;
        }

        let modalTitle = 'Foto del Bien';
        if (type === 'location') modalTitle = `Foto de la Ubicación: ${id}`;
        title.textContent = modalTitle;

        input.dataset.type = type;
        input.dataset.id = id;
        deleteBtn.dataset.type = type;
        deleteBtn.dataset.id = id;
        
        let photoExists = false;
        if (type === 'inventory') photoExists = state.photos[id];
        else if (type === 'additional') photoExists = state.additionalPhotos[id];
        else if (type === 'location') photoExists = state.locationPhotos[id];
        
        viewContainer.classList.add('hidden');
        uploadContainer.classList.add('hidden');
        cameraViewContainer.classList.add('hidden');
        stopCamera();

        if (photoExists) {
            viewContainer.classList.remove('hidden');
            
            let storeKey = '';
            if (type === 'inventory') storeKey = `inventory-${id}`;
            else if (type === 'additional') storeKey = `additional-${id}`;
            else if (type === 'location') storeKey = `location-${id}`;

            photoDB.getItem('photos', storeKey).then(imageBlob => {
                if (imageBlob) {
                    activeModalPhotoUrl = URL.createObjectURL(imageBlob);
                    img.src = activeModalPhotoUrl;
                } else {
                    img.src = '';
                    img.alt = 'Error al cargar la imagen desde la base de datos.';
                }
            }).catch(() => {
                img.src = '';
                img.alt = 'Error al cargar la imagen.';
            });
        } else {
            if (!state.readOnlyMode) {
                uploadContainer.classList.remove('hidden');
            } else {
                // Caso: Read only sin foto
                viewContainer.classList.remove('hidden');
                img.style.display = 'none';
                deleteBtn.style.display = 'none';
                const msg = document.createElement('p');
                msg.textContent = 'Sin fotografía disponible.';
                msg.className = 'text-gray-500 italic mt-4';
                viewContainer.appendChild(msg);
            }
        }
        
        modal.classList.add('show');
    }
    
    function showItemDetailsModal(clave) {
         const item = state.inventory.find(i => i['CLAVE UNICA'] === clave);
         if (!item) return;

         const { modal, title, content } = elements.itemDetailsModal;
         title.textContent = `Detalles: ${item['CLAVE UNICA']}`;
         
         const userData = state.resguardantes.find(u => u.name === item['NOMBRE DE USUARIO']);
         let loc = 'N/A';
         if (userData) {
            loc = (userData.locations && userData.locations.length > 0) ? userData.locations.join(', ') : userData.locationWithId;
         }

         const userDetails = userData 
            ? `<strong>Usuario:</strong> ${userData.name}<br>
               <strong>Área:</strong> ${userData.area}<br>
               <strong>Ubicación:</strong> ${loc}`
            : `<strong>Usuario:</strong> ${item['NOMBRE DE USUARIO'] || 'No asignado'}`;

         content.innerHTML = `
            <p><strong>Descripción:</strong> ${item.DESCRIPCION || 'N/A'}</p>
            <p><strong>Marca:</strong> ${item.MARCA || 'N/A'}</p>
            <p><strong>Modelo:</strong> ${item.MODELO || 'N/A'}</p>
            <p><strong>Serie:</strong> ${item.SERIE || 'N/A'}</p>
            <hr class="my-2">
            <p>${userDetails}</p>
         `;
         modal.classList.add('show');
    }

    function showItemDetailView(clave) {
        const item = state.inventory.find(i => i['CLAVE UNICA'] === clave);
        if (!item) return;

        const { detailView } = elements;
        const areaName = state.areaNames[item.areaOriginal] || `Área ${item.areaOriginal}`;

        detailView.clave.textContent = item['CLAVE UNICA'];
        detailView.descripcion.textContent = item['DESCRIPCION'] || 'N/A';
        detailView.marca.textContent = item['MARCA'] || 'N/A';
        detailView.modelo.textContent = item['MODELO'] || 'N/A';
        detailView.serie.textContent = item['SERIE'] || 'N/A';
        detailView.usuario.textContent = item['NOMBRE DE USUARIO'] || 'Sin Asignar';
        detailView.ubicacionEspecifica.textContent = item.ubicacionEspecifica || 'No registrada (Ubicación general)';
        detailView.area.textContent = areaName;
        
        const warningContainer = detailView.areaWarning;
        warningContainer.innerHTML = '';
        warningContainer.className = 'mt-3 p-3 rounded-lg text-sm hidden';

        const activeUser = state.activeResguardante;
        if (activeUser && item.areaOriginal !== activeUser.area) {
            warningContainer.classList.remove('hidden');
            warningContainer.classList.add('bg-yellow-100', 'dark:bg-yellow-900/50', 'text-yellow-800', 'dark:text-yellow-200');
            warningContainer.innerHTML = `<i class="fa-solid fa-triangle-exclamation mr-2"></i><strong>Aviso:</strong> Este bien pertenece al <strong>área ${item.areaOriginal}</strong>, pero el usuario activo está en el <strong>área ${activeUser.area}</strong>.`;
        } else if (!activeUser) {
            warningContainer.classList.remove('hidden');
            warningContainer.classList.add('bg-blue-100', 'dark:bg-blue-900/50', 'text-blue-800', 'dark:text-blue-200');
            warningContainer.innerHTML = `<i class="fa-solid fa-info-circle mr-2"></i>Para ubicar este bien, primero activa un usuario en la pestaña "Usuarios".`;
        }

        updateDetailViewPhoto(clave);

        const closeModal = () => detailView.modal.classList.remove('show');

        const ubicarHandler = () => {
            if (!state.activeResguardante) return showToast('Activa un usuario para poder ubicar.', 'error');
            const checkbox = document.querySelector(`tr[data-clave="${clave}"] .inventory-item-checkbox`);
            if(checkbox) checkbox.checked = true;
            handleInventoryActions('ubicar');
            if(checkbox) checkbox.checked = false; 
            closeModal();
        };

        const reetiquetarHandler = () => {
             if (!state.activeResguardante) return showToast('Activa un usuario para poder re-etiquetar.', 'error');
            const checkbox = document.querySelector(`tr[data-clave="${clave}"] .inventory-item-checkbox`);
            if(checkbox) checkbox.checked = true;
            handleInventoryActions('re-etiquetar');
            if(checkbox) checkbox.checked = false; 
            closeModal();
        };

        const notaHandler = () => { showNotesModal(clave); };
        const fotoHandler = () => { showPhotoModal('inventory', clave); };
        
        detailView.ubicarBtn.onclick = ubicarHandler;
        detailView.reetiquetarBtn.onclick = reetiquetarHandler;
        detailView.notaBtn.onclick = notaHandler;
        detailView.fotoBtn.onclick = fotoHandler;

        detailView.modal.classList.add('show');
    }

    function showUserDetailView(userId) {
        const user = state.resguardantes.find(u => u.id === userId);
        if (!user) return;

        const { modal, title, name, area, location, photo, noPhoto } = elements.userDetailView;
        
        title.textContent = 'Detalles del Usuario';
        name.textContent = user.name;
        area.textContent = state.areaNames[user.area] || `Área ${user.area}`;
        
        // Mostrar lista completa de ubicaciones
        if (user.locations && user.locations.length > 0) {
            location.innerHTML = `<ul class="list-disc list-inside">${user.locations.map(l => `<li>${l}</li>`).join('')}</ul>`;
        } else {
            location.textContent = user.locationWithId || 'Sin ubicación';
        }
        
        photo.classList.add('hidden');
        noPhoto.classList.remove('hidden');
        photo.src = '';

        // Nota: La foto se carga usando la ubicación principal
        if (state.locationPhotos[user.locationWithId]) {
            photoDB.getItem('photos', `location-${user.locationWithId}`).then(imageBlob => {
                if (imageBlob) {
                    const objectURL = URL.createObjectURL(imageBlob);
                    photo.src = objectURL;
                    photo.onload = () => URL.revokeObjectURL(objectURL);
                    photo.classList.remove('hidden');
                    noPhoto.classList.add('hidden');
                }
            }).catch(() => {
                photo.classList.add('hidden');
                noPhoto.classList.remove('hidden');
                photo.src = '';
            });
        }

        modal.classList.add('show');
    }

    function showAdicionalDetailView(itemId) {
        const item = state.additionalItems.find(i => i.id === itemId);
        if (!item) return;

        // Usamos referencia directa para evitar errores si falta una coma arriba
        const view = elements.adicionalDetailView;
        
        view.title.textContent = 'Detalles del Bien Adicional';
        view.descripcion.textContent = item.descripcion || 'N/A';
        view.clave.textContent = item.clave || 'N/A';
        view.claveAsignada.textContent = item.claveAsignada || 'N/A';
        view.marca.textContent = item.marca || 'N/A';
        view.modelo.textContent = item.modelo || 'N/A';
        view.serie.textContent = item.serie || 'N/A';
        view.area.textContent = item.area || 'N/A';
        view.usuario.textContent = item.usuario || 'N/A';
        
        // Lógica segura para la ubicación específica
        if (view.ubicacionEspecifica) {
            view.ubicacionEspecifica.textContent = item.ubicacionEspecifica || 'No registrada (Ubicación general)';
        }

        let tipoText = 'Institucional';
        if (item.personal === 'Si') {
            if (item.tieneFormatoEntrada === true) {
                tipoText = 'Personal (Con Formato)';
            } else if (item.tieneFormatoEntrada === false) {
                tipoText = 'Personal (Sin Formato - Regularizar)';
            } else {
                tipoText = 'Personal (Estado de formato no registrado)';
            }
        }
        view.tipo.textContent = tipoText;

        view.photo.classList.add('hidden');
        view.noPhoto.classList.remove('hidden');
        view.photo.src = '';

        if (state.additionalPhotos[item.id]) {
            photoDB.getItem('photos', `additional-${item.id}`).then(imageBlob => {
                if (imageBlob) {
                    const objectURL = URL.createObjectURL(imageBlob);
                    view.photo.src = objectURL;
                    view.photo.onload = () => URL.revokeObjectURL(objectURL);
                    view.photo.classList.remove('hidden');
                    view.noPhoto.classList.add('hidden');
                }
            }).catch(() => {
                view.photo.classList.add('hidden');
                view.noPhoto.classList.remove('hidden');
                view.photo.src = '';
            });
        }
        
        view.modal.classList.add('show');
    }
    function showNotesModal(clave) {
        const selectedClaves = clave ? [clave] : Array.from(document.querySelectorAll('.inventory-item-checkbox:checked')).map(cb => cb.closest('tr').dataset.clave);
        if (selectedClaves.length === 0) {
            if(!clave) return showToast('Seleccione al menos un bien.', 'error');
            return;
        }

        if (selectedClaves.length > 1) {
            elements.noteTextarea.value = '';
            elements.noteTextarea.placeholder = `Añadir una nota a los ${selectedClaves.length} bienes seleccionados...`;
        } else {
            elements.noteTextarea.value = state.notes[selectedClaves[0]] || '';
            elements.noteTextarea.placeholder = 'Escribe tu nota aquí...';
        }

        elements.noteSaveBtn.dataset.claves = JSON.stringify(selectedClaves);
        elements.notesModal.classList.add('show');
    }

    function showQrModal(clave) {
        const { modal, container, title } = elements.qrDisplayModal;
        container.innerHTML = '';
        title.textContent = `Código QR del Bien: ${clave}`;
        new QRCode(container, {
            text: clave,
            width: 200,
            height: 200,
            correctLevel: QRCode.CorrectLevel.H
        });
        modal.classList.add('show');
    }

    // Función helper para renderizar los chips en el modal de edición
    const renderEditTempLocations = () => {
        const container = elements.editUserModal.locationsList;
        if (!tempEditUserLocations || tempEditUserLocations.length === 0) {
            container.innerHTML = '<p class="text-xs text-gray-400 italic w-full text-center self-center empty-msg">Añade al menos una ubicación con el botón +</p>';
            return;
        }
        
        container.innerHTML = tempEditUserLocations.map((loc, index) => `
            <span class="location-chip">
                ${loc}
                <i class="fa-solid fa-xmark remove-edit-loc-btn cursor-pointer text-red-500 hover:text-red-700 ml-2" data-index="${index}"></i>
            </span>
        `).join('');
    };

    // NUEVO: Función showEditUserModal actualizada para Multi-Ubicación
    function showEditUserModal(index) {
        if (state.readOnlyMode) return;
        const user = state.resguardantes[index];
        
        // Llenar campos básicos
        elements.editUserModal.name.value = user.name;
        elements.editUserModal.areaSelect.value = user.area;
        
        // Cargar ubicaciones existentes a la variable temporal
        // Si el usuario ya tiene array, úsalo. Si es antiguo, usa la propiedad legacy.
        if (user.locations && user.locations.length > 0) {
            tempEditUserLocations = [...user.locations];
        } else {
            // Fallback para usuarios migrados
            tempEditUserLocations = user.locationWithId ? [user.locationWithId] : [];
        }
        
        // Renderizar los chips
        renderEditTempLocations();
        
        // Resetear inputs de agregar ubicación en el modal
        elements.editUserModal.locationType.value = 'OFICINA';
        elements.editUserModal.locationManual.value = '';
        elements.editUserModal.locationManual.classList.add('hidden');
        
        // Guardar referencia al usuario que se edita
        elements.editUserModal.saveBtn.dataset.userIndex = index;
        
        elements.editUserModal.modal.classList.add('show');
    }
    
    function showAreaClosureModal(areaId) {
         logActivity('Preparación Acta Cierre', `Área ${areaId} lista para generar acta.`);
    }

    async function startQrScanner() {
        if (state.readOnlyMode) return;
        elements.qrScannerModal.classList.add('show');
        if (html5QrCode && html5QrCode.isScanning) {
            await html5QrCode.stop();
        }
        html5QrCode = new Html5Qrcode("qr-reader");

        const qrCodeSuccessCallback = (decodedText, decodedResult) => {
            stopQrScanner();
            elements.inventory.searchInput.value = decodedText;
            currentPage = 1;
            filterAndRenderInventory();
            showToast(`Bien con clave ${decodedText} encontrado.`);
            logActivity('Escaneo QR', `Se encontró la clave: ${decodedText}.`);
            changeTab('inventory');
        };

        const config = { fps: 10, qrbox: { width: 250, height: 250 } };
        
        html5QrCode.start(
            { facingMode: "environment" }, 
            config,
            qrCodeSuccessCallback
        ).catch(err => {
            showToast('Error al iniciar la cámara. Revisa los permisos.', 'error');
            console.error("Error al iniciar el escaner QR: ", err);
            stopQrScanner();
        });
    }

    function stopQrScanner() {
        if (html5QrCode && html5QrCode.isScanning) {
            html5QrCode.stop().then(ignore => {
                elements.qrScannerModal.classList.remove('show');
            }).catch(err => {
                elements.qrScannerModal.classList.remove('show');
            });
        } else {
            elements.qrScannerModal.classList.remove('show');
        }
    }
    function populateAdicionalesFilters() {
        const areaSelect = elements.adicionales.areaFilter;
        const userSelect = elements.adicionales.userFilter;
        const selectedArea = areaSelect.value;
        
        areaSelect.innerHTML = '<option value="all">Todas las áreas</option>' + 
            state.areas.map(area => `<option value="${area}">${state.areaNames[area] || area}</option>`).join('');
        areaSelect.value = selectedArea; 
        
        let usersToList = state.resguardantes;
        if (selectedArea !== 'all') {
            usersToList = usersToList.filter(user => user.area === selectedArea);
        }

        const selectedUser = userSelect.value; 

        userSelect.innerHTML = '<option value="all">Todos los usuarios</option>' +
            usersToList.sort((a,b) => a.name.localeCompare(b.name)).map(user => `<option value="${user.name}">${user.name}</option>`).join('');
        
        if (usersToList.some(user => user.name === selectedUser)) {
            userSelect.value = selectedUser;
        } else {
            userSelect.value = 'all';
        }
    }

    function populateAreaSelects() {
        const areasFromInventory = state.inventory.map(item => item.areaOriginal);
        const areasFromUsers = state.resguardantes.map(user => user.area);
        const persistentAreas = state.persistentAreas || [];
        state.areas = [...new Set([...areasFromInventory, ...areasFromUsers, ...persistentAreas])].filter(Boolean).sort();

        [elements.userForm.areaSelect, elements.reports.areaFilter, elements.inventory.areaFilter, elements.editUserModal.areaSelect, elements.adicionales.areaFilter].forEach(select => {
            const selectedValue = select.value;
            const firstOpt = select.id.includes('user-area-select') ? '<option value="">Seleccione</option>' : '<option value="all">Todas</option>';
            select.innerHTML = firstOpt + state.areas.map(area => `<option value="${area}" ${selectedValue === area ? 'selected' : ''}>${state.areaNames[area] || area}</option>`).join('');
            if (selectedValue && !select.querySelector(`option[value="${selectedValue}"]`)) {
                select.value = 'all'; 
            }
        });
    }
    
    function populateReportFilters() {
        const areaSelect = elements.reports.areaFilter;
        const userSelect = elements.reports.userFilter;
        const selectedArea = areaSelect.value;

        areaSelect.innerHTML = '<option value="all">Todas las áreas</option>' + 
            state.areas.map(area => `<option value="${area}">${state.areaNames[area] || area}</option>`).join('');
        areaSelect.value = selectedArea; 
        
        let usersToList = state.resguardantes;
        if (selectedArea !== 'all') {
            usersToList = usersToList.filter(user => user.area === selectedArea);
        }

        const selectedUser = userSelect.value; 

        userSelect.innerHTML = '<option value="all">Todos los usuarios</option>' +
            usersToList.sort((a,b) => a.name.localeCompare(b.name)).map(user => `<option value="${user.name}">${user.name}</option>`).join('');
        
        if (usersToList.some(user => user.name === selectedUser)) {
            userSelect.value = selectedUser;
        } else {
            userSelect.value = 'all';
        }
    }

    function populateBookTypeFilter() {
        const bookTypes = [...new Set(state.inventory.map(item => item.listadoOriginal))].filter(Boolean).sort();
        const select = elements.inventory.bookTypeFilter;
        const staticOptions = Array.from(select.querySelectorAll('option[value]:not([value="all"])')).map(opt => opt.value);
        const allTypes = [...new Set([...staticOptions, ...bookTypes])].sort();
        
        select.innerHTML = '<option value="all">Todos los tipos</option>' + 
            allTypes.map(type => `<option value="${type}">${type}</option>`).join('');
    }
    function exportLabelsToXLSX() {
        const itemsToLabel = state.inventory.filter(item => item['IMPRIMIR ETIQUETA'] === 'SI');
        const additionalItemsToLabel = state.additionalItems.filter(item => item.claveAsignada);

        if (itemsToLabel.length === 0 && additionalItemsToLabel.length === 0) {
            return showToast('No hay bienes marcados para etiquetar.', 'info');
        }
        
        showToast('Generando reporte de etiquetas XLSX...');
        logActivity('Exportación XLSX', `Exportando ${itemsToLabel.length} etiquetas de inventario y ${additionalItemsToLabel.length} de adicionales.`);

        try {
            const inventoryData = itemsToLabel.map(item => {
                const claveUnica = String(item['CLAVE UNICA']);
                
                // Lógica de ubicación precisa
                let locationDisplay = item.ubicacionEspecifica;
                if (!locationDisplay) {
                    const userData = state.resguardantes.find(u => u.name === item['NOMBRE DE USUARIO']);
                    // Si no tiene ubicación específica, intentamos poner la ubicación general del usuario
                    locationDisplay = userData ? (userData.locationWithId || userData.area) : 'N/A';
                }

                return {
                    'Clave única': claveUnica.startsWith('0.') ? claveUnica.substring(1) : claveUnica,
                    'Descripción': item['DESCRIPCION'],
                    'Usuario': item['NOMBRE DE USUARIO'] || 'Sin Asignar',
                    'Ubicación': locationDisplay, // <--- CAMBIADO: Antes era 'Área'
                    'Área': state.resguardantes.find(u => u.name === item['NOMBRE DE USUARIO'])?.area || 'N/A'
                };
            });

            const additionalData = additionalItemsToLabel.map(item => {
                 return {
                    'Clave única': item.claveAsignada,
                    'Descripción': item.descripcion,
                    'Usuario': item.usuario || 'Sin Asignar',
                    'Área': state.resguardantes.find(u => u.name === item.usuario)?.area || 'N/A'
                };
            });

            const combinedData = [...inventoryData, ...additionalData];

            const worksheet = XLSX.utils.json_to_sheet(combinedData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Etiquetas");

            worksheet['!cols'] = [
                { wch: 15 }, { wch: 50 }, { wch: 30 }, { wch: 15 }
            ];

            XLSX.writeFile(workbook, "reporte_etiquetas_combinado.xlsx");
            showToast('Reporte de etiquetas generado con éxito.', 'success');
        } catch (error) {
            console.error("Error generating labels XLSX file:", error);
            showToast('Hubo un error al generar el reporte de etiquetas.', 'error');
        }
    }

    function getDetailedStats() {
        const selectedArea = elements.reports.areaFilter.value;
        const selectedUser = elements.reports.userFilter.value;
        
        let inventoryForStats = state.inventory;
        let additionalForStats = state.additionalItems;

        if (selectedArea !== 'all') {
            inventoryForStats = inventoryForStats.filter(i => i.areaOriginal === selectedArea);
            const usersInArea = state.resguardantes.filter(u => u.area === selectedArea).map(u => u.name);
            additionalForStats = additionalForStats.filter(item => usersInArea.includes(item.usuario));
        }
        if (selectedUser !== 'all') {
            inventoryForStats = inventoryForStats.filter(i => i['NOMBRE DE USUARIO'] === selectedUser);
            additionalForStats = additionalForStats.filter(i => i.usuario === selectedUser);
        }

        const stats = {};
        const groupBy = (arr, key) => arr.reduce((acc, item) => {
            (acc[item[key]] = acc[item[key]] || []).push(item);
            return acc;
        }, {});

        stats.pendingByArea = groupBy(inventoryForStats.filter(i => i.UBICADO === 'NO'), 'areaOriginal');
        stats.assignedByUser = groupBy(inventoryForStats.filter(i => i.UBICADO === 'SI'), 'NOMBRE DE USUARIO');
        const pendingLabels = inventoryForStats.filter(i => i['IMPRIMIR ETIQUETA'] === 'SI');
        stats.labelsByArea = groupBy(pendingLabels, 'areaOriginal');
        stats.labelsByUser = groupBy(pendingLabels, 'NOMBRE DE USUARIO');
        stats.additionalCount = additionalForStats.length; 

        return stats;
    }
    
    function renderReportStats() {
        const stats = getDetailedStats();
        
        let html = `<p class="font-bold">Bienes Adicionales Registrados: <span class="font-normal">${stats.additionalCount}</span></p><hr class="my-2 border-gray-300 dark:border-gray-600">`;

        const generateHtmlList = (title, data) => {
            let listHtml = `<div class="mb-2"><p class="font-bold">${title}</p>`;
            const entries = Object.entries(data);
            if (entries.length === 0) {
                listHtml += `<p class="text-gray-500 text-xs">No hay datos.</p></div>`;
                return listHtml;
            }
            listHtml += '<ul class="list-disc list-inside">';
            entries.forEach(([key, value]) => {
                listHtml += `<li><strong>${key || 'Sin Asignar'}:</strong> ${value.length}</li>`;
            });
            listHtml += '</ul></div>';
            return listHtml;
        };

        html += generateHtmlList('Bienes Asignados por Usuario:', stats.assignedByUser);
        html += '<hr class="my-2 border-gray-300 dark:border-gray-600">';
        html += generateHtmlList('Bienes Pendientes por Área:', stats.pendingByArea);
        html += '<hr class="my-2 border-gray-300 dark:border-gray-600">';
        html += generateHtmlList('Etiquetas Pendientes por Usuario:', stats.labelsByUser);
        html += '<hr class="my-2 border-gray-300 dark:border-gray-600">';
        html += generateHtmlList('Etiquetas Pendientes por Área:', stats.labelsByArea);

        elements.reports.stats.innerHTML = html;
    }

    function generateSimplePendingReport(options = {}) {
        const { areaDisplay = 'Todas las Áreas', entrega, recibe, date } = options;
        const selectedArea = elements.reports.areaFilter.value;
        let pendingItems = state.inventory.filter(item => item.UBICADO === 'NO');

        if (selectedArea !== 'all') {
            pendingItems = pendingItems.filter(item => item.areaOriginal === selectedArea);
        }

        if (pendingItems.length === 0) {
            return showToast('No hay bienes pendientes para los filtros seleccionados.', 'info');
        }
        
        logActivity('Reporte Impreso', `Impresión de reporte de ${pendingItems.length} pendientes.`);

        const template = elements.printTemplates.simplePending;
        
        document.getElementById('print-simple-pending-area').textContent = areaDisplay;
        document.getElementById('print-simple-pending-author-name').textContent = entrega;
        document.getElementById('print-simple-pending-responsible-name').textContent = recibe;

        const tableHead = template.querySelector('thead');
        const tableBody = template.querySelector('tbody');
        
        tableHead.innerHTML = `<tr>
            <th class="col-num">#</th>
            <th class="col-clave">Clave</th>
            <th class="col-desc">Descripción</th>
            <th class="col-marca">Marca</th>
            <th class="col-modelo">Modelo</th>
            <th class="col-serie">Serie</th>
            <th class="col-area">Área Orig.</th>
        </tr>`;

        tableBody.innerHTML = pendingItems.map((item, index) => {
            return `<tr>
            <td class="col-num"></td>
            <td class="col-clave">${item['CLAVE UNICA'] || ''}</td>
            <td class="col-desc">${item['DESCRIPCION'] || ''}</td>
            <td class="col-marca">${item['MARCA'] || ''}</td>
            <td class="col-modelo">${item['MODELO'] || ''}</td>
            <td class="col-serie">${item['SERIE'] || ''}</td>
            <td class="col-area">${item.areaOriginal || ''}</td>
        </tr>`;
        }).join('');
        
        preparePrint('print-simple-pending', { date });
    }
    
    function generatePrintableResguardo(title, user, items, isAdicional = false, options = {}) {
        const {
            areaFullName = 'Área no especificada',
            entrega,
            recibe,
            recibeCargo, 
            isForArea, 
            isForUser,  
            date
        } = options;

         if (!user || user === 'all') {
            return showToast('Por favor, selecciona un usuario o área para generar el informe.', 'error');
        }
        if (items.length === 0) {
            return showToast(`No se encontraron bienes para el filtro seleccionado.`, 'error');
        }
        
        logActivity('Resguardo Impreso', `Resguardo para ${user} con ${items.length} bienes.`);
        
        const template = elements.printTemplates.resguardo;
        const signaturesContainer = document.getElementById('print-resguardo-signatures');
        const responsibleTitleEl = document.getElementById('print-resguardo-responsible-title');

        const responsibleNames = Object.values(state.areaDirectory).map(area => area.name);
        const isResponsableDeArea = responsibleNames.includes(user);

        let recibeTitulo = 'Usuario Resguardante'; 
        if (isAdicional && isForArea) {
            recibeTitulo = 'Responsable de Área';
        } else if (isResponsableDeArea) {
            recibeTitulo = 'Responsable de Área';
        }

        if ((isAdicional && isForArea) || isResponsableDeArea) {
            signaturesContainer.classList.add('center-single');
            responsibleTitleEl.textContent = recibeTitulo;
        } else {
            signaturesContainer.classList.remove('center-single');
            responsibleTitleEl.textContent = recibeTitulo;
        }

        document.getElementById('print-resguardo-title').textContent = title;
        document.getElementById('print-resguardo-area').textContent = areaFullName;
        
        const responsibleName = (isAdicional && isForArea) ? areaFullName : user;
        
        let introText = '';
        if (isAdicional && isForArea) { 
            introText = `Por medio de la presente, se hace constar que <strong>${responsibleName}</strong> recibe para su uso, resguardo y custodia los bienes que se detallan en el presente documento, comprometiéndose a su correcto uso y cuidado.`;
        } else { 
            if (isResponsableDeArea) {
                introText = `Quedo enterado, <strong>${user}</strong> que los Bienes Muebles que se encuentran listados en el presente resguardo, están a partir de la firma del mismo, bajo mi buen uso, custodia, vigilancia y conservación, en caso de daño, robo o extravío, se deberá notificar de inmediato a el Área Administrativa o Comisión para realizar el trámite administrativo correspondiente, por ningún motivo se podrá cambiar o intercambiar los bienes sin previa solicitud y autorización del Área Administrativa o Comisión.`;
            } else {
                introText = `Quedo enterado, <strong>${user}</strong> que los Bienes Muebles que se encuentran listados en el presente resguardo, están a partir de la firma del mismo, bajo mi buen uso, custodia, vigilancia y conservación, en caso de daño, robo o extravio, deberé notificar al jefe inmediato del Área Administrativa o Comisión para realizar el trámite administrativo correspondiente. Por ningún motivo se podra cambiar o intercambiar los bienes sin previa solicitud y autorización del jefe inmediato y/o responsable del inventario.`;
            }
        }
        
        document.getElementById('print-resguardo-text').innerHTML = introText;

        const tableHead = template.querySelector('thead');
        const tableBody = template.querySelector('tbody');
        
        const headerHtml = `<tr>
            <th class="col-num">#</th>
            <th class="col-clave">Clave</th>
            <th class="col-desc">Descripción</th>
            <th class="col-marca">Marca</th>
            <th class="col-modelo">Modelo</th>
            <th class="col-serie">Serie</th>
            <th class="col-area">Área Orig./Proc.</th>
            <th class="col-usuario">Usuario</th>
            <th class="col-status">Tipo de Bien</th>
        </tr>`;
        tableHead.innerHTML = headerHtml;

        const showAsteriskNote = items.some(item => 
            !!item.id && item.personal === 'Si' && item.tieneFormatoEntrada === false
        );

        tableBody.innerHTML = items.map((item, index) => {
            const isItemAdicional = !!item.id; 
            let asteriskHtml = '';
            if (isItemAdicional && item.personal === 'Si' && item.tieneFormatoEntrada === false) {
                asteriskHtml = ' <strong>*</strong>'; 
            }

            const desc = String(item.descripcion || item.DESCRIPCION || '');
            const clave = isItemAdicional ? (item.claveAsignada || item.clave || 'S/C') : item['CLAVE UNICA'];
            const marca = item.marca || item.MARCA;
            const modelo = item.modelo || item.MODELO;
            const serie = item.serie || item.SERIE;
            const areaCol = isItemAdicional ? (item.area || 'N/A') : item.areaOriginal;
            const usuarioCol = item.usuario || item['NOMBRE DE USUARIO'] || 'N/A';
            
            let statusCol = '';
            const itemArea = String(item.area || '').toUpperCase(); 

            if (isItemAdicional) {
                if (itemArea.includes('CONTRATO')) {
                    statusCol = 'Arrendamiento';
                } else if (item.personal === 'Si') {
                    statusCol = 'Personal';
                } else {
                    statusCol = 'Cámara/Controlable';
                }
            } else {
                statusCol = 'Institucional'; 
            }

            return `<tr>
            <td class="col-num"></td>
            <td class="col-clave">${clave || ''}</td>
            <td class="col-desc">${desc}${asteriskHtml}</td>
            <td class="col-marca">${marca || ''}</td>
            <td class="col-modelo">${modelo || ''}</td>
            <td class="col-serie">${serie || ''}</td>
            <td class="col-area">${areaCol}</td>
            <td class="col-usuario">${usuarioCol}</td>
            <td class="col-status">${statusCol}</td>
        </tr>`;
        }).join('');

        const noteElement = document.getElementById('print-resguardo-note');
        if (showAsteriskNote) {
            noteElement.innerHTML = '<strong>* Favor de realizar entrada</strong>';
        } else {
            noteElement.innerHTML = ''; 
        }

        document.getElementById('print-resguardo-count').textContent = `Total de Bienes: ${items.length}`;
        
        document.getElementById('print-resguardo-author-name').textContent = entrega;
        document.getElementById('print-resguardo-author-title').textContent = recibeCargo; 
        document.getElementById('print-resguardo-responsible-name').textContent = recibe;

        preparePrint('print-resguardo', { date });
    }
    function generateAreaClosureReport(options = {}) {
        const { areaId, responsible, location, areaFullName, entrega, recibe, recibeCargo, date } = options;

        const areaItems = state.inventory.filter(item => item.areaOriginal === areaId);
        const usersInArea = state.resguardantes.filter(user => user.area === areaId).map(user => user.name);
        const additionalItemsInArea = state.additionalItems.filter(item => usersInArea.includes(item.usuario));
        const personalItemsInArea = additionalItemsInArea.filter(i => i.personal === 'Si').length;
        const labelsToPrintInArea = areaItems.filter(i => i['IMPRIMIR ETIQUETA'] === 'SI').length;
        
        const statsByUser = areaItems.reduce((acc, item) => {
            const user = item['NOMBRE DE USUARIO'];
            if (user) acc[user] = (acc[user] || 0) + 1;
            return acc;
        }, {});

        let userStatsHtml = '';
        const userEntries = Object.entries(statsByUser);
        if (userEntries.length > 0) {
            const userMap = new Map(state.resguardantes.map(user => [user.name, user]));
            userStatsHtml = '<h2>Estadísticas Detalladas por Usuario</h2><ul>';
            userEntries.sort((a,b) => b[1] - a[1]).forEach(([user, count]) => {
                const userData = userMap.get(user);
                const userDetails = userData ? ` (Área: ${userData.area}, Ubicación: ${userData.locationWithId})` : '';
                userStatsHtml += `<li><strong>${user}</strong>${userDetails}: ${count} bien(es) asignado(s)</li>`;
            });
            userStatsHtml += '</ul>';
        }

        const itemsARegularizar = additionalItemsInArea.filter(item => item.personal === 'Si' && item.tieneFormatoEntrada === false);
        let regularizarHtml = '';
        if(itemsARegularizar.length > 0) {
            regularizarHtml = '<h2>Acciones de Seguimiento</h2><ul>';
            itemsARegularizar.forEach(item => {
                regularizarHtml += `<li>Recuerda a <strong>${item.usuario}</strong> que debe regularizar la entrada de: <em>${item.descripcion}</em>.</li>`;
            });
            regularizarHtml += '</ul>';
        }

        document.getElementById('print-area-closure-name').textContent = areaFullName;
        document.getElementById('print-area-closure-info').innerHTML = `
            <div><b>Responsable del Área:</b> ${recibeCargo || 'No especificado'}</div>
            <div><b>Cargo:</b> ${recibeCargo || 'No especificado'}</div>
            <div><b>Responsable que Recibe:</b> ${responsible}</div>
            <div><b>Ubicación de Firma:</b> ${location}</div>
        `;
        
        document.getElementById('print-area-closure-summary').innerHTML = `
            <h2>Resumen Estadístico del Área</h2>
            <div class="print-summary-grid">
                <div class="print-summary-item"><strong>${areaItems.length}</strong><span>Bienes del Inventario Original</span></div>
                <div class="print-summary-item"><strong>${areaItems.length}</strong><span>Bienes Ubicados</span></div>
                <div class="print-summary-item"><strong>${additionalItemsInArea.length}</strong><span>Bienes Adicionales Encontrados</span></div>
                <div class="print-summary-item"><strong>${personalItemsInArea}</strong><span>Bienes Adicionales (Personales)</span></div>
                <div class="print-summary-item"><strong>${labelsToPrintInArea}</strong><span>Bienes por Re-etiquetar</span></div>
                <div class="print-summary-item"><strong>${usersInArea.length}</strong><span>Usuarios en el Área</span></div>
            </div>
        `;

        document.getElementById('print-area-closure-users').innerHTML = userStatsHtml;
        document.getElementById('print-area-closure-actions').innerHTML = regularizarHtml;
        
        document.getElementById('print-area-closure-author-name').textContent = entrega;
        document.getElementById('print-area-closure-responsible-name').textContent = recibe;

        state.closedAreas[areaId] = { responsible, location, date: new Date().toISOString() };
        logActivity('Acta de área cerrada', `Área: ${areaId}, Responsable que recibe: ${responsible}`);
        saveState();
        renderLoadedLists(); 

        preparePrint('print-area-closure', { date });
    }
    function generateTasksReport(options = {}) {
        const { date } = options;
        
        // 1. Recolectar datos
        const itemsWithNotes = Object.keys(state.notes).filter(key => state.notes[key].trim() !== '');
        const itemsWithPendingLabels = state.inventory.filter(item => item['IMPRIMIR ETIQUETA'] === 'SI');
        const mismatchedItems = state.inventory.filter(item => item.areaIncorrecta);
        const itemsARegularizar = state.additionalItems.filter(item => item.personal === 'Si' && item.tieneFormatoEntrada === false);
        
        // 2. Agrupar Etiquetas por Usuario (Simplificación masiva)
        const labelsByUser = itemsWithPendingLabels.reduce((acc, item) => {
            const user = item['NOMBRE DE USUARIO'] || 'Sin Asignar';
            if (!acc[user]) acc[user] = 0;
            acc[user]++;
            return acc;
        }, {});

        // 3. Agrupar Notas por Clave (Lista compacta)
        const notesList = itemsWithNotes.map(clave => {
            const item = state.inventory.find(i => i['CLAVE UNICA'] === clave);
            const desc = item ? item['DESCRIPCION'].substring(0, 30) + '...' : 'Desconocido';
            return `<li><b>${clave} (${desc}):</b> ${state.notes[clave]}</li>`;
        }).join('');

        let contentHtml = '';

        if (itemsWithNotes.length === 0 && itemsWithPendingLabels.length === 0 && mismatchedItems.length === 0 && itemsARegularizar.length === 0) {
            contentHtml = '<div style="padding:20px; text-align:center; border: 2px dashed #ccc;"><h2>¡Excelente! No hay acciones pendientes.</h2><p>El inventario está limpio.</p></div>';
        } else {
            contentHtml += `<div class="print-section" style="margin-top:0;">`;
            
            // SECCIÓN 1: URGENTE - REGULARIZACIÓN (Tabla compacta)
            if (itemsARegularizar.length > 0 || mismatchedItems.length > 0) {
                contentHtml += `<h3>🔴 Prioridad Alta: Regularización y Ubicación</h3>`;
                contentHtml += `<table class="print-table" style="margin-bottom: 15px;">
                    <thead><tr><th style="width:150px;">Tipo</th><th>Detalle y Acción</th></tr></thead>
                    <tbody>`;
                
                itemsARegularizar.forEach(item => {
                    contentHtml += `<tr>
                        <td><strong>Bien Personal s/Formato</strong></td>
                        <td>Usuario: <strong>${item.usuario}</strong>. Bien: ${item.descripcion}. <br><em>Acción: Solicitar formato de entrada o retirar bien.</em></td>
                    </tr>`;
                });

                mismatchedItems.forEach(item => {
                    const user = item['NOMBRE DE USUARIO'];
                    contentHtml += `<tr>
                        <td><strong>Bien Fuera de Área</strong></td>
                        <td>Clave: <strong>${item['CLAVE UNICA']}</strong> (${item.DESCRIPCION}). <br>Área Orig: ${item.areaOriginal} -> Actual: ${state.resguardantes.find(u=>u.name===user)?.area || '?'}. <br><em>Acción: Transferir administrativamente o devolver físicamente.</em></td>
                    </tr>`;
                });
                contentHtml += `</tbody></table>`;
            }

            // SECCIÓN 2: ADMINISTRATIVO - ETIQUETAS (Resumido)
            if (Object.keys(labelsByUser).length > 0) {
                contentHtml += `<h3>🟡 Administrativo: Etiquetado Pendiente</h3>`;
                contentHtml += `<div style="display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 15px;">`;
                Object.entries(labelsByUser).forEach(([user, count]) => {
                    contentHtml += `<div style="border:1px solid #ccc; padding: 5px 10px; background: #f9f9f9;">
                        <strong>${user}:</strong> ${count} etiquetas
                    </div>`;
                });
                contentHtml += `</div>`;
            }

            // SECCIÓN 3: NOTAS DE SEGUIMIENTO
            if (notesList.length > 0) {
                contentHtml += `<h3>🔵 Notas de Seguimiento</h3>`;
                contentHtml += `<ul style="font-size: 0.85em; margin-top:5px;">${notesList}</ul>`;
            }

            contentHtml += `</div>`;
        }

        if (options.returnAsHtml) {
            return `<h2>Plan de Acción Ejecutivo</h2>${contentHtml}`;
        }

        logActivity('Plan de Acción', 'Generado plan ejecutivo simplificado.');
        
        document.getElementById('print-tasks-content').innerHTML = contentHtml;
        preparePrint('print-tasks-report', { date });
    }
    

    function generateInventoryReport(options = {}) {
        const selectedUser = elements.reports.userFilter.value;
        const selectedArea = elements.reports.areaFilter.value;

        let reportedItems = state.inventory;

        if (selectedUser !== 'all') {
            reportedItems = reportedItems.filter(item => item['NOMBRE DE USUARIO'] === selectedUser);
        }
        if (selectedArea !== 'all') {
            reportedItems = reportedItems.filter(item => item.areaOriginal === selectedArea);
        }

        renderReportTable(reportedItems, 'Reporte de Inventario', { 
            headers: ['Clave Única', 'Descripción', 'Marca', 'Modelo', 'Serie', 'Usuario', 'Ubicado', 'Área Original'] 
        });
        showToast(`Reporte generado con ${reportedItems.length} bienes.`);
    }
    
    function exportInventoryToXLSX() {
        const selectedArea = elements.reports.areaFilter.value;
        let inventoryToExport = state.inventory;
        let additionalToExport = state.additionalItems;
        let fileName = "inventario_completo.xlsx";
        let logMessage = "Exportando inventario completo.";

        // Filtros de área
        if (selectedArea !== 'all') {
            inventoryToExport = state.inventory.filter(item => item.areaOriginal === selectedArea);
            
            const usersInArea = state.resguardantes
                .filter(user => user.area === selectedArea)
                .map(user => user.name);
            additionalToExport = state.additionalItems.filter(item => usersInArea.includes(item.usuario));
            
            fileName = `inventario_area_${selectedArea}.xlsx`;
            logMessage = `Exportando inventario y adicionales para el área ${selectedArea}.`;
        }

        if (inventoryToExport.length === 0 && additionalToExport.length === 0) {
            return showToast('No hay datos para exportar con los filtros actuales.', 'warning');
        }

        showToast('Generando archivo XLSX con ubicaciones específicas...');
        logActivity('Exportación XLSX', logMessage);

        try {
            const workbook = XLSX.utils.book_new();

            // --- 1. PROCESAR INVENTARIO PRINCIPAL ---
            const inventoryData = inventoryToExport.map(item => {
                // LÓGICA DE UBICACIÓN:
                // Si tiene ubicación específica (anclaje), úsala.
                // Si no (es un bien viejo), intenta usar la ubicación principal del usuario.
                let locationDisplay = item.ubicacionEspecifica;
                
                if (!locationDisplay) {
                    const userData = state.resguardantes.find(u => u.name === item['NOMBRE DE USUARIO']);
                    if (userData) {
                        locationDisplay = userData.locationWithId || 'Ubicación General';
                    } else {
                        locationDisplay = 'N/A';
                    }
                }

                return {
                    'Clave Unica': String(item['CLAVE UNICA']).startsWith('0.') ? item['CLAVE UNICA'].substring(1) : item['CLAVE UNICA'],
                    'Descripcion': item['DESCRIPCION'],
                    'Marca': item['MARCA'],
                    'Modelo': item['MODELO'],
                    'Serie': item['SERIE'],
                    'Area Original': item.areaOriginal,
                    'Usuario Asignado': item['NOMBRE DE USUARIO'],
                    'Ubicación': locationDisplay, // <--- CAMBIO AQUÍ (Antes decía 'Ubicaciones Usuario')
                    'Ubicado': item['UBICADO'],
                    'Requiere Etiqueta': item['IMPRIMIR ETIQUETA'],
                    'Tiene Foto': state.photos[item['CLAVE UNICA']] ? 'Si' : 'No',
                    'Nota': state.notes[item['CLAVE UNICA']] || ''
                };
            });

            const inventoryWorksheet = XLSX.utils.json_to_sheet(inventoryData);
            inventoryWorksheet['!cols'] = [
                { wch: 15 }, { wch: 50 }, { wch: 20 }, { wch: 20 }, { wch: 25 },
                { wch: 15 }, { wch: 30 }, { wch: 35 }, { wch: 10 }, { wch: 15 }, { wch: 10 }, { wch: 50 }
            ];
            XLSX.utils.book_append_sheet(workbook, inventoryWorksheet, "Inventario Principal");

            // --- 2. PROCESAR BIENES ADICIONALES ---
            if (additionalToExport.length > 0) {
                const additionalData = additionalToExport.map(item => {
                     // Misma lógica de ubicación para adicionales
                     let locationDisplay = item.ubicacionEspecifica;
                     if (!locationDisplay) {
                         const userData = state.resguardantes.find(u => u.name === item.usuario);
                         if (userData) {
                             locationDisplay = userData.locationWithId || 'Ubicación General';
                         } else {
                             locationDisplay = 'N/A';
                         }
                     }

                    return {
                        'Descripcion': item.descripcion,
                        'Clave Original': item.clave || 'N/A',
                        'Marca': item.marca || 'N/A',
                        'Modelo': item.modelo || 'N/A',
                        'Serie': item.serie || 'N/A',
                        'Area Procedencia': item.area || 'N/A',
                        'Usuario Asignado': item.usuario,
                        'Ubicación': locationDisplay, // <--- CAMBIO AQUÍ
                        'Es Personal': item.personal,
                        'Clave Asignada (Regularizado)': item.claveAsignada || 'N/A'
                    };
                });
                const additionalWorksheet = XLSX.utils.json_to_sheet(additionalData);
                additionalWorksheet['!cols'] = [
                    { wch: 50 }, { wch: 15 }, { wch: 20 }, { wch: 20 }, { wch: 25 },
                    { wch: 20 }, { wch: 30 }, { wch: 35 }, { wch: 12 }, { wch: 25 }
                ];
                XLSX.utils.book_append_sheet(workbook, additionalWorksheet, "Bienes Adicionales");
            }

            XLSX.writeFile(workbook, fileName);
            showToast('Archivo XLSX generado con ubicaciones específicas.', 'success');
        } catch (error) {
            console.error("Error generating XLSX file:", error);
            showToast('Hubo un error al generar el archivo XLSX.', 'error');
        }
    }
// CONTINÚA EN LA PARTE 4
// ... VIENE DE LA PARTE 3

    function generateSessionSummary(options = {}) {
        const { author, areaResponsible, location, date } = options;
        const userMap = new Map(state.resguardantes.map(user => [user.name, user]));
        
        logActivity('Resumen de Sesión', 'Generado resumen de sesión.');

        const involvedAreas = [...new Set(state.inventory.map(i => i.areaOriginal))];
        const totalAdditional = state.additionalItems.length;
        const personalAdditional = state.additionalItems.filter(i => i.personal === 'Si').length;
        const labelsToPrint = state.inventory.filter(i => i['IMPRIMIR ETIQUETA'] === 'SI').length;
        const itemsLocated = state.inventory.filter(i => i.UBICADO === 'SI').length;
        const itemsPending = state.inventory.length - itemsLocated;

        // --- INICIO NUEVA LÓGICA DE TIEMPO (DÍAS ACTIVOS) ---
        const activeDates = new Set();
        
        // 1. Buscar fechas en bienes ubicados del inventario
        state.inventory.forEach(item => {
            if (item.UBICADO === 'SI' && item.fechaUbicado) {
                // Extraer solo la parte de la fecha YYYY-MM-DD
                const day = item.fechaUbicado.split('T')[0];
                activeDates.add(day);
            }
        });

        // 2. Buscar fechas en bienes adicionales registrados
        state.additionalItems.forEach(item => {
            if (item.fechaRegistro) {
                const day = item.fechaRegistro.split('T')[0];
                activeDates.add(day);
            }
        });

        const daysCount = activeDates.size;
        // Ordenar las fechas para mostrarlas (opcional, visualmente útil)
        const sortedDates = Array.from(activeDates).sort();
        const firstDate = sortedDates[0] || 'N/A';
        const lastDate = sortedDates[sortedDates.length - 1] || 'N/A';

        // Texto a mostrar en el reporte
        const duration = daysCount > 0 
            ? `${daysCount} día(s) de actividad efectiva (${firstDate} al ${lastDate})` 
            : 'Sin actividad registrada aún';
        // --- FIN NUEVA LÓGICA ---

        const itemsARegularizar = state.additionalItems.filter(item => item.personal === 'Si' && item.tieneFormatoEntrada === false);
        let regularizarHtml = '';
        if(itemsARegularizar.length > 0) {
            regularizarHtml = '<h2>Acciones de Seguimiento</h2><ul>';
            itemsARegularizar.forEach(item => {
                regularizarHtml += `<li>Recuerda a <strong>${item.usuario}</strong> que debe regularizar la entrada de: <em>${item.descripcion}</em>.</li>`;
            });
            regularizarHtml += '</ul>';
        }
        
        document.getElementById('print-session-location').innerHTML = `<b>Ubicación Física del Inventario:</b> ${location}`;
        
        let areasHtml = '';
        if (involvedAreas.length > 0) {
            areasHtml = '<h2>Detalle de Áreas</h2><ul>';
            involvedAreas.forEach(area => {
                const areaName = state.areaNames[area] || `Área ${area}`;
                const totalBienes = state.inventory.filter(i => i.areaOriginal === area).length;
                areasHtml += `<li><strong>${areaName}:</strong> Cuenta con <strong>${totalBienes}</strong> bienes.</li>`;
            });
            areasHtml += '</ul>';
        }
        document.getElementById('print-session-areas').innerHTML = areasHtml;

        const locationCounts = state.resguardantes.reduce((acc, user) => {
            // Conteo robusto
            if (user.locations && user.locations.length > 0) {
                user.locations.forEach(loc => {
                     const base = loc.replace(/\s\d+$/, '');
                     acc[base] = (acc[base] || 0) + 1;
                });
            } else {
                const loc = user.location || 'Sin Ubicación';
                acc[loc] = (acc[loc] || 0) + 1;
            }
            return acc;
        }, {});

        let locationsHtml = '';
        if (Object.keys(locationCounts).length > 0) {
            locationsHtml = '<h2>Distribución de Ubicaciones Físicas</h2><ul>';
            Object.entries(locationCounts).forEach(([location, count]) => {
                locationsHtml += `<li><strong>${location}:</strong> ${count} ${count > 1 ? 'registros' : 'registro'}</li>`;
            });
            locationsHtml += '</ul>';
        }
        document.getElementById('print-session-locations').innerHTML = locationsHtml;

        document.getElementById('print-session-stats-general').innerHTML = `
            <h2>Estadísticas Generales</h2>
            <div class="print-summary-grid">
                <div class="print-summary-item"><strong>${itemsLocated}</strong><span>Bienes Encontrados</span></div>
                <div class="print-summary-item"><strong>${itemsPending}</strong><span>Bienes Pendientes de Ubicar</span></div>
                <div class="print-summary-item"><strong>${totalAdditional}</strong><span>Bienes Adicionales Encontrados</span></div>
                <div class="print-summary-item"><strong>${personalAdditional}</strong><span>Bienes Personales</span></div>
                <div class="print-summary-item"><strong>${labelsToPrint}</strong><span>Bienes por etiquetar</span></div>
                <div class="print-summary-item"><strong>${duration}</strong><span>Tiempo de Trabajo</span></div>
            </div>
        `;

        const stats = getDetailedStats();
        let statsHtml = '<h2>Estadísticas Detalladas de la Sesión</h2>';
        let assignedUsersHtml = `<div class="print-section"><p><strong>Bienes Asignados por Usuario</strong></p>`;
        const assignedEntries = Object.entries(stats.assignedByUser);
        if (assignedEntries.length === 0) {
            assignedUsersHtml += `<p style="color: #666;">No hay datos.</p></div>`;
        } else {
            assignedUsersHtml += '<ul>';
            assignedEntries.forEach(([user, items]) => {
                const userData = userMap.get(user);
                const userDetails = userData ? ` (Área: ${userData.area}, Ubicación: ${userData.locationWithId})` : '';
                assignedUsersHtml += `<li><strong>${user}</strong>${userDetails}: ${items.length}</li>`;
            });
            assignedUsersHtml += '</ul></div>';
        }
        statsHtml += assignedUsersHtml;
        
        const generateHtmlList = (title, data) => {
            let listHtml = `<div class="print-section"><p><strong>${title}</strong></p>`;
            const entries = Object.entries(data);
            if (entries.length === 0) {
                listHtml += `<p style="color: #666;">No hay datos.</p></div>`;
                return listHtml;
            }
            listHtml += '<ul>';
            entries.forEach(([key, value]) => {
                listHtml += `<li><strong>${key || 'Sin Asignar'}:</strong> ${value.length}</li>`;
            });
            listHtml += '</ul></div>';
            return listHtml;
        };

        statsHtml += generateHtmlList('Bienes Pendientes por Área:', stats.pendingByArea);
        statsHtml += generateHtmlList('Etiquetas Pendientes por Usuario:', stats.labelsByUser);
        statsHtml += generateHtmlList('Etiquetas Pendientes por Área:', stats.labelsByArea);

        document.getElementById('print-session-stats-detailed').innerHTML = statsHtml;
        document.getElementById('print-session-actions').innerHTML = regularizarHtml;

        document.getElementById('print-session-author-name').textContent = author;
        document.getElementById('print-session-responsible-name').textContent = areaResponsible;
        
        preparePrint('print-session-summary', { date });
    }

    async function exportSession(isFinal = false) {
        const { overlay, text } = elements.loadingOverlay;
        const type = isFinal ? 'FINALIZADO' : 'backup-editable';
        text.textContent = 'Generando archivo de respaldo...';
        overlay.classList.add('show');
    
        try {
            const zip = new JSZip();
    
            const stateToSave = { ...state };
            if (isFinal) {
                stateToSave.readOnlyMode = true; 
            }
            delete stateToSave.serialNumberCache;
            delete stateToSave.cameraStream;
            zip.file("session.json", JSON.stringify(stateToSave));
    
            text.textContent = 'Empaquetando fotos...';
            const allPhotos = await photoDB.getAllItems('photos');
            if (allPhotos.length > 0) {
                const photoFolder = zip.folder("photos");
                for (const { key, value } of allPhotos) {
                    photoFolder.file(key, value);
                }
            }
            
            text.textContent = 'Empaquetando imágenes de croquis...';
            const allLayoutImages = await photoDB.getAllItems('layoutImages');
             if (allLayoutImages.length > 0) {
                const layoutImageFolder = zip.folder("layoutImages");
                for (const { key, value } of allLayoutImages) {
                    layoutImageFolder.file(key, value);
                }
            }
            
            text.textContent = 'Comprimiendo archivo...';
            const content = await zip.generateAsync({ type: "blob" });
    
            const a = document.createElement('a');
            const date = new Date().toISOString().slice(0, 10);
            a.href = URL.createObjectURL(content);
            a.download = `inventario-${type}-${date}.zip`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(a.href);
    
            logActivity('Sesión exportada', `Tipo: ${type}`);
            showToast(`Sesión ${isFinal ? 'finalizada y' : ''} exportada como .zip`);
        } catch (e) {
            console.error('Error al exportar la sesión como .zip:', e);
            showToast('Error al exportar la sesión.', 'error');
        } finally {
            overlay.classList.remove('show');
        }
    }
    function renderLoadedLists() {
        const container = elements.settings.loadedListsContainer;
        const countEl = document.getElementById('loaded-lists-count');
        container.innerHTML = '';

        const loadedListsMap = new Map();
        state.inventory.forEach(item => {
            if (!loadedListsMap.has(item.listId)) {
                loadedListsMap.set(item.listId, {
                    listId: item.listId,
                    fileName: item.fileName,
                    areaOriginal: item.areaOriginal,
                    listadoOriginal: item.listadoOriginal,
                    printDate: item.printDate || 'N/D' // Agregado
                });
            }
        });
        const loadedLists = Array.from(loadedListsMap.values());


        countEl.textContent = `Total: ${loadedLists.length}`;

        if (loadedLists.length === 0) {
            container.innerHTML = '<p class="text-gray-500">No hay listados cargados.</p>';
            return;
        }

        loadedLists.forEach(list => {
            const item = document.createElement('div');
            item.className = 'flex items-center justify-between p-3 rounded-lg bg-gray-100 dark:bg-slate-800';
            
            const areaId = list.areaOriginal;
            const isAreaCompleted = !!state.completedAreas[areaId];
            const isAreaClosed = !!state.closedAreas[areaId];
            let areaActionButtonHtml = '';

            if (isAreaClosed) {
                areaActionButtonHtml = `
                    <button data-area-id="${areaId}" class="reprint-area-report-btn px-3 py-1 rounded-lg text-xs font-bold text-white bg-blue-500 hover:bg-blue-600">Reimprimir Acta</button>
                `;
            } else if (isAreaCompleted) {
                 areaActionButtonHtml = `
                    <button data-area-id="${areaId}" class="generate-area-report-btn px-3 py-1 rounded-lg text-xs font-bold text-white bg-green-500 hover:bg-green-600">Generar Acta Cierre</button>
                `;
            }


            item.innerHTML = `
                <div class="flex-grow">
                    <p class="font-semibold text-sm text-gray-500 dark:text-slate-400">Área: <span class="text-gray-900 dark:text-slate-100">${state.areaNames[list.areaOriginal] || list.areaOriginal}</span></p>
                    <p class="font-semibold text-sm text-gray-500 dark:text-slate-400">Impresión: <span class="text-indigo-600 dark:text-indigo-400 font-bold">${list.printDate}</span></p>
                    <p class="font-semibold text-sm text-gray-500 dark:text-slate-400">Tipo de Libro: <span class="text-gray-900 dark:text-slate-100">${list.listadoOriginal}</span></p>
                    <p class="font-semibold text-sm text-gray-500 dark:text-slate-400">Archivo: <span class="text-gray-700 dark:text-slate-300 italic">${list.fileName}</span></p>
                </div>
                <div class="flex flex-col space-y-2 items-end">
                    ${areaActionButtonHtml} 
                    <button data-list-id="${list.listId}" class="delete-list-btn px-3 py-1 rounded-lg text-xs font-bold text-white bg-red-500 hover:bg-red-600">Eliminar Listado</button>
                </div>
            `;
            container.appendChild(item);
        });
    }
    function renderDirectory() {
        const container = elements.settings.directoryContainer;
        const countEl = elements.settings.directoryCount;
        const areas = Object.keys(state.areaDirectory);

        countEl.textContent = `Total: ${areas.length}`;
        
        if (areas.length === 0) {
            container.innerHTML = '<p class="text-gray-500">No se han cargado áreas con información de responsable.</p>';
            return;
        }
        
        container.innerHTML = areas.sort().map((areaKey, index) => {
            const areaInfo = state.areaDirectory[areaKey];
            return `
                <div class="p-3 rounded-lg bg-white dark:bg-slate-800 text-gray-800 border-l-4 border-indigo-400 shadow-sm">
                    <div class="flex-grow">
                        <p class="font-bold text-sm text-gray-900 dark:text-slate-100">
                            ${index + 1}. ${areaInfo.fullName || `ÁREA ${areaKey}`}
                        </p>
                        
                        <p class="text-sm mt-1 text-gray-700 dark:text-slate-300">
                            <strong>Responsable:</strong> 
                            <span class="editable-field" 
                                  data-area-key="${areaKey}" 
                                  data-field="name" 
                                  tabindex="0"
                                  contenteditable="false">
                                  ${areaInfo.name || '(clic para editar)'}
                                  <i class="fa-solid fa-pencil editable-icon"></i>
                            </span>
                        </p>
                         <p class="text-sm text-gray-700 dark:text-slate-300">
                            <strong>Cargo:</strong> 
                            <span class="editable-field" 
                                  data-area-key="${areaKey}" 
                                  data-field="title" 
                                  tabindex="0"
                                  contenteditable="false">
                                  ${areaInfo.title || '(clic para editar)'}
                                  <i class="fa-solid fa-pencil editable-icon"></i>
                            </span>
                        </p>
                    </div>
                </div>
            `;
        }).join('');
    }
    function deleteListAndRefresh(listId) {
        const listToDelete = state.inventory.find(i => i.listId === listId);
        if (!listToDelete) return;

        logActivity('Listado eliminado', `Archivo: ${listToDelete.fileName}, Área: ${listToDelete.areaOriginal}`);
        state.inventory = state.inventory.filter(item => item.listId !== listId);
        showToast(`Listado "${listToDelete.fileName}" eliminado.`);
        updateSerialNumberCache();
        saveState();
        renderDashboard();
        populateAreaSelects();
        populateReportFilters();
        currentPage = 1;
        filterAndRenderInventory();
        renderLoadedLists();
    }

    function showReassignModal(listId, areaOriginal, affectedUsers, affectedAdicionales) {
        const { modal, text, areaSelect, confirmBtn, keepBtn, deleteAllBtn, cancelBtn } = elements.reassignModal;
        
        text.textContent = `El listado del área ${areaOriginal} tiene ${affectedUsers.length} usuario(s) y ${affectedAdicionales.length} bien(es) asociado(s). Elige una acción.`;
        
        areaSelect.innerHTML = state.areas
            .filter(area => area !== areaOriginal)
            .map(area => `<option value="${area}">Área ${state.areaNames[area] || area}</option>`) 
            .join('');
        
        if (areaSelect.options.length === 0) {
            areaSelect.disabled = true;
            confirmBtn.disabled = true;
            areaSelect.innerHTML = '<option>No hay otras áreas disponibles</option>';
        } else {
            areaSelect.disabled = false;
            confirmBtn.disabled = false;
        }
        
        modal.classList.add('show');
        
        const confirmHandler = () => {
            const newArea = areaSelect.value;
            if (!newArea) return showToast('Por favor, selecciona un área para reasignar.', 'error');
            
            affectedUsers.forEach(user => user.area = newArea);
            logActivity('Usuarios reasignados', `${affectedUsers.length} usuarios del área ${areaOriginal} movidos al área ${newArea}.`);
            showToast(`${affectedUsers.length} usuario(s) reasignado(s) al área ${newArea}.`);
            deleteListAndRefresh(listId);
            closeModal();
        };
        
        const keepHandler = () => {
            if (!state.persistentAreas) state.persistentAreas = [];
            if (!state.persistentAreas.includes(areaOriginal)) {
                state.persistentAreas.push(areaOriginal);
            }
            logActivity('Área mantenida', `El área ${areaOriginal} se mantuvo a pesar de eliminar el listado.`);
            showToast(`Los usuarios y bienes se mantendrán en el área ${areaOriginal}.`);
            deleteListAndRefresh(listId);
            closeModal();
        };
        
        const deleteAllHandler = () => {
            showConfirmationModal(
                '¡ADVERTENCIA!', 
                `Esto eliminará permanentemente ${affectedUsers.length} usuario(s) y ${affectedAdicionales.length} bien(es) asociado(s). Esta acción no se puede deshacer. ¿Continuar?`, 
                () => {
                    const affectedUserIds = affectedUsers.map(u => u.id);
                    const affectedUserNames = affectedUsers.map(u => u.name);
                    state.resguardantes = state.resguardantes.filter(u => !affectedUserIds.includes(u.id));
                    state.additionalItems = state.additionalItems.filter(item => !affectedUserNames.includes(item.usuario));
                    logActivity('Eliminación masiva', `Se eliminaron ${affectedUsers.length} usuarios y ${affectedAdicionales.length} bienes del área ${areaOriginal}.`);
                    showToast(`Se eliminaron usuarios y bienes del área ${areaOriginal}.`, 'warning');
                    deleteListAndRefresh(listId);
                    closeModal();
                }
            );
        };

        const closeModal = () => {
            modal.classList.remove('show');
            confirmBtn.removeEventListener('click', confirmHandler);
            keepBtn.removeEventListener('click', keepHandler);
            deleteAllBtn.removeEventListener('click', deleteAllHandler);
            cancelBtn.removeEventListener('click', closeModal);
        };

        confirmBtn.addEventListener('click', confirmHandler, { once: true });
        keepBtn.addEventListener('click', keepHandler, { once: true });
        deleteAllBtn.addEventListener('click', deleteAllHandler, { once: true });
        cancelBtn.addEventListener('click', closeModal, { once: true });
    }
    function showMainApp() {
        elements.loginPage.classList.add('hidden');
        elements.mainApp.classList.remove('hidden');
        elements.currentUserDisplay.textContent = state.currentUser.name;
        elements.settings.summaryAuthor.value = state.currentUser.name;

        updateTheme(state.theme);
        renderDashboard();
        populateAreaSelects();
        populateReportFilters();
        populateBookTypeFilter();
        currentPage = 1;
        filterAndRenderInventory();
        startAutosave();
        renderLoadedLists();
        renderDirectory();
        checkReadOnlyMode();
        changeTab('users');
    }
    function handleEmployeeLogin() {
        const employeeNumber = elements.employeeNumberInput.value;
        const employeeName = verifiers[employeeNumber];
        
        if (employeeName) {
            const newCurrentUser = { number: employeeNumber, name: employeeName };

            if (state.loggedIn && state.currentUser.number !== newCurrentUser.number) {
                 showConfirmationModal(
                    'Cambio de Usuario',
                    `Actualmente hay un inventario en progreso. ¿Deseas continuar con el inventario actual como ${employeeName} o iniciar uno nuevo?`,
                    () => {
                        logActivity('Cambio de usuario', `Sesión continuada por ${employeeName}.`);
                        state.currentUser = newCurrentUser;
                        showToast(`Bienvenido de nuevo, ${employeeName}. Continuando con la sesión actual.`);
                        saveState();
                        showMainApp();
                    },
                    { 
                        confirmText: 'Continuar', 
                        cancelText: 'Iniciar Nuevo',
                        onCancel: () => {
                             state.currentUser = newCurrentUser;
                             resetInventoryState();
                        }
                    }
                );
            } else {
                state.loggedIn = true;
                state.currentUser = newCurrentUser;
                if (!state.sessionStartTime) {
                    state.sessionStartTime = new Date().toISOString();
                    logActivity('Inicio de sesión', `Usuario ${employeeName} ha iniciado sesión.`);
                } else {
                    logActivity('Reanudación de sesión', `Usuario ${employeeName} ha reanudado la sesión.`);
                }
                showToast(`Bienvenido, ${employeeName}`);
                saveState();
                showMainApp();
            }

        } else {
            showToast('Este sistema solo puede ser utilizado por personal de Almacén e inventarios.', 'error');
        }
        elements.employeeNumberInput.value = '';
    }
    
    async function startCamera() {
        if (state.readOnlyMode) return;
        const { cameraStream, uploadContainer, cameraViewContainer } = elements.photo;
        
        stopCamera(); 

        if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
            try {
                const constraints = {
                    video: { 
                        facingMode: "environment" 
                    }
                };
                state.cameraStream = await navigator.mediaDevices.getUserMedia(constraints);
                cameraStream.srcObject = state.cameraStream;
                uploadContainer.classList.add('hidden');
                cameraViewContainer.classList.remove('hidden');
            } catch (err) {
                showToast('No se pudo acceder a la cámara trasera. Revisa los permisos.', 'error');
                console.error("Error al acceder a la cámara: ", err);
            }
        } else {
            showToast('Tu navegador no soporta el acceso a la cámara.', 'error');
        }
    }
    
    function stopCamera() {
        if (state.cameraStream) {
            state.cameraStream.getTracks().forEach(track => track.stop());
            state.cameraStream = null;
        }
    }
    
    function generateInstitutionalAdicionalesReport(options = {}) {
        const selectedUser = elements.reports.userFilter.value;
        const selectedArea = elements.reports.areaFilter.value;

        let institutionalItems = state.additionalItems.filter(item => item.personal === 'No');

        if (selectedArea !== 'all') {
            const usersInArea = state.resguardantes
                .filter(user => user.area === selectedArea)
                .map(user => user.name);
            institutionalItems = institutionalItems.filter(item => usersInArea.includes(item.usuario));
        }
        if (selectedUser !== 'all') {
            institutionalItems = institutionalItems.filter(item => item.usuario === selectedUser);
        }
        
        const reportOptions = {
            isInstitutionalReport: true, 
            headers: ['✓', 'Descripción', 'Clave Original', 'Área Procedencia', 'Marca', 'Serie', 'Usuario', 'Clave Asignada', 'Acción']
        };
        
        renderReportTable(institutionalItems, 'Regularización de Bienes Adicionales Institucionales', reportOptions);
        showToast(`Generado con ${institutionalItems.length} bienes institucionales.`);
        logActivity('Reporte Adicionales (Institucionales)', `Generado con ${institutionalItems.length} bienes.`);
    }
    function renderReportTable(data, title, options = {}) {
        const { withCheckboxes = false, headers = [], isInstitutionalReport = false, reportType = null } = options; 
        
        const { modal, title: modalTitle, tableHead, tableBody } = elements.reports.reportViewModal;

        modalTitle.textContent = title;
        
        tableHead.innerHTML = `<tr>${headers.map(h => `<th class="px-4 py-3 text-left text-xs font-medium text-gray-500 dark:text-gray-400 uppercase tracking-wider">${h}</th>`).join('')}</tr>`;
        tableBody.innerHTML = '';

        if (data.length === 0) {
            tableBody.innerHTML = `<tr><td colspan="${headers.length}" class="text-center py-4 text-gray-500 dark:text-gray-300">No se encontraron bienes.</td></tr>`;
            modal.classList.add('show');
            return;
        }
        
        data.forEach(item => {
            const row = document.createElement('tr');
            let cells = '';
            const clave = item['CLAVE UNICA'];

            if (isInstitutionalReport) {
                const isChecked = state.institutionalReportCheckboxes[item.id] || false;
                cells = `
                    <td class="px-4 py-4"><input type="checkbox" class="rounded institutional-report-checkbox" data-id="${item.id}" ${isChecked ? 'checked' : ''}></td>
                    <td class="px-4 py-4 text-sm">${item.descripcion}</td>
                    <td class="px-4 py-4 text-sm">${item.clave || 'N/A'}</td>
                    <td class="px-4 py-4 text-sm">${item.area || 'N/A'}</td>
                    <td class="px-4 py-4 text-sm">${item.marca || 'N/A'}</td>
                    <td class="px-4 py-4 text-sm">${item.serie || 'N/A'}</td>
                    <td class="px-4 py-4 text-sm">${item.usuario}</td>
                    <td class="px-4 py-4">
                        <input type="text" value="${item.claveAsignada || ''}" placeholder="Asignar..." class="new-clave-input w-24 rounded-md border-gray-300 dark:border-slate-600 shadow-sm p-2 text-sm" data-id="${item.id}" autocomplete="off">
                    </td>
                    <td class="px-4 py-4">
                        <button class="save-new-clave-btn px-3 py-2 rounded-lg text-xs font-bold text-white transition-colors bg-indigo-500 hover:bg-indigo-600" data-id="${item.id}">
                            <i class="fa-solid fa-save mr-1"></i> Guardar
                        </button>
                    </td>
                `;
            } else { 
                if (reportType === 'labels') {
                    cells += `<td class="px-4 py-4">
                        <button class="report-label-done-btn px-3 py-1 rounded-lg text-xs font-bold text-white transition-colors bg-green-500 hover:bg-green-600" data-clave="${clave}">
                            HECHO
                        </button>
                    </td>`;
                } 
                else if (withCheckboxes && reportType) {
                    const isChecked = state.reportCheckboxes[reportType] ? (state.reportCheckboxes[reportType][clave] || false) : false;
                    cells += `<td class="px-4 py-4"><input type="checkbox" class="rounded report-item-checkbox" data-clave="${clave}" data-report-type="${reportType}" ${isChecked ? 'checked' : ''}></td>`;
                }
                
                if (headers.includes('Clave Única')) cells += `<td class="px-4 py-4 text-sm">${clave}</td>`;
                if (headers.includes('Descripción')) cells += `<td class="px-4 py-4 text-sm">${item['DESCRIPCION']}</td>`;
                if (headers.includes('Serie')) cells += `<td class="px-4 py-4 text-sm">${item['SERIE'] || 'N/A'}</td>`;
                if (headers.includes('Usuario')) cells += `<td class="px-4 py-4 text-sm">${item['NOMBRE DE USUARIO'] || 'N/A'}</td>`;
                if (headers.includes('Marca')) cells += `<td class="px-4 py-4 text-sm">${item['MARCA'] || 'N/A'}</td>`;
                if (headers.includes('Modelo')) cells += `<td class="px-4 py-4 text-sm">${item['MODELO'] || 'N/A'}</td>`;
                if (headers.includes('Ubicado')) cells += `<td classs="px-4 py-4 text-sm">${item['UBICADO'] || 'NO'}</td>`;
                if (headers.includes('Área Original')) cells += `<td class="px-4 py-4 text-sm">${item.areaOriginal}</td>`;
                if (headers.includes('Nota')) cells += `<td class" class="px-4 py-4 text-sm">${state.notes[clave] || 'N/A'}</td>`;
                if (headers.includes('Usuario/Área Actual')) {
                    const currentUser = state.resguardantes.find(u => u.name === item['NOMBRE DE USUARIO']);
                    cells += `<td class="px-4 py-4 text-sm">${item['NOMBRE DE USUARIO']} (Área: ${currentUser?.area || 'N/A'})</td>`;
                }
            }
            
            row.innerHTML = cells;
            tableBody.appendChild(row);
        });

        modal.classList.add('show');
    }
    function deleteAdditionalItem(itemId, transferredPhoto = false) {
        const item = state.additionalItems.find(i => i.id === itemId);
        if (!item) return;

        state.additionalItems = state.additionalItems.filter(i => i.id !== itemId);
        
        if (!transferredPhoto) {
            photoDB.deleteItem('photos', `additional-${itemId}`);
        }
        delete state.additionalPhotos[itemId];
        
        renderAdicionalesList(); 
        renderDashboard(); 
        saveState(); 
        updateSerialNumberCache();
        logActivity('Bien adicional eliminado', `Descripción: ${item.descripcion}`);
        showToast('Bien adicional eliminado.');
    }

    async function showTransferPhotoModal(item) {
        const { modal, preview, search, select, confirmBtn, skipBtn } = elements.transferPhotoModal;
        
        try {
            const blob = await photoDB.getItem('photos', `additional-${item.id}`);
            if (blob) {
                const url = URL.createObjectURL(blob);
                preview.src = url;
                preview.onload = () => URL.revokeObjectURL(url);
            } else {
                preview.src = ''; 
            }
        } catch (e) {
            preview.src = '';
        }

        const getPendingItems = (searchTerm = '') => {
            const term = searchTerm.toLowerCase().trim();
            return state.inventory.filter(invItem => 
                invItem.UBICADO === 'NO' && 
                (!term || 
                 (invItem['CLAVE UNICA'] && String(invItem['CLAVE UNICA']).toLowerCase().includes(term)) ||
                 (invItem['DESCRIPCION'] && invItem['DESCRIPCION'].toLowerCase().includes(term)) ||
                 (invItem['SERIE'] && String(invItem['SERIE']).toLowerCase().includes(term))
                )
            );
        };

        const populateSelect = (searchTerm = '') => {
            const items = getPendingItems(searchTerm);
            select.innerHTML = items.length === 0
                ? '<option value="">-- No se encontraron pendientes --</option>'
                : items.map(i => `<option value="${i['CLAVE UNICA']}">(${i['CLAVE UNICA']}) ${i['DESCRIPCION'].substring(0, 40)}...</option>`).join('');
            confirmBtn.disabled = select.value === '';
        };

        search.value = '';
        populateSelect(); 
        
        const searchHandler = debounce(() => populateSelect(search.value), 300);
        search.addEventListener('input', searchHandler);

        select.onchange = () => {
            confirmBtn.disabled = select.value === '';
        };

        const closeModal = () => {
            modal.classList.remove('show');
            search.removeEventListener('input', searchHandler);
            confirmBtn.onclick = null;
            skipBtn.onclick = null;
        };

        confirmBtn.onclick = async () => {
            const targetClave = select.value;
            if (!targetClave) return;

            try {
                const blob = await photoDB.getItem('photos', `additional-${item.id}`);
                if (blob) {
                    await photoDB.setItem('photos', `inventory-${targetClave}`, blob);
                    state.photos[targetClave] = true;
                    logActivity('Foto transferida', `De Adicional (ID: ${item.id}) a Inventario (Clave: ${targetClave})`);
                    showToast('Foto transferida con éxito.', 'success');
                    filterAndRenderInventory(); 
                    deleteAdditionalItem(item.id, true); 
                }
            } catch (e) {
                console.error("Error al transferir foto:", e);
                showToast('Error al transferir la foto.', 'error');
            }
            closeModal();
        };

        skipBtn.onclick = () => {
            showConfirmationModal(
                'Eliminar sin transferir',
                '¿Seguro que quieres eliminar este bien adicional Y su foto permanentemente?',
                () => {
                    deleteAdditionalItem(item.id, false); 
                    closeModal();
                }
            );
        };

        modal.classList.add('show');
    }

    function getLocalDate() {
        const date = new Date();
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0'); 
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
    }

    function showPreprintModal(reportType, data = {}) {
        const { modal, title, fieldsContainer, confirmBtn, dateInput } = elements.preprintModal;
        let fieldsHtml = '';
        let defaultValues = {};
        let titleText = '';

        const selectedArea = data.filterArea || elements.reports.areaFilter.value;
        const selectedUser = data.filterUser || elements.reports.userFilter.value;
        
        const areaId = data.areaId || (selectedArea !== 'all' ? selectedArea : (state.resguardantes.find(u => u.name === selectedUser)?.area || null));
        const areaResponsibleData = areaId ? state.areaDirectory[areaId] : null;

        dateInput.value = getLocalDate();

        switch (reportType) {
            case 'session_summary':
                titleText = 'Generar Resumen de Sesión';
                defaultValues = {
                    author: elements.settings.summaryAuthor.value.trim(),
                    areaResponsible: elements.settings.summaryAreaResponsible.value.trim(),
                    location: elements.settings.summaryLocation.value.trim()
                };
                fieldsHtml = `
                    <div><label class="block text-sm font-medium">Ubicación Física del Inventario:</label><input type="text" id="preprint-location" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.location}"></div>
                    <div><label class="block text-sm font-medium">Realizado por (Entrega):</label><input type="text" id="preprint-author" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.author}"></div>
                    <div><label class="block text-sm font-medium">Responsable del Área (Recibe):</label><input type="text" id="preprint-areaResponsible" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.areaResponsible}"></div>
                `;
                break;
            
            // --- CASO AÑADIDO PARA PLAN DE ACCIÓN ---
            case 'tasks_report':
                titleText = 'Generar Plan de Acción';
                fieldsHtml = `
                    <div class="p-3 bg-blue-50 text-blue-800 rounded-lg text-sm">
                        <i class="fa-solid fa-info-circle mr-2"></i>
                        Se generará un documento PDF con las tareas pendientes (notas, etiquetas, bienes fuera de lugar y regularizaciones).
                    </div>
                `;
                break;
            // ----------------------------------------

            case 'area_closure':
                titleText = 'Generar Acta de Cierre de Área';
                defaultValues = {
                    areaId: data.areaId,
                    responsible: data.responsible || (areaResponsibleData?.name || ''), 
                    location: data.location || '',
                    areaFullName: state.areaNames[data.areaId] || `Área ${data.areaId}`,
                    entrega: state.currentUser.name,
                    recibe: data.responsible || (areaResponsibleData?.name || ''), 
                    recibeCargo: areaResponsibleData?.title || 'Responsable de Área'
                };
                fieldsHtml = `
                    <div><label class="block text-sm font-medium">Nombre Completo del Área:</label><input type="text" id="preprint-areaFullName" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.areaFullName}"></div>
                    <div><label class="block text-sm font-medium">Ubicación de Firma:</label><input type="text" id="preprint-location" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.location}" placeholder="Ej. Oficina 1..."></div>
                    <div><label class="block text-sm font-medium">Entrega (Inventario):</label><input type="text" id="preprint-entrega" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.entrega}"></div>
                    <div><label class="block text-sm font-medium">Recibe de Conformidad:</label><input type="text" id="preprint-recibe" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.recibe}" placeholder="Nombre completo de quien recibe"></div>
                    <div><label class="block text-sm font-medium">Cargo de Quien Recibe:</label><input type="text" id="preprint-recibeCargo" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.recibeCargo}"></div>
                `;
                break;
             case 'simple_pending':
                titleText = 'Imprimir Reporte de Pendientes';
                defaultValues = {
                    areaDisplay: selectedArea !== 'all' ? `${state.areaNames[selectedArea] || selectedArea}` : 'Todas las Áreas', 
                    entrega: state.currentUser.name,
                    recibe: "_________________________"
                };
                fieldsHtml = `
                    <div><label class="block text-sm font-medium">Reporte para:</label><input type="text" id="preprint-areaDisplay" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.areaDisplay}"></div>
                    <div><label class="block text-sm font-medium">Realizó (Entrega):</label><input type="text" id="preprint-entrega" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.entrega}"></div>
                    <div><label class="block text-sm font-medium">Recibe Copia:</label><input type="text" id="preprint-recibe" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.recibe}"></div>
                `;
                break;
            case 'individual_resguardo':
            case 'adicionales_informe':
                titleText = 'Imprimir Resguardo';
                const isAdicional = reportType === 'adicionales_informe';
                const hasMixedItems = data.isAdicional || false;
                
                const isForArea = data.isForArea || false;
                const isForUser = data.isForUser || false;

                let userForReport = 'Usuario';
                if (isAdicional) {
                    if (isForUser) {
                        userForReport = selectedUser;
                    } else if (isForArea) {
                        userForReport = `Responsables del Área ${areaId}`;
                    } else {
                        userForReport = 'Todas las Áreas';
                    }
                } else { 
                    if (selectedUser !== 'all') {
                        userForReport = selectedUser;
                    } else {
                         userForReport = '_________________________'; 
                    }
                }
                
                const entregaValue = areaResponsibleData?.name || '_________________________';
                const receiverName = (isForArea && isAdicional) ? entregaValue : userForReport;

                defaultValues = {
                    areaFullName: areaId ? (state.areaNames[areaId] || `Área ${areaId}`) : 'Todas las Áreas', 
                    entrega: entregaValue,
                    recibe: receiverName, 
                    recibeCargo: areaResponsibleData?.title || 'Responsable de Área',
                    isForArea: isForArea,
                    isForUser: isForUser
                };
                
                fieldsHtml = `
                    <div><label class="block text-sm font-medium">Nombre Completo del Área:</label><input type="text" id="preprint-areaFullName" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.areaFullName}"></div>
                    <div><label class="block text-sm font-medium">Responsable del Área (Entrega):</label><input type="text" id="preprint-entrega" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.entrega}"></div>
                    <div><label class="block text-sm font-medium">Firma de Conformidad (Recibe):</label><input type="text" id="preprint-recibe" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.recibe}"></div>
                    <div><label class="block text-sm font-medium">Cargo de Quien Entrega:</label><input type="text" id="preprint-recibeCargo" class="mt-1 block w-full p-2 border rounded-md" value="${defaultValues.recibeCargo}"></div>
                `;
                break;
            default:
                dateInput.value = '';
                console.warn("showPreprintModal llamado con tipo inesperado:", reportType);
                break;
        }

        title.textContent = titleText;
        fieldsContainer.innerHTML = fieldsHtml; 
        modal.classList.add('show');
        
        // CORRECCIÓN: Gestión robusta de eventos para evitar zombies
        const cleanupNav = handleModalNavigation(modal);

        // Función centralizada de limpieza
        const safeClose = () => {
            modal.classList.remove('show');
            cleanupNav(); // Matamos el listener del teclado
            // Limpiamos referencias para evitar fugas de memoria
            elements.preprintModal.cancelBtn.onclick = null;
            confirmBtn.onclick = null;
        };

        // Asignar cierre seguro al botón Cancelar
        elements.preprintModal.cancelBtn.onclick = safeClose;
        
        // IMPORTANTE: Asignar cierre seguro TAMBIÉN a la X (closeBtn) si existe en este modal
        // (Esto evita que cerrar con X deje el teclado "escuchando")
        const xBtn = modal.querySelector('.fa-xmark')?.closest('button');
        if(xBtn) {
            // Usamos onlick directo para sobrescribir comportamientos previos y asegurar limpieza
            xBtn.onclick = safeClose;
        }

        confirmBtn.onclick = () => {
            // Limpiamos listeners antes de procesar
            cleanupNav(); 
            
            const updatedOptions = { ...defaultValues };
            updatedOptions.date = dateInput.value.trim() || getLocalDate(); 

            const inputs = fieldsContainer.querySelectorAll('input');
            inputs.forEach(input => {
                const key = input.id.replace('preprint-', '');
                updatedOptions[key] = input.value;
            });

            if (reportType === 'area_closure') {
                if (!updatedOptions.responsible || !updatedOptions.location || !updatedOptions.recibe) {
                     return showToast('Para el Acta de Cierre, todos los campos son obligatorios.', 'error');
                }
            }
            
            switch (reportType) {
                case 'session_summary':
                    generateSessionSummary(updatedOptions);
                    break;
                // --- CASO AÑADIDO PARA EJECUTAR EL PLAN DE ACCIÓN ---
                case 'tasks_report':
                    generateTasksReport(updatedOptions);
                    break;
                // ----------------------------------------------------
                case 'area_closure':
                    generateAreaClosureReport(updatedOptions);
                    break;
                 case 'simple_pending':
                    generateSimplePendingReport(updatedOptions);
                    break;
                case 'individual_resguardo':
                     if (!selectedUser || selectedUser === 'all') return showToast('Selecciona un usuario en la pestaña Reportes para generar un resguardo individual.', 'error');
                     
                     const items = data.items || state.inventory.filter(item => item['NOMBRE DE USUARIO'] === selectedUser);
                     const title = data.title || 'Resguardo Individual de Bienes';
                     const isAdicional = data.isAdicional || false; 
                     
                     generatePrintableResguardo(title, updatedOptions.recibe, items, isAdicional, updatedOptions);
                     break;
                case 'adicionales_informe':
                    let itemsToPrint = state.additionalItems;
                    
                    const reportSelectedArea = selectedArea; 
                    const reportSelectedUser = selectedUser;

                    if (reportSelectedArea !== 'all') {
                        const usersInArea = state.resguardantes.filter(u => u.area === reportSelectedArea).map(u => u.name);
                        itemsToPrint = itemsToPrint.filter(item => usersInArea.includes(item.usuario));
                    }
                    if (reportSelectedUser !== 'all') {
                        itemsToPrint = itemsToPrint.filter(item => item.usuario === reportSelectedUser);
                    } 
                    
                    generatePrintableResguardo('Mobiliario y Equipo Ubicado de Manera Adicional Global', updatedOptions.recibe, itemsToPrint, true, updatedOptions);
                    break;
            }
            modal.classList.remove('show');
        };
    }

    function getAreaColor(areaId) {
        if (!state.layoutItemColors[areaId]) {
            let hash = 0;
            for (let i = 0; i < String(areaId).length; i++) {
                hash = String(areaId).charCodeAt(i) + ((hash << 5) - hash);
            }
            const h = hash % 360; 
            const s = 70 + (hash % 20); 
            const l = 55 + (hash % 10); 
            state.layoutItemColors[areaId] = `hsl(${h}, ${s}%, ${l}%)`;
            saveState(); 
        }
        return state.layoutItemColors[areaId];
    }

    function getLocationIcon(locationBase) {
        const base = String(locationBase).toUpperCase();
        if (base.includes('OFICINA')) return 'fa-solid fa-building';
        if (base.includes('CUBICULO') || base.includes('CUBÍCULO')) return 'fa-solid fa-user';
        if (base.includes('BODEGA')) return 'fa-solid fa-box-archive';
        if (base.includes('PASILLO')) return 'fa-solid fa-road';
        if (base.includes('SALA DE JUNTAS')) return 'fa-solid fa-users';
        if (base.includes('SECRETARIAL')) return 'fa-solid fa-keyboard';
        if (base.includes('FOTOCOPIADO')) return 'fa-solid fa-print';
        return 'fa-solid fa-location-dot'; 
    }

    function populateLayoutSidebar() {
        const container = elements.layoutEditor.sidebar;
        container.innerHTML = '';
        
        // Mapa para agrupar usuarios por ubicación física
        const locationsMap = new Map();

        state.resguardantes.forEach(user => {
            // 1. Obtenemos TODAS las ubicaciones del usuario, no solo la principal
            const userLocations = (user.locations && user.locations.length > 0) 
                                  ? user.locations 
                                  : [user.locationWithId || 'Sin Ubicación'];

            // 2. Recorremos cada una de sus ubicaciones
            userLocations.forEach(locId => {
                if (!locId) return;

                if (!locationsMap.has(locId)) {
                    // Intentamos deducir el tipo base (ej: "OFICINA" de "OFICINA 01") para el ícono
                    const baseMatch = locId.match(/^(.*)\s\d+$/);
                    const locationBase = baseMatch ? baseMatch[1] : locId;

                    locationsMap.set(locId, {
                        locationBase: locationBase,
                        areaId: user.area, 
                        users: []
                    });
                }
                // Agregamos al usuario a la lista de esta ubicación
                locationsMap.get(locId).users.push(user.name);
            });
        });

        // Obtenemos los elementos que YA están en el lienzo para no duplicarlos visualmente
        const itemsOnCurrentPage = state.mapLayout[state.currentLayoutPage] || {};

        // 3. Renderizamos la lista completa
        // Convertimos el mapa a array y ordenamos alfabéticamente para que sea fácil buscar
        const sortedLocations = Array.from(locationsMap.entries()).sort((a, b) => a[0].localeCompare(b[0]));

        sortedLocations.forEach(([locId, data]) => {
            const el = document.createElement('div');
            el.className = 'layout-shape draggable-item';
            el.dataset.locationId = locId;
            el.dataset.areaId = data.areaId; 
            
            // Si ya está en el lienzo, lo ocultamos de la barra lateral
            if (itemsOnCurrentPage[locId]) {
                el.classList.add('hidden'); 
            }
            
            // Limitamos la lista de usuarios visibles en la tarjeta para no hacerla gigante
            const maxUsersToShow = 3;
            let usersHtml = data.users.slice(0, maxUsersToShow).map(name => `<li>${name}</li>`).join('');
            if (data.users.length > maxUsersToShow) {
                usersHtml += `<li><em class="text-xs text-gray-500">+ ${data.users.length - maxUsersToShow} más...</em></li>`;
            }

            const iconClass = getLocationIcon(data.locationBase);
            const areaColor = getAreaColor(data.areaId);
            const colorDotHtml = `<div class="area-color-dot" style="background-color: ${areaColor};"></div>`;

            el.innerHTML = `
                ${colorDotHtml}
                <h5><i class="${iconClass} location-icon"></i>${locId}</h5>
                <ul>${usersHtml}</ul>
            `;
            container.appendChild(el);
        });

        if (sortedLocations.length === 0) {
            container.innerHTML = '<p class="text-xs text-gray-400 text-center p-4">No hay ubicaciones registradas en los usuarios.</p>';
        }
    }
    
    function updateLayoutPagination() {
        const { pageName, pagePrev, pageNext, pageRemove } = elements.layoutEditor;
        pageName.value = state.layoutPageNames[state.currentLayoutPage] || state.currentLayoutPage;
        
        const pageKeys = Object.keys(state.layoutPageNames);
        const currentIndex = pageKeys.indexOf(state.currentLayoutPage);
        
        pagePrev.disabled = currentIndex <= 0;
        pageNext.disabled = currentIndex >= pageKeys.length - 1;
        pageRemove.disabled = pageKeys.length <= 1; 
    }

    function switchLayoutPage(newPageKey) {
        state.currentLayoutPage = newPageKey;
        loadSavedLayout(); 
        populateLayoutSidebar(); 
        updateLayoutPagination(); 
    }

    // Variable para rastrear memoria de imágenes del croquis
    let activeLayoutUrls = [];

    async function loadSavedLayout() {
        const canvas = elements.layoutEditor.canvas;
        canvas.innerHTML = ''; 
        
        // 1. LIMPIEZA DE MEMORIA: Revocar URLs anteriores para liberar RAM
        if (activeLayoutUrls.length > 0) {
            activeLayoutUrls.forEach(url => URL.revokeObjectURL(url));
            activeLayoutUrls = [];
        }
        
        const layoutData = state.mapLayout[state.currentLayoutPage] || {};
        
        if (layoutData) {
            for (const id in layoutData) {
                if (layoutData.hasOwnProperty(id)) {
                    const item = layoutData[id];
                    let dataUrl = null;
                    
                    if (item.type === 'image' && item.imageId) {
                        try {
                            const blob = await photoDB.getItem('layoutImages', item.imageId);
                            if (blob) {
                                dataUrl = URL.createObjectURL(blob);
                                // Agregamos la nueva URL a la lista de rastreo
                                activeLayoutUrls.push(dataUrl);
                            }
                        } catch(e) {
                            console.error('Error al cargar imagen de croquis desde DB', e);
                        }
                    }
                    createShapeOnCanvas(id, item.x, item.y, item.width, item.height, item.type, (item.text || ''), dataUrl, (item.rotation || 0), item.areaId);
                }
            }
        }
    }
    
    function createShapeOnCanvas(id, x, y, width, height, type = 'location', text = '', imageDataUrl = null, rotation = 0, areaId = null) {
        const canvas = elements.layoutEditor.canvas;
        
        if (canvas.querySelector(`[data-id="${id}"]`)) {
            return;
        }
        
        const el = document.createElement('div');
        el.className = 'layout-shape layout-on-canvas';
        el.dataset.id = id;
        
        let innerHtml = '';
        let colorDotHtml = ''; 
        let currentAreaId = areaId; 

        if (type === 'location') {
            const user = state.resguardantes.find(u => u.locationWithId === id);
            if (!user) return; 

            currentAreaId = user.area; 
            el.dataset.areaId = currentAreaId; 

            const usersInLoc = state.resguardantes
                .filter(u => u.locationWithId === id)
                .map(u => `<li>${u.name} (Área ${u.area})</li>`)
                .join('');
            const iconClass = getLocationIcon(user.location);
            
            const areaColor = getAreaColor(currentAreaId);
            colorDotHtml = `<div class="area-color-dot" style="background-color: ${areaColor};"></div>`;

            innerHtml = `
                <h5><i class="${iconClass} location-icon"></i>${id}</h5>
                <ul>${usersInLoc}</ul>
            `;
        } 
        else if (type === 'tool') {
            const toolIconClass = {
                'arrow': 'fa-solid fa-arrow-up',
            }[id.split('-')[0]] || 'fa-solid fa-square';
            
            el.classList.add('tool-shape');
            innerHtml = `<i class="${toolIconClass} tool-icon"></i>`;
            width = width || 50;
            height = height || 50;
        }
        else if (type === 'note') {
            el.classList.add('tool-note');
            innerHtml = `<textarea class="layout-shape-note-textarea" placeholder="Escribe una nota...">${text}</textarea>`;
            width = width || 200;
            height = height || 100;
        }
        else if (type === 'text') {
            el.classList.add('tool-text');
            innerHtml = `<textarea class="layout-shape-text-textarea" placeholder="Texto...">${text}</textarea>`;
            width = width || 150;
            height = height || 40;
        }
        else if (type === 'image') {
            el.classList.add('tool-image');
            if (imageDataUrl) {
                el.style.backgroundImage = `url(${imageDataUrl})`;
            } else {
                innerHtml = `<span>Cargando...</span>`; 
            }
            width = width || 300;
            height = height || 200;
        }

        el.style.transform = `translate(${x}px, ${y}px) rotate(${rotation}deg)`;
        if (width) el.style.width = `${width}px`;
        if (height) el.style.height = `${height}px`;
        
        el.dataset.x = x;
        el.dataset.y = y;
        el.dataset.rotation = rotation; 
        el.dataset.type = type; 
        
        const controlsHtml = `
            <div class="layout-delete-btn" title="Eliminar"><i class="fa-solid fa-xmark"></i></div>
            <div class="layout-rotate-handle" title="Rotar"><i class="fa-solid fa-rotate-right"></i></div>
        `;
        
        el.innerHTML = colorDotHtml + innerHtml + controlsHtml; 
        
        canvas.appendChild(el);
        
        if (type === 'note' || type === 'text') { 
            el.querySelector('textarea').addEventListener('input', (e) => {
                saveLayoutPositions(); 
            });
        }
    }

    function saveLayoutPositions() {
        const currentPageLayout = {};
        document.querySelectorAll('#layout-canvas .layout-on-canvas').forEach(el => {
            const id = el.dataset.id;
            const x = parseFloat(el.dataset.x) || 0;
            const y = parseFloat(el.dataset.y) || 0;
            const width = parseFloat(el.style.width) || (el.classList.contains('tool-shape') ? 50 : (el.classList.contains('tool-text') ? 150 : (el.classList.contains('tool-note') ? 200 : (el.classList.contains('tool-image') ? 300 : 180))));
            const height = parseFloat(el.style.height) || (el.classList.contains('tool-shape') ? 50 : (el.classList.contains('tool-text') ? 40 : (el.classList.contains('tool-note') ? 100 : (el.classList.contains('tool-image') ? 200 : 60))));
            const type = el.dataset.type;
            const rotation = parseFloat(el.dataset.rotation) || 0;
            
            const itemData = { x, y, width, height, type, rotation };

            if (type === 'note' || type === 'text') { 
                itemData.text = el.querySelector('textarea').value;
            }
            
            if (type === 'image') {
                itemData.imageId = state.layoutImages[id];
            }
            
            if (type === 'location') {
                itemData.areaId = el.dataset.areaId;
            }
            
            currentPageLayout[id] = itemData;
        });
        state.mapLayout[state.currentLayoutPage] = currentPageLayout;
        logActivity('Croquis guardado', `Se guardó la página ${state.currentLayoutPage} con ${Object.keys(currentPageLayout).length} elementos.`);
    }

    async function printLayout() {
        logActivity('Impresión de Croquis', 'Generando impresión de croquis...');

        // Limpieza previa
        document.querySelectorAll('.print-page.layout-clone').forEach(el => el.remove());
        document.querySelectorAll('.print-page').forEach(page => page.classList.remove('active'));

        const masterTemplate = elements.printTemplates.layout;
        
        // Generar lista de usuarios para la primera página
        const userListContainer = masterTemplate.querySelector('#print-layout-user-list');
        const usersByArea = state.resguardantes.reduce((acc, user) => {
            const areaKey = user.area || 'Sin Área';
            if (!acc[areaKey]) acc[areaKey] = [];
            acc[areaKey].push(user);
            return acc;
        }, {});
        const sortedAreas = Object.keys(usersByArea).sort();
        let userHtml = '<h2>Listado de Usuarios por Área</h2>';
        for (const area of sortedAreas) {
            const areaName = state.areaNames[area] || `Área ${area}`;
            userHtml += `<h3>${areaName}</h3><ul>`;
            userHtml += usersByArea[area]
                // Mostrar todas las ubicaciones en la lista de referencia
                .map(user => {
                    const locs = (user.locations && user.locations.length > 0) 
                                 ? user.locations.join(', ') 
                                 : user.locationWithId;
                    return `<li><strong>${user.name}</strong> (${locs})</li>`;
                })
                .join('');
            userHtml += '</ul>';
        }
        userListContainer.innerHTML = userHtml;

        const allPageKeys = Object.keys(state.layoutPageNames);
        const blobToBase64 = (blob) => new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
        
        const printDate = getLocalDate(); 

        // Iterar por cada página del croquis
        for (let index = 0; index < allPageKeys.length; index++) {
            const pageKey = allPageKeys[index];
            const isFirstPage = index === 0;
            const pageTemplate = isFirstPage ? masterTemplate : masterTemplate.cloneNode(true);
            
            if (!isFirstPage) {
                pageTemplate.id = `print-layout-page-${index}`;
                pageTemplate.classList.add('layout-clone');
                pageTemplate.querySelector('#print-layout-user-list').innerHTML = '';
                elements.printContainer.appendChild(pageTemplate);
            }
            
            const pageName = state.layoutPageNames[pageKey] || pageKey;
            pageTemplate.querySelector('#print-layout-page-number').textContent = `Página ${index + 1}: ${pageName}`;
            pageTemplate.querySelector('#print-layout-date').textContent = `Fecha: ${printDate}`;

            const printCanvasContainer = pageTemplate.querySelector('.print-layout-canvas-container');
            const printCanvas = pageTemplate.querySelector('#print-layout-canvas');
            printCanvas.innerHTML = ''; 
            
            // Ajustar dimensiones para impresión carta
            printCanvasContainer.style.width = '720px'; 
            printCanvasContainer.style.height = '960px'; 
            printCanvas.style.width = '100%';
            printCanvas.style.height = '100%';
            
            const layoutData = state.mapLayout[pageKey] || {};
            let hasShapes = false;

            // Renderizar elementos en el canvas de impresión
            for (const id in layoutData) {
                if (layoutData.hasOwnProperty(id)) {
                    hasShapes = true;
                    const item = layoutData[id];
                    const el = document.createElement('div');
                    el.className = 'layout-shape'; 
                    
                    el.style.position = 'absolute';
                    el.style.left = `${item.x}px`;
                    el.style.top = `${item.y}px`;
                    el.style.width = `${item.width}px`;
                    el.style.height = `${item.height}px`;
                    el.style.transform = `rotate(${item.rotation || 0}deg)`;
                    el.style.transformOrigin = 'center center'; 
                    el.style.fontSize = '0.8em'; 

                    let innerHtml = '';
                    let colorDotHtml = ''; 

                    if (item.type === 'location') {
                        // --- CORRECCIÓN APLICADA AQUÍ ---
                        // Buscar usuarios que tengan esta ubicación (primaria o secundaria)
                        const usersInLocData = state.resguardantes.filter(u => {
                            if (u.locations && u.locations.length > 0) {
                                return u.locations.includes(id);
                            }
                            return u.locationWithId === id;
                        });

                        if (usersInLocData.length === 0) continue; 

                        const usersInLoc = usersInLocData
                            .map(u => `<li>${u.name} (Área ${u.area})</li>`)
                            .join('');
                            
                        const baseMatch = id.match(/^(.*)\s\d+$/);
                        const baseName = baseMatch ? baseMatch[1] : id;
                        const iconClass = getLocationIcon(baseName);
                        
                        const effectiveAreaId = item.areaId || usersInLocData[0].area;
                        const areaColor = getAreaColor(effectiveAreaId);
                        
                        colorDotHtml = `<div class="area-color-dot" style="background-color: ${areaColor}; border-color: #555;"></div>`;
                        innerHtml = `<h5><i class="${iconClass} location-icon"></i>${id}</h5><ul>${usersInLoc}</ul>`;
                        // --------------------------------
                    } 
                    else if (item.type === 'tool') {
                        const toolIconClass = 'fa-solid fa-arrow-up'; 
                        el.classList.add('tool-shape');
                        innerHtml = `<i class="${toolIconClass} tool-icon"></i>`;
                    } else if (item.type === 'note') {
                        el.classList.add('tool-note');
                        innerHtml = `<textarea class="layout-shape-note-textarea" readonly>${item.text || ''}</textarea>`;
                    } else if (item.type === 'text') { 
                        el.classList.add('tool-text');
                        innerHtml = `<textarea class="layout-shape-text-textarea" readonly>${item.text || ''}</textarea>`;
                    } else if (item.type === 'image' && item.imageId) {
                        el.classList.add('tool-image');
                        try {
                            const blob = await photoDB.getItem('layoutImages', item.imageId);
                            if(blob) {
                                const dataUrl = await blobToBase64(blob);
                                el.style.backgroundImage = `url(${dataUrl})`;
                            }
                        } catch(e) { console.error('Error al cargar imagen para impresión', e); }
                    }
                    
                    el.innerHTML = colorDotHtml + innerHtml; 
                    printCanvas.appendChild(el); 
                }
            }
            
            if (!hasShapes) {
                printCanvas.innerHTML = '<p style="text-align:center; padding-top: 40px; color: #888;">Página vacía</p>';
            }
            
            pageTemplate.classList.add('active'); 
        }
        
        preparePrint('print-layout-view', { date: printDate });
    }
    function resetCurrentLayoutPage() {
        if (state.readOnlyMode) return;
        showConfirmationModal('Restablecer Lienzo', `¿Seguro que quieres eliminar todos los elementos de la página "${state.layoutPageNames[state.currentLayoutPage]}"? Las ubicaciones volverán a la barra lateral.`, () => {
            state.mapLayout[state.currentLayoutPage] = {};
            elements.layoutEditor.canvas.innerHTML = '';
            populateLayoutSidebar();
            saveState();
            showToast('Lienzo restablecido.');
            logActivity('Croquis', `Lienzo de la página ${state.currentLayoutPage} restablecido.`);
        });
    }
    window.dragMoveListener = function (event) {
        var target = event.target;
        var x = (parseFloat(target.dataset.x) || 0) + event.dx;
        var y = (parseFloat(target.dataset.y) || 0) + event.dy;
        var rotation = (parseFloat(target.dataset.rotation) || 0);
        
        target.style.transform = 'translate(' + x + 'px, ' + y + 'px) rotate(' + rotation + 'deg)';
        target.dataset.x = x;
        target.dataset.y = y;
    }

    function initialize() {
        photoDB.init().catch(err => console.error('No se pudo iniciar la base de datos de fotos:', err));

        elements.employeeLoginBtn.addEventListener('click', handleEmployeeLogin);
        elements.employeeNumberInput.addEventListener('keydown', e => {
            if (e.key === 'Enter') {
                e.preventDefault();
                handleEmployeeLogin();
            }
        });
        
        elements.dashboard.toggleBtn.addEventListener('click', () => {
            elements.dashboard.headerAndDashboard.classList.toggle('hidden');
        });
        
        elements.logo.title.addEventListener('click', () => {
            logoClickCount++;
            if (logoClickCount >= 5) {
                const logText = state.activityLog.join('\n');
                const blob = new Blob([logText], { type: 'text/plain' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `log_inventario_${new Date().toISOString().slice(0,10)}.txt`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                showToast('Registro de actividad descargado.');
                logActivity('Log descargado', 'El usuario ha descargado el registro de actividad.');
                logoClickCount = 0;
            }
        });


        elements.clearSessionLink.addEventListener('click', (e) => {
            e.preventDefault();
            showConfirmationModal('Limpiar Sesión Completa', 'Esto borrará TODO el progreso, incluyendo usuarios e inventario guardado en este navegador. ¿Estás seguro?', () => {
                localStorage.removeItem('inventarioProState');
                deleteDB('InventarioProPhotosDB').finally(() => {
                   window.location.reload();
                });
            });
        });
        elements.logoutBtn.addEventListener('click', () => {
             logActivity('Cierre de sesión', `Usuario ${state.currentUser.name} ha salido.`);
             saveState();
             elements.mainApp.classList.add('hidden');
             elements.loginPage.classList.remove('hidden');
        });

        elements.uploadBtn.addEventListener('click', () => {
            elements.fileInput.value = ''; // FIX: Resetear para permitir cargar el mismo archivo de nuevo
            elements.fileInput.click();
        });
        elements.fileInput.addEventListener('change', (e) => {
            [...e.target.files].forEach(file => processFile(file));
            e.target.value = '';
        });
        elements.tabsContainer.addEventListener('click', e => {
            const tabBtn = e.target.closest('.tab-btn');
            if(tabBtn && tabBtn.dataset.tab) changeTab(tabBtn.dataset.tab);
        });
        
        const debouncedSearch = debounce(() => {
            currentPage = 1;
            filterAndRenderInventory();
        }, 300);
        elements.inventory.searchInput.addEventListener('input', debouncedSearch);

        elements.inventory.tableBody.addEventListener('click', (e) => {
            if (state.inventoryEditMode) {
                // Si estamos editando, NO abrir el modal al hacer click en la fila
                // Permitir solo clicks en iconos específicos si es necesario
                if (!e.target.closest('.action-btn') && !e.target.closest('.camera-icon') && !e.target.closest('.note-icon')) {
                    return; 
                }
            }
            const target = e.target;
            const row = target.closest('tr');
            const clave = row?.dataset.clave;
            if (!clave) return;

            if (target.closest('.note-icon, .camera-icon, .view-qr-btn, .view-details-btn')) {
                if (target.closest('.note-icon')) showNotesModal(clave);
                else if (target.closest('.camera-icon')) showPhotoModal('inventory', clave);
                else if (target.closest('.view-qr-btn')) showQrModal(clave);
                else if (target.closest('.view-details-btn')) showItemDetailsModal(clave);
            } 
            else if (!target.classList.contains('inventory-item-checkbox')) { 
                 showItemDetailView(clave);
            }
        });

        elements.inventory.statusFilter.addEventListener('change', () => { currentPage = 1; filterAndRenderInventory(); });
        elements.inventory.areaFilter.addEventListener('change', () => { currentPage = 1; filterAndRenderInventory(); });
        elements.inventory.bookTypeFilter.addEventListener('change', () => { currentPage = 1; filterAndRenderInventory(); });
        elements.inventory.selectAllCheckbox.addEventListener('change', e =>
            document.querySelectorAll('.inventory-item-checkbox').forEach(cb => cb.checked = e.target.checked));
        elements.inventory.ubicadoBtn.addEventListener('click', () => handleInventoryActions('ubicar'));
        elements.inventory.reEtiquetarBtn.addEventListener('click', () => handleInventoryActions('re-etiquetar'));
        elements.inventory.desubicarBtn.addEventListener('click', () => handleInventoryActions('desubicar'));
        elements.inventory.addNoteBtn.addEventListener('click', () => showNotesModal());
        elements.inventory.qrScanBtn.addEventListener('click', startQrScanner);
        elements.inventory.clearSearchBtn.addEventListener('click', () => {
            elements.inventory.searchInput.value = '';
            elements.inventory.statusFilter.value = 'all';
            elements.inventory.areaFilter.value = 'all';
            elements.inventory.bookTypeFilter.value = 'all';
            currentPage = 1;
            filterAndRenderInventory();
            elements.inventory.searchInput.focus();
        });
        elements.inventory.prevPageBtn.addEventListener('click', () => { if (currentPage > 1) { currentPage--; renderInventoryTable(); }});
        elements.inventory.nextPageBtn.addEventListener('click', () => {
            const totalPages = Math.ceil(filteredItems.length / itemsPerPage) || 1;
            if (currentPage < totalPages) { currentPage++; renderInventoryTable(); }
        });
        // --- LÓGICA DEL MODO EDICIÓN (AJUSTES) ---
        const editModeToggle = document.getElementById('inventory-edit-mode-toggle');
        if (editModeToggle) {
            editModeToggle.addEventListener('change', (e) => {
                state.inventoryEditMode = e.target.checked;
                
                if (state.inventoryEditMode) {
                    showToast('Modo Edición ACTIVADO. Puedes editar directamente en la tabla.', 'warning');
                } else {
                    showToast('Modo Edición DESACTIVADO.', 'info');
                }
                
                // Recargar la tabla para aplicar/quitar la editabilidad
                renderInventoryTable();
            });
        }

        // --- GUARDADO AUTOMÁTICO AL EDITAR CELDAS ---
        elements.inventory.tableBody.addEventListener('focusout', (e) => {
            // Detectar cuando se sale de una celda editable (evento blur/focusout)
            const cell = e.target;
            if (!state.inventoryEditMode || !cell.classList.contains('inventory-editable-cell')) return;

            const row = cell.closest('tr');
            const clave = row.dataset.clave;
            const field = cell.dataset.field; // DESCRIPCION, MARCA, etc.
            const newValue = cell.textContent.trim();

            const item = state.inventory.find(i => i['CLAVE UNICA'] === clave);
            
            if (item && item[field] !== newValue) {
                const oldValue = item[field];
                item[field] = newValue;
                
                // Actualizar también caché de series si se cambió la serie
                if (field === 'SERIE') updateSerialNumberCache();
                
                saveState();
                logActivity('Edición rápida', `Bien ${clave}: ${field} cambiado de "${oldValue}" a "${newValue}"`);
                showToast('Cambio guardado automáticamente.');
            }
        });

        // --- PREVENIR ENTER (SALTOS DE LÍNEA) EN CELDAS ---
        elements.inventory.tableBody.addEventListener('keydown', (e) => {
            if (state.inventoryEditMode && e.target.classList.contains('inventory-editable-cell')) {
                if (e.key === 'Enter') {
                    e.preventDefault(); // Evitar salto de línea
                    e.target.blur(); // Disparar el guardado (focusout)
                }
            }
        });

        // --- INICIO MEJORA MULTI-UBICACIÓN: Lógica UI de creación de usuario ---
        
        const renderTempLocations = () => {
            const container = elements.userForm.locationsList;
            if (tempUserLocations.length === 0) {
                container.innerHTML = '<p class="text-xs text-gray-400 italic w-full text-center self-center empty-msg">Añade al menos una ubicación con el botón +</p>';
                return;
            }
            
            container.innerHTML = tempUserLocations.map((loc, index) => `
                <span class="location-chip">
                    ${loc}
                    <i class="fa-solid fa-xmark remove-loc-btn" data-index="${index}"></i>
                </span>
            `).join('');
        };
        
        // Listener para botón agregar ubicación
        elements.userForm.addLocationBtn.addEventListener('click', () => {
            if (state.readOnlyMode) return;
            const locationType = elements.userForm.locationSelect.value;
            const locationManual = elements.userForm.locationManual.value.trim();
            
            if (locationType === 'OTRA' && !locationManual) {
                return showToast('Por favor escribe la ubicación manual.', 'warning');
            }
            
            // Normalizamos la base para buscar en el estado
            let locationBaseRaw = locationType === 'OTRA' ? locationManual : locationType;
            let locationBase = locationBaseRaw.trim().toUpperCase();
            
            // Contamos cuántas hay temporales que coincidan con esta base de forma robusta
            const countInTemp = tempUserLocations.filter(l => {
                const match = l.match(/^(.*?)\s*(\d+)$/);
                const lBase = match ? match[1].trim().toUpperCase() : l.trim().toUpperCase();
                return lBase === locationBase;
            }).length;

            // Obtenemos el contador global actual
            const currentGlobalCount = state.locations[locationBase] || 0;
            
            // Calculamos el siguiente ID
            const nextCount = currentGlobalCount + 1 + countInTemp;
            const finalId = String(nextCount).padStart(2, '0');
            
            // Construimos el string final usando el nombre original (raw) para mantener formato visual
            const locationWithId = `${locationBaseRaw.trim()} ${finalId}`;
            
            tempUserLocations.push(locationWithId);
            renderTempLocations();
            
            if (locationType === 'OTRA') {
                elements.userForm.locationManual.value = '';
            }
        });
        
        elements.userForm.locationsList.addEventListener('click', (e) => {
            if (e.target.classList.contains('remove-loc-btn')) {
                const index = parseInt(e.target.dataset.index);
                tempUserLocations.splice(index, 1);
                renderTempLocations();
            }
        });

        elements.userForm.createBtn.addEventListener('click', () => {
            if (state.readOnlyMode) return;
            const name = elements.userForm.name.value.trim();
            const area = elements.userForm.areaSelect.value;
            
            if (!name || !area) {
                return showToast('Nombre y Área son obligatorios.', 'error');
            }
            
            if (tempUserLocations.length === 0) {
                const locationType = elements.userForm.locationSelect.value;
                const locationManual = elements.userForm.locationManual.value.trim();
                
                if (locationType === 'OTRA' && !locationManual) {
                    return showToast('Debes añadir al menos una ubicación.', 'error');
                }
                
                const locationBase = locationType === 'OTRA' ? locationManual : locationType;
                const locationId = String((state.locations[locationBase] || 0) + 1).padStart(2, '0');
                tempUserLocations.push(`${locationBase} ${locationId}`);
            }

            const createUserAction = () => {
                tempUserLocations.forEach(locFull => {
                    const baseMatch = locFull.match(/^(.*)\s\d+$/);
                    if (baseMatch) {
                        const base = baseMatch[1];
                        state.locations[base] = (state.locations[base] || 0) + 1;
                    }
                });
                
                const primaryLocation = tempUserLocations[0];
                const locationBase = primaryLocation.replace(/\s\d+$/, '');

                const newUser = { 
                    name, 
                    area, 
                    location: locationBase, 
                    locationWithId: primaryLocation, 
                    locations: [...tempUserLocations], 
                    id: generateUUID() 
                };
                
                state.resguardantes.push(newUser);
                state.activeResguardante = newUser;
                
                logActivity('Usuario creado', `Nombre: ${name}, Área: ${area}, Ubicaciones: ${tempUserLocations.join(', ')}`);

                renderUserList();
                populateAreaSelects();
                populateReportFilters();
                saveState();
                showToast(`Usuario ${name} creado y activado.`);
                updateActiveUserBanner();
                
                elements.userForm.name.value = '';
                elements.userForm.locationManual.value = '';
                elements.userForm.locationSelect.value = 'OFICINA';
                elements.userForm.locationManual.classList.add('hidden');
                tempUserLocations = []; 
                renderTempLocations();
            };

            const existingUser = state.resguardantes.find(u => u.name.trim().toLowerCase() === name.toLowerCase());
            if (existingUser) {
                showConfirmationModal('Usuario Existente', `El usuario "${name}" ya existe. ¿Confirmas que deseas crear este registro nuevo?`, createUserAction);
            } else {
                createUserAction();
            }
        });
        // --- FIN MEJORA MULTI-UBICACIÓN ---

        elements.userForm.locationSelect.addEventListener('change', e => elements.userForm.locationManual.classList.toggle('hidden', e.target.value !== 'OTRA'));
        
        // ============================================================
        // NUEVOS LISTENERS PARA EL MODAL DE EDICIÓN (Multi-Ubicación)
        // ============================================================
        
        elements.editUserModal.locationType.addEventListener('change', e => {
             elements.editUserModal.locationManual.classList.toggle('hidden', e.target.value !== 'OTRA');
        });
        
        elements.editUserModal.addLocationBtn.addEventListener('click', () => {
             if (state.readOnlyMode) return;
             const locType = elements.editUserModal.locationType.value;
             const locManual = elements.editUserModal.locationManual.value.trim();
             
             if (locType === 'OTRA' && !locManual) {
                 return showToast('Por favor escribe la ubicación manual.', 'warning');
             }
             
             const base = locType === 'OTRA' ? locManual : locType;
             
             // Calcular nuevo ID
             // Obtenemos cuántas veces existe esta base en las ubicaciones globales
             const currentGlobalCount = state.locations[base] || 0;
             
             // Obtenemos cuántas de este tipo ya tiene el usuario ACTUALMENTE en edición
             // OJO: Aquí hay un detalle lógico. Si estamos editando, y agregamos una nueva,
             // deberíamos incrementar el contador global. 
             // Para simplificar visualmente en el modal, calculamos el siguiente ID disponible globalmente.
             // PERO, si el usuario ya tiene "OFICINA 01", y el global va en 05, la nueva debería ser 06.
             
             // Estrategia segura: Calcular ID basado en el contador global + 1 + las que hayamos agregado en esta sesión temporal
             const addedInSession = tempEditUserLocations.filter(l => l.startsWith(base) && !state.resguardantes[elements.editUserModal.saveBtn.dataset.userIndex].locations?.includes(l)).length;
             
             const nextIdNum = currentGlobalCount + 1 + addedInSession;
             const finalId = String(nextIdNum).padStart(2, '0');
             const newLoc = `${base} ${finalId}`;
             
             tempEditUserLocations.push(newLoc);
             renderEditTempLocations();
             
             if (locType === 'OTRA') elements.editUserModal.locationManual.value = '';
        });
        
        elements.editUserModal.locationsList.addEventListener('click', (e) => {
             if (e.target.classList.contains('remove-edit-loc-btn')) {
                 const index = parseInt(e.target.dataset.index);
                 tempEditUserLocations.splice(index, 1);
                 renderEditTempLocations();
             }
        });

        // Listener de GUARDAR en el modal de edición actualizado
        elements.editUserModal.saveBtn.addEventListener('click', e => {
            const index = e.target.dataset.userIndex;
            const user = state.resguardantes[index];
            const oldName = user.name;
            const newName = elements.editUserModal.name.value.trim();
            const newArea = elements.editUserModal.areaSelect.value;
            
            if (tempEditUserLocations.length === 0) {
                return showToast('El usuario debe tener al menos una ubicación.', 'error');
            }
            
            // Actualizar datos básicos
            user.name = newName;
            user.area = newArea;
            user.locations = [...tempEditUserLocations]; // Guardar el array completo
            
            // La primera es la principal
            user.locationWithId = tempEditUserLocations[0];
            user.location = tempEditUserLocations[0].replace(/\s\d+$/, '');
            
            // Actualizar nombres en inventario si cambió
            if (oldName !== newName) {
                state.inventory.forEach(i => { if(i['NOMBRE DE USUARIO'] === oldName) i['NOMBRE DE USUARIO'] = newName; });
                state.additionalItems.forEach(i => { if(i.usuario === oldName) i.usuario = newName; });
            }

            if (state.activeResguardante && state.activeResguardante.id === user.id) {
                state.activeResguardante = user;
            }
            
            // Recalcular contadores globales basados en la nueva realidad
            recalculateLocationCounts();

            elements.editUserModal.modal.classList.remove('show');
            renderUserList(); 
            populateReportFilters();
            saveState(); 
            logActivity('Usuario editado', `Nombre: ${newName}, Ubicaciones: ${user.locations.length}`);
            showToast('Usuario actualizado correctamente.');
        });

        // ============================================================
        
        elements.userForm.list.addEventListener('click', e => {
            const button = e.target.closest('button');
            const icon = e.target.closest('i.location-photo-btn');
            const userInfoClick = e.target.closest('.user-info-clickable');
            
            if (userInfoClick) {
                const userId = userInfoClick.dataset.userId;
                if (userId) {
                    showUserDetailView(userId);
                    return; 
                }
            }

            if(icon) {
                const locationId = icon.dataset.locationId;
                if (locationId) showPhotoModal('location', locationId);
                return;
            }

            if (!button || state.readOnlyMode) return;
            const index = parseInt(button.dataset.index, 10);
            
            if (button.classList.contains('activate-user-btn')) {
                const user = state.resguardantes[index];
                state.activeResguardante = user;
                logActivity('Usuario activado', `Usuario: ${user.name}`);
                showToast(`Usuario ${user.name} activado.`);
                renderUserList();
                updateActiveUserBanner();
            } else if (button.classList.contains('edit-user-btn')) {
                showEditUserModal(index);
            } else if (button.classList.contains('delete-user-btn')) {
                const user = state.resguardantes[index];
                const assignedItemsCount = state.inventory.filter(item => item['NOMBRE DE USUARIO'] === user.name).length;
                
                let title = '¿Eliminar Usuario?';
                let text = `¿Estás seguro de que quieres eliminar a ${user.name}?`;

                if (assignedItemsCount > 0) {
                    title = '¡Advertencia! Usuario con Bienes Asignados';
                    text = `El usuario ${user.name} tiene ${assignedItemsCount} bien(es) bajo su resguardo. Si lo eliminas, estos bienes quedarán sin un responsable válido. ¿Estás seguro de que quieres continuar?`;
                }

                showConfirmationModal(title, text, () => {
                    const recentlyDeleted = { item: user, originalIndex: index };
                    state.resguardantes.splice(index, 1);
                    if (state.activeResguardante?.id === user.id) {
                        state.activeResguardante = null;
                        updateActiveUserBanner();
                    }

                    recalculateLocationCounts();
                    renderUserList();
                    populateReportFilters();
                    logActivity('Usuario eliminado', `Nombre: ${user.name} (tenía ${assignedItemsCount} bienes)`);
                    
                    showUndoToast('Usuario eliminado.', () => {
                        state.resguardantes.splice(recentlyDeleted.originalIndex, 0, recentlyDeleted.item);
                        recalculateLocationCounts();
                        renderUserList(); 
                        saveState(); 
                        showToast('Acción deshecha.');
                        logActivity('Acción deshecha', `Restaurado usuario eliminado: ${user.name}`);
                    });

                    saveState();
                });
            }
        });
        
        elements.activeUserBanner.deactivateBtn.addEventListener('click', () => {
            if (state.activeResguardante) {
                logActivity('Usuario desactivado', `Usuario: ${state.activeResguardante.name}`);
                showToast(`Usuario ${state.activeResguardante.name} desactivado.`);
                state.activeResguardante = null;
                updateActiveUserBanner();
                renderUserList(); 
            }
        });
        
        elements.adicionales.form.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                const form = e.target.form;
                const focusable = Array.from(form.querySelectorAll('input, select, button, textarea'));
                const index = focusable.indexOf(e.target);
                const nextElement = focusable[index + 1];
                if (nextElement) {
                    nextElement.focus();
                } else {
                     elements.adicionales.addBtn.click();
                }
            }
        });

        elements.adicionales.addBtn.addEventListener('click', () => {
            if (state.readOnlyMode) return;
            if (!state.activeResguardante) return showToast('Debe activar un usuario resguardante para registrar bienes.', 'error');
            
            const formData = new FormData(elements.adicionales.form);
            const newItem = Object.fromEntries(formData.entries());
            if (!newItem.descripcion) return showToast('La descripción es obligatoria.', 'error');
            
            const newSerie = newItem.serie.trim();
            if (newSerie && state.serialNumberCache.has(newSerie.toLowerCase())) {
                return showToast('Advertencia: La serie de este bien ya existe en el inventario.', 'warning');
            }

            newItem.usuario = state.activeResguardante.name;
            newItem.id = generateUUID();
            newItem.fechaRegistro = new Date().toISOString();

            const finalizeItemAddition = (item) => {
                state.additionalItems.push(item);
                logActivity('Bien adicional registrado', `Descripción: ${item.descripcion}, Usuario: ${item.usuario}`);
                showToast('Bien adicional registrado.');
                elements.adicionales.form.reset();
                document.getElementById('ad-clave').value = '';
                document.querySelector('#adicional-form input[name="personal"][value="No"]').checked = true;
                renderAdicionalesList(); saveState(); renderDashboard(); updateSerialNumberCache();
                document.getElementById('ad-clave').focus();
            };

            if (newItem.personal === 'Si') {
                const { modal, siBtn, noBtn } = elements.formatoEntradaModal;
                modal.classList.add('show');
                const siHandler = () => { newItem.tieneFormatoEntrada = true; finalizeItemAddition(newItem); closeModal(); };
                const noHandler = () => { newItem.tieneFormatoEntrada = false; finalizeItemAddition(newItem); closeModal(); };
                const closeModal = () => { modal.classList.remove('show'); siBtn.removeEventListener('click', siHandler, { once: true }); noBtn.removeEventListener('click', noHandler, { once: true }); };
                siBtn.addEventListener('click', siHandler, { once: true });
                noBtn.addEventListener('click', noHandler, { once: true });
            } else {
                finalizeItemAddition(newItem);
            }
        });

        elements.adicionales.areaFilter.addEventListener('change', () => {
            populateAdicionalesFilters(); 
            renderAdicionalesList(); 
        });
        elements.adicionales.userFilter.addEventListener('change', renderAdicionalesList);
        
        elements.adicionales.list.addEventListener('click', e => {
            const itemClickable = e.target.closest('.adicional-item-clickable');
            const isButton = e.target.closest('button');

            if (itemClickable && !isButton) {
                const id = itemClickable.dataset.id;
                if (id) {
                    showAdicionalDetailView(id);
                    return; 
                }
            }

            if (state.readOnlyMode) return;

            const editBtn = e.target.closest('.edit-adicional-btn');
            const deleteBtn = e.target.closest('.delete-adicional-btn');
            const photoBtn = e.target.closest('.adicional-photo-btn');
            
            const id = editBtn?.dataset.id || deleteBtn?.dataset.id || photoBtn?.dataset.id;
            
            if (!id) return;
            
            const item = state.additionalItems.find(i => i.id === id);
            if (!item) return;

            if (editBtn) {
                showEditAdicionalModal(id);
            }
            
            if (deleteBtn) {
                if (state.additionalPhotos[id]) {
                    showTransferPhotoModal(item);
                } else {
                    showConfirmationModal('Eliminar Bien Adicional', `¿Seguro que quieres eliminar "${item.descripcion}"?`, () => {
                        deleteAdditionalItem(item.id, false); 
                    });
                }
            }
            
            if (photoBtn) {
                showPhotoModal('additional', id);
            }
        });

        elements.adicionales.printResguardoBtn.addEventListener('click', () => {
            const selectedArea = elements.adicionales.areaFilter.value;
            const selectedUser = elements.adicionales.userFilter.value;

            const data = {
                isForArea: selectedArea !== 'all' && selectedUser === 'all',
                isForUser: selectedUser !== 'all',
                filterArea: selectedArea,
                filterUser: selectedUser
            };

            showPreprintModal('adicionales_informe', data);
        });
        elements.reports.exportXlsxBtn.addEventListener('click', exportInventoryToXLSX);
        elements.reports.exportLabelsXlsxBtn.addEventListener('click', exportLabelsToXLSX);
        
        elements.reports.areaFilter.addEventListener('change', () => {
            populateReportFilters(); 
            renderAreaProgress();   
            renderReportStats();     
        });

        elements.reports.userFilter.addEventListener('change', () => {
            renderAreaProgress(); 
            renderReportStats();  
        });

        elements.reports.reportButtons.forEach(button => {
            button.addEventListener('click', () => {
                const reportType = button.dataset.reportType;
                if (!reportType) return;
                
                let data = {};
                if (reportType === 'adicionales_informe') {
                    const selectedArea = elements.reports.areaFilter.value;
                    const selectedUser = elements.reports.userFilter.value;
                    data.isForArea = selectedArea !== 'all' && selectedUser === 'all';
                    data.isForUser = selectedUser !== 'all';
                }
                
                const preprintReports = [
                    'session_summary', 'area_closure', 'simple_pending', 
                    'individual_resguardo', 'adicionales_informe'
                ];
                
                const directReports = {
                    'tasks_report': () => showPreprintModal('tasks_report', data), 
                    'inventory': () => generateInventoryReport(),
                    'labels': () => {
                        const selectedUser = elements.reports.userFilter.value;
                        const selectedArea = elements.reports.areaFilter.value;
                        let items = state.inventory.filter(item => item['IMPRIMIR ETIQUETA'] === 'SI');

                        if (selectedArea !== 'all') {
                            items = items.filter(item => item.areaOriginal === selectedArea);
                        }
                        if (selectedUser !== 'all') {
                            items = items.filter(item => item['NOMBRE DE USUARIO'] === selectedUser);
                        }
                        renderReportTable(items, 'Reporte de Etiquetas', { withCheckboxes: true, reportType: 'labels', headers: ['Acción', 'Clave Única', 'Descripción', 'Usuario'] });
                    },
                    'pending': () => {
                        let pendingItems = state.inventory.filter(item => item.UBICADO !== 'SI');
                        const selectedArea = elements.reports.areaFilter.value;
                        if (selectedArea !== 'all') {
                            pendingItems = pendingItems.filter(item => item.areaOriginal === selectedArea);
                        }
                        renderReportTable(pendingItems, 'Reporte de Bienes Pendientes', { withCheckboxes: false, headers: ['Clave Única', 'Descripción', 'Serie', 'Área Original'] });
                    },
                    'notes': () => {
                        const selectedUser = elements.reports.userFilter.value;
                        const selectedArea = elements.reports.areaFilter.value;
                        let items = state.inventory.filter(item => state.notes[item['CLAVE UNICA']]);
                        
                        if (selectedArea !== 'all') {
                            items = items.filter(item => item.areaOriginal === selectedArea);
                        }
                        if (selectedUser !== 'all') {
                            items = items.filter(item => item['NOMBRE DE USUARIO'] === selectedUser);
                        }
                        renderReportTable(items, 'Reporte de Notas', { withCheckboxes: true, reportType: 'notes', headers: ['Acción', 'Clave Única', 'Descripción', 'Nota'] });
                    },
                    'mismatched': () => {
                        const selectedUser = elements.reports.userFilter.value;
                        const selectedArea = elements.reports.areaFilter.value;
                        let items = state.inventory.filter(item => item.areaIncorrecta);

                        if (selectedArea !== 'all') {
                            items = items.filter(item => item.areaOriginal === selectedArea);
                        }
                        if (selectedUser !== 'all') {
                            items = items.filter(item => item['NOMBRE DE USUARIO'] === selectedUser);
                        }
                        renderReportTable(items, 'Reporte de Bienes Fuera de Área', { withCheckboxes: true, reportType: 'mismatched', headers: ['Acción', 'Clave Única', 'Descripción', 'Área Original', 'Usuario/Área Actual'] });
                    },
                    'institutional_adicionales': () => generateInstitutionalAdicionalesReport()
                };

                if (reportType === 'individual_resguardo') {
                    const selectedUser = elements.reports.userFilter.value;
                    const selectedArea = elements.reports.areaFilter.value;

                    // CASO 1: Si hay usuario seleccionado -> Comportamiento Clásico (Uno por uno)
                    if (selectedUser !== 'all') {
                        const userInventoryItems = state.inventory.filter(item => item['NOMBRE DE USUARIO'] === selectedUser);
                        const userAdditionalItems = state.additionalItems.filter(item => item.usuario === selectedUser);
                        
                        data = { ...data, isForUser: true }; 

                        if (userAdditionalItems.length > 0) {
                            // Lógica existente de preguntar por adicionales...
                            const { modal, yesBtn, noBtn } = elements.addAdicionalesConfirm;
                            modal.classList.add('show');
                            const yesHandler = () => {
                                const combinedItems = [...userInventoryItems, ...userAdditionalItems];
                                data.items = combinedItems;
                                data.title = 'Resguardo Individual (Inventario y Adicionales)';
                                data.isAdicional = true; 
                                modal.classList.remove('show');
                                showPreprintModal(reportType, data);
                            };
                            const noHandler = () => {
                                data.items = userInventoryItems;
                                data.title = 'Resguardo Individual de Bienes';
                                data.isAdicional = false;
                                modal.classList.remove('show');
                                showPreprintModal(reportType, data);
                            };
                            yesBtn.onclick = yesHandler;
                            noBtn.onclick = noHandler;
                        } else {
                            data.items = userInventoryItems;
                            data.title = 'Resguardo Individual de Bienes';
                            data.isAdicional = false;
                            showPreprintModal(reportType, data);
                        }
                    } 
                    // CASO 2: Si NO hay usuario, pero SÍ hay Área -> NUEVO MODO MASIVO
                    else if (selectedArea !== 'all') {
                        openBatchPrintModal(selectedArea);
                    }
                    // CASO 3: Nada seleccionado
                    else {
                        return showToast('Por favor, selecciona un Área para impresión masiva o un Usuario específico.', 'error');
                    }
                } else if (preprintReports.includes(reportType)) {
                    showPreprintModal(reportType, data);
                } else if (directReports[reportType]) {
                    directReports[reportType]();
                }
            });
        });

        // --- Listeners para Impresión Masiva (NUEVO) ---
        if (elements.batchModal && elements.batchModal.closeBtn) {
            elements.batchModal.closeBtn.addEventListener('click', () => elements.batchModal.modal.classList.remove('show'));
            elements.batchModal.cancelBtn.addEventListener('click', () => elements.batchModal.modal.classList.remove('show'));
            
            elements.batchModal.usersListContainer.addEventListener('change', updateBatchCount);
            
            elements.batchModal.selectAllBtn.addEventListener('click', () => {
                document.querySelectorAll('.batch-user-checkbox').forEach(cb => cb.checked = true);
                updateBatchCount();
            });
            
            elements.batchModal.deselectAllBtn.addEventListener('click', () => {
                document.querySelectorAll('.batch-user-checkbox').forEach(cb => cb.checked = false);
                updateBatchCount();
            });

            elements.batchModal.generateBtn.addEventListener('click', generateBatchReport);
        } 

        elements.reports.reportViewModal.modal.addEventListener('click', (e) => {
            if(state.readOnlyMode) return;
            const saveBtn = e.target.closest('.save-new-clave-btn');
            
            const doneBtn = e.target.closest('.report-label-done-btn');
            if (doneBtn) {
                const clave = doneBtn.dataset.clave;
                const item = state.inventory.find(i => i['CLAVE UNICA'] === clave);
                if (item) {
                    item['IMPRIMIR ETIQUETA'] = 'NO';
                    logActivity('Etiqueta marcada como HECHA', `Clave: ${clave}`);
                    showToast(`Se quitó la marca de etiqueta para la clave ${clave}.`);
                    saveState();
                    doneBtn.closest('tr').remove();
                }
                return; 
            }
            
            if (saveBtn) {
                const itemId = saveBtn.dataset.id;
                const row = saveBtn.closest('tr');
                const input = row.querySelector('.new-clave-input');
                const newClave = input.value.trim();

                if (newClave && state.serialNumberCache.has(newClave.toLowerCase())) {
                    const item = state.additionalItems.find(i => i.id === itemId);
                    if (!item || item.claveAsignada !== newClave) {
                        return showToast('Error: Esa clave o número de serie ya existe en el inventario.', 'error');
                    }
                }

                const itemIndex = state.additionalItems.findIndex(i => i.id === itemId);
                if (itemIndex !== -1) {
                    state.additionalItems[itemIndex].claveAsignada = newClave;
                    updateSerialNumberCache();
                    saveState();
                    logActivity('Clave Asignada a Bien Adicional actualizada', `ID: ${itemId}, Nueva Clave: ${newClave || 'NINGUNA'}`);
                    showToast('Clave actualizada con éxito.', 'success');
                }
            }
        });

        elements.reports.reportViewModal.modal.addEventListener('change', (e) => {
            const checkbox = e.target;
            if (checkbox.classList.contains('report-item-checkbox')) {
                const clave = checkbox.dataset.clave;
                const reportType = checkbox.dataset.reportType;
                if (clave && reportType && state.reportCheckboxes[reportType]) {
                    state.reportCheckboxes[reportType][clave] = checkbox.checked;
                    saveState();
                }
            }
            else if (checkbox.classList.contains('institutional-report-checkbox')) {
                const itemId = checkbox.dataset.id;
                if (itemId) {
                    state.institutionalReportCheckboxes[itemId] = checkbox.checked;
                    saveState();
                }
            }
        });


        elements.areaClosure.cancelBtn.addEventListener('click', () => elements.areaClosure.modal.classList.remove('show'));
        elements.noteSaveBtn.addEventListener('click', e => {
            if (state.readOnlyMode) return;
            const claves = JSON.parse(e.target.dataset.claves);
            const noteText = elements.noteTextarea.value.trim();
            claves.forEach(clave => state.notes[clave] = noteText);
            logActivity('Nota guardada', `Nota para clave(s): ${claves.join(', ')}`);
            showToast('Nota(s) guardada(s).');
            elements.notesModal.classList.remove('show');
            filterAndRenderInventory(); saveState();
        });
        elements.editAdicionalModal.saveBtn.addEventListener('click', () => {
            const id = elements.editAdicionalModal.saveBtn.dataset.id;
            const itemIndex = state.additionalItems.findIndex(i => i.id === id);
            if (itemIndex === -1) return;
            const formData = new FormData(elements.editAdicionalModal.form);
            const updatedData = Object.fromEntries(formData.entries());
            state.additionalItems[itemIndex] = { ...state.additionalItems[itemIndex], ...updatedData };
            elements.editAdicionalModal.modal.classList.remove('show');
            renderAdicionalesList(); updateSerialNumberCache(); saveState();
            logActivity('Bien adicional editado', `ID: ${id}`);
            showToast('Bien adicional actualizado.');
        });
        
        elements.photo.useCameraBtn.addEventListener('click', startCamera);

        elements.photo.switchToUploadBtn.addEventListener('click', () => {
            stopCamera();
            elements.photo.cameraViewContainer.classList.add('hidden');
            elements.photo.uploadContainer.classList.remove('hidden');
        });
        elements.photo.captureBtn.addEventListener('click', () => {
            const { cameraStream, photoCanvas, input } = elements.photo;
            const context = photoCanvas.getContext('2d');
            photoCanvas.width = cameraStream.videoWidth;
            photoCanvas.height = cameraStream.videoHeight;
            context.drawImage(cameraStream, 0, 0, photoCanvas.width, photoCanvas.height);
            
            photoCanvas.toBlob(blob => {
                if (blob) {
                    const type = input.dataset.type;
                    const id = input.dataset.id;
                    if (blob.size > 2 * 1024 * 1024) return showToast('La imagen es demasiado grande (máx 2MB).', 'error');

                    photoDB.setItem('photos', `${type}-${id}`, blob).then(() => {
                        if (type === 'inventory') { state.photos[id] = true; filterAndRenderInventory(); updateDetailViewPhoto(id); } 
                        else if (type === 'additional') { state.additionalPhotos[id] = true; renderAdicionalesList(); }
                        else if (type === 'location') { state.locationPhotos[id] = true; renderUserList(); }
                        logActivity('Foto capturada', `Tipo: ${type}, ID: ${id}`);
                        showToast(`Foto adjuntada.`);
                        elements.photo.modal.classList.remove('show');
                        stopCamera(); saveState();
                    }).catch(err => showToast('Error al guardar la foto.', 'error'));
                }
            }, 'image/jpeg', 0.9);
        });

        elements.photo.input.addEventListener('change', e => {
            const file = e.target.files[0];
            const type = e.target.dataset.type;
            const id = e.target.dataset.id;
            if (file && type && id) {
                if (file.size > 2 * 1024 * 1024) return showToast('La imagen es demasiado grande (máx 2MB).', 'error');
                photoDB.setItem('photos', `${type}-${id}`, file).then(() => {
                    if (type === 'inventory') { state.photos[id] = true; filterAndRenderInventory(); updateDetailViewPhoto(id); } 
                    else if (type === 'additional') { state.additionalPhotos[id] = true; renderAdicionalesList(); }
                    else if (type === 'location') { state.locationPhotos[id] = true; renderUserList(); }
                    logActivity('Foto subida', `Tipo: ${type}, ID: ${id}`);
                    showToast(`Foto adjuntada.`);
                    stopCamera();
                    elements.photo.modal.classList.remove('show'); saveState();
                }).catch(err => showToast('Error al guardar la foto.', 'error'));
            }
        });
        elements.photo.deleteBtn.addEventListener('click', e => {
            const type = e.target.dataset.type;
            const id = e.target.dataset.id;
            showConfirmationModal('Eliminar Foto', `¿Seguro que quieres eliminar la foto?`, () => {
                photoDB.deleteItem('photos', `${type}-${id}`).then(() => {
                    if (type === 'inventory') { delete state.photos[id]; filterAndRenderInventory(); updateDetailViewPhoto(id); } 
                    else if (type === 'additional') { delete state.additionalPhotos[id]; renderAdicionalesList(); }
                    else if (type === 'location') { delete state.locationPhotos[id]; renderUserList(); }
                    logActivity('Foto eliminada', `Tipo: ${type}, ID: ${id}`);
                    showToast(`Foto eliminada.`);
                    stopCamera();
                    elements.photo.modal.classList.remove('show'); saveState();
                }).catch(err => showToast('Error al eliminar la foto.', 'error'));
            });
        });

        // --- CORRECCIÓN FINAL: Lista completa de botones de cierre ---
const closeButtons = [
elements.noteCancelBtn, 
elements.photo.closeBtn, 
elements.editUserModal.cancelBtn, 
elements.editAdicionalModal.cancelBtn, 
elements.qrDisplayModal.closeBtn, 
elements.itemDetailsModal.closeBtn, 
elements.preprintModal.cancelBtn, 
elements.layoutEditor.closeBtn, 
elements.transferPhotoModal.cancelBtn, 
elements.addAdicionalesConfirm.noBtn,
elements.detailView.closeBtn, 
elements.reports.reportViewModal.closeBtn,       
elements.reports.reportViewModal.closeFooterBtn, 
elements.userDetailView.closeBtn,                
elements.userDetailView.closeFooterBtn,          
elements.adicionalDetailView.closeBtn,           
elements.adicionalDetailView.closeFooterBtn      
];

closeButtons.forEach(btn => {
if (btn) { 
    btn.addEventListener('click', () => {
        stopCamera(); 
        const modalOverlay = btn.closest('.modal-overlay');
        if (modalOverlay) {
            modalOverlay.classList.remove('show');
        }
    });
}
});
        elements.qrScannerCloseBtn.addEventListener('click', stopQrScanner);

        elements.settings.themes.forEach(btn => btn.addEventListener('click', () => updateTheme(btn.dataset.theme)));
        
        elements.settings.exportSessionBtn.addEventListener('click', () => exportSession(false));
        elements.settings.finalizeInventoryBtn.addEventListener('click', () => {
            showConfirmationModal('Finalizar Inventario', 'Esto creará un archivo de respaldo final de solo lectura. No podrás realizar más cambios en este inventario. ¿Estás seguro?', () => {
                exportSession(true);
            });
        });
        
        elements.settings.importSessionBtn.addEventListener('click', () => elements.settings.importFileInput.click());
        
        elements.settings.importFileInput.addEventListener('change', async (event) => {
            const file = event.target.files[0];
            if (!file || !file.name.endsWith('.zip')) {
                return showToast('Por favor, selecciona un archivo de sesión .zip válido.', 'error');
            }
            
            logActivity('Importación de sesión', `Archivo: ${file.name}`);
            const { overlay, text } = elements.loadingOverlay;
            text.textContent = 'Abriendo archivo de sesión...';
            overlay.classList.add('show');

            try {
                const jszip = new JSZip();
                const zip = await jszip.loadAsync(file);
                
                const sessionFile = zip.file('session.json');
                if (!sessionFile) throw new Error('El archivo .zip no contiene un session.json válido.');

                const sessionData = await sessionFile.async('string');
                const importedState = JSON.parse(sessionData);
                
                const photoFolder = zip.folder("photos");
                if (photoFolder) {
                    overlay.classList.remove('show');
                    elements.importProgress.modal.classList.add('show');
                    
                    const photoFiles = [];
                    photoFolder.forEach((relativePath, file) => {
                        if (!file.dir) {
                            photoFiles.push(file);
                        }
                    });
                    
                    const totalPhotos = photoFiles.length;
                    let processedPhotos = 0;
                    
                    for (const file of photoFiles) {
                        const key = file.name.split('/').pop();
                        const blob = await file.async("blob");
                        await photoDB.setItem('photos', key, blob);
                        processedPhotos++;
                        
                        const percent = Math.round((processedPhotos / totalPhotos) * 100);
                        elements.importProgress.bar.style.width = `${percent}%`;
                        elements.importProgress.bar.textContent = `${percent}%`;
                        elements.importProgress.text.textContent = `Restaurando foto ${processedPhotos} de ${totalPhotos}...`;
                    }
                    
                    elements.importProgress.modal.classList.remove('show');
                }
                
                
                const layoutImageFolder = zip.folder("layoutImages");
                if (layoutImageFolder) {
                     const layoutImageFiles = [];
                    layoutImageFolder.forEach((relativePath, file) => {
                        if (!file.dir) layoutImageFiles.push(file);
                    });
                    for (const file of layoutImageFiles) {
                        const key = file.name.split('/').pop();
                        const blob = await file.async("blob");
                        await photoDB.setItem('layoutImages', key, blob);
                    }
                }
                
                localStorage.setItem('inventarioProState', JSON.stringify(importedState));
                showToast('Sesión importada con éxito. Recargando aplicación...', 'success');
                setTimeout(() => window.location.reload(), 1500);

            } catch (err) {
                console.error("Error al importar la sesión:", err);
                showToast('Error fatal al importar el archivo de sesión.', 'error');
                overlay.classList.remove('show');
                elements.importProgress.modal.classList.remove('show');
            } finally {
                event.target.value = '';
            }
        });

        elements.settings.loadedListsContainer.addEventListener('click', (e) => {
            if (state.readOnlyMode) return;
            const deleteBtn = e.target.closest('.delete-list-btn');
            const generateBtn = e.target.closest('.generate-area-report-btn'); 
            const reprintBtn = e.target.closest('.reprint-area-report-btn');   

            if (deleteBtn) {
                const listId = Number(deleteBtn.dataset.listId);
                const listToDelete = state.inventory.find(i => i.listId === listId);
                if (!listToDelete) return;
                const areaOriginal = listToDelete.areaOriginal;
                const affectedUsers = state.resguardantes.filter(u => u.area === areaOriginal);
                const affectedAdicionales = state.additionalItems.filter(item => 
                    affectedUsers.some(user => user.name === item.usuario)
                );
                if (affectedUsers.length > 0 || affectedAdicionales.length > 0) {
                    showReassignModal(listId, areaOriginal, affectedUsers, affectedAdicionales);
                } else {
                    showConfirmationModal('Eliminar Listado', `¿Seguro que quieres eliminar el listado del archivo "${listToDelete.fileName}"?`, () => {
                        deleteListAndRefresh(listId);
                    });
                }
            }
            
            if (generateBtn || reprintBtn) {
                const areaId = generateBtn?.dataset.areaId || reprintBtn?.dataset.areaId;
                let data = { areaId };
                
                if (reprintBtn) { 
                     const closedInfo = state.closedAreas[areaId];
                     if (closedInfo) {
                        data = { ...data, responsible: closedInfo.responsible, location: closedInfo.location };
                     } else {
                         data = { ...data, responsible: '', location: '' }; 
                     }
                } else { 
                    data = { ...data, responsible: state.areaDirectory[areaId]?.name || '', location: '' };
                }

                showPreprintModal('area_closure', data);
            }
        });
        
        let aboutClickCount = 0;
        elements.settings.aboutHeader.addEventListener('click', () => {
            aboutClickCount++;
            if (aboutClickCount >= 5) {
                elements.settings.aboutContent.classList.remove('hidden');
            }
        });
        
        elements.settings.directoryContainer.addEventListener('click', (e) => {
            const field = e.target.closest('.editable-field');
            if (!field || state.readOnlyMode) return;
            
            field.dataset.originalText = field.textContent.trim();
            
            field.contentEditable = "true";
            field.focus();
            
            try {
                document.execCommand('selectAll', false, null);
            } catch(err) {
                window.getSelection().selectAllChildren(field);
            }
        });

        elements.settings.directoryContainer.addEventListener('blur', (e) => {
            const field = e.target.closest('.editable-field[contenteditable="true"]');
            if (!field) return; 
            
            field.contentEditable = "false";
            
            const areaKey = field.dataset.areaKey;
            const fieldName = field.dataset.field;
            const newText = field.textContent.trim();
            
            if (newText !== state.areaDirectory[areaKey][fieldName]) {
                state.areaDirectory[areaKey][fieldName] = newText;
                saveState();
                logActivity('Directorio editado (en línea)', `Área: ${areaKey}, Campo: ${fieldName}`);
                showToast('Responsable de área actualizado.');
            }
        }, true); 

        elements.settings.directoryContainer.addEventListener('keydown', (e) => {
            const field = e.target.closest('.editable-field[contenteditable="true"]');
            if (!field) return;
            
            if (e.key === 'Enter') {
                e.preventDefault(); 
                field.blur(); 
            } 
            else if (e.key === 'Escape') {
                e.preventDefault();
                field.textContent = field.dataset.originalText || '';
                field.blur(); 
            }
        });
        elements.log.showBtn.addEventListener('click', () => {
            elements.log.content.textContent = state.activityLog.join('\n');
            elements.log.modal.classList.add('show');
        });
        elements.log.closeBtn.addEventListener('click', () => {
            elements.log.modal.classList.remove('show');
        });

        elements.layoutEditor.openBtn.addEventListener('click', () => {
            if(state.readOnlyMode) return showToast('Modo de solo lectura: no se puede editar el croquis.', 'warning');
            switchLayoutPage(state.currentLayoutPage || 'page1');
            elements.layoutEditor.modal.classList.add('show');
            handleModalNavigation(elements.layoutEditor.modal);
        });
        elements.layoutEditor.saveBtn.addEventListener('click', () => {
            if(state.readOnlyMode) return;
            saveLayoutPositions();
            saveState();
            showToast('Croquis guardado con éxito.');
        });
        elements.layoutEditor.printBtn.addEventListener('click', printLayout);

        
        elements.layoutEditor.addImageBtn.addEventListener('click', () => {
            if(state.readOnlyMode) return;
            elements.layoutEditor.imageInput.click();
        });
        
        elements.layoutEditor.imageInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = async (event) => {
                const dataUrl = event.target.result;
                const imageId = `img_${Date.now()}`;
                const shapeId = `image-${imageId}`;
                
                try {
                    await photoDB.setItem('layoutImages', imageId, file);
                    
                    state.layoutImages[shapeId] = imageId;
                    
                    createShapeOnCanvas(shapeId, 20, 20, 300, 200, 'image', '', dataUrl, 0);
                    
                    saveLayoutPositions();
                    saveState();
                } catch (err) {
                    console.error('Error al guardar la imagen del croquis:', err);
                    showToast('Error al guardar la imagen.', 'error');
                }
            };
            reader.readAsDataURL(file);
            
            e.target.value = '';
        });
        

        
        elements.layoutEditor.sidebar.addEventListener('click', (e) => {
            const item = e.target.closest('.draggable-item');
            if (!item || state.readOnlyMode) return;
            
            const x = (elements.layoutEditor.canvasWrapper.scrollLeft + elements.layoutEditor.canvasWrapper.clientWidth / 2) - 90; 
            const y = (elements.layoutEditor.canvasWrapper.scrollTop + elements.layoutEditor.canvasWrapper.clientHeight / 2) - 30; 
            const snappedX = Math.round(x / 10) * 10;
            const snappedY = Math.round(y / 10) * 10;
            
            const locId = item.dataset.locationId;
            const areaId = item.dataset.areaId; 
            createShapeOnCanvas(locId, snappedX, snappedY, null, null, 'location', '', null, 0, areaId); 
            item.classList.add('hidden'); 
            
            saveLayoutPositions();
            showToast(`Ubicación ${locId} añadida al lienzo.`);
        });
        
        elements.layoutEditor.canvas.addEventListener('click', (e) => {
            const deleteBtn = e.target.closest('.layout-delete-btn');
            if (!deleteBtn || state.readOnlyMode) return;

            const shape = deleteBtn.closest('.layout-on-canvas');
            if (!shape) return;

            const id = shape.dataset.id;
            const type = shape.dataset.type;

            shape.remove();

            if (type === 'location') {
                const sidebarItem = document.querySelector(`.draggable-item[data-location-id="${id}"]`);
                if (sidebarItem) {
                    sidebarItem.classList.remove('hidden');
                }
            }
            
            if (type === 'image') {
                delete state.layoutImages[id];
            }
            
            saveLayoutPositions();
            saveState();
            showToast('Elemento eliminado del croquis.');
        });

        elements.layoutEditor.pageAdd.addEventListener('click', () => {
            const newPageKey = `page${Date.now()}`;
            const newPageName = `Página ${Object.keys(state.layoutPageNames).length + 1}`;
            state.mapLayout[newPageKey] = {};
            state.layoutPageNames[newPageKey] = newPageName;
            switchLayoutPage(newPageKey);
            saveState();
        });
        elements.layoutEditor.pageRemove.addEventListener('click', () => {
            const pageKeys = Object.keys(state.layoutPageNames);
            if (pageKeys.length <= 1) return showToast('No se puede eliminar la última página.', 'warning');
            
            showConfirmationModal('Eliminar Página', `¿Seguro que quieres eliminar la "${state.layoutPageNames[state.currentLayoutPage]}"?`, () => {
                delete state.mapLayout[state.currentLayoutPage];
                delete state.layoutPageNames[state.currentLayoutPage];
                const newPageKeys = Object.keys(state.layoutPageNames);
                switchLayoutPage(newPageKeys[0]); 
                saveState();
            });
        });
        elements.layoutEditor.pageReset.addEventListener('click', resetCurrentLayoutPage);
        elements.layoutEditor.pagePrev.addEventListener('click', () => {
            const pageKeys = Object.keys(state.layoutPageNames);
            const currentIndex = pageKeys.indexOf(state.currentLayoutPage);
            if (currentIndex > 0) switchLayoutPage(pageKeys[currentIndex - 1]);
        });
        elements.layoutEditor.pageNext.addEventListener('click', () => {
            const pageKeys = Object.keys(state.layoutPageNames);
            const currentIndex = pageKeys.indexOf(state.currentLayoutPage);
            if (currentIndex < pageKeys.length - 1) switchLayoutPage(pageKeys[currentIndex + 1]);
        });
        elements.layoutEditor.pageName.addEventListener('change', (e) => {
            const newName = e.target.value.trim();
            if (newName) {
                state.layoutPageNames[state.currentLayoutPage] = newName;
                saveState();
                showToast('Nombre de página actualizado.');
            }
        });
        
        const claveInput = document.getElementById('ad-clave');
        const serieInput = document.getElementById('ad-serie');
        const claveFeedback = document.getElementById('ad-clave-feedback');
        const serieFeedback = document.getElementById('ad-serie-feedback');

        const checkDuplicate = (value, feedbackElement) => {
            if (!value.trim()) {
                feedbackElement.textContent = '';
                return;
            }
            if (state.serialNumberCache.has(String(value).trim().toLowerCase())) {
                feedbackElement.textContent = 'Esta clave/serie ya existe en el inventario.';
            } else {
                feedbackElement.textContent = '';
            }
        };

        claveInput.addEventListener('input', debounce(() => {
            checkDuplicate(claveInput.value, claveFeedback);
        }, 400));

        serieInput.addEventListener('input', debounce(() => {
            checkDuplicate(serieInput.value, serieFeedback);
        }, 400));
        const inventoryTableBody = elements.inventory.tableBody;
        const photoPreviewPopover = document.getElementById('photo-preview-popover');
        const photoPreviewImg = document.getElementById('photo-preview-img');
        let currentPhotoUrl = null;
        let popoverTimeout;

        inventoryTableBody.addEventListener('mouseover', (e) => {
            const cameraIcon = e.target.closest('.camera-icon');
            if (!cameraIcon) return;
            
            clearTimeout(popoverTimeout); 

            const row = cameraIcon.closest('tr');
            const clave = row.dataset.clave;

            if (state.photos[clave]) {
                // CORRECCIÓN: Limpiar memoria previa ANTES de pedir una nueva
                if (currentPhotoUrl) {
                    URL.revokeObjectURL(currentPhotoUrl);
                    currentPhotoUrl = null;
                }

                photoPreviewPopover.classList.remove('hidden');
                // Placeholder ligero mientras carga
                photoPreviewImg.src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
                
                photoDB.getItem('photos', `inventory-${clave}`).then(imageBlob => {
                    if (!imageBlob) {
                        photoPreviewPopover.classList.add('hidden');
                        return;
                    };
                    // Doble verificación por si el usuario ya movió el mouse
                    if (currentPhotoUrl) URL.revokeObjectURL(currentPhotoUrl);
                    
                    currentPhotoUrl = URL.createObjectURL(imageBlob);
                    photoPreviewImg.src = currentPhotoUrl;
                }).catch(err => {
                    console.error("Error al cargar la previsualización de la foto:", err);
                    photoPreviewPopover.classList.add('hidden');
                });
            }
        });

        inventoryTableBody.addEventListener('mouseout', (e) => {
            const cameraIcon = e.target.closest('.camera-icon');
            if (!cameraIcon) return;

            popoverTimeout = setTimeout(() => {
                photoPreviewPopover.classList.add('hidden');
                photoPreviewImg.src = '';
                if (currentPhotoUrl) {
                    URL.revokeObjectURL(currentPhotoUrl);
                    currentPhotoUrl = null;
                }
            }, 100);
        });
        const userSearchInput = document.getElementById('user-search-input');
        if (userSearchInput) {
            userSearchInput.addEventListener('input', renderUserList);
        }

        elements.settings.importPhotosBtn.addEventListener('click', () => {
            if (state.readOnlyMode) return showToast('Modo de solo lectura: no se pueden importar fotos.', 'warning');
            if (state.inventory.length === 0) return showToast('Carga un inventario antes de importar fotos.', 'error');
            elements.settings.importPhotosInput.click();
        });

        elements.settings.importPhotosInput.addEventListener('change', async (e) => {
            const files = e.target.files;
            if (!files.length) return;

            const { modal, text, bar } = elements.importProgress;
            modal.classList.add('show');
            text.textContent = 'Iniciando importación de fotos...';
            bar.style.width = '0%';
            bar.textContent = '0%';

            const inventoryClaves = new Set(state.inventory.map(item => String(item['CLAVE UNICA'])));
            let successCount = 0;
            let errorCount = 0;
            const totalFiles = files.length;

            for (let i = 0; i < totalFiles; i++) {
                const file = files[i];
                const fileName = file.name;
                const clave = fileName.substring(0, fileName.lastIndexOf('.'));

                const percent = Math.round(((i + 1) / totalFiles) * 100);
                bar.style.width = `${percent}%`;
                bar.textContent = `${percent}%`;
                text.textContent = `Procesando ${i + 1} de ${totalFiles}: ${fileName}`;

                if (inventoryClaves.has(clave)) {
                    if (file.size > 2 * 1024 * 1024) {
                        console.warn(`Archivo ignorado (muy grande): ${fileName}`);
                        errorCount++;
                        continue;
                    }
                    try {
                        await photoDB.setItem('photos', `inventory-${clave}`, file);
                        state.photos[clave] = true;
                        successCount++;
                    } catch (err) {
                        console.error(`Error al guardar la foto ${fileName}:`, err);
                        errorCount++;
                    }
                } else {
                    console.warn(`Archivo ignorado (clave no encontrada): ${fileName}`);
                    errorCount++;
                }
            }
            
            modal.classList.remove('show');
            
            if (successCount > 0) {
                saveState();
                filterAndRenderInventory();
                showToast(`Importación completa: ${successCount} fotos guardadas con éxito.`, 'success');
            }
            if (errorCount > 0) {
                showToast(`${errorCount} archivos fueron ignorados (clave no encontrada o archivo muy grande). Revisa la consola para más detalles.`, 'warning');
            }
            
            elements.settings.importPhotosInput.value = '';
        });
        
        elements.settings.restorePhotosBtn.addEventListener('click', () => {
            if (state.readOnlyMode) return showToast('Modo de solo lectura activado.', 'warning');
            if (state.inventory.length === 0) return showToast('Carga un inventario antes de restaurar fotos.', 'error');
            elements.settings.restorePhotosInput.click();
        });

        elements.settings.restorePhotosInput.addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (!file) return;

            const { modal, text, bar } = elements.importProgress;
            modal.classList.add('show');

            const inventoryClaves = new Set(state.inventory.map(item => String(item['CLAVE UNICA'])));
            let successCount = 0;
            let ignoredCount = 0;

            try {
                const jszip = new JSZip();
                const zip = await jszip.loadAsync(file);

                const photoFolder = zip.folder("photos");
                if (!photoFolder) {
                    modal.classList.remove('show');
                    return showToast('Error: El archivo .zip no contiene una carpeta de fotos válida.', 'error');
                }

                const photoFiles = [];
                photoFolder.forEach((relativePath, file) => {
                    if (!file.dir) photoFiles.push(file);
                });
                const totalPhotos = photoFiles.length;

                for (let i = 0; i < totalPhotos; i++) {
                    const photoFile = photoFiles[i];
                    const key = photoFile.name.split('/').pop();
                    const clave = key.replace('inventory-', '').replace('additional-', '').replace('location-','');
                    
                    const percent = Math.round(((i + 1) / totalPhotos) * 100);
                    bar.style.width = `${percent}%`;
                    bar.textContent = `${percent}%`;
                    text.textContent = `Restaurando foto ${i + 1} de ${totalPhotos}...`;

                    if (inventoryClaves.has(clave)) {
                        const blob = await photoFile.async("blob");
                        await photoDB.setItem('photos', key, blob);
                        if (key.startsWith('inventory-')) state.photos[clave] = true;
                        if (key.startsWith('additional-')) state.additionalPhotos[clave] = true;
                        successCount++;
                    } else {
                        ignoredCount++;
                    }
                }

                modal.classList.remove('show');

                if (successCount > 0) {
                    saveState();
                    filterAndRenderInventory();
                    showToast(`${successCount} fotos restauradas y asociadas con éxito.`, 'success');
                }
                if (ignoredCount > 0) {
                    showToast(`${ignoredCount} fotos del backup fueron ignorados porque sus claves no se encontraron en el inventario actual.`, 'warning');
                }
                if (successCount === 0 && ignoredCount === 0) {
                     showToast('No se encontraron fotos en el backup para procesar.', 'info');
                }

            } catch (err) {
                modal.classList.remove('show');
                console.error("Error al restaurar fotos desde el backup:", err);
                showToast('Error al procesar el archivo .zip. Asegúrate de que es un backup válido.', 'error');
            } finally {
                elements.settings.restorePhotosInput.value = '';
            }
        });

        let tooltipTimeout; 
        const card = elements.dashboard.dailyProgressCard;
        const tooltip = elements.dashboard.progressTooltip;

        card.addEventListener('mouseenter', e => {
            clearTimeout(tooltipTimeout); 
            tooltip.style.pointerEvents = 'auto'; 

            const progressByDate = [...state.inventory, ...state.additionalItems].reduce((acc, item) => {
                const dateStr = item.fechaUbicado || item.fechaRegistro;
                if (!dateStr) return acc;
                
                const date = dateStr.slice(0, 10);
                if (!acc[date]) {
                    acc[date] = { inventory: 0, additional: 0 };
                }
                if (item.fechaUbicado) acc[date].inventory++;
                if (item.fechaRegistro) acc[date].additional++;
                
                return acc;
            }, {});

            const sortedDates = Object.keys(progressByDate).sort((a, b) => new Date(b) - new Date(a));

            if (sortedDates.length === 0) return;

            let tooltipContent = '<h4>Progreso por Fecha</h4><ul>';
            sortedDates.forEach(date => {
                const { inventory, additional } = progressByDate[date];
                const formattedDate = new Date(date + 'T00:00:00').toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit', year: '2-digit' });
                tooltipContent += `<li><strong>${formattedDate}:</strong> ${inventory} (Inv) + ${additional} (Adic)</li>`;
            });
            tooltipContent += '</ul>';
            
            tooltip.innerHTML = tooltipContent;
            tooltip.style.display = 'block';

            const rect = e.currentTarget.getBoundingClientRect();
            tooltip.style.left = `${rect.left}px`;
            tooltip.style.top = `${rect.bottom + 10}px`;
        });
        
        const hideTooltip = () => {
            tooltipTimeout = setTimeout(() => {
                tooltip.style.display = 'none';
                tooltip.style.pointerEvents = 'none';
            }, 300); 
        };

        card.addEventListener('mouseleave', hideTooltip);
        tooltip.addEventListener('mouseleave', hideTooltip);

        tooltip.addEventListener('mouseenter', () => {
            clearTimeout(tooltipTimeout);
        });
        
        window.addEventListener('beforeunload', (event) => {
            if (state.loggedIn && !state.readOnlyMode) {
                event.preventDefault();
                event.returnValue = '';
            }
        });

        interact('.layout-on-canvas')
            .draggable({
                listeners: {
                    move(event) {
                        if(state.readOnlyMode) return;
                        const target = event.target;
                        const x = (parseFloat(target.dataset.x) || 0) + event.dx;
                        const y = (parseFloat(target.dataset.y) || 0) + event.dy;
                        const rotation = (parseFloat(target.dataset.rotation) || 0);

                        target.style.transform = `translate(${x}px, ${y}px) rotate(${rotation}deg)`;
                        target.dataset.x = x;
                        target.dataset.y = y;
                    }
                },
                modifiers: [
                    interact.modifiers.snap({
                        targets: [ interact.snappers.grid({ x: 10, y: 10 }) ],
                        range: Infinity,
                        relativePoints: [ { x: 0, y: 0 } ]
                    })
                ],
                inertia: false
            })
            .resizable({
                edges: { left: true, right: true, bottom: true, top: true },
                listeners: {
                    move (event) {
                        if(state.readOnlyMode) return;
                        let target = event.target;
                        let x = (parseFloat(target.dataset.x) || 0);
                        let y = (parseFloat(target.dataset.y) || 0);
                        const rotation = (parseFloat(target.dataset.rotation) || 0);

                        target.style.width = event.rect.width + 'px';
                        target.style.height = event.rect.height + 'px';

                        x += event.deltaRect.left;
                        y += event.deltaRect.top;

                        target.style.transform = `translate(${x}px, ${y}px) rotate(${rotation}deg)`;
                        target.dataset.x = x;
                        target.dataset.y = y;
                    }
                },
                modifiers: [
                    interact.modifiers.snap({
                        targets: [ interact.snappers.grid({ x: 10, y: 10 }) ],
                        range: Infinity
                    }),
                    interact.modifiers.restrictSize({
                        min: { width: 50, height: 50 }
                    })
                ],
                inertia: false
            });

        interact('.draggable-item')
            .draggable({
                listeners: { move: window.dragMoveListener },
                inertia: true
            });

        interact('.draggable-tool')
            .draggable({
                listeners: { move: window.dragMoveListener },
                inertia: true
            });

        interact('#layout-canvas').dropzone({
            accept: '.draggable-item, .draggable-tool',
            ondrop: function(event) {
                if(state.readOnlyMode) return;
                const draggableElement = event.relatedTarget;
                
                const canvasWrapperRect = elements.layoutEditor.canvasWrapper.getBoundingClientRect();
                const itemRect = draggableElement.getBoundingClientRect();
                
                const x = (itemRect.left - canvasWrapperRect.left) + elements.layoutEditor.canvasWrapper.scrollLeft;
                const y = (itemRect.top - canvasWrapperRect.top) + elements.layoutEditor.canvasWrapper.scrollTop;
                const snappedX = Math.round(x / 10) * 10;
                const snappedY = Math.round(y / 10) * 10;
                
                if (draggableElement.classList.contains('draggable-item')) {
                    const locId = draggableElement.dataset.locationId;
                    const areaId = draggableElement.dataset.areaId; 
                    createShapeOnCanvas(locId, snappedX, snappedY, null, null, 'location', '', null, 0, areaId); 
                    draggableElement.classList.add('hidden'); 
                } 
                else if (draggableElement.classList.contains('draggable-tool')) {
                    const toolType = draggableElement.dataset.toolType;
                    const toolId = `${toolType}-${Date.now()}`; 
                    
                    if (toolType === 'note') {
                        createShapeOnCanvas(toolId, snappedX, snappedY, 200, 100, 'note');
                    } else if (toolType === 'arrow') { 
                        createShapeOnCanvas(toolId, snappedX, snappedY, 50, 50, 'tool');
                    } else if (toolType === 'text') { 
                        createShapeOnCanvas(toolId, snappedX, snappedY, 150, 40, 'text');
                    } 
                }
                
                draggableElement.style.transform = 'none';
                draggableElement.dataset.x = 0;
                draggableElement.dataset.y = 0;
                
                saveLayoutPositions();
            }
        });

        interact('.layout-rotate-handle').draggable({
            onmove: (event) => {
                if(state.readOnlyMode) return;
                
                const handle = event.target;
                const shape = handle.closest('.layout-on-canvas');
                if (!shape) return;

                const rect = shape.getBoundingClientRect();
                const canvasRect = elements.layoutEditor.canvasWrapper.getBoundingClientRect();
                
                const centerX = rect.left + (rect.width / 2) + window.scrollX;
                const centerY = rect.top + (rect.height / 2) + window.scrollY;

                const angle = Math.atan2(event.pageY - centerY, event.pageX - centerX) * (180 / Math.PI);
                
                let rotation = Math.round(angle + 90); 
                
                rotation = Math.round(rotation / 15) * 15;

                const x = parseFloat(shape.dataset.x) || 0;
                const y = parseFloat(shape.dataset.y) || 0;

                shape.style.transform = `translate(${x}px, ${y}px) rotate(${rotation}deg)`;
                shape.dataset.rotation = rotation;
            },
            onend: (event) => {
                if(state.readOnlyMode) return;
                saveLayoutPositions();
                saveState();
            }
        });
        // ============================================================
        // LÓGICA DE CONCILIACIÓN (COMPARADOR INTELIGENTE)
        // ============================================================
        
        const compareBtn = document.getElementById('compare-inventory-btn');
        const compareInput = document.getElementById('compare-file-input');
        const reconModal = document.getElementById('reconciliation-modal');
        const reconContent = document.getElementById('diff-content-container');
        let currentDiffData = { newItems: [], modItems: [], delItems: [] };

        if (compareBtn) {
            compareBtn.addEventListener('click', () => {
                if (state.readOnlyMode) return showToast('Modo lectura activo.', 'warning');
                compareInput.value = ''; // FIX: Resetear input
                compareInput.click();
            });
        }

   if (compareInput) {
            compareInput.addEventListener('change', (e) => {
                const file = e.target.files[0];
                if (!file) return;
                processComparisonFile(file);
                e.target.value = ''; // Reset
            });
        }

        function processComparisonFile(file) {
            elements.loadingOverlay.overlay.classList.add('show');
            elements.loadingOverlay.text.textContent = 'Analizando diferencias...';
            
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = e.target.result;
                    const workbook = XLSX.read(data, { type: 'binary' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    
                    // NUEVO: Capturar fecha del archivo de comparación
                    const comparisonDate = findReportDateSmart(sheet);
                    
                    // Actualizar título del modal de conciliación con la fecha
                    const reconTitle = document.querySelector('#reconciliation-modal h3');
                    if (reconTitle) {
                        reconTitle.innerHTML = `<i class="fa-solid fa-scale-balanced mr-2"></i>Conciliación de Inventario <span class="text-sm font-normal text-indigo-200 ml-2">(${comparisonDate})</span>`;
                    }

                    // Usamos la misma lógica de lectura que en la carga inicial
                    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 11 });
                    const claveUnicaRegex = /^(?:\d{5,6}|0\.\d+)$/;

                    const newInventoryList = rawData.map(row => {
                        const clave = String(row[0] || '').trim();
                        if (!claveUnicaRegex.test(clave)) return null;
                        return {
                            'CLAVE UNICA': clave, 'DESCRIPCION': String(row[1] || ''), 
                            'MARCA': row[4] || '', 'MODELO': row[5] || '', 'SERIE': row[6] || '',
                            'listadoOriginal': sheet['B7']?.v || 'Actualizado', // Dato extra
                            'areaOriginal': (sheet['A10']?.v || '').match(/AREA\s(\d+)/)?.[1] || 'Sin Área'
                        };
                    }).filter(Boolean);

                    runComparisonAlgorithm(newInventoryList);

                } catch (error) {
                    console.error(error);
                    showToast('Error al leer el archivo de comparación.', 'error');
                } finally {
                    elements.loadingOverlay.overlay.classList.remove('show');
                }
            };
            reader.readAsBinaryString(file);
        }

        function runComparisonAlgorithm(newList) {
            // 1. Indexar inventario ACTUAL
            const currentMap = new Map();
            state.inventory.forEach(item => currentMap.set(item['CLAVE UNICA'], item));

            const diff = { newItems: [], modItems: [], delItems: [] };
            const processedKeys = new Set();

            // 2. Recorrer NUEVA lista
            newList.forEach(newItem => {
                const clave = newItem['CLAVE UNICA'];
                processedKeys.add(clave);

                // CÓDIGO CORREGIDO (DENTRO DE runComparisonAlgorithm)

if (currentMap.has(clave)) {
    // Ya existe: Verificar cambios
    const currentItem = currentMap.get(clave);
    const modifications = [];

    // Helper para normalizar: quita acentos (NFD), espacios extra y pasa a mayúsculas
    const norm = (val) => String(val || '')
        .trim()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Quita acentos
        .toUpperCase()
        .replace(/\s+/g, ' ');

    // Campos a auditar (Comparación insensible a mayúsculas/espacios/acentos)
    if (norm(currentItem['DESCRIPCION']) !== norm(newItem['DESCRIPCION'])) {
        modifications.push({ field: 'DESCRIPCION', old: currentItem['DESCRIPCION'], new: newItem['DESCRIPCION'] });
    }
    if (norm(currentItem['MARCA']) !== norm(newItem['MARCA'])) {
        modifications.push({ field: 'MARCA', old: currentItem['MARCA'], new: newItem['MARCA'] });
    }
    if (norm(currentItem['MODELO']) !== norm(newItem['MODELO'])) {
        modifications.push({ field: 'MODELO', old: currentItem['MODELO'], new: newItem['MODELO'] });
    }
    if (norm(currentItem['SERIE']) !== norm(newItem['SERIE'])) {
        modifications.push({ field: 'SERIE', old: currentItem['SERIE'], new: newItem['SERIE'] });
    }

    if (modifications.length > 0) {
        diff.modItems.push({ clave, newItem, modifications });
    }
} else {
    // No existe: Es NUEVO
    diff.newItems.push(newItem);
}
            });

            // 3. Detectar ELIMINADOS (Estaban en currentMap pero no en processedKeys)
            currentMap.forEach((item, clave) => {
                if (!processedKeys.has(clave)) {
                    // OJO: Solo considerar eliminados si pertenecen al ÁREA del archivo que subimos
                    // (Para evitar borrar todo el inventario si subes solo una hoja de un área)
                    // Implementación segura: Si el nuevo archivo tiene área X, solo revisamos bajas de área X.
                    // Si el nuevo archivo es general, revisamos todo.
                    // Por simplicidad en V1: Marcamos como 'Posible Baja'.
                    diff.delItems.push(item);
                }
            });

            currentDiffData = diff;
            renderReconciliationUI(diff);
        }

        function renderReconciliationUI(diff) {
            document.getElementById('diff-count-new').textContent = diff.newItems.length;
            document.getElementById('diff-count-mod').textContent = diff.modItems.length;
            document.getElementById('diff-count-del').textContent = diff.delItems.length;

            // Render inicial (Nuevos)
            renderDiffTab('new');
            reconModal.classList.add('show');
        }

        function renderDiffTab(type) {
            reconContent.innerHTML = '';
            let items = [];
            let html = '';

            // Actualizar estilos de pestañas
            ['new', 'mod', 'del'].forEach(t => {
                const btn = document.getElementById(`tab-diff-${t}`);
                if (t === type) {
                    btn.className = 'px-4 py-2 font-bold text-indigo-600 border-b-2 border-indigo-600';
                } else {
                    btn.className = 'px-4 py-2 text-gray-500 hover:text-gray-700';
                }
            });

            if (type === 'new') {
                items = currentDiffData.newItems;
                if (items.length === 0) return reconContent.innerHTML = '<p class="text-center text-gray-500 p-4">No hay bienes nuevos.</p>';
                
                html = items.map(item => `
                    <div class="diff-card diff-new flex items-center">
                        <input type="checkbox" class="mr-3 w-5 h-5 rounded diff-check" checked data-type="new" data-clave="${item['CLAVE UNICA']}">
                        <div>
                            <p class="font-bold text-green-700">ALTA: ${item['CLAVE UNICA']}</p>
                            <p class="text-sm">${item['DESCRIPCION']}</p>
                            <p class="text-xs text-gray-500">Área: ${item.areaOriginal}</p>
                        </div>
                    </div>
                `).join('');
            } 
            else if (type === 'mod') {
                items = currentDiffData.modItems;
                if (items.length === 0) return reconContent.innerHTML = '<p class="text-center text-gray-500 p-4">No hay modificaciones detectadas.</p>';

                html = items.map(data => {
                    const modsHtml = data.modifications.map(m => 
                        `<div><span class="font-semibold text-xs">${m.field}:</span> <span class="diff-old-val">${m.old}</span> <i class="fa-solid fa-arrow-right text-gray-400 text-xs mx-1"></i> <span class="diff-new-val">${m.new}</span></div>`
                    ).join('');

                    return `
                    <div class="diff-card diff-mod flex items-center">
                        <input type="checkbox" class="mr-3 w-5 h-5 rounded diff-check" checked data-type="mod" data-clave="${data.clave}">
                        <div>
                            <p class="font-bold text-orange-700">CAMBIO: ${data.clave}</p>
                            <div class="pl-2 border-l-2 border-orange-200 mt-1">${modsHtml}</div>
                        </div>
                    </div>`;
                }).join('');
            }
            else if (type === 'del') {
                items = currentDiffData.delItems;
                if (items.length === 0) return reconContent.innerHTML = '<p class="text-center text-gray-500 p-4">No hay bajas detectadas.</p>';

                html = items.map(item => `
                    <div class="diff-card diff-del flex items-center">
                        <input type="checkbox" class="mr-3 w-5 h-5 rounded diff-check" data-type="del" data-clave="${item['CLAVE UNICA']}"> 
                        <div>
                            <p class="font-bold text-red-700">BAJA: ${item['CLAVE UNICA']}</p>
                            <p class="text-sm text-gray-600">${item['DESCRIPCION']}</p>
                            <p class="text-xs text-red-500 font-bold">Advertencia: Si confirmas, se eliminará el bien y su historial.</p>
                        </div>
                    </div>
                `).join('');
                // Nota: Checkbox de bajas DESMARCADO por defecto por seguridad
            }

            reconContent.innerHTML = html;
        }

        // Listeners de Pestañas
        document.getElementById('tab-diff-new')?.addEventListener('click', () => renderDiffTab('new'));
        document.getElementById('tab-diff-mod')?.addEventListener('click', () => renderDiffTab('mod'));
        document.getElementById('tab-diff-del')?.addEventListener('click', () => renderDiffTab('del'));

        // Botón de cierre y cancelar
        document.getElementById('reconciliation-close-btn')?.addEventListener('click', () => reconModal.classList.remove('show'));
        document.getElementById('reconciliation-cancel-btn')?.addEventListener('click', () => reconModal.classList.remove('show'));

        // BOTÓN MÁGICO: APLICAR CAMBIOS
        document.getElementById('reconciliation-apply-btn')?.addEventListener('click', () => {
            const checkboxes = document.querySelectorAll('.diff-check:checked');
            if (checkboxes.length === 0) return showToast('No has seleccionado ninguna acción.', 'warning');

            let processedCount = 0;
            const activeUser = state.currentUser ? state.currentUser.name : 'Sistema';

            checkboxes.forEach(cb => {
                const type = cb.dataset.type;
                const clave = cb.dataset.clave;

                if (type === 'new') {
                    // Insertar nuevo
                    const newItemData = currentDiffData.newItems.find(i => i['CLAVE UNICA'] === clave);
                    if (newItemData) {
                        // Inicializar campos app-specific
                        newItemData['UBICADO'] = 'NO';
                        newItemData['IMPRIMIR ETIQUETA'] = 'NO';
                        newItemData['NOMBRE DE USUARIO'] = '';
                        // Usamos un ID negativo temporal o timestamp para listId si no tenemos uno
                        newItemData.listId = Date.now(); 
                        state.inventory.push(newItemData);
                    }
                }
                else if (type === 'mod') {
                    // Actualizar existente (MERGE)
                    const modData = currentDiffData.modItems.find(i => i.clave === clave);
                    const targetIndex = state.inventory.findIndex(i => i['CLAVE UNICA'] === clave);
                    if (modData && targetIndex !== -1) {
                        // Solo sobrescribimos los campos que cambiaron en el excel
                        // PRESERVANDO: Fotos, Ubicación, Usuario, Notas
                        state.inventory[targetIndex]['DESCRIPCION'] = modData.newItem['DESCRIPCION'];
                        state.inventory[targetIndex]['MARCA'] = modData.newItem['MARCA'];
                        state.inventory[targetIndex]['MODELO'] = modData.newItem['MODELO'];
                        state.inventory[targetIndex]['SERIE'] = modData.newItem['SERIE'];
                    }
                }
                else if (type === 'del') {
                    // Eliminar
                    state.inventory = state.inventory.filter(i => i['CLAVE UNICA'] !== clave);
                    // Opcional: Limpiar fotos de eliminados
                    delete state.photos[clave];
                    photoDB.deleteItem('photos', `inventory-${clave}`);
                }
                processedCount++;
            });

            logActivity('Conciliación de Inventario', `Se aplicaron ${processedCount} cambios (Altas/Bajas/Mods).`);
            showToast(`Sincronización completada: ${processedCount} cambios aplicados.`, 'success');
            
            saveState();
            updateSerialNumberCache();
            populateAreaSelects();
            renderDashboard();
            currentPage = 1;
            filterAndRenderInventory();
            
            reconModal.classList.remove('show');
        });

        if (loadState()) {

            recalculateLocationCounts(); 

            if (state.loggedIn) {
                showMainApp();
            } else {
                elements.loginPage.classList.remove('hidden');
                elements.mainApp.classList.add('hidden');
            }
        } else {
            elements.loginPage.classList.remove('hidden');
            elements.mainApp.classList.add('hidden');
        }
    }
    // --- LÓGICA DE IMPRESIÓN MASIVA (BATCH) ---
    
    function openBatchPrintModal(areaId) {
        if (!areaId || areaId === 'all') return showToast('Error: Área no válida.', 'error');

        const usersInArea = state.resguardantes.filter(u => u.area === areaId);
        if (usersInArea.length === 0) return showToast('No hay usuarios registrados en esta área.', 'warning');

        const areaName = state.areaNames[areaId] || `Área ${areaId}`;
        const areaResponsable = state.areaDirectory[areaId];

        // 1. Configurar valores iniciales
        elements.batchModal.areaNameDisplay.textContent = areaName;
        elements.batchModal.dateInput.value = getLocalDate();
        elements.batchModal.entregaInput.value = areaResponsable ? areaResponsable.name : '_________________________';
        elements.batchModal.cargoInput.value = areaResponsable ? areaResponsable.title : 'Responsable de Área'; // Default
        elements.batchModal.includeAdditionals.checked = false;

        // 2. Renderizar lista de usuarios
        const container = elements.batchModal.usersListContainer;
        container.innerHTML = '';
        
        usersInArea.forEach(user => {
            const itemDiv = document.createElement('div');
            itemDiv.className = 'flex items-center p-3 border-b border-gray-100 dark:border-slate-800 hover:bg-indigo-50 dark:hover:bg-slate-800 transition-colors';
            
            // Contar bienes para mostrar info útil
            const inventoryCount = state.inventory.filter(i => i['NOMBRE DE USUARIO'] === user.name).length;
            const additionalCount = state.additionalItems.filter(i => i.usuario === user.name).length;
            
            const hasItems = inventoryCount > 0 || additionalCount > 0;
            const styleClass = hasItems ? 'text-gray-800 dark:text-gray-200' : 'text-gray-400 dark:text-gray-600 italic';
            
            itemDiv.innerHTML = `
                <label class="flex items-center w-full cursor-pointer">
                    <input type="checkbox" class="batch-user-checkbox w-5 h-5 text-indigo-600 rounded focus:ring-indigo-500 mr-3" 
                           value="${user.name}" ${hasItems ? 'checked' : ''}>
                    <div class="flex-grow">
                        <p class="font-medium ${styleClass}">${user.name}</p>
                        <p class="text-xs text-gray-500">
                            Inventario: <b>${inventoryCount}</b> | Adicionales: <b>${additionalCount}</b>
                        </p>
                    </div>
                    ${user.name === (areaResponsable?.name) ? '<span class="text-xs bg-yellow-100 text-yellow-800 px-2 py-1 rounded-full font-bold">Responsable</span>' : ''}
                </label>
            `;
            container.appendChild(itemDiv);
        });

        updateBatchCount();
        elements.batchModal.modal.classList.add('show');
        handleModalNavigation(elements.batchModal.modal);
    }

    function updateBatchCount() {
        const checked = document.querySelectorAll('.batch-user-checkbox:checked').length;
        elements.batchModal.countDisplay.textContent = checked;
        elements.batchModal.generateBtn.disabled = checked === 0;
        if(checked === 0) {
             elements.batchModal.generateBtn.classList.add('opacity-50', 'cursor-not-allowed');
        } else {
             elements.batchModal.generateBtn.classList.remove('opacity-50', 'cursor-not-allowed');
        }
    }

    async function generateBatchReport() {
        const selectedCheckboxes = Array.from(document.querySelectorAll('.batch-user-checkbox:checked'));
        if (selectedCheckboxes.length === 0) return;

        const globalDate = elements.batchModal.dateInput.value;
        const globalEntrega = elements.batchModal.entregaInput.value;
        const globalCargoEntrega = elements.batchModal.cargoInput.value;
        const includeAdditionals = elements.batchModal.includeAdditionals.checked;
        const areaId = elements.reports.areaFilter.value;
        const areaFullName = state.areaNames[areaId] || `Área ${areaId}`;
        const areaResponsableName = state.areaDirectory[areaId]?.name; // Para detectar quién es el jefe

        logActivity('Impresión Masiva', `Generando ${selectedCheckboxes.length} resguardos para el área ${areaId}.`);
        showToast('Generando documento masivo, por favor espera...', 'info');

        // Preparar contenedor de impresión
        const printContainer = document.getElementById('print-view-container');
        // Limpiar clones anteriores si existen
        document.querySelectorAll('.print-page.batch-clone').forEach(el => el.remove());
        // Desactivar páginas activas previas
        document.querySelectorAll('.print-page').forEach(page => page.classList.remove('active'));

        const masterTemplate = elements.printTemplates.resguardo;
        
        // --- BUCLE MAESTRO DE GENERACIÓN ---
        for (let i = 0; i < selectedCheckboxes.length; i++) {
            const userName = selectedCheckboxes[i].value;
            
            // 1. Obtener bienes
            let items = state.inventory.filter(item => item['NOMBRE DE USUARIO'] === userName);
            if (includeAdditionals) {
                const additionals = state.additionalItems.filter(item => item.usuario === userName);
                items = [...items, ...additionals];
            }

            if (items.length === 0) continue; // Saltar usuarios vacíos si se seleccionaron por error

            // 2. Clonar plantilla
            const pageClone = masterTemplate.cloneNode(true);
            pageClone.id = `batch-page-${i}`;
            pageClone.classList.add('batch-clone', 'active', 'batch-mode'); // batch-mode para CSS específico
            
            // 3. Lógica de Responsable (Tu requerimiento clave)
            const isUserResponsable = (userName === areaResponsableName);
            const signaturesContainer = pageClone.querySelector('.print-signatures');
            const responsibleTitleEl = pageClone.querySelector('#print-resguardo-responsible-title');
            const introTextEl = pageClone.querySelector('#print-resguardo-text');

            if (isUserResponsable) {
                signaturesContainer.classList.add('center-single');
                responsibleTitleEl.textContent = 'Responsable de Área';
                introTextEl.innerHTML = `Quedo enterado, <strong>${userName}</strong> que los Bienes Muebles que se encuentran listados en el presente resguardo, están a partir de la firma del mismo, bajo mi buen uso, custodia, vigilancia y conservación, en caso de daño, robo o extravío, se deberá notificar de inmediato a el Área Administrativa o Comisión para realizar el trámite administrativo correspondiente, por ningún motivo se podrá cambiar o intercambiar los bienes sin previa solicitud y autorización del Área Administrativa o Comisión.`;
            } else {
                signaturesContainer.classList.remove('center-single');
                responsibleTitleEl.textContent = 'Usuario Resguardante';
                introTextEl.innerHTML = `Quedo enterado, <strong>${userName}</strong> que los Bienes Muebles que se encuentran listados en el presente resguardo, están a partir de la firma del mismo, bajo mi buen uso, custodia, vigilancia y conservación, en caso de daño, robo o extravio, deberé notificar al jefe inmediato del Área Administrativa o Comisión para realizar el trámite administrativo correspondiente. Por ningún motivo se podra cambiar o intercambiar los bienes sin previa solicitud y autorización del jefe inmediato y/o responsable del inventario.`;
            }

            // 4. Llenar Datos
            pageClone.querySelector('#print-resguardo-title').textContent = 'Resguardo Individual de Bienes';
            pageClone.querySelector('#print-resguardo-area').textContent = areaFullName;
            pageClone.querySelector('.print-header-date').textContent = `Fecha: ${globalDate}`;
            
            pageClone.querySelector('#print-resguardo-author-name').textContent = globalEntrega;
            pageClone.querySelector('#print-resguardo-author-title').textContent = globalCargoEntrega;
            
            // Recibe: Siempre es el usuario dueño del resguardo
            pageClone.querySelector('#print-resguardo-responsible-name').textContent = userName;

            // 5. Llenar Tabla (Reutilizando lógica simple)
            const tbody = pageClone.querySelector('tbody');
            tbody.innerHTML = items.map(item => {
                const isAd = !!item.id; // Es adicional si tiene ID
                let type = 'Institucional';
                if (isAd) {
                     if (item.personal === 'Si') type = 'Personal';
                     else if ((item.area || '').includes('CONTRATO')) type = 'Arrendamiento';
                     else type = 'Cámara/Controlable';
                }
                
                const clave = isAd ? (item.claveAsignada || item.clave || 'S/C') : item['CLAVE UNICA'];
                const asterisk = (isAd && item.personal === 'Si' && item.tieneFormatoEntrada === false) ? ' <strong>*</strong>' : '';

                // CÓDIGO CORREGIDO (SEGURO)
                // Aseguramos que los textos no rompan el HTML
            const safeDesc = escapeHTML(item.descripcion || item.DESCRIPCION || '');
            const safeMarca = escapeHTML(item.marca || item.MARCA || '');
            const safeModelo = escapeHTML(item.modelo || item.MODELO || '');
            const safeSerie = escapeHTML(item.serie || item.SERIE || '');
            const safeArea = escapeHTML(isAd ? (item.area || 'N/A') : item.areaOriginal);
            const safeUser = escapeHTML(userName);

            return `<tr>
                <td class="col-num"></td>
                <td class="col-clave">${escapeHTML(clave)}</td>
                <td class="col-desc">${safeDesc}${asterisk}</td>
                <td class="col-marca">${safeMarca}</td>
                <td class="col-modelo">${safeModelo}</td>
                <td class="col-serie">${safeSerie}</td>
                <td class="col-area">${safeArea}</td>
                <td class="col-usuario">${safeUser}</td>
                <td class="col-status">${type}</td>
            </tr>`;
            }).join('');

            pageClone.querySelector('#print-resguardo-count').textContent = `Total de Bienes: ${items.length}`;
            
            // Nota al pie
            const showNote = items.some(i => !!i.id && i.personal === 'Si' && i.tieneFormatoEntrada === false);
            pageClone.querySelector('#print-resguardo-note').innerHTML = showNote ? '<strong>* Favor de realizar entrada</strong>' : '';

            if (i < selectedCheckboxes.length - 1) {
                pageClone.classList.add('batch-page-break-after'); 
            }
            
            printContainer.appendChild(pageClone);
        }
        // CÓDIGO CORREGIDO
        elements.batchModal.modal.classList.remove('show');

        // Esperar al siguiente ciclo de pintado del navegador (doble frame para asegurar)
        requestAnimationFrame(() => {
            requestAnimationFrame(() => {
                // Pequeño timeout de seguridad adicional variable según cantidad de hojas
                const dynamicDelay = Math.min(selectedCheckboxes.length * 50, 2000); // Mínimo 50ms por hoja, máx 2 seg
                setTimeout(() => window.print(), dynamicDelay); 
            });
        });
    }        

    initialize();
});