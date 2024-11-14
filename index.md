---
layout: default
title: "Bon de Commande MSV"
---

<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Bon de Commande MSV</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.7.2/font/bootstrap-icons.css" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        .status-draft {
            background-color: #fff3cd !important;
            border-left: 4px solid #ffc107;
        }
        .status-confirmed {
            background-color: #d4edda !important;
            border-left: 4px solid #28a745;
        }
        .card.status-draft .card-header {
            background-color: #fff3cd;
            color: #856404;
        }
        .card.status-confirmed .card-header {
            background-color: #d4edda;
            color: #155724;
        }
        .search-results {
            position: absolute;
            top: 100%;
            left: 0;
            right: 0;
            background: white;
            border: 1px solid #ddd;
            border-radius: 4px;
            z-index: 1000;
            max-height: 400px;
            overflow-y: auto;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .search-item {
            padding: 10px;
            cursor: pointer;
            border-bottom: 1px solid #f0f0f0;
        }
        .search-item:hover {
            background-color: #f8f9fa;
        }
        .search-item:last-child {
            border-bottom: none;
        }
        .item-row {
            transition: all 0.2s ease;
            padding: 12px;
            border-radius: 6px;
            background-color: white;
            margin-bottom: 8px;
        }
        .item-row:hover {
            background-color: #f8f9fa;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }
        .item-row.dragging {
            opacity: 0.5;
            background: #e9ecef;
            cursor: move;
        }
        .drag-handle {
            cursor: move;
            color: #adb5bd;
            padding: 8px;
            display: inline-block;
        }
        .drag-handle:hover {
            color: #6c757d;
        }
        .stock-status {
            display: block;
            margin-top: 4px;
            font-size: 0.85em;
        }
        .notifications {
            position: fixed;
            bottom: 20px;
            right: 20px;
            z-index: 1050;
        }
        .notification {
            padding: 12px 24px;
            margin-bottom: 10px;
            border-radius: 6px;
            color: white;
            animation: slideIn 0.3s ease-out;
            box-shadow: 0 2px 8px rgba(0,0,0,0.2);
            min-width: 250px;
        }
        .notification.success { background-color: rgba(40, 167, 69, 0.95); }
        .notification.error { background-color: rgba(220, 53, 69, 0.95); }
        .notification.info { background-color: rgba(23, 162, 184, 0.95); }

        @keyframes slideIn {
            from { transform: translateX(100%); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }

        @media (max-width: 768px) {
            .item-row .row { margin-bottom: 8px; }
            .notifications {
                left: 20px;
                right: 20px;
            }
            .notification { width: auto; }
        }

        .item-row.light-green { background-color: rgba(40, 167, 69, 0.1) !important; }
        .item-row.light-orange { background-color: rgba(255, 193, 7, 0.1) !important; }
        .item-row.light-red { background-color: rgba(220, 53, 69, 0.1) !important; }
        .item-row.light-blue { background-color: rgba(23, 162, 184, 0.1) !important; }

        .item-row.light-green:hover { background-color: rgba(40, 167, 69, 0.2) !important; }
        .item-row.light-orange:hover { background-color: rgba(255, 193, 7, 0.2) !important; }
        .item-row.light-red:hover { background-color: rgba(220, 53, 69, 0.2) !important; }
        .item-row.light-blue:hover { background-color: rgba(23, 162, 184, 0.2) !important; }
    </style>
</head>
<body>
    <div class="container py-4">
        <h2 class="mb-4">Bon de Commande MSV</h2>

        <!-- Import du Stock -->
        <div class="card mb-4">
            <div class="card-header d-flex justify-content-between align-items-center">
                <h5 class="mb-0">Import du Stock</h5>
                <button class="btn btn-sm btn-outline-secondary" id="clearStock">
                    <i class="bi bi-trash"></i> Effacer le stock
                </button>
            </div>
            <div class="card-body">
                <div class="row mb-3">
                    <div class="col">
                        <input type="file" id="stockFile" class="form-control" accept=".csv,.xlsx,.xls">
                        <small class="text-muted">
                            Formats acceptés: CSV (séparateur point-virgule), Excel (.xlsx, .xls)
                            <br>Colonnes requises: ID, sku, item_name, Qty, Price
                        </small>
                    </div>
                </div>
                <div id="stockPreview" class="table-responsive" style="max-height: 400px; overflow-y: auto;">
                    <table class="table table-sm table-hover">
                        <thead>
                            <tr>
                                <th>SKU</th>
                                <th>Nom</th>
                                <th class="text-end">Stock</th>
                                <th class="text-end">Prix</th>
                            </tr>
                        </thead>
                        <tbody id="stockPreviewBody">
                            <tr>
                                <td colspan="4" class="text-center text-muted">
                                    Aucun stock importé
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- Détails de la commande -->
        <div class="card mb-4" id="orderCard">
            <div class="card-header d-flex justify-content-between align-items-center">
                <div class="d-flex align-items-center gap-3">
                    <h5 class="mb-0">Détails de la commande</h5>
                    <select id="orderStatus" class="form-select form-select-sm" style="width: 150px;">
                        <option value="draft" class="bg-warning text-dark">Brouillon</option>
                        <option value="confirmed" class="bg-success text-white">Confirmé</option>
                    </select>
                </div>
                <div class="btn-group">
                    <button class="btn btn-outline-primary btn-sm dropdown-toggle" data-bs-toggle="dropdown">
                        <i class="bi bi-save"></i> Sauvegarder
                    </button>
                    <ul class="dropdown-menu">
                        <li>
                            <a class="dropdown-item" href="#" id="saveDraftLocal">
                                <i class="bi bi-bookmark"></i> Sauvegarder en local
                            </a>
                        </li>
                        <li>
                            <a class="dropdown-item" href="#" id="saveDraftFile">
                                <i class="bi bi-download"></i> Télécharger le brouillon
                            </a>
                        </li>
                    </ul>
                    <button class="btn btn-outline-secondary btn-sm" onclick="document.getElementById('loadDraftFile').click()">
                        <i class="bi bi-folder-open"></i> Charger
                    </button>
                    <input type="file" id="loadDraftFile" accept=".json" class="d-none">
                </div>
            </div>
            <div class="card-body">
                <!-- En-tête de commande -->
                <div class="row mb-3">
                    <div class="col-md-4">
                        <label class="form-label">Nom du client</label>
                        <input type="text" id="supplier" class="form-control" required>
                    </div>
                    <div class="col-md-4">
                        <label class="form-label">WorkOrder</label>
                        <input type="text" id="orderNumber" class="form-control" required>
                    </div>
                    <div class="col-md-4">
                        <label class="form-label">Date</label>
                        <input type="date" id="orderDate" class="form-control" required>
                    </div>
                </div>

                <!-- Boutons d'action -->
                <div class="mb-3">
                    <button id="showImport" class="btn btn-secondary" disabled>
                        <i class="bi bi-upload"></i> Import Multiple
                    </button>
                    <button id="clearAll" class="btn btn-outline-danger">
                        <i class="bi bi-trash"></i> Tout effacer
                    </button>
                </div>

                <!-- Section Import Multiple -->
                <div id="importSection" class="card mb-3 d-none">
                    <div class="card-body">
                        <h6>Import en masse</h6>
                        <div class="mb-3">
                            <input type="file" id="bulkFileInput" class="form-control" accept=".csv,.xlsx,.xls">
                            <small class="text-muted">Format: fichier avec colonnes SKU et Quantité</small>
                        </div>
                        <textarea id="bulkImport" class="form-control mb-2" rows="4" 
                            placeholder="SKU, Quantité&#10;Un article par ligne"></textarea>
                        <div class="form-check mb-2">
                            <input class="form-check-input" type="checkbox" id="addK2Prefix">
                            <label class="form-check-label" for="addK2Prefix">
                                Ajouter le préfixe "K2-" aux SKUs
                            </label>
                        </div>
                        <button id="processImport" class="btn btn-primary">
                            <i class="bi bi-check-lg"></i> Importer
                        </button>
                    </div>
                </div>

                <!-- Liste des articles -->
                <div id="itemsList" class="mb-3"></div>

                <button id="addItem" class="btn btn-primary" disabled>
                    <i class="bi bi-plus-lg"></i> Ajouter un article
                </button>
            </div>
        </div>

        <!-- Barre d'actions -->
        <div class="d-flex justify-content-between align-items-center">
            <h4 id="totalAmount">Total: 0,00 €</h4>
            <div class="btn-group">
                <button id="exportXls" class="btn btn-success">
                    <i class="bi bi-file-excel"></i> Export XLS
                </button>
                <button id="exportCsv" class="btn btn-success">
                    <i class="bi bi-file-text"></i> Export CSV
                </button>
            </div>
        </div>
    </div>

    <!-- Zone de notifications -->
    <div id="notifications" class="notifications"></div>

    <script>
// Variables globales
let stockItems = [];
let items = [];
let lastSavedState = null;

// Colonnes pour l'export
const EXPORT_COLUMNS = [
    'Order Date',
    'CF.Date de livraison souhaitée',
    'PurchaseOrder',
    'SKU',
    'Item Name',
    'QuantityOrdered',
    'Delivery Method',
    'Shipping Address',
    'Shipping Street2',
    'Shipping City',
    'Shipping State',
    'Shipping Country',
    'Shipping Code',
    'Shipping Phone',
    'CF.REMARQUES POUR PREPARATION',
    'SalesOrder Number',
    'Customer Name',
    'Currency Code',
    'Warehouse Name',
    'Item Price',
    'Item Tax',
    'Item Tax %',
    'Sales Person'
];

// Valeurs fixes pour l'export
const FIXED_VALUES = {
    'SalesOrder Number': 'DVSAS-4730',
    'Customer Name': 'MAISON SOLAIRE VOLTALIA',
    'Currency Code': 'EUR',
    'Warehouse Name': 'Distribution Voltalia',
    'Item Price': '1',
    'Item Tax': 'V.A.T',
    'Item Tax %': '20.000000',
    'Sales Person': 'ECHARIF Younes'
};

// Fonctions utilitaires
function formatPrice(price) {
    return new Intl.NumberFormat('fr-FR', {
        style: 'currency',
        currency: 'EUR',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(price);
}

function parsePrice(priceStr) {
    if (typeof priceStr === 'number') return priceStr;
    if (!priceStr) return 0;
    
    let cleanedStr = priceStr
        .toString()
        .replace('€', '')
        .replace(/\s/g, '')
        .trim();
        
    if (cleanedStr.includes(',')) {
        cleanedStr = cleanedStr.replace(',', '.');
    }

    const parsed = parseFloat(cleanedStr);
    return isNaN(parsed) ? 0 : parsed;
}

function showNotification(message, type = 'success') {
    const notif = document.createElement('div');
    notif.className = `notification ${type}`;
    notif.innerHTML = `
        <i class="bi ${type === 'success' ? 'bi-check-circle' : 
                      type === 'error' ? 'bi-exclamation-circle' : 
                      'bi-info-circle'}"></i>
        <span class="ms-2">${message}</span>
    `;
    document.getElementById('notifications').appendChild(notif);
    setTimeout(() => {
        notif.style.opacity = '0';
        setTimeout(() => notif.remove(), 300);
    }, 3000);
}


function updateTotal() {
    const total = items.reduce((sum, item) => sum + (item.quantity * item.price), 0);
    document.getElementById('totalAmount').textContent = `Total: ${formatPrice(total)}`;
    document.getElementById('totalAmount').title = `${items.length} article(s)`;
}
function getRowColor(colorName) {
    switch (colorName) {
        case 'light-green':
            return 'rgba(40, 167, 69, 0.1)';
        case 'light-orange':
            return 'rgba(255, 193, 7, 0.1)';
        case 'light-red':
            return 'rgba(220, 53, 69, 0.1)';
        case 'light-blue':
            return 'rgba(23, 162, 184, 0.1)';
        default:
            return '';
    }
}

// Import du stock
async function importStock(file) {
    try {
        const content = await readFileContent(file);
        let result;

        if (file.name.endsWith('.csv')) {
            result = parseCSVStock(content);
        } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            result = parseExcelStock(content);
        } else {
            throw new Error('Format de fichier non supporté');
        }

        if (!result.success) {
            throw new Error(result.error);
        }

        stockItems = result.data;
        localStorage.setItem('stockItems', JSON.stringify(stockItems));
        updateStockPreview(stockItems);
        document.getElementById('addItem').disabled = false;
        document.getElementById('showImport').disabled = false;
        showNotification(`${stockItems.length} articles importés avec succès`);

    } catch (error) {
        console.error('Erreur import:', error);
        showNotification(error.message, 'error');
    }
}

function readFileContent(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(new Error('Erreur de lecture du fichier'));
        
        if (file.name.endsWith('.csv')) {
            reader.readAsText(file, 'UTF-8');
        } else {
            reader.readAsBinaryString(file);
        }
    });
}

function parseCSVStock(content) {
    try {
        const lines = content.split(/\r?\n/).filter(line => line.trim());
        const headers = lines[0].split(';').map(h => h.trim());
        
        // Mapping plus flexible des colonnes
        const columnMapping = {
            id: ['ID', 'id', 'reference', 'ref'],
            sku: ['sku', 'SKU', 'code', 'article_code', 'code_article'],
            item_name: ['item_name', 'name', 'nom', 'designation', 'article', 'Item Name'],
            qty: ['qty', 'Qty', 'quantity', 'stock', 'Stock'],
            price: ['price', 'Price', 'prix', 'unit_price', 'prix_unitaire']
        };

        const colIndexes = {};
        for (const [key, possibleNames] of Object.entries(columnMapping)) {
            const foundIndex = possibleNames.findIndex(name => headers.includes(name));
            if (foundIndex === -1) {
                throw new Error(`Colonne manquante pour ${key}. Noms possibles: ${possibleNames.join(', ')}`);
            }
            colIndexes[key] = headers.indexOf(possibleNames[foundIndex]);
        }

        const data = lines.slice(1)
            .filter(line => line.trim())
            .map((line, index) => {
                const values = line.split(';');
                if (values.length < headers.length) {
                    console.warn(`Ligne ${index + 2} ignorée: données incomplètes`);
                    return null;
                }

                return {
                    id: values[colIndexes.id].trim(),
                    sku: values[colIndexes.sku].trim(),
                    item_name: values[colIndexes.item_name].trim(),
                    qty: parseInt(values[colIndexes.qty]) || 0,
                    price: parsePrice(values[colIndexes.price])
                };
            })
            .filter(item => item !== null && item.sku && item.item_name);

        return { success: true, data };
    } catch (error) {
        return { success: false, error: `Erreur format CSV: ${error.message}` };
    }
}

function parseExcelStock(content) {
    try {
        const workbook = XLSX.read(content, { type: 'binary' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        if (jsonData.length < 2) {
            throw new Error('Fichier vide ou invalide');
        }

        const headers = jsonData[0].map(h => h?.toString().trim());
        
        // Même mapping que pour CSV
        const columnMapping = {
            id: ['ID', 'id', 'reference', 'ref'],
            sku: ['sku', 'SKU', 'code', 'article_code', 'code_article'],
            item_name: ['item_name', 'name', 'nom', 'designation', 'article', 'Item Name'],
            qty: ['qty', 'Qty', 'quantity', 'stock', 'Stock'],
            price: ['price', 'Price', 'prix', 'unit_price', 'prix_unitaire']
        };

        const colIndexes = {};
        for (const [key, possibleNames] of Object.entries(columnMapping)) {
            const foundIndex = possibleNames.findIndex(name => headers.includes(name));
            if (foundIndex === -1) {
                throw new Error(`Colonne manquante pour ${key}. Noms possibles: ${possibleNames.join(', ')}`);
            }
            colIndexes[key] = headers.indexOf(possibleNames[foundIndex]);
        }

        const data = jsonData.slice(1)
            .filter(row => row.length >= headers.length)
            .map((row, index) => {
                try {
                    let priceStr = row[colIndexes.price]?.toString() || '0';
                    if (typeof row[colIndexes.price] === 'number') {
                        priceStr = row[colIndexes.price].toString();
                    }

                    return {
                        id: row[colIndexes.id]?.toString().trim(),
                        sku: row[colIndexes.sku]?.toString().trim(),
                        item_name: row[colIndexes.item_name]?.toString().trim(),
                        qty: parseInt(row[colIndexes.qty]) || 0,
                        price: parsePrice(priceStr)
                    };
                } catch (e) {
                    console.warn(`Ligne ${index + 2} ignorée:`, e);
                    return null;
                }
            })
            .filter(item => item !== null && item.sku && item.item_name);

        return { success: true, data };
    } catch (error) {
        return { success: false, error: `Erreur format Excel: ${error.message}` };
    }
}

function updateStockPreview(items) {
    const tbody = document.getElementById('stockPreviewBody');
    
    if (!items.length) {
        tbody.innerHTML = `
            <tr>
                <td colspan="4" class="text-center text-muted">
                    Aucun stock importé
                </td>
            </tr>`;
        return;
    }

    tbody.innerHTML = items.map(item => `
        <tr>
            <td>${item.sku}</td>
            <td>${item.item_name}</td>
            <td class="text-end ${item.qty === 0 ? 'text-danger' : 'text-success'}">${item.qty}</td>
            <td class="text-end">${formatPrice(item.price)}</td>
        </tr>
    `).join('');
}
// Gestion des articles
function createItemRow(itemData = {}) {
    if (document.getElementById('orderStatus').value === 'confirmed') {
        showNotification('Impossible de modifier une commande confirmée', 'error');
        return null;
    }

    const row = document.createElement('div');
    row.className = 'item-row mb-2 p-2 border rounded';
    row.draggable = true;

    // Notez la nouvelle colonne pour la couleur ajoutée après le drag-handle
    row.innerHTML = `
        <div class="row g-2 align-items-center">
            <div class="col-auto">
                <i class="bi bi-grip-vertical drag-handle"></i>
            </div>
            <div class="col-auto">
                <select class="form-select form-select-sm row-color" style="width: 100px;">
                    <option value="none">Couleur</option>
                    <option value="light-green" ${itemData.color === 'light-green' ? 'selected' : ''}>🟢 Module et micro</option>
                    <option value="light-orange" ${itemData.color === 'light-orange' ? 'selected' : ''}>🟡 Partie éléctrique</option>
                    <option value="light-red" ${itemData.color === 'light-red' ? 'selected' : ''}>🔴 Sy Fixation</option>
                    <option value="light-blue" ${itemData.color === 'light-blue' ? 'selected' : ''}>🔵 Partie finale</option>
                </select>
            </div>
            <div class="col-md-3">
                <div class="position-relative">
                    <input type="text" class="form-control sku-input" 
                           placeholder="SKU" value="${itemData.sku || ''}"
                           autocomplete="off">
                    <div class="search-results d-none"></div>
                </div>
            </div>
            <div class="col-md-3">
                <input type="text" class="form-control item-name" 
                       placeholder="Nom de l'article" value="${itemData.item_name || ''}" readonly>
            </div>
            <div class="col-md-1">
                <input type="number" class="form-control quantity-input" 
                       min="1" value="${itemData.quantity || 1}">
            </div>
            <div class="col-md-2">
                <input type="number" class="form-control price-input text-end" step="0.01"
                       value="${itemData.price || 0}" readonly>
            </div>
            <div class="col-md-1 d-flex gap-1">
                <button class="btn btn-outline-secondary btn-sm duplicate-item" title="Dupliquer">
                    <i class="bi bi-files"></i>
                </button>
                <button class="btn btn-outline-danger btn-sm remove-item" title="Supprimer">
                    <i class="bi bi-trash"></i>
                </button>
            </div>
            <div class="col-12">
                <small class="stock-status ${itemData.qty > 0 ? 'text-success' : 'text-danger'}">
                    Stock disponible: ${itemData.qty || 0}
                </small>
            </div>
        </div>
    `;

    setupItemRowEvents(row, itemData);
    setupDragAndDrop(row);

    // Gestion des couleurs
    if (itemData.color) {
        row.classList.add(itemData.color);
    }

    const colorSelect = row.querySelector('.row-color');
    colorSelect.addEventListener('change', (e) => {
        const color = e.target.value;
        
        // Supprimer toutes les classes de couleur existantes
        row.classList.remove('light-green', 'light-orange', 'light-red', 'light-blue');
        
        // Ajouter la nouvelle classe si une couleur est sélectionnée
        if (color !== 'none') {
            row.classList.add(color);
        }
        
        // Mise à jour des données
        const index = Array.from(row.parentNode.children).indexOf(row);
        items[index].color = color === 'none' ? null : color;
        saveDraftLocal();
    });

    return row;
}

function setupItemRowEvents(row, itemData) {
    const skuInput = row.querySelector('.sku-input');
    const searchResults = row.querySelector('.search-results');
    let currentStockItem = itemData.sku ? stockItems.find(item => item.sku === itemData.sku) : null;

    skuInput.addEventListener('input', (e) => {
        const search = e.target.value.toLowerCase();
        if (search.length < 2) {
            searchResults.classList.add('d-none');
            return;
        }

        const matches = stockItems.filter(item => 
            item.sku.toLowerCase().includes(search) || 
            item.item_name.toLowerCase().includes(search)
        ).slice(0, 15);

        if (matches.length > 0) {
            searchResults.innerHTML = `
                <div class="search-results-container" style="max-height: 400px; overflow-y: auto;">
                    ${matches.map(item => `
                        <div class="search-item p-2 border-bottom" data-sku="${item.sku}">
                            <div class="d-flex justify-content-between align-items-center">
                                <div>
                                    <div class="fw-bold">${item.sku}</div>
                                    <div class="small text-muted">${item.item_name}</div>
                                </div>
                                <span class="badge ${item.qty > 0 ? 'bg-success' : 'bg-danger'}">
                                    Stock: ${item.qty}
                                </span>
                            </div>
                        </div>
                    `).join('')}
                </div>`;
            searchResults.classList.remove('d-none');
        } else {
            searchResults.innerHTML = '<div class="p-2 text-muted">Aucun résultat</div>';
            searchResults.classList.remove('d-none');
        }
    });

    searchResults.addEventListener('click', (e) => {
        const searchItem = e.target.closest('.search-item');
        if (!searchItem) return;

        const sku = searchItem.dataset.sku;
        currentStockItem = stockItems.find(item => item.sku === sku);
        
        if (currentStockItem) {
            skuInput.value = currentStockItem.sku;
            row.querySelector('.item-name').value = currentStockItem.item_name;
            row.querySelector('.price-input').value = currentStockItem.price;
            row.querySelector('.stock-status').className = 
                `stock-status ${currentStockItem.qty > 0 ? 'text-success' : 'text-danger'}`;
            row.querySelector('.stock-status').textContent = `Stock disponible: ${currentStockItem.qty}`;
            
            const index = Array.from(row.parentNode.children).indexOf(row);
            items[index] = {
                ...currentStockItem,
                quantity: parseInt(row.querySelector('.quantity-input').value)
            };
            
            updateTotal();
            saveDraftLocal();
        }
        
        searchResults.classList.add('d-none');
    });

    const quantityInput = row.querySelector('.quantity-input');
    quantityInput.addEventListener('change', (e) => {
        if (!currentStockItem) return;
        
        const quantity = parseInt(e.target.value);
        if (quantity > currentStockItem.qty) {
            showNotification(`Stock insuffisant. Maximum: ${currentStockItem.qty}`, 'error');
            e.target.value = currentStockItem.qty;
        } else if (quantity < 1) {
            e.target.value = 1;
        }
        
        const index = Array.from(row.parentNode.children).indexOf(row);
        items[index] = {
            ...currentStockItem,
            quantity: parseInt(e.target.value)
        };
        
        updateTotal();
        saveDraftLocal();
    });

    row.querySelector('.remove-item').addEventListener('click', () => {
        const index = Array.from(row.parentNode.children).indexOf(row);
        items.splice(index, 1);
        row.remove();
        updateTotal();
        saveDraftLocal();
    });

    // Fermer la recherche quand on clique ailleurs
    document.addEventListener('click', (e) => {
        if (!row.contains(e.target)) {
            searchResults.classList.add('d-none');
        }
    });
}

function setupDragAndDrop(row) {
    row.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('text/plain', '');
        row.classList.add('dragging');
    });

    row.addEventListener('dragend', () => {
        row.classList.remove('dragging');
    });

    row.addEventListener('dragover', (e) => {
        e.preventDefault();
        const dragging = document.querySelector('.dragging');
        if (dragging && dragging !== row) {
            const container = row.parentNode;
            const afterElement = getDragAfterElement(container, e.clientY);
            if (afterElement) {
                container.insertBefore(dragging, afterElement);
            } else {
                container.appendChild(dragging);
            }
            // Mise à jour du tableau items
            items = Array.from(container.children).map((row, index) => {
                const itemIndex = items.findIndex(item => 
                    item.sku === row.querySelector('.sku-input').value);
                return items[itemIndex];
            });
            saveDraftLocal();
        }
    });
}

function getDragAfterElement(container, y) {
    const draggableElements = [...container.querySelectorAll('.item-row:not(.dragging)')];
    
    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        if (offset < 0 && offset > closest.offset) {
            return { offset: offset, element: child };
        } else {
            return closest;
        }
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

// Import multiple
function handleBulkFileImport(file) {
    const reader = new FileReader();
    
    reader.onload = async (e) => {
        try {
            let data;
            if (file.name.endsWith('.csv')) {
                data = parseBulkCSV(e.target.result);
            } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                data = parseBulkExcel(e.target.result);
            } else {
                throw new Error('Format de fichier non supporté');
            }

            processBulkImport(data);
        } catch (error) {
            console.error('Erreur import:', error);
            showNotification(error.message, 'error');
        }
    };

    if (file.name.endsWith('.csv')) {
        reader.readAsText(file, 'UTF-8');
    } else {
        reader.readAsBinaryString(file);
    }
}

function parseBulkCSV(content) {
    const lines = content.split(/\r?\n/).filter(line => line.trim());
    const headers = lines[0].split(',').map(h => h.trim().toLowerCase());  // Changé ; en , pour le séparateur
    
    // Liste étendue des noms possibles pour les colonnes
    const skuHeaders = [
        'sku', 
        'art. no', 
        'article nr.', 
        'article nr',  // Ajouté sans point
        'article', 
        'référence'
    ];
    
    const qtyHeaders = [
        'quantity', 
        'qty', 
        'amount',      // Important pour votre cas
        'nombre', 
        'quantité'
    ];

    // Recherche des index des colonnes avec gestion plus souple des espaces et de la casse
    const skuIndex = headers.findIndex(h => 
        skuHeaders.some(sh => h.replace(/[.,\s]/g, '').toLowerCase() === sh.replace(/[.,\s]/g, '').toLowerCase())
    );
    
    const qtyIndex = headers.findIndex(h => 
        qtyHeaders.some(qh => h.replace(/[.,\s]/g, '').toLowerCase() === qh.replace(/[.,\s]/g, '').toLowerCase())
    );

    if (skuIndex === -1) {
        throw new Error(`Colonne SKU manquante. Noms possibles: ${skuHeaders.join(', ')}`);
    }

    console.log('Headers trouvés:', {
        headers: headers,
        sku: headers[skuIndex],
        qty: qtyIndex !== -1 ? headers[qtyIndex] : 'non trouvé'
    });

    return lines.slice(1)
        .map(line => {
            // Gestion correcte des champs entre guillemets
            const values = line.match(/(".*?"|[^,]+)(?=\s*,|\s*$)/g)
                .map(val => val.replace(/^"|"$/g, '').trim());

            const rawSku = values[skuIndex]?.trim() || '';
            const addPrefix = document.getElementById('addK2Prefix').checked;
            const sku = addPrefix ? `K2-${rawSku}` : rawSku;
            return {
                sku,
                quantity: qtyIndex !== -1 ? parseInt(values[qtyIndex]) || 1 : 1
            };
        })
        .filter(item => item.sku);
}

function parseBulkExcel(content) {
    const workbook = XLSX.read(content, { type: 'binary' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    if (jsonData.length < 2) {
        throw new Error('Fichier vide ou invalide');
    }

    const headers = jsonData[0].map(h => h?.toString().toLowerCase().trim());
    
    // Liste étendue des noms possibles pour les colonnes
    const skuHeaders = [
        'sku', 
        'art. no', 
        'article nr.', 
        'article nr',  // Ajouté sans point
        'article', 
        'référence'
    ];
    
    const qtyHeaders = [
        'quantity', 
        'qty', 
        'amount',      // Important pour votre cas
        'nombre', 
        'quantité'
    ];

    // Recherche des index des colonnes avec gestion plus souple des espaces et de la casse
    const skuIndex = headers.findIndex(h => 
        skuHeaders.some(sh => h.replace(/[.,\s]/g, '').toLowerCase() === sh.replace(/[.,\s]/g, '').toLowerCase())
    );
    
    const qtyIndex = headers.findIndex(h => 
        qtyHeaders.some(qh => h.replace(/[.,\s]/g, '').toLowerCase() === qh.replace(/[.,\s]/g, '').toLowerCase())
    );

    if (skuIndex === -1) {
        throw new Error(`Colonne SKU manquante. Noms possibles: ${skuHeaders.join(', ')}`);
    }

    console.log('Headers trouvés:', {
        headers: headers,
        sku: headers[skuIndex],
        qty: qtyIndex !== -1 ? headers[qtyIndex] : 'non trouvé'
    });

    return jsonData.slice(1)
        .map(row => {
            const rawSku = row[skuIndex]?.toString().trim() || '';
            const addPrefix = document.getElementById('addK2Prefix').checked;
            const sku = addPrefix ? `K2-${rawSku}` : rawSku;
            return {
                sku,
                quantity: qtyIndex !== -1 ? parseInt(row[qtyIndex]) || 1 : 1
            };
        })
        .filter(item => item.sku);
}

function processBulkImport(importData) {
    let imported = 0;
    let errors = [];

    importData.forEach((row, index) => {
        const stockItem = stockItems.find(item => item.sku === row.sku);

        if (stockItem) {
            const quantity = row.quantity;
            if (quantity > stockItem.qty) {
                errors.push(`Ligne ${index + 1}: Stock insuffisant pour ${row.sku} (demandé: ${quantity}, disponible: ${stockItem.qty})`);
            } else {
                const newItem = {
                    ...stockItem,
                    quantity
                };
                items.push(newItem);
                const rowElement = createItemRow(newItem);
                document.getElementById('itemsList').appendChild(rowElement);
                imported++;
            }
        } else {
            errors.push(`Ligne ${index + 1}: SKU ${row.sku} non trouvé`);
        }
    });

    updateTotal();
    saveDraftLocal();

    if (imported > 0) {
        showNotification(`${imported} article(s) importé(s) avec succès`);
    }
    if (errors.length > 0) {
        showNotification(`${errors.length} erreur(s) durant l'import`, 'error');
        alert('Détail des erreurs:\n\n' + errors.join('\n'));
    }
}
// Fonctions d'export
function prepareExportData(orderData) {
    return items.map(item => {
        // Création de l'objet avec toutes les colonnes initialisées à vide
        const rowData = EXPORT_COLUMNS.reduce((acc, col) => {
            acc[col] = '';
            return acc;
        }, {});

        // Ajout des valeurs spécifiques
        rowData['Order Date'] = orderData.orderDate;
        rowData['CF.Date de livraison souhaitée'] = '';
        rowData['PurchaseOrder'] = `${orderData.supplier} - ${orderData.orderNumber} `;
        rowData['SKU'] = item.sku;
        rowData['Item Name'] = item.item_name;
        rowData['QuantityOrdered'] = item.quantity;

        // Colonnes d'expédition explicitement vides
        const shippingColumns = [
            'Delivery Method', 'Shipping Address', 'Shipping Street2',
            'Shipping City', 'Shipping State', 'Shipping Country',
            'Shipping Code', 'Shipping Phone', 'CF.REMARQUES POUR PREPARATION'
        ];
        shippingColumns.forEach(col => rowData[col] = '');

        // Ajout des valeurs fixes
        Object.assign(rowData, FIXED_VALUES);

        return rowData;
    });
}

function exportToExcel() {
    try {
        const orderData = getCurrentOrderData();
        if (!validateOrderData(orderData)) return;

        const exportData = prepareExportData(orderData);
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(exportData, { header: EXPORT_COLUMNS });

        ws['!cols'] = EXPORT_COLUMNS.map(col => ({
            wch: Math.max(
                col.length,
                ...exportData.map(row => String(row[col] || '').length),
                20
            )
        }));

        XLSX.utils.book_append_sheet(wb, ws, "Commande");
        XLSX.writeFile(wb, `${orderData.supplier}-${orderData.orderNumber}_${orderData.orderDate}.xlsx`);
        showNotification('Export Excel réussi');
    } catch (error) {
        showNotification('Erreur lors de l\'export Excel', 'error');
        console.error('Export Excel error:', error);
    }
}

function exportToCSV() {
    try {
        const orderData = getCurrentOrderData();
        if (!validateOrderData(orderData)) return;

        const exportData = prepareExportData(orderData);
        const rows = [EXPORT_COLUMNS.join(';')];
        
        exportData.forEach(row => {
            const rowValues = EXPORT_COLUMNS.map(col => {
                const value = row[col]?.toString() || '';
                return value.includes(';') ? `"${value}"` : value;
            });
            rows.push(rowValues.join(';'));
        });

        // Ajout de lignes vides à la fin
        for (let i = 0; i < 15; i++) {
            rows.push(EXPORT_COLUMNS.map(() => '').join(';'));
        }

        const csv = '\ufeff' + rows.join('\r\n');
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${orderData.supplier}-${orderData.orderNumber}_${orderData.orderDate}.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        showNotification('Export CSV réussi');
    } catch (error) {
        showNotification('Erreur lors de l\'export CSV', 'error');
        console.error('Export CSV error:', error);
    }
}

// Fonctions de sauvegarde
function saveDraftLocal() {
    const draftData = getCurrentOrderData();
    try {
        localStorage.setItem('currentDraft', JSON.stringify(draftData));
        lastSavedState = JSON.stringify(draftData);
    } catch (error) {
        showNotification('Erreur lors de la sauvegarde locale', 'error');
        console.error('Erreur sauvegarde:', error);
    }
}

function saveDraftFile() {
    const draftData = getCurrentOrderData();
    const orderNumber = draftData.orderNumber || 'brouillon';
    const date = new Date().toISOString().split('T')[0];
    const fileName = `commande_${orderNumber}_${date}.json`;

    try {
        const blob = new Blob([JSON.stringify(draftData, null, 2)], {
            type: 'application/json'
        });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        showNotification('Brouillon sauvegardé avec succès');
    } catch (error) {
        showNotification('Erreur lors de la sauvegarde du fichier', 'error');
    }
}

function getCurrentOrderData() {
    return {
        orderNumber: document.getElementById('orderNumber').value,
        supplier: document.getElementById('supplier').value,
        orderDate: document.getElementById('orderDate').value,
        status: document.getElementById('orderStatus').value,
        items: items,
        lastModified: new Date().toISOString()
    };
}

function validateOrderData(orderData) {
    const errors = [];
    if (!orderData.orderNumber) errors.push("WorkOrder manquant");
    if (!orderData.supplier) errors.push("Nom du client manquant");
    if (!orderData.orderDate) errors.push("Date manquante");
    if (items.length === 0) errors.push("Aucun article dans la commande");

    if (errors.length > 0) {
        showNotification(`Erreurs de validation:\n${errors.join('\n')}`, 'error');
        return false;
    }
    return true;
}

function restoreDraft(draftData) {
    document.getElementById('orderNumber').value = draftData.orderNumber || '';
    document.getElementById('supplier').value = draftData.supplier || '';
    document.getElementById('orderDate').value = draftData.orderDate || '';
    document.getElementById('orderStatus').value = draftData.status || 'draft';
    
    items = draftData.items || [];
    document.getElementById('itemsList').innerHTML = '';
    items.forEach(item => {
        const row = createItemRow(item);
        document.getElementById('itemsList').appendChild(row);
    });
    
    updateTotal();
    updateOrderStatus();
}

function updateOrderStatus() {
    const status = document.getElementById('orderStatus').value;
    const orderCard = document.getElementById('orderCard');
    
    orderCard.classList.remove('status-draft', 'status-confirmed');
    orderCard.classList.add(`status-${status}`);

    const inputs = document.querySelectorAll('#orderCard input, #orderCard button, #orderCard select');
    inputs.forEach(input => {
        if (!input.classList.contains('status-control')) {
            input.disabled = status === 'confirmed';
        }
    });

    document.getElementById('addItem').disabled = status === 'confirmed' || stockItems.length === 0;
    document.getElementById('showImport').disabled = status === 'confirmed' || stockItems.length === 0;
    document.getElementById('clearAll').disabled = status === 'confirmed';
}

// Événements
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('orderDate').valueAsDate = new Date();

    const savedStock = localStorage.getItem('stockItems');
    if (savedStock) {
        try {
            stockItems = JSON.parse(savedStock);
            updateStockPreview(stockItems);
            document.getElementById('addItem').disabled = false;
            document.getElementById('showImport').disabled = false;
            showNotification('Stock restauré', 'info');
        } catch (error) {
            console.error('Erreur restauration stock:', error);
        }
    }

    const savedDraft = localStorage.getItem('currentDraft');
    if (savedDraft) {
        try {
            const draftData = JSON.parse(savedDraft);
            if (confirm('Un brouillon non terminé a été trouvé. Voulez-vous le restaurer ?')) {
                restoreDraft(draftData);
                showNotification('Brouillon restauré', 'info');
            } else {
                localStorage.removeItem('currentDraft');
            }
        } catch (error) {
            console.error('Erreur restauration brouillon:', error);
        }
    }

    lastSavedState = JSON.stringify(getCurrentOrderData());

    // Événements d'import
    document.getElementById('stockFile').addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        await importStock(file);
        e.target.value = '';
    });

    document.getElementById('clearStock').addEventListener('click', () => {
        if (confirm('Voulez-vous vraiment effacer tout le stock ?')) {
            stockItems = [];
            localStorage.removeItem('stockItems');
            updateStockPreview([]);
            document.getElementById('addItem').disabled = true;
            document.getElementById('showImport').disabled = true;
            showNotification('Stock effacé');
        }
    });

    // Événements des articles
    document.getElementById('addItem').addEventListener('click', () => {
        if (document.getElementById('orderStatus').value === 'confirmed') {
            showNotification('Impossible de modifier une commande confirmée', 'error');
            return;
        }
        const row = createItemRow();
        if (row) {
            document.getElementById('itemsList').appendChild(row);
            items.push({});
            updateTotal();
            saveDraftLocal();
        }
    });

    // Import multiple
    document.getElementById('showImport').addEventListener('click', () => {
        if (document.getElementById('orderStatus').value === 'confirmed') {
            showNotification('Impossible de modifier une commande confirmée', 'error');
            return;
        }
        document.getElementById('importSection').classList.toggle('d-none');
    });

    document.getElementById('bulkFileInput').addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        handleBulkFileImport(file);
        e.target.value = '';
        document.getElementById('importSection').classList.add('d-none');
    });

    document.getElementById('processImport').addEventListener('click', () => {
        const text = document.getElementById('bulkImport').value.trim();
        if (!text) {
            showNotification('Aucune donnée à importer', 'error');
            return;
        }

        const lines = text.split('\n')
            .map(line => {
                const [rawSku, quantityStr = "1"] = line.split(',').map(v => v.trim());
                const addPrefix = document.getElementById('addK2Prefix').checked;
                const sku = addPrefix ? `K2-${rawSku}` : rawSku;
                return { sku, quantity: parseInt(quantityStr) || 1 };
            })
            .filter(item => item.sku);

        processBulkImport(lines);
        document.getElementById('bulkImport').value = '';
        document.getElementById('importSection').classList.add('d-none');
    });

    // Gestion du statut
    document.getElementById('orderStatus').addEventListener('change', () => {
        const status = document.getElementById('orderStatus').value;
        if (status === 'confirmed' && !confirm('Attention : Confirmer la commande la rendra non modifiable. Continuer ?')) {
            document.getElementById('orderStatus').value = 'draft';
            return;
        }
        updateOrderStatus();
        saveDraftLocal();
    });

    // Gestion des sauvegardes
    document.getElementById('saveDraftLocal').addEventListener('click', () => {
        saveDraftLocal();
        showNotification('Brouillon sauvegardé localement');
    });

    document.getElementById('saveDraftFile').addEventListener('click', saveDraftFile);

    document.getElementById('loadDraftFile').addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        
        if (items.length > 0 && !confirm('Cette action remplacera la commande en cours. Continuer ?')) {
            e.target.value = '';
            return;
        }
        
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const draftData = JSON.parse(e.target.result);
                restoreDraft(draftData);
                showNotification('Brouillon chargé avec succès');
            } catch (error) {
                showNotification('Erreur lors du chargement du brouillon', 'error');
            }
        };
        reader.readAsText(file);
        e.target.value = '';
    });

    // Export
    document.getElementById('exportXls').addEventListener('click', exportToExcel);
    document.getElementById('exportCsv').addEventListener('click', exportToCSV);

    // Suppression
    document.getElementById('clearAll').addEventListener('click', () => {
        if (document.getElementById('orderStatus').value === 'confirmed') {
            showNotification('Impossible de modifier une commande confirmée', 'error');
            return;
        }
        if (items.length === 0) {
            showNotification('Aucun article à effacer', 'info');
            return;
        }
        if (confirm('Voulez-vous vraiment effacer tous les articles ?')) {
            items = [];
            document.getElementById('itemsList').innerHTML = '';
            updateTotal();
            saveDraftLocal();
            showNotification('Tous les articles ont été effacés');
        }
    });

    // Vérification des modifications non sauvegardées
    window.addEventListener('beforeunload', (e) => {
        const currentState = JSON.stringify(getCurrentOrderData());
        if (currentState !== lastSavedState) {
            e.preventDefault();
            e.returnValue = '';
            return 'Vous avez des modifications non sauvegardées. Voulez-vous vraiment quitter ?';
        }
    });
});
    </script>
</body>
</html>
