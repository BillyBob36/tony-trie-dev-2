// Utiliser la configuration externe
const CONFIG = window.CONFIG || {};

// Valider la configuration au démarrage
if (typeof validateConfig === 'function') {
    const configErrors = validateConfig();
    if (configErrors.length > 0) {
        console.error('Erreurs de configuration:', configErrors);
        alert('Erreurs de configuration détectées. Consultez la console pour plus de détails.');
    }
}

// Variables globales
let isSignedIn = false;
let spreadsheets = [];
let criteria = [];
let results = []; // Stockage des résultats de traitement
let isProcessing = false;
let currentBatch = 0;
let totalBatches = 0;
let processedCount = 0;
let matchedCount = 0;

// Initialisation
document.addEventListener('DOMContentLoaded', function() {
    initializeGoogleAPI();
    setupEventListeners();
    generateDefaultPrompt();
    setupAIConfigTabs();
    setupAIConfigEventListeners();
});

// Initialisation de l'API Google
function initializeGoogleAPI() {
    gapi.load('client', initClient);
}

// Variables pour Google Identity Services
let tokenClient;
let accessToken = null;

function initClient() {
    // Vérifier si le Client ID est configuré
    if (!CONFIG.GOOGLE.CLIENT_ID || CONFIG.GOOGLE.CLIENT_ID.includes('YOUR_GOOGLE_CLIENT_ID') || CONFIG.GOOGLE.CLIENT_ID.includes('test-client-id')) {
        console.warn('Google Client ID non configuré. Mode démo activé.');
        showDemoMode();
        return;
    }
    
    // Attendre que la bibliothèque Google Identity Services soit chargée
    function waitForGoogleLibrary() {
        if (typeof google !== 'undefined' && google.accounts && google.accounts.oauth2) {
            try {
                // Initialiser le client de token pour OAuth 2.0
                tokenClient = google.accounts.oauth2.initTokenClient({
                    client_id: CONFIG.GOOGLE.CLIENT_ID,
                    scope: CONFIG.GOOGLE.SCOPES,
                    callback: (tokenResponse) => {
                        accessToken = tokenResponse.access_token;
                        isSignedIn = true;
                        updateSigninStatus();
                        loadSpreadsheets();
                    },
                });
                
                // Configurer le bouton de connexion Google
                renderSignInButton();
                
            } catch (error) {
                console.error('Erreur d\'initialisation:', error);
                showError('Erreur d\'initialisation de l\'API Google');
            }
        } else {
            // Réessayer après 100ms
            setTimeout(waitForGoogleLibrary, 100);
        }
    }
    
    waitForGoogleLibrary();
}

// Mode démo pour tester l'interface sans Google OAuth
function showDemoMode() {
    document.getElementById('signin-button').style.display = 'none';
    document.getElementById('user-info').style.display = 'none';
    document.getElementById('main-content').style.display = 'block';
    
    // Afficher un message de mode démo
    const demoMessage = document.createElement('div');
    demoMessage.className = 'demo-notice';
    demoMessage.innerHTML = `
        <h3>🔧 Mode Démo</h3>
        <p>Pour utiliser l'application complète, configurez votre Google Client ID dans config.js</p>
        <p>En mode démo, vous pouvez explorer l'interface utilisateur.</p>
    `;
    
    const container = document.querySelector('.container');
    container.insertBefore(demoMessage, container.firstChild);
    
    // Simuler des données de test
    populateDemoData();
}

// Données de démonstration
function populateDemoData() {
    // Simuler des spreadsheets de démonstration
    spreadsheets = [
        { id: 'demo1', name: 'Fichier Demo - Entreprises' },
        { id: 'demo2', name: 'Fichier Demo - Profils' },
        { id: 'demo3', name: 'Fichier Demo - Postes' }
    ];
    
    // Mettre à jour les sélecteurs
    updateSpreadsheetSelectors();
    
    showSuccess('Mode démo activé - Interface prête à être testée');
}

function renderSignInButton() {
    // Créer un bouton personnalisé pour OAuth 2.0
    const signInBtn = document.createElement('button');
    signInBtn.className = 'btn btn-primary btn-large';
    signInBtn.innerHTML = '🔐 Se connecter avec Google';
    signInBtn.onclick = signIn;
    
    const container = document.getElementById('signin-button');
    container.innerHTML = '';
    container.appendChild(signInBtn);
}

function signIn() {
    if (tokenClient) {
        tokenClient.requestAccessToken();
    } else {
        console.error('Token client non initialisé');
        showError('Erreur d\'initialisation de l\'authentification');
    }
}

function signOut() {
    if (accessToken) {
        google.accounts.oauth2.revoke(accessToken, () => {
            console.log('Déconnexion réussie');
            showSuccess('Déconnexion réussie');
        });
        accessToken = null;
        isSignedIn = false;
        updateSigninStatus();
    }
}

function updateSigninStatus() {
    if (isSignedIn) {
        // Pour la nouvelle API, nous n'avons pas accès direct au profil utilisateur
        // Nous pouvons utiliser l'API People pour obtenir ces informations si nécessaire
        document.getElementById('user-name').textContent = 'Utilisateur connecté';
        document.getElementById('signin-button').style.display = 'none';
        document.getElementById('user-info').style.display = 'block';
        document.getElementById('main-content').style.display = 'block';
    } else {
        document.getElementById('signin-button').style.display = 'block';
        document.getElementById('user-info').style.display = 'none';
        document.getElementById('main-content').style.display = 'none';
    }
}

// Configuration des événements
function setupEventListeners() {
    // Bouton de déconnexion
    document.getElementById('signout-button').addEventListener('click', signOut);
    
    // Bouton d'ajout de critère
    document.getElementById('add-criteria').addEventListener('click', addCriteria);
    
    // Formulaire de critère
    document.getElementById('criteria-form').addEventListener('submit', addCriteria);
    
    // Fermeture du modal
    document.querySelector('.close').addEventListener('click', closeCriteriaModal);
    window.addEventListener('click', function(event) {
        const modal = document.getElementById('criteria-modal');
        if (event.target === modal) {
            closeCriteriaModal();
        }
    });
    
    // Changements de sélection pour les critères (modal)
    document.getElementById('criteria-spreadsheet').addEventListener('change', loadCriteriaSheets);
    document.getElementById('criteria-sheet').addEventListener('change', loadCriteriaColumns);
    
    // Boutons de traitement
    document.getElementById('start-processing').addEventListener('click', startProcessing);
    document.getElementById('stop-processing').addEventListener('click', stopProcessing);
    
    // Configurer les gestionnaires d'événements pour les sélecteurs principaux
    setupTargetEventHandlers();
    setupOutputEventHandlers();
}

// Gestionnaires d'événements pour les sélecteurs de données cibles
function setupTargetEventHandlers() {
    const spreadsheetSelect = document.getElementById('target-spreadsheet');
    const sheetSelect = document.getElementById('target-sheet');
    
    spreadsheetSelect.addEventListener('change', async function() {
        if (this.value) {
            clearSelect(sheetSelect, 'Chargement des feuilles...');
            
            try {
                const sheets = await loadSheets(this.value, 'target-sheet');
                if (sheets.length > 0) {
                    populateSelect(sheetSelect, sheets.map(sheet => ({
                        id: sheet.properties.title,
                        name: sheet.properties.title
                    })), 'id', 'name', 'Sélectionner une feuille...');
                } else {
                    disableSelect(sheetSelect, 'Aucune feuille trouvée');
                }
            } catch (error) {
                disableSelect(sheetSelect, 'Erreur de chargement');
            }
        } else {
            clearSelect(sheetSelect, 'Sélectionner d\'abord un fichier');
        }
    });
}

// Gestionnaires d'événements pour les sélecteurs de sortie
function setupOutputEventHandlers() {
    const spreadsheetSelect = document.getElementById('output-spreadsheet');
    const sheetSelect = document.getElementById('output-sheet');
    
    spreadsheetSelect.addEventListener('change', async function() {
        if (this.value) {
            clearSelect(sheetSelect, 'Chargement des feuilles...');
            
            try {
                const sheets = await loadSheets(this.value, 'output-sheet');
                if (sheets.length > 0) {
                    populateSelect(sheetSelect, sheets.map(sheet => ({
                        id: sheet.properties.title,
                        name: sheet.properties.title
                    })), 'id', 'name', 'Sélectionner une feuille...');
                } else {
                    disableSelect(sheetSelect, 'Aucune feuille trouvée');
                }
            } catch (error) {
                disableSelect(sheetSelect, 'Erreur de chargement');
            }
        } else {
            clearSelect(sheetSelect, 'Sélectionner d\'abord un fichier');
        }
    });
}

// Chargement des Google Sheets
async function loadSpreadsheets() {
    // En mode démo, ne pas charger les vrais spreadsheets
    if (!CONFIG.GOOGLE.CLIENT_ID || CONFIG.GOOGLE.CLIENT_ID.includes('YOUR_GOOGLE_CLIENT_ID') || CONFIG.GOOGLE.CLIENT_ID.includes('test-client-id')) {
        console.log('Mode démo - Utilisation des données simulées');
        return;
    }
    
    // Vérifier si l'utilisateur est connecté
    if (!accessToken) {
        showError('Veuillez vous connecter d\'abord');
        return;
    }
    
    try {
        showLoading(CONFIG.MESSAGES.LOADING.SPREADSHEETS);
        
        console.log('Chargement des spreadsheets avec accessToken:', accessToken ? 'présent' : 'absent');
        
        const url = `https://www.googleapis.com/drive/v3/files?pageSize=${CONFIG.UI.MAX_SPREADSHEETS_DISPLAY}&q=mimeType%3D%27application%2Fvnd.google-apps.spreadsheet%27&fields=files(id%2C%20name)&key=${CONFIG.GOOGLE.API_KEY}`;
        console.log('URL de requête:', url);
        
        const response = await fetch(url, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        console.log('Statut de la réponse:', response.status);
        
        if (!response.ok) {
            const errorText = await response.text();
            console.error('Erreur API:', errorText);
            throw new Error(`HTTP error! status: ${response.status} - ${errorText}`);
        }
        
        const data = await response.json();
        console.log('Données reçues:', data);
        
        spreadsheets = data.files || [];
        console.log('Nombre de spreadsheets trouvés:', spreadsheets.length);
        
        if (spreadsheets.length === 0) {
            console.warn('Aucun fichier Google Sheets trouvé');
            showError(CONFIG.MESSAGES.ERROR.NO_SPREADSHEETS_FOUND);
        } else {
            console.log('Spreadsheets trouvés:', spreadsheets.map(s => s.name));
        }
        
        // Mettre à jour tous les sélecteurs de fichiers
        updateSpreadsheetSelectors();
        
        hideLoading();
        if (spreadsheets.length > 0) {
            showSuccess(`${spreadsheets.length} fichiers Google Sheets chargés`);
        }
    } catch (error) {
        console.error('Erreur lors du chargement des fichiers:', error);
        
        // Vérifier si c'est une erreur d'API non activée
        if (error.message && error.message.includes('Google Drive API has not been used')) {
            showError(CONFIG.MESSAGES.ERROR.API_NOT_ENABLED + '\n\nLien d\'activation: https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=802942339545');
        } else {
            showError(CONFIG.MESSAGES.ERROR.SPREADSHEETS_LOAD_FAILED + ': ' + error.message);
        }
        
        hideLoading();
    }
}

function updateSpreadsheetSelectors() {
    const selectors = [
        'criteria-spreadsheet',
        'target-spreadsheet', 
        'output-spreadsheet'
    ];
    
    selectors.forEach(selectorId => {
        const selector = document.getElementById(selectorId);
        populateSelect(selector, spreadsheets, 'id', 'name', 'Sélectionner un fichier...');
    });
}

function populateSelect(selectElement, items, valueKey, textKey, defaultText = 'Sélectionner...') {
    selectElement.innerHTML = `<option value="">${defaultText}</option>`;
    
    items.forEach(item => {
        const option = document.createElement('option');
        option.value = item[valueKey];
        option.textContent = item[textKey];
        selectElement.appendChild(option);
    });
    
    // Activer le select
    selectElement.disabled = false;
    selectElement.classList.remove('disabled');
    
    // Ajouter la fonctionnalité de recherche si il y a plus de 3 options
    // SAUF pour les colonnes cibles dans les critères (target-column-*) et les feuilles (sheet) mais PAS les fichiers (spreadsheet)
    if (items.length > 3 && 
        !selectElement.id.startsWith('target-column-') && 
        !selectElement.id.includes('sheet') || 
        (items.length > 3 && selectElement.id.includes('spreadsheet'))) {
        addSearchToSelect(selectElement);
    }
}

function disableSelect(selectElement, message = 'Aucune option disponible') {
    selectElement.innerHTML = `<option value="">${message}</option>`;
    selectElement.disabled = true;
    selectElement.classList.add('disabled');
}

function clearSelect(selectElement, message = 'Sélectionner d\'abord l\'option précédente') {
    selectElement.innerHTML = `<option value="">${message}</option>`;
    selectElement.disabled = true;
    selectElement.classList.add('disabled');
}

// Fonction pour ajouter la fonctionnalité de recherche à un élément select
function addSearchToSelect(selectElement) {
    // Vérifier si la recherche n'est pas déjà ajoutée
    if (selectElement.dataset.searchEnabled) {
        return;
    }
    
    selectElement.dataset.searchEnabled = 'true';
    
    // Stocker les options originales
    let originalOptions = [];
    
    function updateOriginalOptions() {
        originalOptions = Array.from(selectElement.options).map(option => ({
            value: option.value,
            text: option.textContent,
            selected: option.selected
        }));
    }
    
    // Créer un conteneur pour le select avec recherche
    const container = document.createElement('div');
    container.className = 'select-with-search';
    container.style.display = 'flex';
    container.style.alignItems = 'flex-start';
    container.style.gap = '8px';
    container.style.width = '100%';
    
    // Créer un sous-conteneur pour l'input et le select
    const inputSelectContainer = document.createElement('div');
    inputSelectContainer.style.flex = '1';
    inputSelectContainer.style.display = 'flex';
    inputSelectContainer.style.flexDirection = 'column';
    
    // Créer l'input de recherche
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.className = 'search-input';
    searchInput.placeholder = 'Tapez pour rechercher...';
    searchInput.style.width = '100%';
    searchInput.style.padding = '8px';
    searchInput.style.border = '1px solid #ddd';
    searchInput.style.borderRadius = '4px';
    searchInput.style.marginBottom = '5px';
    searchInput.style.display = 'none';
    
    // Créer le bouton de recherche
    const searchButton = document.createElement('button');
    searchButton.type = 'button';
    searchButton.innerHTML = '🔍';
    searchButton.title = 'Activer la recherche';
    searchButton.style.border = '1px solid #ddd';
    searchButton.style.background = '#f8f9fa';
    searchButton.style.cursor = 'pointer';
    searchButton.style.fontSize = '16px';
    searchButton.style.padding = '8px 12px';
    searchButton.style.borderRadius = '4px';
    searchButton.style.height = 'fit-content';
    searchButton.style.flexShrink = '0';
    
    // Remplacer le select par le conteneur
    selectElement.parentNode.insertBefore(container, selectElement);
    inputSelectContainer.appendChild(searchInput);
    inputSelectContainer.appendChild(selectElement);
    container.appendChild(inputSelectContainer);
    container.appendChild(searchButton);
    
    // Ajuster le style du select
    selectElement.style.width = '100%';
    
    let isSearchMode = false;
    
    // Fonction pour basculer le mode recherche
    function toggleSearchMode() {
        isSearchMode = !isSearchMode;
        
        if (isSearchMode) {
            updateOriginalOptions();
            searchInput.style.display = 'block';
            searchButton.innerHTML = '✕';
            searchButton.title = 'Fermer la recherche';
            searchInput.focus();
        } else {
            searchInput.style.display = 'none';
            searchInput.value = '';
            searchButton.innerHTML = '🔍';
            searchButton.title = 'Activer la recherche';
            restoreAllOptions();
        }
    }
    
    // Fonction pour activer automatiquement la recherche au clavier
    function activateSearchOnKeyboard() {
        if (!isSearchMode) {
            toggleSearchMode();
        }
    }
    
    // Fonction pour restaurer toutes les options
    function restoreAllOptions() {
        selectElement.innerHTML = '';
        originalOptions.forEach(optionData => {
            const option = document.createElement('option');
            option.value = optionData.value;
            option.textContent = optionData.text;
            if (optionData.selected) {
                option.selected = true;
            }
            selectElement.appendChild(option);
        });
    }
    
    // Fonction pour filtrer les options
    function filterOptions(searchTerm) {
        const filteredOptions = originalOptions.filter(optionData => 
            optionData.text.toLowerCase().includes(searchTerm.toLowerCase())
        );
        
        selectElement.innerHTML = '';
        
        // Toujours inclure l'option par défaut (première option)
        if (originalOptions.length > 0) {
            const defaultOption = document.createElement('option');
            defaultOption.value = originalOptions[0].value;
            defaultOption.textContent = originalOptions[0].text;
            selectElement.appendChild(defaultOption);
        }
        
        // Ajouter les options filtrées (sauf la première qui est déjà ajoutée)
        filteredOptions.slice(1).forEach(optionData => {
            const option = document.createElement('option');
            option.value = optionData.value;
            option.textContent = optionData.text;
            selectElement.appendChild(option);
        });
    }
    
    // Événements
    searchButton.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        toggleSearchMode();
    });
    
    searchInput.addEventListener('input', (e) => {
        filterOptions(e.target.value);
    });
    
    searchInput.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            toggleSearchMode();
        }
    });
    
    // Pas d'activation automatique de la recherche
    
    // Observer les changements dans le select pour mettre à jour les options
    const observer = new MutationObserver(() => {
        if (!isSearchMode) {
            updateOriginalOptions();
        }
    });
    
    observer.observe(selectElement, {
        childList: true,
        subtree: true
    });
    
    // Mettre à jour les options initiales
    updateOriginalOptions();
}

// Chargement des feuilles d'un spreadsheet
async function loadSheets(spreadsheetId, targetSelectId) {
    // En mode démo, simuler des feuilles
    if (!CONFIG.GOOGLE.CLIENT_ID || CONFIG.GOOGLE.CLIENT_ID.includes('YOUR_GOOGLE_CLIENT_ID') || CONFIG.GOOGLE.CLIENT_ID.includes('test-client-id')) {
        const selector = document.getElementById(targetSelectId);
        const demoSheets = [
            { properties: { title: 'Feuille Demo 1' } },
            { properties: { title: 'Feuille Demo 2' } },
            { properties: { title: 'Données' } }
        ];
        
        selector.innerHTML = '<option value="">Sélectionner une feuille...</option>';
        
        demoSheets.forEach(sheet => {
            const option = document.createElement('option');
            option.value = sheet.properties.title;
            option.textContent = sheet.properties.title;
            selector.appendChild(option);
        });
        
        return demoSheets;
    }
    
    // Vérifier si l'utilisateur est connecté
    if (!accessToken) {
        showError('Veuillez vous connecter d\'abord');
        return [];
    }
    
    try {
        showLoading(CONFIG.MESSAGES.LOADING.SHEETS);
        
        const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        const sheets = data.sheets || [];
        const selector = document.getElementById(targetSelectId);
        
        if (sheets.length === 0) {
            showError(CONFIG.MESSAGES.ERROR.NO_SHEETS_FOUND);
        }
        
        selector.innerHTML = '<option value="">Sélectionner une feuille...</option>';
        
        sheets.forEach(sheet => {
            const option = document.createElement('option');
            option.value = sheet.properties.title;
            option.textContent = sheet.properties.title;
            selector.appendChild(option);
        });
        
        hideLoading();
        return sheets;
    } catch (error) {
        console.error('Erreur lors du chargement des feuilles:', error);
        showError(CONFIG.MESSAGES.ERROR.SHEETS_LOAD_FAILED);
        hideLoading();
        return [];
    }
}

function loadCriteriaSheets() {
    const spreadsheetId = document.getElementById('criteria-spreadsheet').value;
    if (spreadsheetId) {
        loadSheets(spreadsheetId, 'criteria-sheet');
    }
}

function loadTargetSheets() {
    const spreadsheetId = document.getElementById('target-spreadsheet').value;
    if (spreadsheetId) {
        loadSheets(spreadsheetId, 'target-sheet');
        loadTargetColumns(); // Charger aussi les colonnes pour la sélection des critères
    }
}

function loadOutputSheets() {
    const spreadsheetId = document.getElementById('output-spreadsheet').value;
    if (spreadsheetId) {
        loadSheets(spreadsheetId, 'output-sheet');
    }
}

// Chargement des colonnes
async function loadColumns(spreadsheetId, sheetName, targetSelectId) {
    // En mode démo, simuler des colonnes
    if (!CONFIG.GOOGLE.CLIENT_ID || CONFIG.GOOGLE.CLIENT_ID.includes('YOUR_GOOGLE_CLIENT_ID') || CONFIG.GOOGLE.CLIENT_ID.includes('test-client-id')) {
        const selector = document.getElementById(targetSelectId);
        const demoHeaders = ['Nom', 'Entreprise', 'Poste', 'Email', 'Téléphone', 'Secteur'];
        
        const columnsData = demoHeaders.map((header, index) => ({
            letter: String.fromCharCode(65 + index),
            name: `${String.fromCharCode(65 + index)} - ${header}`,
            index: index
        }));
        
        populateSelect(selector, columnsData, 'index', 'name', 'Sélectionner une colonne...');
        
        return columnsData;
    }
    
    // Vérifier si l'utilisateur est connecté
    if (!accessToken) {
        showError('Veuillez vous connecter d\'abord');
        return [];
    }
    
    try {
        showLoading(CONFIG.MESSAGES.LOADING.COLUMNS);
        
        const response = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}!1:1`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const data = await response.json();
        const headers = data.values ? data.values[0] : [];
        const selector = document.getElementById(targetSelectId);
        
        if (headers.length === 0) {
            showError(CONFIG.MESSAGES.ERROR.NO_COLUMNS_FOUND);
            return [];
        }
        
        // Préparer les données des colonnes avec lettres et noms
        const columnsData = headers.map((header, index) => ({
            letter: String.fromCharCode(65 + index),
            name: `${String.fromCharCode(65 + index)} - ${header}`,
            index: index
        }));
        
        populateSelect(selector, columnsData, 'index', 'name', 'Sélectionner une colonne...');
        
        if (CONFIG.DEBUG.LOG_COLUMN_LOADING) {
            console.log(`Colonnes chargées pour ${sheetName}:`, headers);
        }
        
        hideLoading();
        return columnsData;
    } catch (error) {
        console.error('Erreur lors du chargement des colonnes:', error);
        showError(CONFIG.MESSAGES.ERROR.COLUMNS_LOAD_FAILED);
        hideLoading();
        return [];
    }
}

function loadCriteriaColumns() {
    const spreadsheetId = document.getElementById('criteria-spreadsheet').value;
    const sheetName = document.getElementById('criteria-sheet').value;
    
    if (spreadsheetId && sheetName) {
        loadColumns(spreadsheetId, sheetName, 'criteria-column');
    }
}

function loadTargetColumns() {
    const spreadsheetId = document.getElementById('target-spreadsheet').value;
    const sheetName = document.getElementById('target-sheet').value;
    
    if (spreadsheetId && sheetName) {
        loadColumns(spreadsheetId, sheetName, 'target-column');
    }
}

// Gestion des critères
function openCriteriaModal() {
    document.getElementById('criteria-modal').style.display = 'block';
    
    // Réinitialiser le formulaire
    document.getElementById('criteria-form').reset();
    
    // Recharger les sélecteurs
    updateSpreadsheetSelectors();
}

function closeCriteriaModal() {
    document.getElementById('criteria-modal').style.display = 'none';
}

function addCriteria(event) {
    event.preventDefault();
    
    const name = document.getElementById('criteria-name').value;
    const spreadsheetId = document.getElementById('criteria-spreadsheet').value;
    const sheetName = document.getElementById('criteria-sheet').value;
    const columnIndex = document.getElementById('criteria-column').value;
    const targetColumnIndex = document.getElementById('target-column').value;
    
    if (!name || !spreadsheetId || !sheetName || columnIndex === '' || targetColumnIndex === '') {
        showError(CONFIG.MESSAGES.ERROR.MISSING_FIELDS);
        return;
    }
    
    const criteriaItem = {
        id: Date.now(),
        name: name,
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        columnIndex: parseInt(columnIndex),
        targetColumnIndex: parseInt(targetColumnIndex),
        values: []
    };
    
    criteria.push(criteriaItem);
    updateCriteriaDisplay();
    closeCriteriaModal();
    
    // Charger les valeurs du critère
    loadCriteriaValues(criteriaItem);
}

// Ajouter un nouveau critère
function addCriteria() {
    const criteriaContainer = document.getElementById('criteria-list');
    const criteriaId = Date.now(); // ID unique basé sur le timestamp
    
    const criteriaDiv = document.createElement('div');
    criteriaDiv.className = 'criteria-item';
    criteriaDiv.id = `criteria-${criteriaId}`;
    
    criteriaDiv.innerHTML = `
        <div class="form-row">
            <div class="form-group">
                <label>Nom du critère:</label>
                <input type="text" id="criteria-name-${criteriaId}" placeholder="Nom du critère" readonly>
            </div>
            <div class="form-group">
                <label>Fichier:</label>
                <select id="criteria-spreadsheet-${criteriaId}">
                    <option value="">Sélectionner un fichier...</option>
                </select>
            </div>
            <div class="form-group">
                <label>Feuille:</label>
                <select id="criteria-sheet-${criteriaId}" disabled>
                    <option value="">Sélectionner d'abord un fichier</option>
                </select>
            </div>
            <div class="form-group">
                <label>Colonne:</label>
                <select id="criteria-column-${criteriaId}" disabled>
                    <option value="">Sélectionner d'abord une feuille</option>
                </select>
            </div>
            <div class="form-group">
                <label>Colonne correspondante dans les données cibles:</label>
                <select id="target-column-${criteriaId}" disabled>
                    <option value="">Sélectionner d'abord les données cibles</option>
                </select>
            </div>
            <button type="button" class="btn btn-danger" onclick="removeCriteria('${criteriaId}')">
                Supprimer
            </button>
        </div>
    `;
    
    criteriaContainer.appendChild(criteriaDiv);
    
    // Peupler le sélecteur de fichiers
    const spreadsheetSelect = document.getElementById(`criteria-spreadsheet-${criteriaId}`);
    if (spreadsheets && spreadsheets.length > 0) {
        populateSelect(spreadsheetSelect, spreadsheets, 'id', 'name', 'Sélectionner un fichier...');
    } else {
        disableSelect(spreadsheetSelect, 'Aucun fichier disponible');
    }
    
    // Configurer les gestionnaires d'événements
    setupCriteriaEventHandlers(criteriaId);
    
    // Mettre à jour le compteur de critères
    updateCriteriaCount();
}

async function loadCriteriaValues(criteriaItem) {
    if (!isSignedIn || !accessToken) {
        console.error('Utilisateur non connecté');
        return;
    }
    
    try {
        const columnLetter = String.fromCharCode(65 + criteriaItem.columnIndex);
        const range = `${criteriaItem.sheetName}!${columnLetter}:${columnLetter}`;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${criteriaItem.spreadsheetId}/values/${encodeURIComponent(range)}`;
        
        const response = await fetch(url, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const data = await response.json();
        const values = data.values || [];
        criteriaItem.values = values.slice(1).map(row => row[0]).filter(val => val && val.trim());
        
        console.log(`Critère "${criteriaItem.name}" chargé avec ${criteriaItem.values.length} valeurs`);
    } catch (error) {
        console.error('Erreur lors du chargement des valeurs du critère:', error);
        showError(`${CONFIG.MESSAGES.ERROR.CRITERIA_LOAD_FAILED}: "${criteriaItem.name}"`);
    }
}

function updateCriteriaDisplay() {
    const container = document.getElementById('criteria-list');
    container.innerHTML = '';
    
    criteria.forEach(criteriaItem => {
        const div = document.createElement('div');
        div.className = 'criteria-item';
        div.innerHTML = `
            <div class="criteria-info">
                <div class="criteria-name">${criteriaItem.name}</div>
                <div class="criteria-details">
                    ${getSpreadsheetName(criteriaItem.spreadsheetId)} > ${criteriaItem.sheetName} > 
                    Colonne ${String.fromCharCode(65 + criteriaItem.columnIndex)}
                    (${criteriaItem.values.length} valeurs)
                </div>
            </div>
            <div class="criteria-actions">
                <button class="btn btn-danger btn-small" onclick="removeCriteria(${criteriaItem.id})">
                    Supprimer
                </button>
            </div>
        `;
        container.appendChild(div);
    });
}

function removeCriteria(id) {
    criteria = criteria.filter(c => c.id !== id);
    updateCriteriaDisplay();
}

// Supprimer un critère dynamique
function removeDynamicCriteria(criteriaId) {
    const criteriaDiv = document.getElementById(`criteria-${criteriaId}`);
    if (criteriaDiv) {
        criteriaDiv.remove();
        updateCriteriaCount();
    }
}

// Mettre à jour le compteur de critères
function updateCriteriaCount() {
    const criteriaItems = document.querySelectorAll('.criteria-item');
    const count = criteriaItems.length;
    
    // Mettre à jour l'affichage du nombre de critères
    const addButton = document.getElementById('add-criteria');
    if (addButton) {
        const countText = count > 0 ? ` (${count})` : '';
        addButton.textContent = `Ajouter un critère${countText}`;
    }
    
    // Activer/désactiver le bouton de traitement selon le nombre de critères
    const startButton = document.getElementById('start-processing');
    if (startButton) {
        startButton.disabled = count === 0;
        if (count === 0) {
            startButton.title = 'Ajoutez au moins un critère pour commencer';
        } else {
            startButton.title = '';
        }
    }
}

function getSpreadsheetName(id) {
    const sheet = spreadsheets.find(s => s.id === id);
    return sheet ? sheet.name : 'Fichier inconnu';
}

// Génération du prompt par défaut
function generateDefaultPrompt() {
    document.getElementById('ai-prompt').value = CONFIG.DEFAULT_AI_PROMPT;
}

// Collecter les critères depuis les éléments DOM dynamiques
function collectDynamicCriteria() {
    const criteriaItems = document.querySelectorAll('.criteria-item[id^="criteria-"]');
    const dynamicCriteria = [];
    
    criteriaItems.forEach(item => {
        const criteriaId = item.id.replace('criteria-', '');
        const nameInput = document.getElementById(`criteria-name-${criteriaId}`);
        const spreadsheetSelect = document.getElementById(`criteria-spreadsheet-${criteriaId}`);
        const sheetSelect = document.getElementById(`criteria-sheet-${criteriaId}`);
        const columnSelect = document.getElementById(`criteria-column-${criteriaId}`);
        const targetColumnSelect = document.getElementById(`target-column-${criteriaId}`);
        
        // Vérifier que tous les champs sont remplis
        if (nameInput && nameInput.value && 
            spreadsheetSelect && spreadsheetSelect.value && 
            sheetSelect && sheetSelect.value && 
            columnSelect && columnSelect.value !== '' &&
            targetColumnSelect && targetColumnSelect.value !== '') {
            
            dynamicCriteria.push({
                id: parseInt(criteriaId),
                name: nameInput.value,
                spreadsheetId: spreadsheetSelect.value,
                sheetName: sheetSelect.value,
                columnIndex: parseInt(columnSelect.value),
                targetColumnIndex: parseInt(targetColumnSelect.value),
                values: [] // Sera chargé plus tard
            });
        }
    });
    
    return dynamicCriteria;
}

// Fonction de validation des champs obligatoires
function validateRequiredFields() {
    // Supprimer tous les messages d'erreur existants
    document.querySelectorAll('.field-error-message').forEach(msg => msg.remove());
    
    const requiredFields = [
        { id: 'target-spreadsheet', label: 'Fichier source' },
        { id: 'target-sheet', label: 'Feuille source' },
        { id: 'output-spreadsheet', label: 'Fichier de destination' },
        { id: 'output-sheet', label: 'Feuille de destination' }
    ];
    
    let firstMissingField = null;
    
    for (const field of requiredFields) {
        const element = document.getElementById(field.id);
        if (!element || !element.value) {
            // Créer le message d'erreur
            const errorMessage = document.createElement('div');
            errorMessage.className = 'field-error-message';
            errorMessage.style.color = 'red';
            errorMessage.style.fontSize = '12px';
            errorMessage.style.marginBottom = '5px';
            errorMessage.textContent = `⚠️ ${field.label} est obligatoire`;
            
            // Insérer le message avant l'élément
            if (element && element.parentNode) {
                element.parentNode.insertBefore(errorMessage, element);
            }
            
            // Marquer le premier champ manquant
            if (!firstMissingField) {
                firstMissingField = element;
            }
        }
    }
    
    // Vérifier les critères et leurs colonnes correspondantes
    const criteriaItems = document.querySelectorAll('.criteria-item[id^="criteria-"]');
    let hasCriteria = false;
    let hasIncompleteCriteria = false;
    let firstIncompleteField = null;
    
    criteriaItems.forEach(item => {
        const criteriaId = item.id.replace('criteria-', '');
        const nameInput = document.getElementById(`criteria-name-${criteriaId}`);
        const spreadsheetSelect = document.getElementById(`criteria-spreadsheet-${criteriaId}`);
        const sheetSelect = document.getElementById(`criteria-sheet-${criteriaId}`);
        const columnSelect = document.getElementById(`criteria-column-${criteriaId}`);
        const targetColumnSelect = document.getElementById(`target-column-${criteriaId}`);
        
        if (nameInput && nameInput.value) {
            hasCriteria = true;
            
            // Vérifier si la colonne correspondante est manquante
            if (!targetColumnSelect || targetColumnSelect.value === '') {
                hasIncompleteCriteria = true;
                
                // Créer le message d'erreur pour la colonne correspondante
                const existingError = targetColumnSelect?.parentNode?.querySelector('.field-error-message');
                if (!existingError && targetColumnSelect) {
                    const errorMessage = document.createElement('div');
                    errorMessage.className = 'field-error-message';
                    errorMessage.style.color = 'red';
                    errorMessage.style.fontSize = '12px';
                    errorMessage.style.marginBottom = '5px';
                    errorMessage.textContent = '⚠️ Sélectionnez la colonne correspondante dans les données cibles';
                    
                    targetColumnSelect.parentNode.insertBefore(errorMessage, targetColumnSelect);
                    
                    if (!firstIncompleteField) {
                        firstIncompleteField = targetColumnSelect;
                    }
                }
            }
        }
    });
    
    // Si aucun critère n'est défini
    if (!hasCriteria) {
        const addCriteriaButton = document.getElementById('add-criteria');
        if (addCriteriaButton) {
            const errorMessage = document.createElement('div');
            errorMessage.className = 'field-error-message';
            errorMessage.style.color = 'red';
            errorMessage.style.fontSize = '12px';
            errorMessage.style.marginBottom = '5px';
            errorMessage.textContent = '⚠️ Au moins un critère est obligatoire';
            
            addCriteriaButton.parentNode.insertBefore(errorMessage, addCriteriaButton);
            
            if (!firstMissingField) {
                firstMissingField = addCriteriaButton;
            }
        }
    }
    
    // Prioriser les champs de critères incomplets
    if (firstIncompleteField && !firstMissingField) {
        firstMissingField = firstIncompleteField;
    }
    
    // Faire défiler vers le premier champ manquant
    if (firstMissingField) {
        firstMissingField.scrollIntoView({ 
            behavior: 'smooth', 
            block: 'center' 
        });
        
        // Optionnel: mettre le focus sur le champ
        setTimeout(() => {
            if (firstMissingField.focus) {
                firstMissingField.focus();
            }
        }, 500);
        
        return false; // Validation échouée
    }
    
    return true; // Validation réussie
}

// Traitement principal
async function startProcessing() {
    // Valider les champs obligatoires
    if (!validateRequiredFields()) {
        return;
    }
    
    // En mode démo, simuler le traitement
    if (!CONFIG.GOOGLE.CLIENT_ID || CONFIG.GOOGLE.CLIENT_ID.includes('YOUR_GOOGLE_CLIENT_ID') || CONFIG.GOOGLE.CLIENT_ID.includes('test-client-id')) {
        simulateDemoProcessing();
        return;
    }
    
    // Collecter les critères depuis les deux systèmes
    const dynamicCriteria = collectDynamicCriteria();
    const allCriteria = [...criteria, ...dynamicCriteria];
    
    // Mettre à jour le tableau global des critères
    criteria = allCriteria;
    
    const targetSpreadsheetId = document.getElementById('target-spreadsheet').value;
    const targetSheetName = document.getElementById('target-sheet').value;
    const outputSpreadsheetId = document.getElementById('output-spreadsheet').value;
    const outputSheetName = document.getElementById('output-sheet').value;
    
    isProcessing = true;
    processedCount = 0;
    matchedCount = 0;
    
    // Mettre à jour l'interface
    document.getElementById('start-processing').style.display = 'none';
    document.getElementById('stop-processing').style.display = 'inline-block';
    document.getElementById('progress-section').style.display = 'block';
    document.getElementById('results-section').style.display = 'none';
    
    try {
        // Sauvegarder les paramètres IA avant de commencer
        saveAISettings();
        
        // Charger les valeurs des critères dynamiques
        updateProgress('Chargement des critères...', 0);
        for (const criteriaItem of dynamicCriteria) {
            if (criteriaItem.values.length === 0) {
                await loadCriteriaValues(criteriaItem);
            }
        }
        
        // Charger toutes les données cibles
        updateProgress(CONFIG.MESSAGES.LOADING.TARGET_DATA, 10);
        const targetData = await loadTargetData(targetSpreadsheetId, targetSheetName);
        
        if (targetData.length === 0) {
            throw new Error(CONFIG.MESSAGES.ERROR.NO_TARGET_DATA);
        }
        
        // Configuration optimisée pour le traitement par lots
        const aiConfig = getAIConfig();
        const batchSize = calculateOptimalBatchSize(targetData.length, criteria.length);
        
        // Calculer le nombre de lots avec la taille optimisée
        totalBatches = Math.ceil(targetData.length / batchSize);
        
        let allMatchedRows = [];
        let apiCallCount = 0;
        let startTime = Date.now();
        let exportedBatches = [];
        let isFirstExport = true;
        
        const exportConfig = CONFIG.EXPORT.PROGRESSIVE_EXPORT;
        const exportBatchSize = exportConfig.ENABLED ? exportConfig.BATCH_SIZE : 1000;
        
        console.log('Configuration exportation progressive:', {
            enabled: exportConfig.ENABLED,
            batchSize: exportBatchSize,
            delay: exportConfig.BATCH_DELAY
        });
        
        showMessage(`Traitement optimisé: ${totalBatches} lots de ${batchSize} éléments avec exportation progressive (export tous les ${exportBatchSize} matchs)`, 'info', 5000);
        
        // Traiter par lots avec exportation progressive
        for (let i = 0; i < totalBatches && isProcessing; i++) {
            currentBatch = i + 1;
            const start = i * batchSize;
            const end = Math.min(start + batchSize, targetData.length);
            const batch = targetData.slice(start, end);
            
            updateProgress(`Traitement du lot ${currentBatch}/${totalBatches}...`, 
                         (currentBatch - 1) / totalBatches * 90); // Réserver 10% pour l'export final
            
            // Traitement avec gestion des erreurs et retry
            const batchResult = await processBatchWithRetry(batch, apiCallCount);
            allMatchedRows = allMatchedRows.concat(batchResult.matchedRows);
            apiCallCount += batchResult.apiCalls;
            
            processedCount += batch.length;
            matchedCount = allMatchedRows.length;
            updateStats();
            
            // Exportation progressive : accumuler les résultats
            if (exportConfig.ENABLED) {
                if (batchResult.matchedRows.length > 0) {
                    exportedBatches = exportedBatches.concat(batchResult.matchedRows);
                    console.log(`Ajout de ${batchResult.matchedRows.length} résultats. Total accumulé: ${exportedBatches.length}/${exportBatchSize}`);
                }
                
                // Exporter quand on atteint la taille de lot d'export ou à la fin
                const shouldExport = exportedBatches.length >= exportBatchSize || (i === totalBatches - 1 && exportedBatches.length > 0);
                console.log(`Vérification export: ${exportedBatches.length} >= ${exportBatchSize} = ${exportedBatches.length >= exportBatchSize}, fin traitement: ${i === totalBatches - 1}, shouldExport: ${shouldExport}`);
                
                if (shouldExport) {
                    try {
                        updateProgress(`Export en cours... (${exportedBatches.length} résultats)`, 
                                     (currentBatch / totalBatches) * 90 + 5);
                        
                        await exportProgressiveBatch(
                            exportedBatches, 
                            outputSpreadsheetId, 
                            outputSheetName, 
                            targetData[0], 
                            isFirstExport
                        );
                        
                        console.log(`Exporté ${exportedBatches.length} résultats (lot ${Math.ceil(matchedCount / exportBatchSize)})`);
                        exportedBatches = []; // Vider le buffer d'export
                        isFirstExport = false;
                        
                        // Délai entre les exports pour éviter la surcharge
                        if (exportConfig.BATCH_DELAY > 0) {
                            await sleep(exportConfig.BATCH_DELAY);
                        }
                        
                    } catch (exportError) {
                        console.error('Erreur lors de l\'export progressif:', exportError);
                        showMessage(`Erreur d'export: ${exportError.message}`, 'warning', 3000);
                        // Continuer le traitement même en cas d'erreur d'export
                    }
                }
            }
            
            // Délai adaptatif entre les lots de traitement
            if (i + 1 < totalBatches && isProcessing) {
                const adaptiveDelay = calculateAdaptiveDelay(apiCallCount, startTime, aiConfig.batchDelay);
                await sleep(adaptiveDelay);
            }
        }
        
        if (isProcessing) {
            // Export final des résultats restants
            if (!exportConfig.ENABLED) {
                // Export classique : tous les résultats en une fois
                updateProgress(CONFIG.MESSAGES.LOADING.EXPORTING, 95);
                await exportEnrichedResults(allMatchedRows, outputSpreadsheetId, outputSheetName, targetData[0]);
            } else if (exportedBatches.length > 0) {
                // Export progressif : exporter les derniers résultats restants
                try {
                    updateProgress(`Export final... (${exportedBatches.length} résultats restants)`, 95);
                    await exportProgressiveBatch(
                        exportedBatches, 
                        outputSpreadsheetId, 
                        outputSheetName, 
                        targetData[0], 
                        isFirstExport
                    );
                    console.log(`Export final: ${exportedBatches.length} résultats`);
                } catch (exportError) {
                    console.error('Erreur lors de l\'export final:', exportError);
                    showMessage(`Erreur d'export final: ${exportError.message}`, 'warning', 3000);
                }
            }
            
            updateProgress('Terminé !', 100);
            showResults(allMatchedRows.length, processedCount);
            
            // Afficher les statistiques de performance
            const totalTime = Math.round((Date.now() - startTime) / 1000);
            const exportMode = exportConfig.ENABLED ? 'progressive' : 'finale';
            showMessage(`Traitement terminé en ${totalTime}s avec ${apiCallCount} appels API (export ${exportMode})`, 'success', 10000);
            showSuccess(CONFIG.MESSAGES.SUCCESS.PROCESSING_COMPLETE);
        }
        
    } catch (error) {
        console.error('Erreur lors du traitement:', error);
        showError(`${CONFIG.MESSAGES.ERROR.PROCESSING_FAILED}: ${error.message}`);
    } finally {
        stopProcessing();
    }
}

function stopProcessing() {
    isProcessing = false;
    document.getElementById('start-processing').style.display = 'inline-block';
    document.getElementById('stop-processing').style.display = 'none';
}

async function loadTargetData(spreadsheetId, sheetName) {
    if (!isSignedIn || !accessToken) {
        throw new Error('Utilisateur non connecté');
    }
    
    try {
        const range = `${sheetName}!A:ZZ`;
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`;
        
        const response = await fetch(url, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        const data = await response.json();
        const values = data.values || [];
        let targetData = values.slice(1); // Exclure l'en-tête
        
        // Appliquer la limitation du nombre de lignes
        const rowsLimit = document.getElementById('rows-limit').value;
        if (rowsLimit !== 'all') {
            const limit = parseInt(rowsLimit);
            targetData = targetData.slice(0, limit);
        }
        
        return targetData;
    } catch (error) {
        console.error('Erreur lors du chargement des données cibles:', error);
        throw new Error(CONFIG.MESSAGES.ERROR.TARGET_DATA_LOAD_FAILED);
    }
}

async function processBatch(batch) {
    // Nouvelle logique de traitement séquentiel
    return await processBatchSequential(batch);
}

// Nouvelle fonction qui implémente le filtrage progressif séquentiel
async function processBatchSequential(batch) {
    const aiConfig = getAIConfig();
    let remainingRows = batch.map((row, index) => ({
        data: row,
        originalIndex: index,
        matchDetails: [],
        totalConfidence: 0,
        matchCount: 0
    }));
    
    // Traiter chaque critère séquentiellement
    for (let criteriaIndex = 0; criteriaIndex < criteria.length; criteriaIndex++) {
        if (!isProcessing) break;
        
        const criteriaItem = criteria[criteriaIndex];
        const newRemainingRows = [];
        
        console.log(`Traitement du critère ${criteriaIndex + 1}/${criteria.length}: "${criteriaItem.name}" sur ${remainingRows.length} lignes restantes`);
        
        for (const rowItem of remainingRows) {
            if (!isProcessing) break;
            
            const targetValue = rowItem.data[criteriaItem.targetColumnIndex] || '';
            
            // Vérifier si la valeur cible existe
            if (!targetValue.trim()) {
                // Ligne éliminée car pas de valeur dans la colonne cible
                continue;
            }
            
            // Vérifier la correspondance avec le seuil de tolérance
            const matchResult = await checkCriteriaMatchWithThreshold(targetValue, criteriaItem.values, aiConfig.threshold);
            
            // Si la correspondance respecte le seuil, garder la ligne
            if (matchResult.isMatch && matchResult.confidence >= aiConfig.threshold) {
                // Ajouter les détails de correspondance
                rowItem.matchDetails.push({
                    criteria: criteriaItem.name,
                    targetValue: targetValue,
                    confidence: matchResult.confidence,
                    matchType: matchResult.matchType,
                    details: matchResult.details
                });
                
                rowItem.totalConfidence += matchResult.confidence;
                rowItem.matchCount++;
                
                // Garder cette ligne pour le critère suivant
                newRemainingRows.push(rowItem);
            }
            // Sinon, la ligne est éliminée et ne sera pas traitée pour les critères suivants
        }
        
        // Mettre à jour les lignes restantes pour le critère suivant
        remainingRows = newRemainingRows;
        
        console.log(`Après critère "${criteriaItem.name}": ${remainingRows.length} lignes restantes`);
        
        // Si plus aucune ligne ne reste, arrêter le traitement
        if (remainingRows.length === 0) {
            console.log('Aucune ligne ne respecte tous les critères traités jusqu\'à présent');
            break;
        }
    }
    
    // Convertir les lignes restantes au format attendu
    const matchedRows = remainingRows.map(rowItem => ({
        data: rowItem.data,
        matchDetails: rowItem.matchDetails,
        averageConfidence: rowItem.matchCount > 0 ? Math.round(rowItem.totalConfidence / rowItem.matchCount) : 0,
        matchTypes: [...new Set(rowItem.matchDetails.map(d => d.matchType))]
    }));
    
    console.log(`Traitement terminé: ${matchedRows.length} lignes correspondent à tous les critères`);
    return matchedRows;
}

async function checkCriteriaMatch(targetValue, criteriaValues) {
    const aiConfig = getAIConfig();
    
    // Vérification rapide d'abord (correspondance exacte)
    const targetLower = aiConfig.caseSensitive ? targetValue.trim() : targetValue.toLowerCase().trim();
    
    for (const criteriaValue of criteriaValues) {
        const criteriaLower = aiConfig.caseSensitive ? criteriaValue.trim() : criteriaValue.toLowerCase().trim();
        if (targetLower === criteriaLower || 
            targetLower.includes(criteriaLower) || 
            criteriaLower.includes(targetLower)) {
            return {
                isMatch: true,
                confidence: 95,
                matchType: 'exact',
                details: 'Correspondance exacte trouvée'
            };
        }
    }
    
    // Si pas de correspondance exacte, utiliser l'IA
    return await checkWithAI(targetValue, criteriaValues);
}

// Nouvelle fonction qui prend en compte le seuil de tolérance utilisateur
async function checkCriteriaMatchWithThreshold(targetValue, criteriaValues, userThreshold) {
    const aiConfig = getAIConfig();
    
    // Vérification rapide d'abord (correspondance exacte)
    const targetLower = aiConfig.caseSensitive ? targetValue.trim() : targetValue.toLowerCase().trim();
    
    for (const criteriaValue of criteriaValues) {
        const criteriaLower = aiConfig.caseSensitive ? criteriaValue.trim() : criteriaValue.toLowerCase().trim();
        if (targetLower === criteriaLower || 
            targetLower.includes(criteriaLower) || 
            criteriaLower.includes(targetLower)) {
            return {
                isMatch: true,
                confidence: 95,
                matchType: 'exact',
                details: 'Correspondance exacte trouvée'
            };
        }
    }
    
    // Si pas de correspondance exacte, utiliser l'IA avec le seuil utilisateur
    if (aiConfig.enabled) {
        try {
            const aiResult = await checkWithAI(targetValue, criteriaValues);
            
            // Appliquer le seuil utilisateur au résultat de l'IA
            if (aiResult.confidence >= userThreshold) {
                return aiResult;
            } else {
                // Confidence insuffisante selon le seuil utilisateur
                return {
                    isMatch: false,
                    confidence: aiResult.confidence,
                    matchType: aiResult.matchType,
                    details: `Confidence ${aiResult.confidence}% < seuil ${userThreshold}%`
                };
            }
        } catch (error) {
            console.warn('Erreur IA, pas de correspondance trouvée:', error);
            return {
                isMatch: false,
                confidence: 0,
                matchType: 'none',
                details: 'Aucune correspondance trouvée (IA indisponible)'
            };
        }
    }
    
    // Si IA désactivée, aucune correspondance
    return {
        isMatch: false,
        confidence: 0,
        matchType: 'none',
        details: 'Aucune correspondance exacte trouvée'
    };
}

async function checkWithAI(targetValue, criteriaValues) {
    const aiConfig = getAIConfig();
    
    // Vérifier que la clé API OpenAI est configurée
    if (!CONFIG.OPENAI?.API_KEY) {
        console.error('Clé API OpenAI non configurée');
        throw new Error('Clé API OpenAI manquante. Veuillez configurer votre clé API dans les variables d\'environnement.');
    }
    
    let retries = 0;
    
    while (retries < aiConfig.retryAttempts) {
        try {
            const criteriaList = criteriaValues.slice(0, CONFIG.PROCESSING?.MAX_CRITERIA_PER_REQUEST || 10).join(', ');
            
            const messages = [
                {
                    role: 'system',
                    content: aiConfig.prompt
                },
                {
                    role: 'user',
                    content: `Valeur à vérifier: "${targetValue}"\nListe des critères: ${criteriaList}\n\nEst-ce que la valeur correspond à l'un des critères de la liste ?`
                }
            ];
            
            const response = await fetch(CONFIG.OPENAI?.API_URL || 'https://api.openai.com/v1/chat/completions', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${CONFIG.OPENAI?.API_KEY || ''}`
                },
                body: JSON.stringify({
                    model: aiConfig.model,
                    messages: messages,
                    max_tokens: aiConfig.maxTokens,
                    temperature: aiConfig.temperature
                })
            });
            
            if (!response.ok) {
                throw new Error(`Erreur API OpenAI: ${response.status} - ${response.statusText}`);
            }
            
            const data = await response.json();
            const result = data.choices[0].message.content.trim();
            
            // Analyser la réponse pour déterminer la correspondance et la confiance
            const confidence = extractConfidenceFromResponse(result.toLowerCase());
            const isMatch = confidence >= aiConfig.threshold || 
                           result.toLowerCase().includes('oui') || 
                           result.toLowerCase().includes('yes') || 
                           result.toLowerCase().includes('match');
            
            if (CONFIG.DEBUG?.LOG_AI_RESPONSES) {
                console.log(`IA Response for "${targetValue}": ${result} (Confidence: ${confidence}%)`);
            }
            
            return {
                isMatch: isMatch,
                confidence: confidence,
                matchType: 'ai',
                details: result
            };
            
        } catch (error) {
            retries++;
            console.error(`Erreur lors de la vérification IA (tentative ${retries}):`, error);
            
            if (retries >= aiConfig.retryAttempts) {
                console.error('Nombre maximum de tentatives atteint pour l\'IA');
                
                // Fallback vers la correspondance basique si activé
                if (aiConfig.fallbackMatching) {
                    const basicResult = checkBasicMatch(targetValue, criteriaValues);
                    return {
                        isMatch: basicResult,
                        confidence: basicResult ? 60 : 40,
                        matchType: 'fallback',
                        details: 'IA indisponible, correspondance basique utilisée'
                    };
                } else {
                    return {
                        isMatch: false,
                        confidence: 0,
                        matchType: 'error',
                        details: `Erreur IA: ${error.message}`
                    };
                }
            }
            
            // Attendre avant de réessayer
            await sleep(aiConfig.batchDelay);
        }
    }
    
    return {
        isMatch: false,
        confidence: 0,
        matchType: 'error',
        details: 'Échec après plusieurs tentatives'
    };
}

// Extraire le niveau de confiance de la réponse IA
function extractConfidenceFromResponse(response) {
    const confidenceMatch = response.match(/(\d+)%/);
    if (confidenceMatch) {
        return parseInt(confidenceMatch[1]);
    }
    
    // Mots-clés pour estimer la confiance
    if (response.includes('très probable') || response.includes('certain') || response.includes('definitely')) return 90;
    if (response.includes('probable') || response.includes('likely') || response.includes('oui')) return 75;
    if (response.includes('possible') || response.includes('maybe') || response.includes('perhaps')) return 50;
    if (response.includes('peu probable') || response.includes('unlikely') || response.includes('doubtful')) return 25;
    if (response.includes('impossible') || response.includes('non') || response.includes('no')) return 10;
    
    return 50; // Valeur par défaut
}

// Fonction de correspondance basique en cas d'échec de l'IA
function checkBasicMatch(targetValue, criteriaValues) {
    const targetLower = targetValue.toLowerCase().trim();
    
    return criteriaValues.some(criteriaValue => {
        const criteriaLower = criteriaValue.toLowerCase().trim();
        return targetLower === criteriaLower || 
               targetLower.includes(criteriaLower) || 
               criteriaLower.includes(targetLower);
    });
}

async function exportResults(matchedRows, outputSpreadsheetId, outputSheetName, headerRow) {
    if (!isSignedIn || !accessToken) {
        throw new Error('Utilisateur non connecté');
    }
    
    try {
        // Préparer les données avec l'en-tête
        const dataToExport = [headerRow, ...matchedRows];
        
        // Effacer la feuille de destination
        const clearRange = `${outputSheetName}!A:ZZ`;
        const clearUrl = `https://sheets.googleapis.com/v4/spreadsheets/${outputSpreadsheetId}/values/${encodeURIComponent(clearRange)}:clear`;
        
        await fetch(clearUrl, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });
        
        // Écrire les nouvelles données
        const updateRange = `${outputSheetName}!A1`;
        const updateUrl = `https://sheets.googleapis.com/v4/spreadsheets/${outputSpreadsheetId}/values/${encodeURIComponent(updateRange)}?valueInputOption=RAW`;
        
        const updateResponse = await fetch(updateUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                values: dataToExport
            })
        });
        
        if (!updateResponse.ok) {
            throw new Error(`HTTP ${updateResponse.status}: ${updateResponse.statusText}`);
        }
        
        console.log(`${matchedRows.length} lignes exportées vers ${outputSheetName}`);
    } catch (error) {
        console.error('Erreur lors de l\'exportation:', error);
        throw new Error(CONFIG.MESSAGES.ERROR.EXPORT_FAILED);
    }
}

// Fonctions utilitaires
function updateProgress(text, percentage) {
    document.getElementById('progress-text').textContent = text;
    document.getElementById('progress-fill').style.width = `${percentage}%`;
}

function updateStats() {
    document.getElementById('processed-count').textContent = processedCount;
    document.getElementById('matched-count').textContent = matchedCount;
}

function showResults(matchedCount, totalCount) {
    const resultsSection = document.getElementById('results-section');
    const resultsSummary = document.getElementById('results-summary');
    
    // Récupérer les informations de destination
    const outputSpreadsheetSelect = document.getElementById('output-spreadsheet');
    const outputSheetSelect = document.getElementById('output-sheet');
    const outputSpreadsheetId = outputSpreadsheetSelect.value;
    const outputSheetName = outputSheetSelect.value;
    const outputSpreadsheetName = outputSpreadsheetSelect.options[outputSpreadsheetSelect.selectedIndex]?.text || 'Fichier de destination';
    
    // Créer le lien vers Google Sheets
    let destinationLink = '';
    if (outputSpreadsheetId && outputSheetName) {
        const sheetUrl = `https://docs.google.com/spreadsheets/d/${outputSpreadsheetId}/edit#gid=0`;
        destinationLink = `
            <div class="destination-link">
                <h4>📊 Résultats Exportés</h4>
                <p>Les résultats ont été automatiquement exportés vers votre fichier Google Sheets.</p>
                <a href="${sheetUrl}" target="_blank" class="btn">
                    <span>📋</span>
                    Ouvrir ${outputSpreadsheetName} - ${outputSheetName}
                </a>
                <small>Cliquez pour ouvrir le fichier dans un nouvel onglet</small>
            </div>
        `;
    } else {
        destinationLink = `
            <div class="destination-link">
                <h4>📊 Résultats Exportés</h4>
                <p>Les résultats ont été exportés vers la destination configurée.</p>
            </div>
        `;
    }
    
    resultsSummary.innerHTML = `
        <div class="results-summary">
            <h3>Résultats du Traitement</h3>
            <div class="stats-grid">
                <div class="stat-item">
                    <span class="stat-number">${matchedCount}</span>
                    <span class="stat-label">Correspondances trouvées</span>
                </div>
                <div class="stat-item">
                    <span class="stat-number">${totalCount}</span>
                    <span class="stat-label">Profils traités</span>
                </div>
                <div class="stat-item">
                    <span class="stat-number">${Math.round((matchedCount / totalCount) * 100)}%</span>
                    <span class="stat-label">Taux de correspondance</span>
                </div>
            </div>
        </div>
        
        ${destinationLink}
    `;
    
    resultsSection.style.display = 'block';
}

function showError(message) {
    showMessage(message, 'error-message', CONFIG.UI.ERROR_MESSAGE_DURATION);
}

function showSuccess(message) {
    showMessage(message, 'success-message', CONFIG.UI.SUCCESS_MESSAGE_DURATION);
}

function showMessage(message, className, duration) {
    const messageDiv = document.createElement('div');
    messageDiv.className = className;
    messageDiv.textContent = message;
    
    // Insérer au début du container
    const container = document.querySelector('.container');
    container.insertBefore(messageDiv, container.firstChild);
    
    // Supprimer après la durée spécifiée
    setTimeout(() => {
        if (messageDiv.parentNode) {
            messageDiv.parentNode.removeChild(messageDiv);
        }
    }, duration);
}

function showLoading(message) {
    updateProgress(message, 0);
    document.getElementById('progress-section').style.display = 'block';
}

function hideLoading() {
    document.getElementById('progress-section').style.display = 'none';
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// Gestionnaires d'événements pour les sélecteurs de critères
function setupCriteriaEventHandlers(criteriaId) {
    const spreadsheetSelect = document.getElementById(`criteria-spreadsheet-${criteriaId}`);
    const sheetSelect = document.getElementById(`criteria-sheet-${criteriaId}`);
    const columnSelect = document.getElementById(`criteria-column-${criteriaId}`);
    const targetColumnSelect = document.getElementById(`target-column-${criteriaId}`);
    const nameInput = document.getElementById(`criteria-name-${criteriaId}`);
    
    // Charger les colonnes cibles quand les données cibles changent
    function loadTargetColumnsForCriteria() {
        const targetSpreadsheetId = document.getElementById('target-spreadsheet').value;
        const targetSheetName = document.getElementById('target-sheet').value;
        
        if (targetSpreadsheetId && targetSheetName) {
            loadColumns(targetSpreadsheetId, targetSheetName, `target-column-${criteriaId}`);
        } else {
            clearSelect(targetColumnSelect, 'Sélectionner d\'abord les données cibles');
        }
    }
    
    // Écouter les changements des données cibles
    const targetSpreadsheetSelect = document.getElementById('target-spreadsheet');
    const targetSheetSelect = document.getElementById('target-sheet');
    
    if (targetSpreadsheetSelect) {
        targetSpreadsheetSelect.addEventListener('change', loadTargetColumnsForCriteria);
    }
    if (targetSheetSelect) {
        targetSheetSelect.addEventListener('change', loadTargetColumnsForCriteria);
    }
    
    // Charger immédiatement si les données cibles sont déjà sélectionnées
    loadTargetColumnsForCriteria();
    
    spreadsheetSelect.addEventListener('change', async function() {
        if (this.value) {
            clearSelect(sheetSelect, 'Chargement des feuilles...');
            clearSelect(columnSelect, 'Sélectionner d\'abord une feuille');
            nameInput.value = '';
            
            try {
                if (!isSignedIn || !accessToken) {
                    throw new Error('Utilisateur non connecté');
                }
                
                const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.value}`;
                const response = await fetch(url, {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`,
                        'Content-Type': 'application/json'
                    }
                });
                
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }
                
                const data = await response.json();
                const sheets = data.sheets || [];
                if (sheets.length > 0) {
                    populateSelect(sheetSelect, sheets.map(sheet => ({
                        id: sheet.properties.title,
                        name: sheet.properties.title
                    })), 'id', 'name', 'Sélectionner une feuille...');
                } else {
                    disableSelect(sheetSelect, 'Aucune feuille trouvée');
                }
            } catch (error) {
                console.error('Erreur lors du chargement des feuilles:', error);
                disableSelect(sheetSelect, 'Erreur de chargement');
                disableSelect(columnSelect, 'Erreur de chargement');
                showError('Erreur lors du chargement des feuilles');
            }
        } else {
            clearSelect(sheetSelect, 'Sélectionner d\'abord un fichier');
            clearSelect(columnSelect, 'Sélectionner d\'abord une feuille');
            nameInput.value = '';
        }
    });
    
    sheetSelect.addEventListener('change', async function() {
        if (this.value && spreadsheetSelect.value) {
            clearSelect(columnSelect, 'Chargement des colonnes...');
            nameInput.value = '';
            
            try {
                if (!isSignedIn || !accessToken) {
                    throw new Error('Utilisateur non connecté');
                }
                
                const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetSelect.value}/values/${encodeURIComponent(this.value + '!1:1')}`;
                
                const response = await fetch(url, {
                    headers: {
                        'Authorization': `Bearer ${accessToken}`
                    }
                });
                
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }
                
                const data = await response.json();
                const values = data.values;
                if (values && values.length > 0 && values[0].length > 0) {
                    const headers = values[0];
                    const columnsData = headers.map((header, index) => ({
                        letter: String.fromCharCode(65 + index),
                        name: `${String.fromCharCode(65 + index)} - ${header}`,
                        index: index,
                        headerText: header
                    }));
                    
                    populateSelect(columnSelect, columnsData, 'index', 'name', 'Sélectionner une colonne...');
                    // Ajouter la recherche si plus de 3 options
                    if (columnsData.length > 3) {
                        addSearchToSelect(columnSelect);
                    }
                } else {
                    disableSelect(columnSelect, 'Aucune colonne trouvée');
                }
            } catch (error) {
                console.error('Erreur lors du chargement des colonnes:', error);
                disableSelect(columnSelect, 'Erreur de chargement');
                showError('Erreur lors du chargement des colonnes');
            }
        } else {
            clearSelect(columnSelect, 'Sélectionner d\'abord une feuille');
            nameInput.value = '';
        }
    });
    
    columnSelect.addEventListener('change', function() {
        if (this.value !== '') {
            // Récupérer le texte de l'en-tête depuis l'option sélectionnée
            const selectedOption = this.selectedOptions[0];
            const optionText = selectedOption.textContent;
            // Extraire le nom de l'en-tête (après le tiret)
            const headerName = optionText.split(' - ')[1] || optionText;
            nameInput.value = headerName;
        } else {
            nameInput.value = '';
        }
    });
}

// Simulation du traitement en mode démo
function simulateDemoProcessing() {
    isProcessing = true;
    processedCount = 0;
    matchedCount = 0;
    
    // Mettre à jour l'interface
    document.getElementById('start-processing').style.display = 'none';
    document.getElementById('stop-processing').style.display = 'inline-block';
    document.getElementById('progress-section').style.display = 'block';
    document.getElementById('results-section').style.display = 'none';
    
    // Générer des données de démonstration réalistes
    generateDemoResults();
    
    // Simuler le traitement avec des données fictives
    const totalRows = results.length;
    const matchedRows = results.filter(r => r.confidence > 0).length;
    
    let progress = 0;
    const interval = setInterval(() => {
        if (!isProcessing) {
            clearInterval(interval);
            return;
        }
        
        progress += 10;
        processedCount = Math.floor((progress / 100) * totalRows);
        matchedCount = Math.floor((progress / 100) * matchedRows);
        
        updateProgress(`Traitement en cours... (Mode Démo)`, progress);
        updateStats();
        
        if (progress >= 100) {
            clearInterval(interval);
            updateProgress('Terminé ! (Mode Démo)', 100);
            showResults(matchedRows, totalRows);
            showSuccess('Traitement démo terminé avec succès !');
            stopProcessing();
        }
    }, 500);
}

// Génération de données de démonstration
function generateDemoResults() {
    const demoProfiles = [
        'Jean Dupont - Développeur',
        'Marie Martin - Chef de Projet',
        'Pierre Durand - Ingénieur',
        'Sophie Bernard - Designer',
        'Luc Moreau - Analyste',
        'Emma Petit - Consultante',
        'Thomas Roux - Architecte',
        'Julie Blanc - Product Manager',
        'Nicolas Garnier - Data Scientist',
        'Camille Faure - UX Designer'
    ];
    
    const demoCompanies = [
        'Google France',
        'Microsoft',
        'Amazon',
        'Apple',
        'Meta',
        'Netflix',
        'Spotify',
        'Airbnb',
        'Uber',
        'Tesla'
    ];
    
    const matchTypes = ['exact', 'partial', 'ai'];
    
    results = [];
    
    // Générer 150 résultats de démonstration
    for (let i = 0; i < 150; i++) {
        const sourceProfile = demoProfiles[Math.floor(Math.random() * demoProfiles.length)];
        const matchedProfile = demoCompanies[Math.floor(Math.random() * demoCompanies.length)];
        const matchType = matchTypes[Math.floor(Math.random() * matchTypes.length)];
        
        // Générer un score de confiance basé sur le type de correspondance
        let confidence;
        switch (matchType) {
            case 'exact':
                confidence = 85 + Math.floor(Math.random() * 15); // 85-100%
                break;
            case 'partial':
                confidence = 50 + Math.floor(Math.random() * 35); // 50-85%
                break;
            case 'ai':
                confidence = 20 + Math.floor(Math.random() * 60); // 20-80%
                break;
        }
        
        results.push({
            sourceProfile,
            matchedProfile,
            confidence,
            matchType,
            timestamp: new Date().toISOString()
        });
    }
    
    // Ajouter quelques non-correspondances (confidence = 0)
    for (let i = 0; i < 50; i++) {
        const sourceProfile = demoProfiles[Math.floor(Math.random() * demoProfiles.length)];
        results.push({
            sourceProfile,
            matchedProfile: 'Aucune correspondance',
            confidence: 0,
            matchType: 'none',
            timestamp: new Date().toISOString()
        });
    }
}

// Fonctions de filtrage et d'aperçu supprimées - les résultats sont maintenant exportés automatiquement

// Fonctions d'export supprimées - l'export se fait automatiquement vers la destination configurée

// Écriture vers Google Sheets
async function writeToGoogleSheet(spreadsheetId, sheetName, data) {
    const range = `${sheetName}!A1:${String.fromCharCode(64 + data[0].length)}${data.length}`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
    
    const response = await makeAuthenticatedRequest(url, {
        method: 'PUT',
        body: JSON.stringify({
            values: data
        })
    });
    
    if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    
    return await response.json();
}

// Téléchargement CSV
// Fonction de téléchargement CSV supprimée - l'export se fait automatiquement

// Fonctions globales pour les événements onclick
window.removeCriteria = removeCriteria;
window.removeDynamicCriteria = removeDynamicCriteria;
window.closeCriteriaModal = closeCriteriaModal;

// Configuration des onglets IA
function setupAIConfigTabs() {
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabPanels = document.querySelectorAll('.tab-panel');
    
    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const targetTab = button.dataset.tab;
            
            // Désactiver tous les onglets
            tabButtons.forEach(btn => {
                if (btn && btn.classList) {
                    btn.classList.remove('active');
                }
            });
            tabPanels.forEach(panel => {
                if (panel && panel.classList) {
                    panel.classList.remove('active');
                }
            });
            
            // Activer l'onglet sélectionné
            if (button && button.classList) {
                button.classList.add('active');
            }
            const targetElement = document.getElementById(targetTab + '-tab');
            if (targetElement && targetElement.classList) {
                targetElement.classList.add('active');
            }
        });
    });
}

// Configuration des événements IA
function setupAIConfigEventListeners() {
    // Boutons de contrôle du prompt
    const resetPromptBtn = document.getElementById('reset-prompt');
    const savePromptBtn = document.getElementById('save-prompt');
    const promptTextarea = document.getElementById('ai-prompt');
    
    if (resetPromptBtn) {
        resetPromptBtn.addEventListener('click', () => {
            generateDefaultPrompt();
            showMessage('Prompt réinitialisé avec les valeurs par défaut', 'success', 3000);
        });
    }
    
    if (savePromptBtn) {
        savePromptBtn.addEventListener('click', () => {
            const prompt = promptTextarea.value.trim();
            if (prompt) {
                localStorage.setItem('aiPrompt', prompt);
                showMessage('Prompt sauvegardé', 'success', 2000);
            }
        });
    }
    
    // Sliders avec valeurs dynamiques
    const temperatureSlider = document.getElementById('ai-temperature');
    const temperatureValue = document.getElementById('temperature-value');
    const thresholdSlider = document.getElementById('match-threshold');
    const thresholdValue = document.getElementById('threshold-value');
    
    if (temperatureSlider && temperatureValue) {
        temperatureSlider.addEventListener('input', (e) => {
            temperatureValue.textContent = e.target.value;
        });
    }
    
    if (thresholdSlider && thresholdValue) {
        thresholdSlider.addEventListener('input', (e) => {
            thresholdValue.textContent = e.target.value;
        });
    }
    
    // Charger les paramètres sauvegardés
    loadAISettings();
}

// Charger les paramètres IA sauvegardés
function loadAISettings() {
    const savedPrompt = localStorage.getItem('aiPrompt');
    const promptTextarea = document.getElementById('ai-prompt');
    
    if (savedPrompt && promptTextarea) {
        promptTextarea.value = savedPrompt;
    } else if (promptTextarea && !promptTextarea.value.trim()) {
        // Si aucun prompt sauvegardé et le textarea est vide, charger le prompt par défaut
        promptTextarea.value = CONFIG.DEFAULT_AI_PROMPT;
    }
    
    // Charger d'autres paramètres depuis localStorage si nécessaire
    const savedTemperature = localStorage.getItem('aiTemperature');
    const savedThreshold = localStorage.getItem('matchThreshold');
    const savedModel = localStorage.getItem('aiModel');
    
    if (savedTemperature) {
        const tempSlider = document.getElementById('temperature');
        const tempValue = document.getElementById('temperature-value');
        if (tempSlider && tempValue) {
            tempSlider.value = savedTemperature;
            tempValue.textContent = savedTemperature;
        }
    }
    
    if (savedThreshold) {
        const thresholdSlider = document.getElementById('match-threshold');
        const thresholdValue = document.getElementById('threshold-value');
        if (thresholdSlider && thresholdValue) {
            thresholdSlider.value = savedThreshold;
            thresholdValue.textContent = savedThreshold + '%';
        }
    }
    
    if (savedModel) {
        const modelSelect = document.getElementById('ai-model');
        if (modelSelect) {
            modelSelect.value = savedModel;
        }
    }
}

// Sauvegarder les paramètres IA
function saveAISettings() {
    const temperature = document.getElementById('temperature')?.value;
    const threshold = document.getElementById('match-threshold')?.value;
    const model = document.getElementById('ai-model')?.value;
    const prompt = document.getElementById('ai-prompt')?.value;
    
    if (temperature) localStorage.setItem('aiTemperature', temperature);
    if (threshold) localStorage.setItem('matchThreshold', threshold);
    if (model) localStorage.setItem('aiModel', model);
    if (prompt) localStorage.setItem('aiPrompt', prompt);
}

// Obtenir la configuration IA actuelle
function getAIConfig() {
    return {
        enabled: true, // IA toujours activée par défaut
        model: document.getElementById('ai-model')?.value || 'gpt-3.5-turbo',
        temperature: parseFloat(document.getElementById('temperature')?.value || '0.3'),
        maxTokens: parseInt(document.getElementById('max-tokens')?.value || '150'),
        timeout: parseInt(document.getElementById('timeout')?.value || '30'),
        threshold: parseInt(document.getElementById('match-threshold')?.value || '70'),
        retryAttempts: parseInt(document.getElementById('retry-attempts')?.value || '3'),
        batchDelay: parseInt(document.getElementById('batch-delay')?.value || '1000'),
        fallbackMatching: document.getElementById('fallback-matching')?.checked || true,
        caseSensitive: document.getElementById('case-sensitive')?.checked || false,
        prompt: document.getElementById('ai-prompt')?.value || ''
    };
}

// Fonctions d'optimisation pour le traitement par lots
function calculateOptimalBatchSize(totalRows, criteriaCount) {
    // Calculer la taille optimale du lot basée sur le nombre de lignes et de critères
    const baseBatchSize = CONFIG.PROCESSING.BATCH_SIZE || 10;
    
    // Réduire la taille du lot si beaucoup de critères
    const criteriaFactor = Math.max(1, criteriaCount / 3);
    
    // Ajuster selon le nombre total de lignes
    let adjustedBatchSize = Math.floor(baseBatchSize / criteriaFactor);
    
    // Limites min/max
    adjustedBatchSize = Math.max(2, Math.min(adjustedBatchSize, 20));
    
    return adjustedBatchSize;
}

function calculateAdaptiveDelay(apiCallCount, startTime, baseDelay) {
    // Calculer un délai adaptatif basé sur l'utilisation de l'API
    const elapsedTime = Date.now() - startTime;
    const callsPerSecond = apiCallCount / (elapsedTime / 1000);
    
    // Augmenter le délai si trop d'appels par seconde
    if (callsPerSecond > 5) {
        return baseDelay * 2;
    } else if (callsPerSecond > 3) {
        return baseDelay * 1.5;
    }
    
    return baseDelay;
}

async function processBatchWithRetry(batch, apiCallCount) {
    const maxRetries = 3;
    let retryCount = 0;
    
    while (retryCount < maxRetries) {
        try {
            const result = await processBatch(batch);
            return {
                matchedRows: result,
                apiCalls: batch.length * criteria.length // Estimation des appels API
            };
        } catch (error) {
            retryCount++;
            
            if (error.message.includes('rate limit') || error.message.includes('quota')) {
                // Délai exponentiel pour les erreurs de limite de taux
                const delay = Math.pow(2, retryCount) * 2000;
                showMessage(`Limite API atteinte, retry dans ${delay/1000}s...`, 'info', 3000);
                await sleep(delay);
            } else if (retryCount >= maxRetries) {
                throw error;
            } else {
                await sleep(1000 * retryCount);
            }
        }
    }
    
    throw new Error('Échec du traitement après plusieurs tentatives');
}

// Fonction pour renouveler le token d'accès
function renewAccessToken() {
    return new Promise((resolve, reject) => {
        if (tokenClient) {
            tokenClient.requestAccessToken({
                callback: (tokenResponse) => {
                    if (tokenResponse.access_token) {
                        accessToken = tokenResponse.access_token;
                        console.log('Token d\'accès renouvelé avec succès');
                        resolve(tokenResponse.access_token);
                    } else {
                        reject(new Error('Échec du renouvellement du token'));
                    }
                },
                error_callback: (error) => {
                    reject(new Error(`Erreur de renouvellement: ${error.error}`));
                }
            });
        } else {
            reject(new Error('Client de token non initialisé'));
        }
    });
}

// Fonction utilitaire pour effectuer des appels API avec gestion automatique du renouvellement de token
async function makeAuthenticatedRequest(url, options = {}) {
    if (!isSignedIn || !accessToken) {
        throw new Error('Utilisateur non connecté');
    }
    
    // Ajouter l'en-tête d'autorisation
    const requestOptions = {
        ...options,
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
            ...options.headers
        }
    };
    
    // Première tentative
    let response = await fetch(url, requestOptions);
    
    // Si erreur 401 (token expiré), tenter de renouveler le token
    if (response.status === 401) {
        console.log('Token expiré, tentative de renouvellement automatique...');
        try {
            await renewAccessToken();
            
            // Mettre à jour le token dans les en-têtes et réessayer
            requestOptions.headers['Authorization'] = `Bearer ${accessToken}`;
            response = await fetch(url, requestOptions);
        } catch (renewError) {
            console.error('Échec du renouvellement automatique du token:', renewError);
            throw new Error('Session expirée. Veuillez vous reconnecter.');
        }
    }
    
    return response;
}

// Exportation progressive par lots pour éviter la surcharge mémoire
async function exportProgressiveBatch(matchedRows, outputSpreadsheetId, outputSheetName, headerRow, isFirstBatch = false) {
    try {
        const exportConfig = CONFIG.EXPORT.PROGRESSIVE_EXPORT;
        
        if (!exportConfig.ENABLED) {
            // Fallback vers l'exportation classique
            return await exportEnrichedResults(matchedRows, outputSpreadsheetId, outputSheetName, headerRow);
        }
        
        // Préparer les données pour ce lot
        const exportData = [];
        
        // Ajouter les en-têtes seulement pour le premier lot
        if (isFirstBatch) {
            exportData.push(headerRow);
        }
        
        // Ajouter les données de ce lot
        matchedRows.forEach(row => {
            exportData.push(row.data);
        });
        
        // Déterminer la plage d'écriture et l'URL
        let url, method;
        if (isFirstBatch) {
            // Premier lot : effacer la feuille et écrire à partir de A1
            const clearUrl = `https://sheets.googleapis.com/v4/spreadsheets/${outputSpreadsheetId}/values/${encodeURIComponent(outputSheetName + '!A:ZZ')}:clear`;
            await makeAuthenticatedRequest(clearUrl, { method: 'POST' });
            
            const range = `${outputSheetName}!A1`;
            url = `https://sheets.googleapis.com/v4/spreadsheets/${outputSpreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`;
            method = 'PUT';
        } else {
            // Lots suivants : ajouter à la fin en utilisant append
            url = `https://sheets.googleapis.com/v4/spreadsheets/${outputSpreadsheetId}/values/${encodeURIComponent(outputSheetName)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
            method = 'POST';
        }
        const response = await makeAuthenticatedRequest(url, {
            method: method,
            body: JSON.stringify({
                values: exportData
            })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        console.log(`Lot exporté: ${matchedRows.length} lignes vers ${outputSheetName}`);
        return response;
        
    } catch (error) {
        console.error('Erreur lors de l\'exportation du lot:', error);
        throw new Error(`Échec de l'exportation du lot: ${error.message}`);
    }
}

// Fonction d'exportation classique (conservée pour compatibilité)
async function exportEnrichedResults(matchedRows, outputSpreadsheetId, outputSheetName, headerRow) {
    try {
        // Préparer les données avec seulement les données originales
        const exportData = [];
        
        // En-têtes originaux seulement
        exportData.push(headerRow);
        
        // Données originales seulement (copié-collé des lignes qui matchent)
        matchedRows.forEach(row => {
            exportData.push(row.data);
        });
        
        // Exporter vers Google Sheets avec gestion automatique du token
        const url = `https://sheets.googleapis.com/v4/spreadsheets/${outputSpreadsheetId}/values/${encodeURIComponent(outputSheetName + '!A1')}?valueInputOption=RAW`;
        
        const response = await makeAuthenticatedRequest(url, {
            method: 'PUT',
            body: JSON.stringify({
                values: exportData
            })
        });
        
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
        
        return response;
        
    } catch (error) {
        console.error('Erreur lors de l\'exportation:', error);
        throw new Error('Échec de l\'exportation des résultats');
    }
}