// Configuration de l'application Google Sheets AI Matcher
// Modifiez ces valeurs selon vos besoins

const APP_CONFIG = {
    // Configuration Google OAuth
    GOOGLE: {
        // Client ID Google
        CLIENT_ID: window.GOOGLE_CLIENT_ID || '802942339545-djuda331o72r3b1f2d6ra29e8hoqea56.apps.googleusercontent.com',
        
        // Clé API Google (requise pour l'accès aux APIs)
        API_KEY: window.GOOGLE_API_KEY || 'AIzaSyCa6k_Jl3XPn76UhlIe_4nasKU2v11-UWg',
        
        // Scopes nécessaires pour accéder aux Google Sheets
        SCOPES: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.readonly',
        
        // APIs Google à charger
        DISCOVERY_DOCS: [
            'https://sheets.googleapis.com/$discovery/rest?version=v4',
            'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest'
        ]
    },
    
    // Configuration OpenAI
    OPENAI: {
        // Clé API OpenAI
        API_KEY: window.OPENAI_API_KEY,
        
        // Modèle à utiliser
        MODEL: 'gpt-5-nano',
        
        // Paramètres du modèle
        MAX_TOKENS: 10,
        TEMPERATURE: 0,
        
        // URL de l'API
        API_URL: 'https://api.openai.com/v1/chat/completions'
    },
    
    // Configuration du traitement
    PROCESSING: {
        // Nombre de profils à traiter par lot
        // Réduisez si vous avez des timeouts, augmentez pour plus de vitesse
        BATCH_SIZE: 50,
        
        // Délai entre les lots (en millisecondes)
        BATCH_DELAY: 1000,
        
        // Nombre maximum de tentatives en cas d'erreur
        MAX_RETRIES: 3,
        
        // Délai entre les tentatives (en millisecondes)
        RETRY_DELAY: 2000
    },
    
    // Configuration de l'interface
    UI: {
        // Nombre maximum de lignes à afficher dans l'aperçu
        MAX_PREVIEW_ROWS: 100,
        
        // Nombre maximum de spreadsheets à afficher
        MAX_SPREADSHEETS_DISPLAY: 100,
        
        // Délai d'affichage des notifications (en millisecondes)
        NOTIFICATION_DURATION: 5000,
        
        // Animation des éléments
        ANIMATION_DURATION: 300
    },
    
    // Messages de l'application
    MESSAGES: {
        LOADING: {
            SPREADSHEETS: 'Chargement des fichiers Google Sheets...',
            SHEETS: 'Chargement des feuilles...',
            COLUMNS: 'Chargement des colonnes...',
            PROCESSING: 'Traitement en cours...'
        },
        ERROR: {
            NO_SPREADSHEETS_FOUND: 'Aucun fichier Google Sheets trouvé',
            SPREADSHEETS_LOAD_FAILED: 'Erreur lors du chargement des fichiers',
            API_NOT_ENABLED: 'L\'API Google Drive n\'est pas activée pour ce projet',
            AUTHENTICATION_FAILED: 'Erreur d\'authentification',
            NETWORK_ERROR: 'Erreur de réseau'
        },
        SUCCESS: {
            SPREADSHEETS_LOADED: 'Fichiers chargés avec succès',
            PROCESSING_COMPLETE: 'Traitement terminé'
        }
    },
    
    // Configuration de l'export
    EXPORT: {
        // Nom par défaut du fichier exporté
        DEFAULT_FILENAME: 'Résultats_Correspondances',
        
        // Format de date pour les noms de fichiers
        DATE_FORMAT: 'YYYY-MM-DD_HH-mm',
        
        // Exportation progressive par lots
        PROGRESSIVE_EXPORT: {
            // Activer l'exportation progressive
            ENABLED: true,
            
            // Nombre de résultats à exporter par lot
            BATCH_SIZE: 10,
            
            // Délai entre les exports de lots (en millisecondes)
            BATCH_DELAY: 500
        },
        
        // Couleurs pour les résultats
        COLORS: {
            MATCH: '#4CAF50',
            NO_MATCH: '#F44336',
            PARTIAL: '#FF9800'
        }
    },
    
    // Configuration du debug
    DEBUG: {
        // Activer les logs détaillés
        VERBOSE_LOGGING: false,
        
        // Afficher les requêtes API
        LOG_API_CALLS: false,
        
        // Simuler les erreurs (pour les tests)
        SIMULATE_ERRORS: false
    },
    
    // Prompt par défaut pour l'IA
    DEFAULT_AI_PROMPT: `Tu es un assistant IA spécialisé dans la correspondance de données. Ta tâche est de déterminer si des valeurs correspondent entre elles, en tenant compte des variations possibles.

Règles de correspondance :
1. Ignore les différences de casse (majuscules/minuscules)
2. Ignore les espaces supplémentaires au début et à la fin
3. Accepte les variations d'orthographe mineures (1-2 caractères différents)
4. Accepte les traductions entre français et anglais pour les termes courants
5. Accepte les abréviations courantes (ex: "Dev" pour "Développeur")

Exemples de correspondances valides :
- "Google" et "google" → OUI
- "Développeur" et "Developer" → OUI
- "Chef de projet" et "Project Manager" → OUI
- "Apple Inc." et "Apple" → OUI
- "Ingénieur logiciel" et "Software Engineer" → OUI
- "Microsoft Corp" et "Microsoft" → OUI

Exemples de non-correspondances :
- "Microsoft" et "Amazon" → NON
- "Développeur" et "Comptable" → NON
- "Paris" et "Londres" → NON

Réponds uniquement par "OUI" si les valeurs correspondent selon ces règles, ou "NON" si elles ne correspondent pas.`
};

// Export de la configuration
if (typeof module !== 'undefined' && module.exports) {
    module.exports = APP_CONFIG;
} else {
    // Pour le navigateur
    window.CONFIG = APP_CONFIG;
}