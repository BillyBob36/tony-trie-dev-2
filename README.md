# Tony Trie Dev - AI-Powered Google Sheets Matcher

🔍 Application web intelligente pour matcher des profils dans Google Sheets avec l'IA.

*Dernière mise à jour: Configuration des secrets pour le déploiement*

## 🚀 Déploiement sur GitHub Pages

### Prérequis
1. **Clés API configurées** dans les secrets GitHub :
   - `GOOGLE_CLIENT_ID` : Client ID Google OAuth
   - `GOOGLE_API_KEY` : Clé API Google pour Sheets/Drive
   - `OPENAI_API_KEY` : Clé API OpenAI

### Configuration des secrets GitHub
1. Allez dans **Settings** → **Secrets and variables** → **Actions**
2. Créez les 3 secrets mentionnés ci-dessus
3. Le déploiement se fera automatiquement via GitHub Actions

### Développement local
1. Clonez le repository
2. Modifiez `secrets.js` avec vos vraies clés API
3. Ouvrez `index.html` dans votre navigateur

### Fonctionnalités
- ✅ Authentification Google OAuth
- ✅ Lecture des Google Sheets
- ✅ Analyse IA des profils
- ✅ Interface utilisateur moderne
- ✅ Déploiement automatique

## 🔧 Technologies
- HTML5, CSS3, JavaScript
- Google Sheets API
- OpenAI API
- GitHub Actions
- GitHub Pages