# 📊 Application de Statistiques AVIA VOLT

Application web Flask pour centraliser et analyser les données de recharge des stations AVIA VOLT.

## 🚀 Fonctionnalités

- 📈 Tableaux de statistiques par jour et par station
- 📊 Graphiques interactifs
- 📥 Export Excel
- 📄 Génération de rapports Word et PowerPoint
- 🔄 Import automatique depuis API Bump et bases PostgreSQL

## 🛠️ Technologies

- **Backend** : Flask (Python)
- **Base de données** : SQLite + PostgreSQL
- **Frontend** : HTML, CSS (Bootstrap), JavaScript (Chart.js)
- **APIs** : Bump API, PostgreSQL

## 📦 Installation locale

1. Clonez le dépôt :
```bash
git clone https://github.com/VOTRE_USERNAME/appli-recharge-avia.git
cd appli-recharge-avia
```

2. Installez les dépendances :
```bash
pip install -r requirements.txt
```

3. Configurez les variables d'environnement (optionnel) :
```bash
export BUMP_APP_ID="votre_id"
export BUMP_APP_KEY="votre_cle"
export PG_HOST="localhost"
export PG_DATABASE="recharge_db"
# ... etc
```

4. Lancez l'application :
```bash
python app.py
```

5. Ouvrez http://localhost:5000

## 🌐 Déploiement sur Render.com

Voir le guide complet : [GUIDE_RENDER.md](GUIDE_RENDER.md)

## 📝 Variables d'environnement

| Variable | Description | Défaut |
|----------|-------------|--------|
| `BUMP_APP_ID` | ID application Bump API | - |
| `BUMP_APP_KEY` | Clé application Bump API | - |
| `PG_HOST` | Hôte PostgreSQL | localhost |
| `PG_PORT` | Port PostgreSQL | 5432 |
| `PG_DATABASE` | Nom base PostgreSQL | recharge_db |
| `PG_USER` | Utilisateur PostgreSQL | postgres |
| `PG_PASSWORD` | Mot de passe PostgreSQL | - |
| `PORT` | Port Flask (production) | 5000 |

## 📁 Structure du projet

```
.
├── app.py                    # Application Flask principale
├── config.py                 # Configuration centralisée
├── bump_import.py            # Import données Bump API
├── EC_import.py              # Import données PostgreSQL
├── requirements.txt          # Dépendances Python
├── Procfile                  # Config Render.com
├── config_stations.xlsx      # Configuration des stations
├── bump_data.db              # Base SQLite
├── index.html                # Page d'accueil
├── graphiques.html           # Page graphiques
└── resultats.html            # Page résultats
```

## 👥 Auteur

Application développée pour la gestion des stations AVIA VOLT.

## 📄 Licence

Usage interne uniquement.
