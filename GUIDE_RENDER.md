# 🚀 Guide de déploiement sur Render.com

## ✅ Pourquoi Render.com ?

- ✅ **PostgreSQL gratuit inclus** (1 GB)
- ✅ Déploiement automatique depuis GitHub
- ✅ HTTPS automatique
- ✅ 750h/mois gratuit
- ✅ Variables d'environnement sécurisées
- ✅ Logs en temps réel

---

## 📦 Étape 1 : Préparer les fichiers

Vous devez avoir ces fichiers dans votre dossier :

```
appli_recharge/
├── app.py
├── requirements.txt    ← ✅ Déjà créé
├── Procfile           ← ✅ Déjà créé
├── .gitignore         ← ✅ Déjà créé
├── bump_import.py
├── EC_import.py
├── config_stations.xlsx
├── bump_data.db
├── index.html
├── graphiques.html
└── resultats.html
```

---

## 🐙 Étape 2 : Mettre votre code sur GitHub

### 2.1 Créer un dépôt GitHub

1. Allez sur https://github.com
2. Connectez-vous (ou créez un compte gratuit)
3. Cliquez sur le **"+"** en haut à droite → **"New repository"**
4. Nommez-le : `appli-recharge-avia`
5. **Laissez en "Private"** si vos données sont sensibles
6. Cliquez sur **"Create repository"**

### 2.2 Pousser votre code sur GitHub

**Sur votre PC, ouvrez PowerShell ou CMD** dans le dossier de votre application :

```bash
cd "C:\Users\camus\Desktop\Appli data recharge"

# Initialiser Git
git init

# Ajouter tous les fichiers
git add .

# Créer un premier commit
git commit -m "Premier commit - Application de statistiques AVIA VOLT"

# Lier à votre dépôt GitHub (remplacez VOTRE_USERNAME)
git remote add origin https://github.com/VOTRE_USERNAME/appli-recharge-avia.git

# Pousser le code
git push -u origin main
```

⚠️ **Si vous avez des erreurs** :
```bash
# Si Git n'est pas installé, installez-le : https://git-scm.com/download/win

# Si la branche s'appelle "master" au lieu de "main" :
git branch -M main
git push -u origin main
```

---

## 🌐 Étape 3 : Créer un compte Render.com

1. Allez sur https://render.com
2. Cliquez sur **"Get Started"**
3. Connectez-vous avec **GitHub** (recommandé)
4. Autorisez Render à accéder à vos dépôts

---

## 🗄️ Étape 4 : Créer une base de données PostgreSQL (Optionnel)

**⚠️ IMPORTANT :** Si vous avez déjà un PostgreSQL quelque part (serveur entreprise, PC local), vous pouvez sauter cette étape et utiliser votre base existante.

### Si vous voulez créer une nouvelle base PostgreSQL sur Render :

1. Dans le dashboard Render, cliquez sur **"New +"**
2. Sélectionnez **"PostgreSQL"**
3. Configurez :
   - **Name** : `appli-recharge-db`
   - **Database** : `recharge_db`
   - **User** : `recharge_user`
   - **Region** : Choisissez le plus proche (ex: Frankfurt)
   - **Plan** : **Free** (1 GB, expire après 90 jours)
4. Cliquez sur **"Create Database"**

⏳ Attendez 1-2 minutes que la base soit créée.

5. Une fois créée, notez les informations de connexion :
   - **Internal Database URL** (pour votre app sur Render)
   - **External Database URL** (pour vous connecter depuis votre PC)

---

## 🚀 Étape 5 : Déployer l'application Web

1. Dans le dashboard Render, cliquez sur **"New +"**
2. Sélectionnez **"Web Service"**
3. Connectez votre dépôt GitHub :
   - Recherchez `appli-recharge-avia`
   - Cliquez sur **"Connect"**

4. Configurez le service :
   - **Name** : `appli-recharge-avia`
   - **Region** : Même que votre base de données
   - **Branch** : `main`
   - **Root Directory** : (laissez vide)
   - **Runtime** : `Python 3`
   - **Build Command** : `pip install -r requirements.txt`
   - **Start Command** : `gunicorn app:app`
   - **Plan** : **Free** (750h/mois)

5. Cliquez sur **"Create Web Service"**

⏳ **Attendez 3-5 minutes** que l'application se déploie.

---

## 🔐 Étape 6 : Configurer les variables d'environnement

Si vous utilisez PostgreSQL, vous devez configurer les credentials de manière sécurisée.

### 6.1 Dans Render.com :

1. Dans votre Web Service, allez dans **"Environment"**
2. Cliquez sur **"Add Environment Variable"**
3. Ajoutez ces variables :

```
# PostgreSQL (si vous utilisez celui de Render)
DATABASE_URL=<copiez l'Internal Database URL de votre base Render>

# OU si vous utilisez votre propre PostgreSQL :
PG_HOST=192.168.1.XXX
PG_PORT=5432
PG_DATABASE=nom_base
PG_USER=utilisateur
PG_PASSWORD=mot_de_passe

# Bump API
BUMP_APP_ID=2727ac86-4f15-4994-91a5-172e3006ee7b
BUMP_APP_KEY=6ad38fdc-c5dc-4f7e-af06-3e44b0c59ea5
```

4. Cliquez sur **"Save Changes"**

### 6.2 Modifier votre code pour utiliser les variables d'environnement :

Dans `EC_import.py`, remplacez :
```python
PG_CONFIG = {
    'host': '192.168.1.50',
    'port': 5432,
    'database': 'recharge_db',
    'user': 'admin',
    'password': 'MonMotDePasse123'
}
```

Par :
```python
import os

PG_CONFIG = {
    'host': os.environ.get('PG_HOST', 'localhost'),
    'port': int(os.environ.get('PG_PORT', 5432)),
    'database': os.environ.get('PG_DATABASE', 'recharge_db'),
    'user': os.environ.get('PG_USER', 'postgres'),
    'password': os.environ.get('PG_PASSWORD', '')
}
```

Dans `bump_import.py`, remplacez :
```python
APPLICATION_ID = "2727ac86-4f15-4994-91a5-172e3006ee7b"
APPLICATION_KEY = "6ad38fdc-c5dc-4f7e-af06-3e44b0c59ea5"
```

Par :
```python
import os

APPLICATION_ID = os.environ.get('BUMP_APP_ID', "2727ac86-4f15-4994-91a5-172e3006ee7b")
APPLICATION_KEY = os.environ.get('BUMP_APP_KEY', "6ad38fdc-c5dc-4f7e-af06-3e44b0c59ea5")
```

### 6.3 Pousser les modifications sur GitHub :

```bash
git add .
git commit -m "Ajout variables d'environnement"
git push
```

⏳ Render va **automatiquement redéployer** votre application !

---

## 🎯 Étape 7 : Tester l'application

1. Dans Render, allez dans votre Web Service
2. En haut, vous verrez l'URL de votre app : `https://appli-recharge-avia.onrender.com`
3. Cliquez dessus pour ouvrir votre application !

🎉 **Votre application est en ligne !**

---

## 🔄 Étape 8 : Mettre à jour l'application

À chaque fois que vous modifiez votre code :

```bash
cd "C:\Users\camus\Desktop\Appli data recharge"

# Ajouter les modifications
git add .

# Créer un commit
git commit -m "Description de vos modifications"

# Pousser sur GitHub
git push
```

⚡ **Render détecte automatiquement** les changements et redéploie !

---

## 📊 Étape 9 : Initialiser la base de données

Si vous utilisez PostgreSQL sur Render, vous devez créer les tables :

### Option A : Via Shell Render

1. Dans votre Web Service, allez dans **"Shell"**
2. Exécutez :
```bash
python EC_import.py
python bump_import.py
```

### Option B : Depuis votre PC

1. Connectez-vous avec l'**External Database URL** :
```bash
psql <External_Database_URL>
```
2. Créez vos tables manuellement ou importez vos données

---

## ⚠️ Limitations du plan gratuit

### Web Service :
- 🕐 **750 heures/mois** (~31 jours)
- ⏸️ **Se met en veille** après 15 minutes d'inactivité
- ⏱️ **Redémarre en 30-60 secondes** à la première visite
- 💾 **512 MB RAM**
- 🌐 Domaine : `votre-app.onrender.com`

### PostgreSQL :
- 💾 **1 GB de stockage**
- 📅 **Expire après 90 jours** (vous devrez recréer une nouvelle base)
- 🔄 **Pas de backups automatiques**

### 💡 Astuce pour éviter la mise en veille :
Utilisez un service comme **UptimeRobot** (gratuit) pour pinger votre app toutes les 5 minutes.

---

## 🐛 Résolution des problèmes

### Problème 1 : "Application failed to start"

**Vérifiez les logs** :
1. Dans Render → votre Web Service → **"Logs"**
2. Lisez les erreurs

**Causes courantes** :
- Dépendance manquante dans `requirements.txt`
- Erreur de syntaxe Python
- Port mal configuré (doit être dynamique sur Render)

**Solution** : Assurez-vous que `app.py` utilise le port dynamique :
```python
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
```

### Problème 2 : "Database connection failed"

**Vérifiez** :
- Les variables d'environnement sont bien configurées
- L'URL de la base de données est correcte
- La base PostgreSQL est bien démarrée

### Problème 3 : "Build failed"

**Solution** :
- Vérifiez que `requirements.txt` est correct
- Essayez de spécifier les versions exactes des packages
- Vérifiez les logs de build dans Render

### Problème 4 : Application lente au premier chargement

**C'est normal !** Le plan gratuit se met en veille. Solutions :
- Attendez 30-60 secondes
- Utilisez UptimeRobot pour garder l'app active
- Passez au plan payant ($7/mois) pour éliminer la mise en veille

---

## 💰 Passer au plan payant (Optionnel)

Si vous voulez :
- ❌ Pas de mise en veille
- ✅ Plus de RAM (2 GB+)
- ✅ Base de données permanente
- ✅ Backups automatiques

**Coût** : ~$7-25/mois selon vos besoins

---

## 🔐 Sécurité

### ⚠️ IMPORTANT : Protéger vos données

1. **Gardez votre dépôt GitHub en PRIVATE**
2. **N'incluez JAMAIS** :
   - Mots de passe en clair dans le code
   - Clés API directement dans les fichiers
   - Base de données avec données sensibles

3. **Utilisez toujours les variables d'environnement** pour les credentials

4. **Ajoutez une authentification** si l'app contient des données sensibles :
```python
from flask_httpauth import HTTPBasicAuth
auth = HTTPBasicAuth()

@auth.verify_password
def verify_password(username, password):
    if username == 'admin' and password == 'votre_mot_de_passe':
        return True
    return False

@app.route('/')
@auth.login_required
def index():
    # ...
```

---

## 📝 Checklist finale

Avant de déployer :
- [ ] `requirements.txt` contient toutes les dépendances
- [ ] `Procfile` existe avec `web: gunicorn app:app`
- [ ] `.gitignore` exclut les fichiers sensibles
- [ ] Variables d'environnement configurées dans Render
- [ ] Code poussé sur GitHub
- [ ] PostgreSQL créé et initialisé (si nécessaire)
- [ ] Application testée localement

---

## 🎯 Résultat final

Votre application sera accessible sur :
```
https://appli-recharge-avia.onrender.com
```

Partagez ce lien avec vos collègues ! 🚀

---

## 📞 Support

- Documentation Render : https://render.com/docs
- Forum communautaire : https://community.render.com/
- Status page : https://status.render.com/

---

## 🚀 Prochaines étapes

1. **Configurez les backups** de votre base de données
2. **Ajoutez un nom de domaine personnalisé** (payant)
3. **Configurez des alertes** pour surveiller votre app
4. **Optimisez les performances** en analysant les logs

Bonne chance avec votre déploiement ! 🍀
