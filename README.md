# ECN Revision Cockpit

Application full-stack pour suivre les revisions ECN avec comptes utilisateurs, sauvegarde automatique et avancement separe par utilisateur.

## Local

```bash
npm install
npm run dev:full
```

Puis ouvrir `http://127.0.0.1:5173`.

En local, l'app utilise SQLite dans `ecn-revisions.db`.

## Deploiement sans nom de domaine

Solution recommandee :

- Render pour heberger le site + API Node/Express.
- Neon pour une base PostgreSQL persistante.
- Pas besoin d'acheter un domaine : Render fournit une URL publique en `onrender.com`.

### 1. Creer la base Neon

1. Creer un projet sur Neon.
2. Copier la chaine de connexion PostgreSQL, aussi appelee `DATABASE_URL`.
3. Garder `sslmode=require` dans l'URL.

### 2. Mettre le projet sur GitHub

```bash
git init
git add .
git commit -m "Initial ECN revision app"
git branch -M main
git remote add origin YOUR_GITHUB_REPO_URL
git push -u origin main
```

### 3. Creer le service Render

1. Render > New > Web Service.
2. Connecter le repo GitHub.
3. Build command:

```bash
npm install && npm run build
```

4. Start command:

```bash
npm start
```

5. Ajouter les variables d'environnement :

```bash
DATABASE_URL=la_connection_string_neon
JWT_SECRET=un_long_secret_aleatoire
NODE_ENV=production
```

Render donnera une URL publique du type `https://ton-app.onrender.com`.

## Persistance

Chaque utilisateur a :

- son compte
- ses parametres
- son tableau de cours
- ses heures Pass1/Series
- ses reviews, flashcards, difficultes et remarques

Les modifications sont sauvegardees automatiquement vers la base via `/api/progress`.

## Notes

- Le fichier SQLite local et les logs sont ignores par Git.
- En production, utiliser PostgreSQL via `DATABASE_URL`; ne pas utiliser SQLite sur Render pour les vraies donnees.
- La dependance `xlsx` sert a importer/exporter Excel et garde une alerte npm connue sans correctif officiel.
