# Captures SAP vers Excel

Convertit des captures d'écran de factures SAP (F2 Sales Invoice) en fichiers Excel,
via l'API Claude (extraction vision) et SheetJS (génération du fichier).

## Structure

```
├── api/extract.ts        # Fonction serveur (Vercel Edge) : appelle Anthropic avec la clé API
├── src/InvoiceToExcel.tsx # Interface (upload, revue, génération du .xlsx)
├── src/main.tsx / index.css
└── index.html
```

La clé API ne circule jamais dans le navigateur : le frontend appelle `/api/extract`,
qui tourne sur le serveur Vercel et détient la clé.

## 1. Tester en local

```bash
npm install
cp .env.example .env.local   # puis renseigner ta clé ANTHROPIC_API_KEY
npm i -g vercel               # une seule fois
vercel dev                    # lance le site + la fonction /api/extract ensemble
```

(`npm run dev` seul ne suffit pas : il ne sait pas exécuter `api/extract.ts`. Utilise `vercel dev`
pour tester en local avec la fonction serveur active.)

## 2. Mettre sur GitHub

```bash
cd invoice-to-excel
git init
git add .
git commit -m "Initial commit"
```

Puis sur https://github.com/new : créer un dépôt vide (sans README/gitignore), sans le rendre public
si tu préfères le garder privé. Ensuite :

```bash
git branch -M main
git remote add origin https://github.com/<ton-compte>/<nom-du-repo>.git
git push -u origin main
```

## 3. Déployer sur Vercel

1. Va sur https://vercel.com et connecte-toi avec ton compte GitHub.
2. « Add New… » → « Project » → sélectionne le dépôt que tu viens de pousser.
3. Vercel détecte automatiquement Vite (Framework Preset : Vite). Laisse les réglages par défaut
   (Build Command `vite build`, Output Directory `dist`).
4. Avant de cliquer sur « Deploy », ouvre « Environment Variables » et ajoute :
   - `ANTHROPIC_API_KEY` = ta clé (créée sur https://console.anthropic.com → Settings → API Keys)
5. Clique sur « Deploy ». Au bout d'une minute, Vercel te donne une URL du type
   `https://ton-projet.vercel.app`.

À partir de là, chaque `git push` sur `main` redéploie automatiquement le site.

## 4. Obtenir une clé API Anthropic

- https://console.anthropic.com → Settings → API Keys → Create Key
- Un compte de facturation (carte bancaire) est nécessaire pour l'usage API — indépendant d'un
  abonnement claude.ai.

## Notes

- Le modèle utilisé par défaut est `claude-sonnet-5` (`api/extract.ts`). Tu peux le remplacer par
  `claude-haiku-4-5-20251001` (moins cher, un peu moins précis) si le volume de captures est élevé.
- Chaque appel d'extraction consomme des crédits API — vérifie ta facturation Anthropic si tu traites
  beaucoup de factures.
- Le fichier `.xlsx` généré reproduit la mise en forme du modèle d'origine (en-tête, formules
  `Total = P.U. × Qté`, ligne `TOTAL HT`).
