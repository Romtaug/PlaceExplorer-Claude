# 🌍 Place Explorer — GitHub Actions

Génère un guide de voyage Excel (31 catégories, top 20 lieux via Google Places) et l'envoie par **email HTML** depuis `romtaug@gmail.com`.

## 🚀 Installation (3 minutes)

### 1. Uploader les fichiers sur GitHub
Glisser-déposer **tout le contenu** de ce dossier dans ton repo (l'arborescence `.github/workflows/` doit être conservée) :

```
.github/workflows/place-explorer.yml
Image/logo.png
Image/travel.png
Image/qrcode.png
place_explorer.py
maps_route.py
README.md
```

> 💡 Sur Windows, le dossier `.github` peut être masqué dans l'explorateur. Si le drag-and-drop ne le prend pas, crée le fichier à la main sur GitHub : **Add file → Create new file** → nom : `.github/workflows/place-explorer.yml` → coller le contenu.

### 2. Créer les 2 secrets
Repo → **Settings → Secrets and variables → Actions → New repository secret** :

| Nom | Valeur |
|---|---|
| `GOOGLE_API_KEY` | Ta clé API Google Places |
| `GMAIL_APP_PASSWORD` | Ton mot de passe d'application Gmail (romtaug@gmail.com) |

⚠️ **Ne jamais mettre ces valeurs dans le code.** Si elles ont déjà été poussées dans un repo public, révoque-les et régénère-les :
- Clé API : https://console.cloud.google.com/ → API et services → Identifiants
- App Password : https://myaccount.google.com/apppasswords

### 3. Activer "Places API (New)" (obligatoire, 2 minutes)
Le script utilise la **nouvelle** API Places de Google (1 appel par catégorie au lieu de ~21) :
1. https://console.cloud.google.com/ → ton projet → **API et services → Bibliothèque**
2. Cherche **"Places API (New)"** → **Activer** (attention : c'est une API distincte de l'ancienne "Places API", les deux coexistent dans la bibliothèque)
3. Si ta clé a des restrictions d'API (page **Identifiants** → ta clé → *API restrictions*), ajoute **Places API (New)** à la liste autorisée (garde aussi Geocoding, Directions et Static Maps, toujours utilisées)

L'enrichissement **Wikipédia** (article + popularité, boost des incontournables dans le parcours) est 100% gratuit et **sans clé** — rien à faire. Désactivable avec la variable `WIKI_ENRICH=0`.

### 4. Lancer
Onglet **Actions → 🌍 Place Explorer → Run workflow** :
- **location** : `Lisbonne, Portugal` (format `Ville, Pays` ou `Pays`)
- **vacation_month** : `Août, 2026`
- **receiver_emails** : destinataires séparés par des virgules (défaut : romtaug@gmail.com)

Le mail part avec l'Excel en pièce jointe, et l'Excel est aussi disponible en **artifact** du run (30 jours).

## 📧 Le mail
- HTML stylé : logo en header, bannière voyage, étapes numérotées, liens utiles, prompt ChatGPT, QR code PayPal
- Fallback texte brut identique à la version d'origine
- Mêmes liens qu'avant (Skyscanner, Rome2Rio, Airbnb, Booking, Revolut parrainage, BordEuro...)

## ⏱ Durée et quota
32 catégories × **1 requête Text Search (New)** ≈ **32 appels API par run** (~1-2 min), soit ~20× moins que l'ancienne version : un usage personnel reste dans le palier gratuit mensuel du SKU. Les enrichissements Wikipédia (boost popularité) et OpenStreetMap (points de vue) sont gratuits et sans clé.

## 🧠 Intelligence du parcours
- Résultats Google en **français** (noms, adresses, descriptions)
- Sélection : 4 filets anti-parasites (catégorie, types Google, enseignes, statut) - fini les supermarchés et campings dans l'itinéraire
- Classement : note bayésienne adaptative × popularité Wikipédia, jours de fermeture affichés
- **Journées calibrées en temps** : durées de visite par type de lieu + marche estimée, budget réglable via `ITINERARY_DAY_MINUTES` (défaut 540 = 9h) ; autres réglages : `ITINERARY_QUALITY_PRIOR`, `ITINERARY_WIKI_BOOST`, `WIKI_ENRICH=0`, `OSM_VIEWPOINTS=0`
