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

### 3. Lancer
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
31 catégories × ~21 requêtes ≈ **650 appels API Places par run** (~5-10 min). Surveille ton quota/facturation GCP.
