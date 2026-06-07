# -*- coding: utf-8 -*-
"""
PlaceExplorer — version GitHub Actions
- Lit la localisation et le mois depuis les variables d'environnement (inputs du workflow)
- Génère l'Excel multi-feuilles (31 catégories, top 20 lieux triés par nombre d'avis)
- Envoie un email HTML stylé depuis romtaug@gmail.com avec l'Excel (+ images si présentes) en pièce jointe

Secrets requis (GitHub > Settings > Secrets and variables > Actions) :
- GOOGLE_API_KEY      : clé API Google Places
- GMAIL_APP_PASSWORD  : mot de passe d'application Gmail de romtaug@gmail.com
"""

import os
import re
import sys
import time
import shutil
import smtplib
import unicodedata

import requests
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from openpyxl import load_workbook

# ----------------------------------------------------------------------------
# Configuration (depuis l'environnement — injecté par le workflow)
# ----------------------------------------------------------------------------
API_KEY = os.environ.get("GOOGLE_API_KEY", "")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
SENDER_EMAIL = "romtaug@gmail.com"
RECEIVER_EMAILS = [e.strip() for e in os.environ.get("RECEIVER_EMAILS", "romtaug@gmail.com").split(",") if e.strip()]

LOCATION = os.environ.get("LOCATION", "").strip()
VACATION_MONTH = os.environ.get("VACATION_MONTH", "").strip()

if not API_KEY:
    print("❌ GOOGLE_API_KEY manquant (secret GitHub).")
    sys.exit(1)
if not GMAIL_APP_PASSWORD:
    print("❌ GMAIL_APP_PASSWORD manquant (secret GitHub).")
    sys.exit(1)
if not LOCATION or not VACATION_MONTH:
    print("❌ LOCATION ou VACATION_MONTH manquant (inputs du workflow).")
    sys.exit(1)


def is_valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)


for email in RECEIVER_EMAILS:
    if not is_valid_email(email):
        print(f"❌ Adresse email invalide : {email}")
        sys.exit(1)
print("✅ Toutes les adresses email sont valides.")


# ----------------------------------------------------------------------------
# Normalisation
# ----------------------------------------------------------------------------
def normalize_location(location):
    """Force le format 'Ville, Pays' ou 'Pays'. Échoue proprement en CI si invalide."""
    location = " ".join(location.strip().split())
    if "," in location:
        parts = location.split(",")
        if len(parts) == 2:
            city, country = parts[0].strip(), parts[1].strip()
            if city and country:
                return f"{city.title()}, {country.title()}"
    elif location.replace(" ", "").isalpha():
        return location.title()
    print("❌ Format incorrect. Utilisez 'Ville, Pays' ou 'Pays' (ex : Paris, France).")
    sys.exit(1)


def normalize_filename(filename):
    return ''.join(
        c for c in unicodedata.normalize('NFD', filename)
        if unicodedata.category(c) != 'Mn'
    )


def extract_city_from_address(full_address):
    address_parts = full_address.split(", ")
    if "Unnamed Road" in full_address:
        city = address_parts[-2] if len(address_parts) > 2 else address_parts[-1]
    elif len(address_parts) > 2:
        city = address_parts[-2]
    elif len(address_parts) == 2:
        city = address_parts[0]
    else:
        city = full_address.split(" ")[0]
    return city


# ----------------------------------------------------------------------------
# Google Places
# ----------------------------------------------------------------------------
def get_place_details(place_id, api_key):
    details_url = "https://maps.googleapis.com/maps/api/place/details/json"
    details_params = {
        'place_id': place_id,
        'fields': 'name,formatted_address,rating,user_ratings_total,price_level,international_phone_number,website',
        'language': 'en',
        'key': api_key
    }
    response = requests.get(details_url, params=details_params, timeout=30)
    if response.status_code == 200:
        return response.json().get('result', {})
    print(f"Failed to fetch details for place_id: {place_id}")
    return None


def search_places(api_key, location, category):
    url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
    params = {'query': f'{category} in {location}', 'language': 'en', 'key': api_key}

    response = requests.get(url, params=params, timeout=30)
    if response.status_code != 200:
        print(f"Erreur {response.status_code} lors de la récupération des données pour {category}")
        return []
    payload = response.json()
    api_status = payload.get('status', 'UNKNOWN')
    if api_status not in ('OK', 'ZERO_RESULTS'):
        print(f"⚠️ API status '{api_status}' pour {category} : {payload.get('error_message', 'pas de détail')}")
        return []
    places = payload.get('results', [])[:20]  # top 20 de la 1ère page

    detailed_places = []
    for place in places:
        place_id = place.get('place_id')
        if not place_id:
            continue
        details = get_place_details(place_id, api_key)
        if details:
            full_address = details.get('formatted_address', 'Not specified')
            detailed_places.append({
                'City': extract_city_from_address(full_address),
                'Address': full_address,
                'Name': details.get('name', 'Not specified'),
                'Total Reviews': details.get('user_ratings_total', 0),
                'Rating (on 5)': details.get('rating', 'Not rated'),
                'Price Level': {0: "Free", 1: "+", 2: "++", 3: "+++", 4: "++++"}.get(
                    details.get('price_level', None), 'Not specified'),
                'Maps': f'=HYPERLINK("https://www.google.com/maps/place/?q=place_id:{place_id}", "📍 Google Maps")',
                'Phone': details.get('international_phone_number', 'Not available')
            })
        time.sleep(0.05)
    return detailed_places


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
def adjust_column_width(file_path):
    wb = load_workbook(file_path)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 2
    wb.save(file_path)



def style_workbook(file_path):
    """
    Stylise tout le classeur : en-têtes navy, lignes alternées, filtres,
    volet figé, liens Maps en bleu, formats de nombres, largeurs ajustées.
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    NAVY = "1F3B63"
    BAND = "F3F6FA"
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    band_fill = PatternFill(start_color=BAND, end_color=BAND, fill_type="solid")
    body_font = Font(name="Calibri", size=10.5, color="3C4043")
    link_font = Font(name="Calibri", size=10.5, color="1A73E8", underline="single", bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=False)
    thin_border = Border(bottom=Side(style="thin", color="E0E5EA"))

    wb = load_workbook(file_path)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row < 1:
            continue

        headers = [c.value for c in ws[1]]
        col_idx = {h: i + 1 for i, h in enumerate(headers)}

        # En-tête
        ws.row_dimensions[1].height = 24
        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center

        # Corps
        for r, row in enumerate(ws.iter_rows(min_row=2), start=2):
            for cell in row:
                cell.font = body_font
                cell.border = thin_border
                cell.alignment = left
                if r % 2 == 0:
                    cell.fill = band_fill
            # Colonnes spécifiques
            if 'Total Reviews' in col_idx:
                c = ws.cell(row=r, column=col_idx['Total Reviews'])
                c.number_format = '#,##0'
                c.alignment = center
            for name in ('Rating (on 5)', 'Price Level', 'City'):
                if name in col_idx:
                    ws.cell(row=r, column=col_idx[name]).alignment = center
            if 'Maps' in col_idx:
                c = ws.cell(row=r, column=col_idx['Maps'])
                c.font = link_font
                c.alignment = center

        # Volet figé + filtres
        ws.freeze_panes = "A2"
        if ws.max_row >= 2:
            ws.auto_filter.ref = ws.dimensions

        # Largeurs : basées sur le contenu, libellé fixe pour Maps, bornées
        for i, col in enumerate(ws.columns, start=1):
            letter = get_column_letter(i)
            header = headers[i - 1] if i - 1 < len(headers) else ""
            if header == 'Maps':
                ws.column_dimensions[letter].width = 18
                continue
            max_length = len(str(header or ""))
            for cell in col:
                if cell.row == 1 or cell.value is None:
                    continue
                max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[letter].width = max(12, min(max_length + 3, 55))

    wb.save(file_path)
    print(f"✅ Classeur stylisé : {file_path}")


def remove_invalid_rows(file_path, location):
    country = location.split(",")[-1].strip()
    print(f"Filtrage des lignes se terminant par : {country}")
    wb = load_workbook(file_path)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows_to_delete = []
        for row in ws.iter_rows(min_row=2):
            cell = row[1]  # colonne Address
            if cell.value and isinstance(cell.value, str):
                cell_value_normalized = " ".join(cell.value.split())
                if not cell_value_normalized.lower().endswith(country.lower()):
                    rows_to_delete.append(cell.row)
        for row_idx in sorted(set(rows_to_delete), reverse=True):
            ws.delete_rows(row_idx)
    wb.save(file_path)
    print(f"✅ Lignes se terminant par '{country}' conservées dans : {file_path}")


CATEGORIES = {
    # Découvertes
    'historical_sites': '🏰 Sites historiques',
    'museums': '🖼️ Musées',
    'churches': '⛪ Églises',
    'cultural_centers': '🎭 Centres culturels',
    'hiking_trails': '🥾 Sentiers de randonnée',
    # Restauration
    'restaurants': '🍴 Restaurants',
    'bars': '🍹 Bars',
    # Détente
    'parks': '🌳 Parcs',
    'beaches': '🏖️ Plages',
    'lakes': '🏞️ Lacs',
    # Divertissements
    'concert_halls': '🎶 Salles de concert',
    'nightclubs': '💃 Boîtes de nuit',
    'movie_theaters': '🎬 Cinémas',
    'stadiums': '🏟️ Stades',
    # Shopping
    'markets': '🌽 Marchés',
    'boutiques': '🛍️ Boutiques',
    'supermarkets': '🛒 Supermarchés',
    # Activités
    'festivals': '🎉 Festivals',
    'amusement_parks': "🎢 Parcs d'attractions",
    'zoos': '🐘 Zoos',
    'aquariums': '🐠 Aquariums',
    'mountain_resorts': '🏔️ Stations de montagne',
    # Sport
    'bike_rentals': '🚴 Locations de vélos',
    'campgrounds': '🏕️ Campings',
    'sports_centers': '🏋️‍♂️ Centres sportifs',
    'spas': '💆‍♀️ Spas',
    'gym': '🏋️‍♀️ Salles de sport',
    # Transports
    'train_stations': '🚆 Gares',
    'airports': '✈️ Aéroports',
    # Éducation
    'schools': '🏫 Écoles',
    # Santé
    'hospitals': '🏥 Hôpitaux',
}


def create_excel_file(api_key, location, vacation_month):
    normalize_location(location)  # validation
    location = " ".join(location.split())
    vacation_month = " ".join(vacation_month.split()).replace(",", "-").replace(" ", "")
    sanitized_location = location.replace(",", "-").replace(" ", "")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_name = f"{sanitized_location}_{vacation_month}.xlsx"
    file_path = os.path.join(script_dir, file_name)

    writer = pd.ExcelWriter(file_path, engine='openpyxl')
    for category, description in CATEGORIES.items():
        print(f"Fetching data for category: {category}")
        data = search_places(api_key, location, category)
        if data:
            df = pd.DataFrame(data).sort_values(by='Total Reviews', ascending=False)
            sheet_name = description if len(description) <= 31 else description[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.close()

    adjust_column_width(file_path)
    print(f"✅ Excel file created with clickable links: {file_path}")
    return file_path, location, vacation_month


# ----------------------------------------------------------------------------
# Email HTML (même contenu, même liens — juste stylé)
# ----------------------------------------------------------------------------
def build_email_bodies(location, location_cleaned, vacation_month_cleaned):
    ia_prompt = (
        f"Je cherche des expériences et activités extraordinaires à {location_cleaned} en {vacation_month_cleaned}, "
        "fais un top 20 des incontournables durant cette période (événements, activités, monuments, restaurants, quartiers) "
        "et un top 10 des villes à visiter autour avec le temps de trajet, indique les démarches administratives nécessaires "
        "(documents, visas, vaccins), les précautions à prendre (arnaques, numéros d'urgence), les coûts approximatifs, "
        "la météo moyenne, les événements locaux, et des astuces pour se déplacer, respecter les coutumes, et profiter au maximum."
    )

    import urllib.parse as _up
    claude_url_plain = "https://claude.ai/new?q=" + _up.quote(ia_prompt)

    # --- Version texte brut (fallback, identique à l'originale) ---
    body_plain = f"""Bienvenue à bord de Place Explorer !

Découvrez {location} ! C'est une destination rêvée pour des aventures inoubliables. Votre guide inclut :
    - Les lieux incontournables à visiter
    - Un plan d'organisation pour votre voyage

🗺 Étapes pour organiser votre voyage :
    1. Ouvrez le fichier Excel attaché avec Google Sheet en cliquant une fois dessus
    2. Consultez chaque feuille pour explorer les meilleurs options par catégorie
    3. Planifiez vos activités (par exemple 2/jours par feuille) sur : https://www.google.com/mymaps un calque par ville
    4. Si le réseau est payant à l’étranger, utilisez Google Maps hors connexion : https://support.google.com/maps/answer/6291838?hl=fr
    5. Envoyez-vous votre parcours du jour (en fonction de la proximité) sur WhatsApp ou via Google Docs pour le garder à portée de main sur votre téléphone : https://web.whatsapp.com, https://wa.me, ou https://docs.google.com

🌍 Liens utiles :
    - ✈ Pour trouver les vols les moins chers et obtenir des indemnisations en cas de retard : https://www.skyscanner.fr/ ou https://www.airhelp.com/fr/
    - 🚅 Pour comparer tous les moyens de transport : https://www.rome2rio.com/
    - 🏠 Pour réserver votre hébergement et véhicule : https://www.airbnb.com et https://www.booking.com
    - 🖍 Pour des avis et recommandations : https://www.tripadvisor.com
    - 🗺 Pour réserver des activités locales : https://www.getyourguide.fr/
    - 📞 Pour trouver des eSIM à moindre coût à l'étranger : https://www.airalo.com/fr
    - 💳 Pour dépenser sans aucuns frais de change et gagner 200 € à l'ouverture : https://revolut.com/referral/?referral-code=romainavh3!DEC1-24-VR-FR

🤖 Ouvrez ce lien pour lancer le prompt ci-dessous dans Claude (déjà pré-rempli, appuyez juste sur Entrée) : {claude_url_plain}\n\nOu copiez-le manuellement sur https://claude.ai :

"{ia_prompt}"

Nous espérons que vous passerez un moment incroyable. Bon voyage ! ✈
N'hésitez pas à faire un don via PayPal à l'adresse romtaug@gmail.com si cela vous a aidé.
Accédez à notre outil pour travailler à l'étranger : https://bordeuroconnect.netlify.app/"""

    # --- Version HTML professionnelle (mêmes liens, même contenu) ---
    NAVY = '#1f3b63'
    RED = '#d32f2f'
    TEXT = '#3c4043'
    MUTED = '#70757a'

    def btn(url, label, bg=NAVY):
        return (f'<a href="{url}" target="_blank" '
                f'style="display:inline-block;background-color:{bg};color:#ffffff;'
                f'padding:11px 22px;border-radius:6px;text-decoration:none;font-weight:600;'
                f'font-size:13px;letter-spacing:0.2px;">{label}</a>')

    def section_title(label, title):
        return (f'<p style="margin:0 0 2px;font-size:11px;font-weight:700;letter-spacing:1.5px;'
                f'text-transform:uppercase;color:{RED};">{label}</p>'
                f'<h2 style="margin:0 0 16px;font-size:19px;font-weight:700;color:{NAVY};">{title}</h2>')

    def step(num, text, buttons=''):
        badge = (f'<span style="display:inline-block;width:26px;height:26px;line-height:26px;'
                 f'border-radius:50%;background-color:{NAVY};color:#ffffff;text-align:center;'
                 f'font-size:13px;font-weight:700;">{num}</span>')
        extra = f'<div style="margin-top:8px;">{buttons}</div>' if buttons else ''
        return (f'<tr><td style="padding:10px 14px 10px 0;vertical-align:top;width:26px;">{badge}</td>'
                f'<td style="padding:12px 0;font-size:14px;color:{TEXT};line-height:1.6;">{text}{extra}</td></tr>')

    def useful(name, text, buttons):
        return (f'<tr><td style="padding:14px 0;border-bottom:1px solid #eceff1;">'
                f'<p style="margin:0 0 2px;font-size:14px;font-weight:700;color:{NAVY};">{name}</p>'
                f'<p style="margin:0 0 9px;font-size:13px;color:{TEXT};line-height:1.55;">{text}</p>'
                f'{buttons}</td></tr>')

    def small_btn(url, label):
        return (f'<a href="{url}" target="_blank" '
                f'style="display:inline-block;background-color:#ffffff;color:{NAVY};'
                f'border:1.5px solid {NAVY};padding:8px 16px;border-radius:6px;text-decoration:none;'
                f'font-weight:600;font-size:12.5px;margin:0 8px 6px 0;">{label} &rarr;</a>')

    preheader = (f"Votre guide de voyage personnalisé pour {location} : lieux incontournables, "
                 f"plan d'organisation et liens utiles.")

    # URL Claude avec le prompt pré-rempli dans la zone de saisie (le destinataire n'a plus qu'à appuyer sur Entrée)
    import urllib.parse
    claude_url = "https://claude.ai/new?q=" + urllib.parse.quote(ia_prompt)

    body_html = f"""\
<!DOCTYPE html>
<html lang="fr">
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background-color:#eef1f4;font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;">
<div style="display:none;max-height:0;overflow:hidden;mso-hide:all;">{preheader}</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#eef1f4;padding:32px 12px;">
<tr><td align="center">
<table role="presentation" width="620" cellpadding="0" cellspacing="0" style="max-width:620px;width:100%;background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 2px 10px rgba(31,59,99,0.10);">

  <!-- En-tête de marque -->
  <tr><td style="padding:28px 36px 20px;text-align:center;">
    <img src="cid:logo" alt="Place Explorer" width="72" style="display:block;margin:0 auto;max-width:72px;height:auto;">
    <p style="margin:10px 0 0;font-size:13px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:{NAVY};">Place&nbsp;Explorer</p>
  </td></tr>

  <!-- Bannière -->
  <tr><td style="padding:0;">
    <img src="cid:travel" alt="Préparation de voyage" width="620" style="display:block;width:100%;height:auto;">
  </td></tr>

  <!-- Titre -->
  <tr><td style="padding:32px 36px 6px;text-align:center;">
    <span style="display:inline-block;background-color:#fdecea;color:{RED};font-size:11.5px;font-weight:700;letter-spacing:1px;text-transform:uppercase;padding:6px 14px;border-radius:20px;margin-bottom:12px;">Guide voyage &middot; {vacation_month_cleaned.replace('-', ' ')}</span>
    <h1 style="margin:0 0 8px;font-size:25px;font-weight:700;color:{NAVY};line-height:1.3;">Votre guide de voyage pour {location}</h1>
    <p style="margin:0;font-size:15px;color:{TEXT};line-height:1.6;">
      Bienvenue à bord&nbsp;! Découvrez <strong>{location}</strong>, une destination rêvée pour des aventures inoubliables.
      Votre guide inclut les lieux incontournables à visiter et un plan d'organisation pour votre voyage.
    </p>
  </td></tr>

  <!-- Encart pièce jointe -->
  <tr><td style="padding:20px 36px 8px;">
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background-color:#f3f6fa;border:1px solid #dde5ee;border-radius:10px;">
      <tr>
        <td style="padding:16px 18px;font-size:24px;width:36px;vertical-align:middle;">📎</td>
        <td style="padding:16px 18px 16px 0;font-size:13.5px;color:{TEXT};line-height:1.55;vertical-align:middle;">
          <strong style="color:{NAVY};">Votre guide Excel est en pièce jointe.</strong><br>
          Plus de 30 catégories (monuments, restaurants, plages, musées…), top 20 des lieux par popularité, liens Google&nbsp;Maps cliquables.
        </td>
      </tr>
    </table>
  </td></tr>

  <!-- Étapes -->
  <tr><td style="padding:26px 36px 8px;">
    {section_title('Organisation', 'Étapes pour organiser votre voyage')}
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%">
      {step('1', "Ouvrez le fichier Excel attaché avec <strong>Google&nbsp;Sheets</strong> en cliquant une fois dessus.")}
      {step('2', "Consultez chaque feuille pour explorer les meilleures options par catégorie.")}
      {step('3', "Planifiez vos activités (par exemple 2 par jour et par feuille), avec un calque par ville&nbsp;:",
            small_btn('https://www.google.com/mymaps', 'Google My Maps'))}
      {step('4', "Si le réseau est payant à l'étranger, utilisez Google Maps hors connexion&nbsp;:",
            small_btn('https://support.google.com/maps/answer/6291838?hl=fr', 'Maps hors connexion'))}
      {step('5', "Envoyez-vous votre parcours du jour (en fonction de la proximité) pour le garder à portée de main sur votre téléphone&nbsp;:",
            small_btn('https://web.whatsapp.com', 'WhatsApp Web') +
            small_btn('https://wa.me', 'wa.me') +
            small_btn('https://docs.google.com', 'Google Docs'))}
    </table>
  </td></tr>

  <!-- Liens utiles -->
  <tr><td style="padding:26px 36px 8px;">
    {section_title('Ressources', 'Liens utiles pour votre voyage')}
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%">
      {useful('Vols', "Trouvez les vols les moins chers et obtenez des indemnisations en cas de retard.",
              small_btn('https://www.skyscanner.fr/', 'Skyscanner') + small_btn('https://www.airhelp.com/fr/', 'AirHelp'))}
      {useful('Transports', "Comparez tous les moyens de transport pour chaque trajet.",
              small_btn('https://www.rome2rio.com/', 'Rome2Rio'))}
      {useful('Hébergement &amp; véhicule', "Réservez votre logement et votre véhicule en quelques clics.",
              small_btn('https://www.airbnb.com', 'Airbnb') + small_btn('https://www.booking.com', 'Booking.com'))}
      {useful('Avis &amp; recommandations', "Consultez les avis des voyageurs avant de choisir.",
              small_btn('https://www.tripadvisor.com', 'Tripadvisor'))}
      {useful('Activités locales', "Réservez des visites et expériences sur place.",
              small_btn('https://www.getyourguide.fr/', 'GetYourGuide'))}
      {useful("Connexion à l'étranger", "Trouvez des eSIM à moindre coût pour rester connecté.",
              small_btn('https://www.airalo.com/fr', 'Airalo'))}
    </table>
  </td></tr>

  <!-- Offre Revolut -->
  <tr><td style="padding:24px 36px 8px;">
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background-color:{NAVY};border-radius:12px;">
      <tr><td style="padding:24px 26px;text-align:center;">
        <p style="margin:0 0 4px;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#9fb3d1;">Offre partenaire</p>
        <p style="margin:0 0 6px;font-size:17px;font-weight:700;color:#ffffff;">Dépensez sans aucuns frais de change</p>
        <p style="margin:0 0 14px;font-size:13.5px;color:#d7e0ec;line-height:1.55;">et gagnez 200&nbsp;€ à l'ouverture de votre compte.</p>
        {btn('https://revolut.com/referral/?referral-code=romainavh3!DEC1-24-VR-FR', 'Ouvrir un compte Revolut', RED)}
      </td></tr>
    </table>
  </td></tr>

  <!-- Prompt Claude -->
  <tr><td style="padding:26px 36px 8px;">
    {section_title('Assistant IA', 'Enrichissez votre expérience')}
    <p style="margin:0 0 12px;font-size:14px;color:{TEXT};line-height:1.6;">
      Cliquez sur le bouton ci-dessous&nbsp;: <strong>Claude</strong> s'ouvre avec ce prompt déjà saisi, il ne reste qu'à appuyer sur Entrée pour obtenir un programme complet et personnalisé&nbsp;:
    </p>
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background-color:#f8f9fa;border-left:4px solid {NAVY};border-radius:0 8px 8px 0;">
      <tr><td style="padding:16px 18px;">
        <p style="margin:0;font-size:12.5px;color:#444746;line-height:1.65;font-family:Consolas,'Courier New',monospace;">"{ia_prompt}"</p>
      </td></tr>
    </table>
    <div style="text-align:center;margin-top:16px;">
      {btn(claude_url, 'Ouvrir Claude — prompt pré-rempli')}
    </div>
  </td></tr>

  <!-- Bon voyage -->
  <tr><td style="padding:28px 36px 24px;text-align:center;">
    <p style="margin:0;font-size:15px;color:{NAVY};font-weight:700;">Nous espérons que vous passerez un moment incroyable. Bon voyage&nbsp;!</p>
  </td></tr>

  <!-- Pied de page -->
  <tr><td style="padding:26px 36px 30px;background-color:#f7f9fb;border-top:1px solid #e8ecf1;text-align:center;">
    <p style="margin:0 0 10px;font-size:13px;color:{TEXT};line-height:1.6;">
      Ce guide vous a été utile&nbsp;? Soutenez le projet d'un don via <strong>PayPal</strong>
      à l'adresse <a href="mailto:romtaug@gmail.com" style="color:{NAVY};font-weight:600;text-decoration:none;">romtaug@gmail.com</a>
      ou en scannant ce QR&nbsp;code&nbsp;:
    </p>
    <img src="cid:qrcode" alt="QR code don PayPal" width="110" style="display:block;margin:8px auto 14px;max-width:110px;height:auto;border-radius:8px;border:1px solid #e0e5ea;">
    <p style="margin:0 0 10px;font-size:13px;color:{TEXT};">Découvrez aussi notre outil pour travailler à l'étranger&nbsp;:</p>
    {small_btn('https://bordeuroconnect.netlify.app/', 'BordEuro Connect')}
    <p style="margin:18px 0 0;font-size:11.5px;color:{MUTED};line-height:1.6;">
      Vous recevez cet e-mail car un guide Place&nbsp;Explorer a été généré pour vous.<br>
      Place&nbsp;Explorer — Guide de voyage automatisé
    </p>
  </td></tr>

</table>
</td></tr>
</table>
</body>
</html>"""

    return body_plain, body_html


def send_email_with_excel(sender_email, password, receiver_emails, subject,
                          body_plain, body_html, file_path, image_paths):
    msg = MIMEMultipart('mixed')
    msg['From'] = sender_email
    msg['To'] = ", ".join(receiver_emails)
    msg['Subject'] = subject

    # Corps : alternative (texte brut + HTML avec images inline)
    alt = MIMEMultipart('alternative')
    alt.attach(MIMEText(body_plain, 'plain', 'utf-8'))

    related = MIMEMultipart('related')
    related.attach(MIMEText(body_html, 'html', 'utf-8'))

    # Images intégrées au corps HTML (cid:logo, cid:travel, cid:qrcode)
    for image_path in image_paths:
        cid = os.path.splitext(os.path.basename(image_path))[0]  # logo / travel / qrcode
        if os.path.exists(image_path):
            with open(image_path, 'rb') as img:
                img_part = MIMEImage(img.read())
            img_part.add_header('Content-ID', f'<{cid}>')
            img_part.add_header('Content-Disposition', 'inline',
                                filename=os.path.basename(image_path))
            related.attach(img_part)
        else:
            print(f"Image non trouvée (ignorée) : {image_path}")

    alt.attach(related)
    msg.attach(alt)

    # Pièce jointe Excel
    try:
        with open(file_path, 'rb') as file:
            part = MIMEBase('application', "octet-stream")
            part.set_payload(file.read())
        encoders.encode_base64(part)
        normalized_name = normalize_filename(os.path.basename(file_path))
        part.add_header('Content-Disposition', f'attachment; filename="{normalized_name}"')
        msg.attach(part)
    except FileNotFoundError:
        print(f"Fichier non trouvé : {file_path}")
        return

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_emails, msg.as_string())
            print(f"✅ E-mail envoyé avec succès à {', '.join(receiver_emails)}")
    except Exception as e:
        print(f"❌ Erreur lors de l'envoi de l'e-mail : {e}")
        sys.exit(1)


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
if __name__ == "__main__":
    try:
        excel_file, location, vacation_month = create_excel_file(API_KEY, LOCATION, VACATION_MONTH)
        remove_invalid_rows(excel_file, location)
        style_workbook(excel_file)
    except SystemExit:
        raise
    except Exception as e:
        print(f"❌ Erreur lors de la création du fichier Excel : {e}")
        sys.exit(1)

    location_cleaned = location.replace(",", "-").replace(" ", "")
    vacation_month_cleaned = vacation_month.replace(" ", "-")

    if not os.path.exists(excel_file):
        print(f"❌ Le fichier Excel n'a pas été trouvé : {excel_file}")
        sys.exit(1)

    # Garde-fou : vérifier que l'Excel contient des données avant d'envoyer
    wb_check = load_workbook(excel_file, read_only=True)
    total_rows = 0
    for sheet_name in wb_check.sheetnames:
        n = sum(1 for _ in wb_check[sheet_name].iter_rows(min_row=2))
        total_rows += n
        print(f"   📄 {sheet_name} : {n} lieux")
    wb_check.close()
    print(f"📊 Total : {total_rows} lieux dans {len(wb_check.sheetnames)} feuilles")
    if total_rows == 0:
        print("❌ Le fichier Excel est vide (0 lieu) — envoi annulé.")
        print("   Causes probables : statut API en erreur (voir ⚠️ ci-dessus) ou filtre pays trop strict.")
        sys.exit(1)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    image_paths = [
        os.path.join(script_dir, "Image", "logo.png"),
        os.path.join(script_dir, "Image", "travel.png"),
        os.path.join(script_dir, "Image", "qrcode.png"),
    ]

    subject = f"🌍 PlaceExplorer : Les Meilleurs Lieux Destination {location_cleaned} en {vacation_month_cleaned}"
    body_plain, body_html = build_email_bodies(location, location_cleaned, vacation_month_cleaned)

    send_email_with_excel(SENDER_EMAIL, GMAIL_APP_PASSWORD, RECEIVER_EMAILS,
                          subject, body_plain, body_html, excel_file, image_paths)

    # Déplacer l'Excel dans Content/ (récupéré ensuite comme artifact par le workflow)
    content_dir = os.path.join(script_dir, "Content")
    os.makedirs(content_dir, exist_ok=True)
    destination = os.path.join(content_dir, os.path.basename(excel_file))
    shutil.move(excel_file, destination)
    print(f"✅ Fichier déplacé dans le dossier Content : {destination}")
