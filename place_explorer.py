# -*- coding: utf-8 -*-
"""
Itinéraire d'UNE JOURNÉE à pied pour PlaceExplorer — un seul lien Google Maps.

Logique :
  - top des lieux les plus populaires (nb d'avis), HORS restaurants/bars/clubs
    et hors catégories non-touristiques (aéroports, gares, hôpitaux, écoles,
    supermarchés, salles de sport) ;
  - ordre optimisé pour minimiser la marche (nearest-neighbor + 2-opt, comme
    thermodata_engine) ;
  - déjeuner = resto le plus populaire (mi-parcours), dîner = 2e resto le plus
    populaire (fin), bar/club = le plus populaire (après) ;
  - 1 seul lien Maps (format api=1, travelmode=walking) ;
  - distance + temps de marche RÉELS via la Directions API (1 seul appel).

Nécessite 'Lat'/'Lng'/'PlaceID' sur chaque lieu (cf. intégration).
"""

import math
import urllib.parse
import requests

# Nombre de visites par défaut, réparties moitié matin / moitié après-midi
# autour du déjeuner. Pas de plafond imposé : monte ce chiffre si tu veux.
N_VISITS = 10

FOOD_CATEGORIES = {"restaurants"}
BAR_CATEGORIES = {"bars", "nightclubs"}            # bar OU club, le plus populaire
NON_TOURIST = {"airports", "train_stations", "hospitals",
               "schools", "supermarkets", "gym"}


# ── Géométrie (reprise de thermodata_engine) ──────────────────────
def _coords(s):
    lat, lng = s.get("Lat"), s.get("Lng")
    if lat is None or lng is None:
        return None
    return (float(lat), float(lng))


def hav(la1, lo1, la2, lo2):
    R = 6371000.0
    p1, p2 = math.radians(la1), math.radians(la2)
    dp, dl = math.radians(la2 - la1), math.radians(lo2 - lo1)
    x = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * R * math.asin(math.sqrt(x))


def _route_len(stops):
    tot = 0.0
    for i in range(len(stops) - 1):
        a, b = _coords(stops[i]), _coords(stops[i + 1])
        if a and b:
            tot += hav(a[0], a[1], b[0], b[1])
    return tot


# ── Optimisation d'ordre : nearest-neighbor + 2-opt ───────────────
def _nearest_neighbor(stops, start):
    rest = set(range(len(stops)))
    order = [start]
    rest.discard(start)
    while rest:
        la, lo = _coords(stops[order[-1]])
        nxt = min(rest, key=lambda j: hav(la, lo, *_coords(stops[j])))
        order.append(nxt)
        rest.discard(nxt)
    return [stops[i] for i in order]


def _two_opt(route):
    best, improved = route[:], True
    while improved:
        improved = False
        for i in range(1, len(best) - 1):
            for k in range(i + 1, len(best)):
                if k - i == 1:
                    continue
                cand = best[:i] + best[i:k + 1][::-1] + best[k + 1:]
                if _route_len(cand) < _route_len(best) - 1e-6:
                    best, improved = cand, True
    return best


def optimize_walking_order(stops, start=0):
    stops = [s for s in stops if _coords(s)]
    if len(stops) <= 2:
        return stops
    return _two_opt(_nearest_neighbor(stops, start))


# ── Sélection ─────────────────────────────────────────────────────
def _top(places, k=None):
    r = sorted(places, key=lambda d: d.get("Total Reviews", 0) or 0, reverse=True)
    return r[:k] if k else r


# ── Lien Google Maps (api=1, marche, place_id fiables) ────────────
def _loc(s):
    return s.get("Address") or s.get("Name") or ""


def build_route_url(stops, travelmode="walking"):
    stops = [s for s in stops if _loc(s)]
    if len(stops) < 2:
        return None
    o, d, mid = stops[0], stops[-1], stops[1:-1]
    p = {"api": "1", "origin": _loc(o), "destination": _loc(d), "travelmode": travelmode}
    if o.get("PlaceID"):
        p["origin_place_id"] = o["PlaceID"]
    if d.get("PlaceID"):
        p["destination_place_id"] = d["PlaceID"]
    if mid:
        p["waypoints"] = "|".join(_loc(s) for s in mid)
        if all(s.get("PlaceID") for s in mid):
            p["waypoint_place_ids"] = "|".join(s["PlaceID"] for s in mid)
    return "https://www.google.com/maps/dir/?" + urllib.parse.urlencode(p, quote_via=urllib.parse.quote_plus)


# ── Distance + temps + tracé RÉELS via Directions API (1 appel) ───
def walking_distance_time(stops, api_key):
    """Retourne (distance_m, duree_s, polyline_encodee) réels à pied,
    ou (None, None, None) si échec."""
    stops = [s for s in stops if _coords(s)]
    if len(stops) < 2 or not api_key:
        return None, None, None

    def ref(s):
        return f"place_id:{s['PlaceID']}" if s.get("PlaceID") else f"{s['Lat']},{s['Lng']}"

    params = {
        "origin": ref(stops[0]),
        "destination": ref(stops[-1]),
        "mode": "walking",
        "key": api_key,
    }
    if len(stops) > 2:
        params["waypoints"] = "|".join(ref(s) for s in stops[1:-1])
    try:
        r = requests.get("https://maps.googleapis.com/maps/api/directions/json",
                         params=params, timeout=30)
        data = r.json()
        if data.get("status") != "OK" or not data.get("routes"):
            print(f"⚠️ Directions API : {data.get('status')} {data.get('error_message','')}")
            return None, None, None
        route = data["routes"][0]
        legs = route["legs"]
        dist = sum(l["distance"]["value"] for l in legs)
        dur = sum(l["duration"]["value"] for l in legs)
        poly = (route.get("overview_polyline") or {}).get("points")
        return dist, dur, poly
    except Exception as e:
        print(f"⚠️ Directions API erreur : {e}")
        return None, None, None


# ── Itinéraire d'une journée ──────────────────────────────────────
def build_day_itinerary(places_by_category, api_key=None, n_visits=N_VISITS, travelmode="walking"):
    """
    Retourne un dict :
      { 'url', 'sequence', 'distance_m', 'duration_s', 'distance_txt', 'duration_txt' }
    ou None si pas assez de données.
    """
    visits, restos, bars = [], [], []
    for cat, places in places_by_category.items():
        for p in places:
            if not _coords(p):
                continue
            if cat in FOOD_CATEGORIES:
                restos.append(p)
            elif cat in BAR_CATEGORIES:
                bars.append(p)
            elif cat in NON_TOURIST:
                continue
            else:
                visits.append(p)

    visits = _top(visits, n_visits)
    if len(visits) < 2:
        return None

    start = max(range(len(visits)), key=lambda i: visits[i].get("Total Reviews", 0) or 0)
    seq = list(optimize_walking_order(visits, start))
    for v in seq:
        v["_step"] = "🎯 Visite"

    # Les 2 restaurants LES PLUS POPULAIRES (par avis) — aucun critère de distance.
    top_restos = _top(restos, 2)
    if len(top_restos) >= 1:
        mid = len(seq) // 2                       # déjeuner après ~la moitié des visites
        seq.insert(mid, dict(top_restos[0], _step="🍴 Déjeuner"))
    if len(top_restos) >= 2:
        seq.append(dict(top_restos[1], _step="🍽️ Dîner"))

    # Le bar / club LE PLUS POPULAIRE (bars + clubs confondus).
    top_bar = _top(bars, 1)
    if top_bar:
        seq.append(dict(top_bar[0], _step="🍹 Bar / Club"))

    url = build_route_url(seq, travelmode=travelmode)
    dist, dur, poly = walking_distance_time(seq, api_key)

    return {
        "url": url,
        "sequence": seq,
        "polyline": poly,
        "distance_m": dist,
        "duration_s": dur,
        "distance_txt": _fmt_dist(dist),
        "duration_txt": _fmt_dur(dur),
        "has_map": False,   # passe à True quand la carte statique est téléchargée
    }


# ── Carte statique (image PNG du parcours) via Static Maps API ────
def static_map_url(itin, api_key, size="640x400", scale=2):
    """Construit l'URL Static Maps : tracé du parcours + marqueurs par type.
    Utilise la polyline réelle (chemin piéton) si dispo, sinon relie les points."""
    seq = itin.get("sequence") or []
    pts = [s for s in seq if _coords(s)]
    if len(pts) < 2 or not api_key:
        return None

    NAVY, ORANGE, PURPLE = "0x1f3b63", "0xe67e22", "0x8e44ad"  # visites / repas / bar

    # Tracé : polyline réelle si on l'a, sinon segments droits entre les points
    if itin.get("polyline"):
        path = f"weight:4|color:{NAVY}cc|enc:{itin['polyline']}"
    else:
        coords = "|".join(f"{s['Lat']},{s['Lng']}" for s in pts)
        path = f"weight:4|color:{NAVY}cc|{coords}"

    # Un marqueur numéroté par arrêt (1-9 puis A-Z), coloré par type.
    # Label Static Maps = 1 seul caractère, d'où le passage aux lettres après 9.
    def _label(i):
        return str(i) if i <= 9 else chr(ord("A") + i - 10)

    params = [("size", size), ("scale", str(scale)), ("path", path), ("key", api_key)]
    for i, s in enumerate(pts, 1):
        step = s.get("_step", "")
        color = ORANGE if ("Déjeuner" in step or "Dîner" in step) else \
                PURPLE if "Bar" in step else NAVY
        params.append(("markers",
                       f"size:mid|color:{color}|label:{_label(i)}|{s['Lat']},{s['Lng']}"))

    return "https://maps.googleapis.com/maps/api/staticmap?" + \
        urllib.parse.urlencode(params, quote_via=urllib.parse.quote)


def download_static_map(itin, api_key, dest_path, size="640x400", scale=2):
    """Télécharge l'image de la carte et l'écrit dans dest_path. True si OK."""
    url = static_map_url(itin, api_key, size=size, scale=scale)
    if not url:
        return False
    try:
        r = requests.get(url, timeout=30)
        ctype = r.headers.get("Content-Type", "")
        if r.status_code == 200 and ctype.startswith("image/"):
            with open(dest_path, "wb") as f:
                f.write(r.content)
            return True
        print(f"⚠️ Static Maps : {r.status_code} {ctype} {r.text[:120]}")
        return False
    except Exception as e:
        print(f"⚠️ Static Maps erreur : {e}")
        return False


def _fmt_dist(m):
    if not m:
        return None
    return f"{m/1000:.1f} km".replace(".", ",")


def _fmt_dur(s):
    if not s:
        return None
    mins = round(s / 60)
    return f"{mins//60}h{mins%60:02d}" if mins >= 60 else f"{mins} min"


def describe(seq):
    return "\n".join(f"{i}. {s.get('_step','•')} — {s.get('Name','?')}"
                     for i, s in enumerate(seq, 1))


# ── Bloc texte pour la version body_plain de l'email ──────────────
def itinerary_plain_block(itin):
    if not itin or not itin.get("url"):
        return ""
    lines = ["", "🗺 VOTRE PARCOURS D'UNE JOURNÉE À PIED", ""]
    infos = []
    if itin.get("distance_txt"):
        infos.append(itin["distance_txt"])
    if itin.get("duration_txt"):
        infos.append(f"{itin['duration_txt']} de marche")
    if infos:
        lines.append("    " + " · ".join(infos))
        lines.append("")
    for i, s in enumerate(itin["sequence"], 1):
        lines.append(f"    {i}. {s.get('_step','')} — {s.get('Name','')}")
    lines += ["", f"    Ouvrir l'itinéraire dans Google Maps : {itin['url']}", ""]
    return "\n".join(lines)


# ── Bloc HTML autonome pour l'email (styles inline, à insérer tel quel) ──
def itinerary_email_block(itin):
    if not itin or not itin.get("url"):
        return ""
    NAVY, RED, TEXT = "#1f3b63", "#d32f2f", "#3c4043"
    steps = "".join(
        f'<tr><td style="padding:6px 0;font-size:13.5px;color:{TEXT};">'
        f'<strong>{i}.</strong> {s.get("_step","")} — {s.get("Name","")}</td></tr>'
        for i, s in enumerate(itin["sequence"], 1)
    )
    infos = []
    if itin.get("distance_txt"):
        infos.append(f"🚶 {itin['distance_txt']}")
    if itin.get("duration_txt"):
        infos.append(f"⏱️ {itin['duration_txt']} de marche")
    info_line = ("&nbsp;&middot;&nbsp;".join(infos)) or "Parcours optimisé à pied"
    map_img = ""
    if itin.get("has_map"):
        map_img = ('<img src="cid:routemap" alt="Carte du parcours" width="100%" '
                   'style="display:block;width:100%;max-width:100%;height:auto;'
                   'border-radius:10px;border:1px solid #dde5ee;margin-bottom:14px;">')
    btn = (f'<a href="{itin["url"]}" target="_blank" '
           f'style="display:inline-block;background-color:{NAVY};color:#fff;padding:12px 24px;'
           f'border-radius:6px;text-decoration:none;font-weight:600;font-size:14px;">'
           f'Ouvrir l\'itinéraire dans Google Maps</a>')
    return f"""
  <tr><td style="padding:26px 36px 8px;">
    <p style="margin:0 0 2px;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:{RED};">Itinéraire</p>
    <h2 style="margin:0 0 6px;font-size:19px;font-weight:700;color:{NAVY};">Parcours d'une journée à pied</h2>
    <p style="margin:0 0 14px;font-size:13.5px;color:{TEXT};">{info_line}</p>
    {map_img}
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="background:#f3f6fa;border:1px solid #dde5ee;border-radius:10px;margin-bottom:14px;">
      <tr><td style="padding:14px 18px;"><table role="presentation" width="100%">{steps}</table></td></tr>
    </table>
    <div style="text-align:center;">{btn}</div>
  </td></tr>"""


if __name__ == "__main__":
    demo = {
        "historical_sites": [
            {"Name": "Tour Eiffel", "Address": "Champ de Mars, 75007 Paris", "PlaceID": "A",
             "Lat": 48.8584, "Lng": 2.2945, "Total Reviews": 350000},
            {"Name": "Arc de Triomphe", "Address": "Pl. Charles de Gaulle, 75008 Paris", "PlaceID": "B",
             "Lat": 48.8738, "Lng": 2.2950, "Total Reviews": 210000}],
        "museums": [
            {"Name": "Louvre", "Address": "Rue de Rivoli, 75001 Paris", "PlaceID": "C",
             "Lat": 48.8606, "Lng": 2.3376, "Total Reviews": 300000},
            {"Name": "Orsay", "Address": "75007 Paris", "PlaceID": "D",
             "Lat": 48.8600, "Lng": 2.3266, "Total Reviews": 150000}],
        "parks": [{"Name": "Luxembourg", "Address": "75006 Paris", "PlaceID": "E",
                   "Lat": 48.8462, "Lng": 2.3372, "Total Reviews": 180000}],
        "restaurants": [
            {"Name": "Le Comptoir", "Address": "75006 Paris", "PlaceID": "R1",
             "Lat": 48.8519, "Lng": 2.3387, "Total Reviews": 5000},
            {"Name": "Bistrot Eiffel", "Address": "75007 Paris", "PlaceID": "R2",
             "Lat": 48.8570, "Lng": 2.2980, "Total Reviews": 4200}],
        "nightclubs": [{"Name": "Le Rex Club", "Address": "75002 Paris", "PlaceID": "N1",
                        "Lat": 48.8703, "Lng": 2.3470, "Total Reviews": 9000}],
        "bars": [{"Name": "Le Perchoir", "Address": "75011 Paris", "PlaceID": "B1",
                  "Lat": 48.8670, "Lng": 2.3780, "Total Reviews": 8000}],
        "airports": [{"Name": "CDG", "Address": "95700 Roissy", "PlaceID": "X",
                      "Lat": 49.0097, "Lng": 2.5479, "Total Reviews": 400000}],
    }
    itin = build_day_itinerary(demo, api_key=None)   # api_key=None -> pas de temps réel en démo
    print(describe(itin["sequence"]))
    print()
    print("Distance/temps réels :", itin["distance_txt"], itin["duration_txt"], "(None car pas d'appel API en démo)")
    print()
    print(itin["url"])
