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
import os
import urllib.parse
import requests

# Nombre de visites par défaut, réparties moitié matin / moitié après-midi
# autour du déjeuner. Pas de plafond imposé : monte ce chiffre si tu veux.
N_VISITS = 10

# ── Garde-fou anti-lieu-excentré ────────────────────────────────────────────
# L'API Places renvoie parfois des lieux ultra-populaires mais hors de portée
# d'une journée à pied (Sintra à 25 km, Cabo da Roca à 40 km), voire des
# résultats parasites carrément dans un autre pays. On écarte du PARCOURS À PIED
# les lieux trop loin du cœur géographique (médiane des positions, robuste aux
# outliers). N'affecte QUE l'itinéraire — l'Excel garde tous les lieux.
# Mettre None pour désactiver (= pur populaire, sans filtre distance).
MAX_KM_FROM_CORE = 30.0

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


def _dedup(places):
    """Retire les doublons (même lieu présent dans plusieurs catégories de
    l'Excel, ex. Castelo de São Jorge à la fois 'Sites historiques' et 'Musées').
    Clé = PlaceID ; fallback nom + coords arrondies si PlaceID absent."""
    seen, out = set(), []
    for p in places:
        key = p.get("PlaceID")
        if not key:
            c = _coords(p)
            key = (p.get("Name"), round(c[0], 5), round(c[1], 5)) if c else p.get("Name")
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


def _median(vals):
    s = sorted(vals)
    n = len(s)
    if n == 0:
        return None
    m = n // 2
    return s[m] if n % 2 else (s[m - 1] + s[m]) / 2.0


def _filter_walkable(groups, max_km, anchor=None):
    """
    Écarte les lieux trop loin du cœur géographique du groupe.
    `groups` = liste de listes de lieux. Le cœur = `anchor` s'il est fourni
    (coordonnées du lieu le plus populaire), sinon la médiane des positions de
    TOUS les lieux (robuste : un lieu isolé à 1000 km ne déplace pas la médiane).
    Renvoie les mêmes listes filtrées. Si max_km est None : ne filtre rien.
    """
    if not max_km:
        return groups, None
    pts = [_coords(p) for g in groups for p in g if _coords(p)]
    if len(pts) < 3:                      # pas assez de points pour un cœur fiable
        return groups, None
    if anchor:
        clat, clng = anchor
    else:
        clat = _median([la for la, _ in pts])
        clng = _median([lo for _, lo in pts])
    max_m = max_km * 1000.0                # hav() renvoie des mètres
    out = []
    for g in groups:
        out.append([p for p in g
                    if _coords(p) and hav(clat, clng, *_coords(p)) <= max_m])
    return out, (clat, clng)


def _densest_anchor(visits, radius_km, pool_size=40):
    """Ancre du parcours = le lieu populaire entouré du PLUS GRAND NOMBRE d'autres
    lieux populaires dans le rayon de marche (densité), départagé par le total
    d'avis cumulé dans ce rayon.

    Pourquoi : un site unique très visité mais isolé (Trakai en Lituanie, Sintra
    au Portugal) ne doit pas l'emporter sur un vrai centre-ville dense (Vilnius,
    Lisbonne). On raisonne sur les `pool_size` lieux les plus populaires pour
    ignorer le bruit. Pour une VILLE en entrée, le centre reste la zone la plus
    dense -> comportement inchangé.
    """
    pts = [p for p in visits if _coords(p)]
    if not pts:
        return None
    pool = _top(pts, pool_size)
    radius_m = radius_km * 1000.0
    best, best_key = None, None
    for c in pool:
        cla, clo = _coords(c)
        near = [v for v in pool if hav(cla, clo, *_coords(v)) <= radius_m]
        key = (len(near), sum(v.get("Total Reviews", 0) or 0 for v in near))
        if best_key is None or key > best_key:
            best_key, best = key, c
    return _coords(best)


# ── Lien Google Maps : format "path" (prend TOUS les arrêts, sans plafond) ──
# Même approche que tour_url de thermodata_engine : .../dir/A/B/C/...
# On utilise les COORDONNÉES GPS (précises et uniques par lieu) plutôt que les
# adresses : les adresses Google Places sont souvent vagues (code postal seul)
# ou dupliquées entre lieux, ce qui empêche Maps de tracer. Les coords, jamais.
def _loc(s):
    c = _coords(s)
    if c:
        return f"{c[0]},{c[1]}"
    if s.get("Address"):
        return urllib.parse.quote(str(s["Address"]).strip())
    return urllib.parse.quote(str(s.get("Name", "")).strip())


def build_route_url(stops, travelmode="walking"):
    stops = [s for s in stops if (s.get("Address") or s.get("Name") or _coords(s))]
    if len(stops) < 2:
        return None
    url = "https://www.google.com/maps/dir/" + "/".join(_loc(s) for s in stops)
    # Mode de transport (suffixe interne Google) : 3e0 voiture, 3e1 vélo, 3e2 marche, 3e3 transports
    modes = {"driving": "3e0", "bicycling": "3e1", "walking": "3e2", "transit": "3e3"}
    if travelmode in modes:
        url += f"/data=!4m2!4m1!{modes[travelmode]}"
    return url


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
def build_day_itinerary(places_by_category, api_key=None, n_visits=N_VISITS,
                        travelmode="walking", anchor=None):
    """
    Retourne un dict :
      { 'url', 'sequence', 'distance_m', 'duration_s', 'distance_txt', 'duration_txt' }
    ou None si pas assez de données.

    `anchor` : (lat, lng) du centre imposé (ex. ville géocodée). S'il est fourni,
    le parcours se centre dessus -> fiable même si Google a ramené des lieux
    d'une autre ville. Sinon, on retombe sur la zone la plus dense des résultats.
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

    # Garde-fou : on retire du parcours à pied les lieux trop loin du cœur.
    # L'Excel, lui, garde tout — ce filtre n'agit que sur l'itinéraire.
    #
    # Cœur du parcours, par ordre de priorité :
    #   1) `anchor` imposé = coordonnées de la ville géocodée (le plus fiable :
    #      on reste sur la ville DEMANDÉE, même si Google a ramené des lieux
    #      d'une autre ville pour une catégorie absente localement) ;
    #   2) sinon, la zone la plus dense en lieux populaires (_densest_anchor),
    #      qui évite qu'un site isolé hyper-visité (Trakai, Sintra) l'emporte.
    # Le rayon MAX_KM_FROM_CORE borne le tout à une distance marchable.
    anchor_src = "ville géocodée" if anchor else "zone la plus dense"
    if anchor is None and MAX_KM_FROM_CORE:
        anchor = _densest_anchor(visits, MAX_KM_FROM_CORE)

    # Rayon ADAPTATIF : on part d'un rayon de marche serré et on l'élargit par
    # paliers UNIQUEMENT s'il manque de quoi remplir le parcours complet
    # (10 visites + 2 restos + 1 bar = 13 étapes). Une grande ville (Vilnius)
    # reste à ~5 km ; une petite ville (Šiauliai) s'élargit pour aller chercher
    # de quoi compléter, sans dépasser le plafond (sinon on sauterait sur une
    # autre ville). Réglable : ITINERARY_BASE_KM / ITINERARY_STEP_KM / ITINERARY_MAX_KM.
    base_km = float(os.environ.get("ITINERARY_BASE_KM", MAX_KM_FROM_CORE or 30.0))
    step_km = float(os.environ.get("ITINERARY_STEP_KM", 5.0))
    max_km = float(os.environ.get("ITINERARY_MAX_KM", 30.0))
    want_v, want_r, want_b = n_visits, 2, 1

    visits, restos, bars, core, used_km = visits, restos, bars, None, base_km
    r = base_km
    while True:
        (fv, fr, fb), c = _filter_walkable([visits, restos, bars], r, anchor=anchor)
        cand = (fv, fr, fb, c, r)
        enough = (len(_dedup(fv)) >= want_v and len(_dedup(fr)) >= want_r
                  and len(_dedup(fb)) >= want_b)
        if enough or r >= max_km or not c:
            visits, restos, bars, core, used_km = cand
            break
        r += step_km

    if core:
        print(f"🚶 Garde-fou parcours : rayon {used_km:g} km autour du centre "
              f"({anchor_src} : {core[0]:.4f}, {core[1]:.4f})")

    # Dédup : un même lieu peut être dans plusieurs catégories de l'Excel.
    visits, restos, bars = _dedup(visits), _dedup(restos), _dedup(bars)
    # Et on évite qu'un lieu déjà classé en visite ressorte en repas/bar.
    vids = {v.get("PlaceID") for v in visits if v.get("PlaceID")}
    restos = [r for r in restos if r.get("PlaceID") not in vids]
    bars = [b for b in bars if b.get("PlaceID") not in vids]

    visits = _top(visits, n_visits)
    if len(visits) < 2:
        return None

    start = max(range(len(visits)), key=lambda i: visits[i].get("Total Reviews", 0) or 0)
    seq = list(optimize_walking_order(visits, start))
    for v in seq:
        v["_step"] = "🎯"

    # Les 2 restaurants LES PLUS POPULAIRES (par avis) parmi les atteignables.
    top_restos = _top(restos, 2)
    if len(top_restos) >= 1:
        # Déjeuner : inséré à mi-parcours, mais à la position de la fenêtre
        # centrale qui rallonge le MOINS le trajet (pas juste l'index du milieu).
        seq = _best_insert(seq, dict(top_restos[0], _step="🍴"))
    if len(top_restos) >= 2:
        seq.append(dict(top_restos[1], _step="🍽️"))

    # Le bar / club LE PLUS POPULAIRE (bars + clubs confondus).
    top_bar = _top(bars, 1)
    if top_bar:
        seq.append(dict(top_bar[0], _step="🍹"))

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


# ── Plan MULTI-JOURS ──────────────────────────────────────────────
# On récupère un large pool de lieux (rayon MAX_KM_FROM_CORE autour de la ville),
# on le découpe en N zones géographiques (clustering équilibré), et chaque zone
# devient une journée : ordre optimisé (2-opt) + déjeuner + dîner + bar du coin.
# Ainsi un grand rayon reste cohérent : l'étalement est absorbé par le découpage.
def _balanced_clusters(pts, k, cap):
    """k zones géographiques ÉQUILIBRÉES (<= cap points chacune).
    Centres en farthest-first, puis affectation capacitée au plus proche."""
    n = len(pts)
    if k <= 1 or n == 0:
        return [0] * n
    centers = [pts[0]]
    while len(centers) < k:
        far = max(pts, key=lambda p: min(hav(p[0], p[1], c[0], c[1]) for c in centers))
        centers.append(far)
    assign = [0] * n
    for _ in range(30):
        counts = [0] * k
        assign = [-1] * n
        pairs = sorted((hav(pts[i][0], pts[i][1], centers[c][0], centers[c][1]), i, c)
                       for i in range(n) for c in range(k))
        for _d, i, c in pairs:
            if assign[i] != -1 or counts[c] >= cap:
                continue
            assign[i] = c
            counts[c] += 1
        for i in range(n):                       # reste éventuel -> zone la moins pleine
            if assign[i] == -1:
                cand = [c for c in range(k) if counts[c] < cap] or list(range(k))
                c = min(cand, key=lambda c: hav(pts[i][0], pts[i][1], centers[c][0], centers[c][1]))
                assign[i] = c
                counts[c] += 1
        newc = []
        for c in range(k):
            cp = [pts[i] for i in range(n) if assign[i] == c]
            newc.append((sum(x for x, _ in cp) / len(cp), sum(y for _, y in cp) / len(cp))
                        if cp else centers[c])
        if newc == centers:
            break
        centers = newc
    return assign


def _order_by_proximity(centroids, anchor):
    """Ordonne les jours du plus proche au plus loin de l'ancre (nearest-neighbor)."""
    k = len(centroids)
    if k <= 1:
        return list(range(k))
    cur = anchor or centroids[0]
    rest, order = set(range(k)), []
    while rest:
        nxt = min(rest, key=lambda c: hav(cur[0], cur[1], centroids[c][0], centroids[c][1]))
        order.append(nxt)
        rest.discard(nxt)
        cur = centroids[nxt]
    return order


def _nearest_unused(places, center, used, k):
    """k lieux les plus proches du centre du jour, non déjà utilisés (repas/bar)."""
    cand = [p for p in places
            if (p.get("PlaceID") or id(p)) not in used and _coords(p)]
    cand.sort(key=lambda p: hav(center[0], center[1], *_coords(p)))
    chosen = cand[:k]
    for p in chosen:
        used.add(p.get("PlaceID") or id(p))
    return chosen


def build_multiday_plan(places_by_category, days, anchor=None,
                        api_key=None, n_visits=N_VISITS):
    """Plan sur `days` jours. Chaque jour vise 13 étapes (n_visits visites +
    déjeuner + dîner + bar), regroupées par zone géographique puis ordre optimisé.
    Renvoie une liste de dicts {day, sequence, url, distance_txt, duration_txt}
    ou [] si pas assez de données. `feasible_days` = limité par le nb de lieux."""
    if not days or days < 1:
        return []

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

    if anchor is None and MAX_KM_FROM_CORE:
        anchor = _densest_anchor(visits, MAX_KM_FROM_CORE)
    (visits, restos, bars), _core = _filter_walkable([visits, restos, bars],
                                                      MAX_KM_FROM_CORE, anchor=anchor)

    visits, restos, bars = _dedup(visits), _dedup(restos), _dedup(bars)
    vids = {v.get("PlaceID") for v in visits if v.get("PlaceID")}
    restos = [r for r in restos if r.get("PlaceID") not in vids]
    bars = [b for b in bars if b.get("PlaceID") not in vids]

    if len(visits) < 2:
        return []

    # On ne crée pas plus de jours que la ville ne peut en remplir (>= 2 visites/jour)
    feasible_days = max(1, min(int(days), len(visits) // 2))
    pool = _top(visits, feasible_days * n_visits)
    pts = [_coords(p) for p in pool]

    assign = _balanced_clusters(pts, feasible_days, cap=n_visits)
    centroids = []
    for c in range(feasible_days):
        cp = [pts[i] for i in range(len(pts)) if assign[i] == c]
        centroids.append((sum(x for x, _ in cp) / len(cp), sum(y for _, y in cp) / len(cp))
                         if cp else (anchor or pts[0]))

    used_resto, used_bar = set(), set()
    plan = []
    for day_no, c in enumerate(_order_by_proximity(centroids, anchor), start=1):
        day_visits = [pool[i] for i in range(len(pool)) if assign[i] == c]
        if not day_visits:
            continue
        start = max(range(len(day_visits)),
                    key=lambda i: day_visits[i].get("Total Reviews", 0) or 0)
        seq = [dict(v, _step="🎯") for v in optimize_walking_order(day_visits, start)]
        cen = centroids[c]
        d_restos = _nearest_unused(restos, cen, used_resto, 2)
        d_bar = _nearest_unused(bars, cen, used_bar, 1)
        if len(d_restos) >= 1:
            seq = _best_insert(seq, dict(d_restos[0], _step="🍴"))
        if len(d_restos) >= 2:
            seq.append(dict(d_restos[1], _step="🍽️"))
        if d_bar:
            seq.append(dict(d_bar[0], _step="🍹"))
        dist, dur, _poly = walking_distance_time(seq, api_key)
        plan.append({
            "day": day_no,
            "sequence": seq,
            "url": build_route_url(seq, travelmode="walking"),
            "distance_txt": _fmt_dist(dist),
            "duration_txt": _fmt_dur(dur),
        })
    if feasible_days < int(days):
        print(f"ℹ️ Plan multi-jours : {feasible_days} jour(s) au lieu de {int(days)} "
              f"demandé(s) - pas assez de lieux dans la zone pour remplir plus.")
    print(f"🗓️ Plan multi-jours : {len(plan)} jour(s) generes.")
    return plan


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
    return "\n".join(f"{i}. {s.get('_step','•')} {s.get('Name','?')}"
                     for i, s in enumerate(seq, 1))


# ── Bloc texte pour la version body_plain de l'email ──────────────
def place_url(s):
    """Lien Google Maps vers UN lieu précis (sa fiche), pour le bouton d'étape."""
    pid = s.get("PlaceID")
    if pid:
        return f"https://www.google.com/maps/place/?q=place_id:{pid}"
    c = _coords(s)
    if c:
        return f"https://www.google.com/maps/search/?api=1&query={c[0]},{c[1]}"
    return ("https://www.google.com/maps/search/?api=1&query="
            + urllib.parse.quote(str(s.get("Name", ""))))


def streetview_url(s):
    """Lien Google Street View (vue panoramique) à l'emplacement du lieu."""
    c = _coords(s)
    if c:
        return (f"https://www.google.com/maps/@?api=1&map_action=pano"
                f"&viewpoint={c[0]},{c[1]}")
    return place_url(s)


def _meta(s):
    """'Parc · ★ 4.6 · 12 345 avis · ++' (catégorie incluse, gère les manquants)."""
    bits = []
    cat = s.get("_category")
    if cat and str(cat) not in ("None", ""):
        bits.append(str(cat))
    try:
        bits.append(f"\u2605 {float(s.get('Rating (on 5)')):.1f}")
    except (TypeError, ValueError):
        pass
    n = int(s.get("Total Reviews") or 0)
    if n:
        bits.append(f"{n:,}".replace(",", "\u202f") + " avis")
    pl = s.get("Price Level")
    if pl and str(pl) not in ("Not specified", "None", ""):
        bits.append(str(pl))
    return " \u00b7 ".join(bits)


def _addr(s):
    """Adresse lisible du lieu (vide si non renseignée)."""
    a = s.get("Address")
    if a and str(a) not in ("Not specified", "None", ""):
        return str(a)
    return str(s.get("City") or "")


def _best_insert(seq, item, lo_frac=0.34, hi_frac=0.66):
    """Insère `item` dans `seq` à la position — bornée à la fenêtre centrale
    [lo_frac, hi_frac] pour rester un déjeuner « de midi » — qui rallonge le
    moins le trajet. Fallback : milieu brut si coordonnées indisponibles."""
    n = len(seq)
    c = _coords(item)
    if not c or n < 2:
        seq.insert(n // 2, item)
        return seq
    lo = max(1, int(n * lo_frac))
    hi = min(n, max(lo, int(n * hi_frac)))
    best_pos, best_cost = n // 2, None
    for pos in range(lo, hi + 1):
        a = _coords(seq[pos - 1])
        b = _coords(seq[pos]) if pos < n else None
        if not a:
            continue
        cost = hav(*a, *c) + (hav(*c, *b) - hav(*a, *b) if b else 0.0)
        if best_cost is None or cost < best_cost:
            best_cost, best_pos = cost, pos
    seq.insert(best_pos, item)
    return seq


def itinerary_plain_block(itin):
    if not itin or not itin.get("url"):
        return ""
    lines = ["", "🗺 VOTRE PARCOURS D'UNE JOURNÉE", "",
             "    Les lieux les plus populaires de votre destination, classés pour",
             "    limiter les trajets. Déjeuner à mi-parcours, dîner en fin de",
             "    journée, bar/club pour finir - suivez simplement les numéros.", ""]
    infos = []
    if itin.get("distance_txt"):
        infos.append(itin["distance_txt"])
    if itin.get("duration_txt"):
        infos.append(f"{itin['duration_txt']}")
    if infos:
        lines.append("    " + " · ".join(infos))
        lines.append("")
    for i, s in enumerate(itin["sequence"], 1):
        lines.append(f"    {i}. {s.get('_step','')} {s.get('Name','')}")
        meta = _meta(s)
        if meta:
            lines.append(f"       {meta}")
        addr = _addr(s)
        if addr:
            lines.append(f"       {addr}")
        lines.append(f"       Maps : {place_url(s)}")
        lines.append(f"       Street View : {streetview_url(s)}")
    lines += ["", f"    Ouvrir tout l'itinéraire dans Google Maps : {itin['url']}", ""]
    return "\n".join(lines)


# ── Bloc HTML autonome pour l'email (styles inline, à insérer tel quel) ──
def itinerary_email_block(itin):
    if not itin or not itin.get("url"):
        return ""
    NAVY, RED, TEXT, MUTED = "#1f3b63", "#d32f2f", "#3c4043", "#7a8699"
    rows = []
    for i, s in enumerate(itin["sequence"], 1):
        meta = _meta(s)
        addr = _addr(s)
        sub = []
        if meta:
            sub.append(f'<span style="font-size:12px;color:{MUTED};">{meta}</span>')
        if addr:
            sub.append(f'<span style="font-size:11.5px;color:{MUTED};">\U0001f4cd {addr}</span>')
        sub_html = ("<br>" + "<br>".join(sub)) if sub else ""
        rows.append(
            f'<tr><td style="padding:9px 0;border-bottom:1px solid #e6edf5;">'
            f'<table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>'
            f'<td style="font-size:13.5px;color:{TEXT};vertical-align:middle;line-height:1.55;">'
            f'<strong>{i}.</strong> {s.get("_step","")} <strong>{s.get("Name","")}</strong>'
            f'{sub_html}</td>'
            f'<td align="right" style="vertical-align:middle;white-space:nowrap;padding-left:10px;">'
            f'<a href="{place_url(s)}" target="_blank" '
            f'style="display:inline-block;background:#eef2f8;color:{NAVY};padding:7px 13px;'
            f'border-radius:6px;text-decoration:none;font-weight:600;font-size:12px;'
            f'border:1px solid #d7e0ec;margin-right:6px;">Maps</a>'
            f'<a href="{streetview_url(s)}" target="_blank" '
            f'style="display:inline-block;background:#eef2f8;color:{NAVY};padding:7px 13px;'
            f'border-radius:6px;text-decoration:none;font-weight:600;font-size:12px;'
            f'border:1px solid #d7e0ec;">Street View</a></td>'
            f'</tr></table></td></tr>'
        )
    steps = "".join(rows)
    infos = []
    if itin.get("distance_txt"):
        infos.append(f"📏 {itin['distance_txt']}")
    if itin.get("duration_txt"):
        infos.append(f"⏱️ {itin['duration_txt']}")
    info_line = ("&nbsp;&middot;&nbsp;".join(infos)) or "Parcours optimisé"
    map_img = ""
    if itin.get("has_map"):
        map_img = ('<img src="cid:routemap" alt="Carte du parcours" width="100%" '
                   'style="display:block;width:100%;max-width:100%;height:auto;'
                   'border-radius:10px;border:1px solid #dde5ee;margin-bottom:14px;">')
    btn = (f'<a href="{itin["url"]}" target="_blank" '
           f'style="display:inline-block;background-color:{NAVY};color:#fff;padding:12px 24px;'
           f'border-radius:6px;text-decoration:none;font-weight:600;font-size:14px;">'
           f'Ouvrir tout l\'itinéraire dans Google Maps</a>')
    return f"""
  <tr><td style="padding:26px 36px 8px;">
    <p style="margin:0 0 2px;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:{RED};">Itinéraire</p>
    <h2 style="margin:0 0 6px;font-size:19px;font-weight:700;color:{NAVY};">Parcours d'une journée</h2>
    <p style="margin:0 0 10px;font-size:13.5px;color:{TEXT};line-height:1.55;">Ce parcours réunit les lieux <strong>les plus populaires</strong> de votre destination, classés dans l'ordre qui <strong>réduit au maximum les trajets</strong>. Le déjeuner tombe à mi-parcours, le dîner en fin de journée, puis un bar ou club pour terminer la soirée&nbsp;: chaque étape a son bouton Maps (avec sa note et son nombre d'avis), ou ouvrez tout le parcours d'un coup en bas.</p>
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
