# link_inserter_v5.py
import argparse, os, re, sys, io, gzip, unicodedata
import pandas as pd, requests, xml.etree.ElementTree as ET

HDRS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X) AppleWebKit/537.36 (KHTML, like Gecko) Chrome Safari",
    "Accept": "application/xml,text/xml,*/*;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
}

def read_table(path, sheet):
    ext = os.path.splitext(path.lower())[1]
    if ext in [".xlsx", ".xls"]: return pd.read_excel(path, sheet_name=sheet)
    if ext in [".csv", ".tsv"]:  return pd.read_csv(path, sep="," if ext==".csv" else "\t")
    raise ValueError(f"Format non supporté: {ext}")

def write_table(df, path):
    parent = os.path.dirname(os.path.abspath(path))
    if parent and not os.path.exists(parent): os.makedirs(parent, exist_ok=True)
    ext = os.path.splitext(path.lower())[1]
    if ext in [".xlsx", ".xls"]: df.to_excel(path, index=False)
    elif ext in [".csv", ".tsv"]: df.to_csv(path, index=False, sep="," if ext==".csv" else "\t")
    else: raise ValueError(f"Format non supporté: {ext}")

def _http_get(url, timeout=30):
    r = requests.get(url, headers=HDRS, timeout=timeout, stream=True); r.raise_for_status()
    b = r.content
    if url.lower().endswith(".gz") or r.headers.get("Content-Encoding","").lower()=="gzip":
        try: b = gzip.decompress(b)
        except Exception: pass
    return b

def _parse_xml_bytes(b):
    txt = b.decode("utf-8", errors="replace").strip()
    try:
        return ET.fromstring(txt), txt
    except ET.ParseError:
        return None, txt

def _extract_locs_from_root(root):
    tag = (root.tag or "").lower()
    locs = [n.text.strip() for n in root.findall(".//{*}loc") if n.text]
    return locs, ("sitemapindex" in tag)

def slugify_city(s: str) -> str:
    if not s: return ""
    # enlever accents → ascii
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.lower()
    # remplacer tout non-alphanum par '-'
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    s = re.sub(r"-{2,}", "-", s)
    return s

def fetch_mappings(sitemap_url: str, verbose=False, max_fetch=100):
    """
    Explore sitemap (index inclus) et renvoie:
      - code2url: { 'v4922': '.../hotels-adelaide-v4922.html', ... }
      - cityslug2url: { 'adelaide': '.../hotels-adelaide-v4922.html', ... }
    On cible les URLs /a/<lang>/destination/city/hotels-<slug>-<code>.html
    """
    seen, queue = set(), [sitemap_url]
    code2url, cityslug2url = {}, {}
    re_city = re.compile(r"/a/[a-z]{2}/destination/city/hotels-([a-z0-9\-]+)-([a-z0-9\-]+)\.html$", re.I)

    while queue and len(seen) < max_fetch:
        url = queue.pop(0)
        if url in seen: continue
        seen.add(url)
        try:
            b = _http_get(url)
        except Exception as e:
            if verbose: print(f"[SITEMAP] GET FAIL {url}: {e}", file=sys.stderr)
            continue
        root, raw = _parse_xml_bytes(b)
        if root is None:
            locs = re.findall(r"<loc>\s*(.*?)\s*</loc>", raw, flags=re.I)
            is_index = any(x.lower().endswith((".xml",".xml.gz")) for x in locs)
        else:
            locs, is_index = _extract_locs_from_root(root)
        if verbose: print(f"[SITEMAP] {'INDEX' if is_index else 'URLSET'} {url} → {len(locs)}")

        if is_index:
            # on garde les sous-sitemaps pertinents (destination, toutes langues)
            for loc in locs:
                low = loc.lower()
                if "/destination" in low: queue.append(loc)
        else:
            for u in locs:
                m = re_city.search(u)
                if m:
                    city_slug, code = m.group(1).lower(), m.group(2).lower()
                    code2url.setdefault(code, u)
                    # ne pas écraser si déjà présent (première occur)
                    cityslug2url.setdefault(city_slug, u)

    if verbose:
        print(f"[SITEMAP] codes: {len(code2url)} | city-slugs: {len(cityslug2url)}")
        ex = list(code2url.items())[:3]
        for k,v in ex: print("  ex:", k, "→", v)
    return code2url, cityslug2url

def normalize_code(raw):
    if raw is None: return ""
    s = str(raw).strip().replace(" ","")
    if not s: return ""
    if s[0] == "V": s = "v" + s[1:]  # cas signalé
    return s.lower()

def already_linked(content, url): return isinstance(content, str) and url in content

def build_city_patterns(city):
    city = (city or "").strip()
    if not city: return []
    city_esc = re.escape(city)
    p1 = re.compile(rf"(?i)\b(h[oô]tel[s]?)\s+(?:à\s+)?({city_esc})\b")
    p2 = re.compile(rf"(?i)\b(h[oô]tel[s]?)\s+({city_esc})\b")
    return [p1, p2]

def link_once(content, url, patterns):
    if not isinstance(content, str) or not content: return content, False, ""
    for pat in patterns:
        m = pat.search(content)
        if m:
            full = m.group(0)
            return content[:m.start()] + f'<a href="{url}">{full}</a>' + content[m.end():], True, full
    return content, False, ""

def maybe_filter_lang(df, col_content, lang_col, lang_value, infer_fr, verbose=False):
    if lang_col and lang_col in df.columns:
        before = len(df)
        df = df[df[lang_col].astype(str).str.lower() == str(lang_value).lower()].copy()
        if verbose: print(f"Filtre langue (col): {before} → {len(df)}")
        return df
    if infer_fr and col_content in df.columns:
        try:
            import langid
            before = len(df)
            mask = df[col_content].fillna("").astype(str).apply(lambda t: langid.classify(t[:1000])[0]=="fr")
            df = df[mask].copy()
            if verbose: print(f"Filtre langue (inférence FR): {before} → {len(df)}")
            return df
        except Exception:
            if verbose: print("[INFO] langid absent → pas de filtre langue")
    return df

def main():
    ap = argparse.ArgumentParser(description="Maillage: ajoute un lien destination (FR) dans le contenu facette.")
    ap.add_argument("--input", required=True); ap.add_argument("--sheet", default=None)
    ap.add_argument("--output", required=True)
    ap.add_argument("--col-code", default="A"); ap.add_argument("--col-city", default="B"); ap.add_argument("--col-content", default="J")
    ap.add_argument("--lang-col", default=None); ap.add_argument("--lang-value", default="fr"); ap.add_argument("--infer-fr", action="store_true")
    ap.add_argument("--sitemap", default="https://all.accor.com/a/fr/destination.sitemap.xml")
    ap.add_argument("--verbose", action="store_true"); ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    df = read_table(args.input, args.sheet)
    # Mapper A/B/J s'ils sont positions
    for attr in ["col_code","col_city","col_content"]:
        col = getattr(args, attr)
        if col not in df.columns:
            idx = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".find(col)
            if idx >= 0 and idx < len(df.columns): setattr(args, attr, df.columns[idx])
    for c in [args.col_code, args.col_city, args.col_content]:
        if c not in df.columns: raise ValueError(f"Colonne manquante: {c}")

    # Normalisation code saisi
    df["__code_norm"] = df[args.col_code].apply(normalize_code)
    # Slug ville
    df["__city_slug"] = df[args.col_city].astype(str).apply(slugify_city)

    # Filtre langue éventuel
    df = maybe_filter_lang(df, args.col_content, args.lang_col, args.lang_value, args.infer_fr, verbose=args.verbose)

    if args.verbose:
        print(f"Lignes à traiter: {len(df)}")
        print(f"Sitemap: {args.sitemap}")

    code2url, cityslug2url = fetch_mappings(args.sitemap, verbose=args.verbose)

    linked, anchors, urls, reasons, codes_found = [], [], [], [], []

    for idx, row in df.iterrows():
        code_norm = (row["__code_norm"] or "").strip()
        city_slug = (row["__city_slug"] or "").strip()
        content = row[args.col_content] if pd.notna(row[args.col_content]) else ""
        city    = str(row[args.col_city]).strip() if pd.notna(row[args.col_city]) else ""

        url = ""; reason = ""
        # 1) match direct par code
        if code_norm and code_norm in code2url:
            url = code2url[code_norm]; reason = "by_code"
            true_code = code_norm
        else:
            # 2) fallback par slug ville
            if city_slug and city_slug in cityslug2url:
                url = cityslug2url[city_slug]; reason = "by_city_slug"
                # extraire le vrai code depuis l'URL
                m = re.search(r"-([a-z0-9\-]+)\.html$", url, flags=re.I)
                true_code = m.group(1).lower() if m else ""
            else:
                true_code = ""

        codes_found.append(true_code)
        if not url or not city or not content:
            linked.append(False); anchors.append(""); urls.append(url); reasons.append(reason or "no_url_or_empty")
            continue

        if already_linked(content, url):
            linked.append(False); anchors.append(""); urls.append(url); reasons.append("already_linked")
            continue

        pats = build_city_patterns(city)
        new_content, did, anchor = link_once(content, url, pats)
        if did:
            df.at[idx, args.col_content] = new_content
            linked.append(True); anchors.append(anchor); urls.append(url); reasons.append(reason or "ok")
        else:
            linked.append(False); anchors.append(""); urls.append(url); reasons.append((reason + "_no_text_match") if reason else "no_text_match")

    df["__dest_url"] = urls
    df["__linked"] = linked
    df["__anchor_matched"] = anchors
    df["__reason"] = reasons
    df["__code_from_url"] = codes_found  # utile pour voir le VRAI code Accor

    if args.dry_run:
        total, done = len(df), sum(linked)
        print(f"[DRY-RUN] {done}/{total} lignes liées.")
        from collections import Counter
        print("Raisons:", Counter(reasons).most_common(8))
    else:
        write_table(df, args.output)
        print(f"✅ Terminé. Liens insérés: {sum(linked)}/{len(df)} → {args.output}")

if __name__ == "__main__":
    main()
