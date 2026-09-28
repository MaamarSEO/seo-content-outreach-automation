# link_inserter_v7.py
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
    locs = [n.text.strip() for n in root.findall(".//{*}loc") if n.text]
    tag = (root.tag or "").lower()
    return locs, ("sitemapindex" in tag)

def slugify_city(s: str) -> str:
    if not s: return ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    s = re.sub(r"-{2,}", "-", s)
    return s

def fetch_mappings(sitemap_url: str, verbose=False, max_fetch=100):
    """
    Renvoie:
      - code2url: { 'v4922': '.../hotels-adelaide-v4922.html', ... }
      - cityslug2url: { 'adelaide': '.../hotels-adelaide-v4922.html', ... }
    Suit les sitemap index et .gz ; retient /destination/(city|region)/hotels-<slug>-<code>.html
    """
    seen, queue = set(), [sitemap_url]
    code2url, cityslug2url = {}, {}
    re_city = re.compile(r"/a/[a-z]{2}/destination/(?:city|region)/hotels-([a-z0-9\-]+)-([a-z0-9\-]+)\.html$", re.I)

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
            for loc in locs:
                if "destin" in loc.lower():  # garder les sitemaps destination
                    queue.append(loc)
        else:
            for u in locs:
                m = re_city.search(u)
                if m:
                    city_slug, code = m.group(1).lower(), m.group(2).lower()
                    code2url.setdefault(code, u)
                    cityslug2url.setdefault(city_slug, u)
    if verbose:
        print(f"[SITEMAP] codes: {len(code2url)} | city-slugs: {len(cityslug2url)}")
    return code2url, cityslug2url

def normalize_code(raw):
    if raw is None: return ""
    s = str(raw).strip().replace(" ","")
    if not s: return ""
    if s[0] == "V": s = "v" + s[1:]  # cas signalé
    return s.lower()

def already_linked_any(row, url, cols):
    if not isinstance(url, str) or not url: return False
    for c in cols:
        val = row[c] if c in row and pd.notna(row[c]) else ""
        if isinstance(val, str) and url in val:
            return True
    return False

def already_linked_cell(content, url):  # par cellule
    return isinstance(content, str) and url in content

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

def maybe_filter_lang(df, col_for_lang, lang_col, lang_value, infer_fr, verbose=False):
    if lang_col and lang_col in df.columns:
        before = len(df)
        df = df[df[lang_col].astype(str).str.lower() == str(lang_value).lower()].copy()
        if verbose: print(f"Filtre langue (col): {before} → {len(df)}")
        return df
    if infer_fr and col_for_lang in df.columns:
        try:
            import langid
            before = len(df)
            mask = df[col_for_lang].fillna("").astype(str).apply(lambda t: langid.classify(t[:1000])[0]=="fr")
            df = df[mask].copy()
            if verbose: print(f"Filtre langue (inférence FR): {before} → {len(df)}")
            return df
        except Exception:
            if verbose: print("[INFO] langid absent → pas de filtre langue")
    return df

def resolve_content_cols(df, cols_arg):
    raw = [c.strip() for c in cols_arg.split(",") if c.strip()]
    resolved = []
    for c in raw:
        if c in df.columns:
            resolved.append(c); continue
        idx = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".find(c.upper())
        if 0 <= idx < len(df.columns):
            resolved.append(df.columns[idx])
        else:
            raise ValueError(f"Colonne contenu introuvable: {c}")
    return resolved

def main():
    ap = argparse.ArgumentParser(description="Maillage: lien destination, priorités colonnes et un seul lien par ligne.")
    ap.add_argument("--input", required=True); ap.add_argument("--sheet", default=None)
    ap.add_argument("--output", required=True)
    ap.add_argument("--col-code", default="A"); ap.add_argument("--col-city", default="B")
    ap.add_argument("--content-cols", default="J", help="Colonnes contenu par ordre de priorité, ex: 'L,N,P,J'")
    ap.add_argument("--one-per-row", action="store_true", help="Insérer au plus un lien par ligne (arrêt dès le 1er succès).")
    ap.add_argument("--lang-col", default=None); ap.add_argument("--lang-value", default="fr"); ap.add_argument("--infer-fr", action="store_true")
    ap.add_argument("--sitemap", default="https://all.accor.com/a/fr/destination.sitemap.xml")
    ap.add_argument("--verbose", action="store_true"); ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    df = read_table(args.input, args.sheet)

    # Mapper A/B s'ils sont des positions
    for attr in ["col_code","col_city"]:
        col = getattr(args, attr)
        if col not in df.columns:
            idx = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".find(col.upper())
            if 0 <= idx < len(df.columns): setattr(args, attr, df.columns[idx])

    # Résoudre les colonnes de contenu (ordre de priorité)
    content_cols = resolve_content_cols(df, args.content_cols)

    for c in [args.col_code, args.col_city]:
        if c not in df.columns: raise ValueError(f"Colonne manquante: {c}")

    # Colonnes techniques
    df["__code_norm"] = df[args.col_code].apply(normalize_code)
    df["__city_slug"] = df[args.col_city].astype(str).apply(slugify_city)

    # Filtre langue éventuel (utilise la 1re col de contenu pour l'inférence FR)
    df = maybe_filter_lang(df, content_cols[0], args.lang_col, args.lang_value, args.infer_fr, verbose=args.verbose)

    if args.verbose:
        print(f"Lignes à traiter: {len(df)}")
        print(f"Sitemap: {args.sitemap}")
        print(f"Priorité colonnes: {content_cols}")
        if args.one_per_row: print("Mode: un seul lien par ligne")

    code2url, cityslug2url = fetch_mappings(args.sitemap, verbose=args.verbose)

    # Colonnes d'audit par contenu
    for c in content_cols:
        base = f"__{c}"
        for sub in ["linked","anchor","reason"]:
            name = f"{base}_{sub}"
            if name not in df.columns:
                df[name] = False if sub=="linked" else ""

    # Colonnes d'audit globales
    if "__dest_url" not in df.columns: df["__dest_url"] = ""
    if "__reason_global" not in df.columns: df["__reason_global"] = ""
    if "__code_from_url" not in df.columns: df["__code_from_url"] = ""

    total_linked = 0

    for idx, row in df.iterrows():
        code_norm = (row["__code_norm"] or "").strip()
        city_slug = (row["__city_slug"] or "").strip()
        city      = str(row[args.col_city]).strip() if pd.notna(row[args.col_city]) else ""

        # Résoudre l'URL une seule fois par ligne (code -> fallback slug)
        if code_norm and code_norm in code2url:
            url = code2url[code_norm]; reason_global = "by_code"
            m = re.search(r"-([a-z0-9\-]+)\.html$", url, flags=re.I)
            true_code = m.group(1).lower() if m else code_norm
        elif city_slug and city_slug in cityslug2url:
            url = cityslug2url[city_slug]; reason_global = "by_city_slug"
            m = re.search(r"-([a-z0-9\-]+)\.html$", url, flags=re.I)
            true_code = m.group(1).lower() if m else ""
        else:
            url = ""; reason_global = "no_url"; true_code = ""

        df.at[idx, "__dest_url"] = url
        df.at[idx, "__reason_global"] = reason_global
        df.at[idx, "__code_from_url"] = true_code

        # Si l'URL est déjà présente dans une des colonnes → skip (et noter la raison)
        if url and already_linked_any(row, url, content_cols):
            for c in content_cols:
                base = f"__{c}"
                df.at[idx, f"{base}_linked"] = False
                df.at[idx, f"{base}_anchor"] = ""
                df.at[idx, f"{base}_reason"] = reason_global + "_already_in_row"
            continue

        inserted = False
        for c in content_cols:
            base = f"__{c}"
            col_linked = f"{base}_linked"
            col_anchor = f"{base}_anchor"
            col_reason = f"{base}_reason"

            content = row[c] if pd.notna(row[c]) else ""
            if not url or not city or not content:
                df.at[idx, col_linked] = False
                df.at[idx, col_anchor] = ""
                df.at[idx, col_reason] = (reason_global or "no_url") + "_no_content_or_city"
                continue

            if already_linked_cell(content, url):
                df.at[idx, col_linked] = False
                df.at[idx, col_anchor] = ""
                df.at[idx, col_reason] = reason_global + "_already_in_cell"
                continue

            patterns = build_city_patterns(city)
            new_content, did, anchor = link_once(content, url, patterns)
            if did:
                df.at[idx, c] = new_content
                df.at[idx, col_linked] = True
                df.at[idx, col_anchor] = anchor
                df.at[idx, col_reason] = reason_global
                total_linked += 1
                inserted = True
                if args.one_per_row:
                    # renseigner les autres colonnes non traitées
                    remaining_cols = [x for x in content_cols if x != c]
                    for rc in remaining_cols:
                        rb = f"__{rc}"
                        df.at[idx, f"{rb}_linked"] = False
                        df.at[idx, f"{rb}_anchor"] = ""
                        df.at[idx, f"{rb}_reason"] = "skipped_after_first_insert"
                    break
            else:
                df.at[idx, col_linked] = False
                df.at[idx, col_anchor] = ""
                df.at[idx, col_reason] = reason_global + "_no_text_match"

        # Si aucun insert et one-per-row, rien de plus à faire (les raisons sont déjà posées)

    if args.dry_run:
        print(f"[DRY-RUN] Liens insérés (toutes colonnes, priorité respectée): {total_linked}")
    else:
        write_table(df, args.output)
        print(f"✅ Terminé. Liens insérés (toutes colonnes, priorité): {total_linked} → {args.output}")

if __name__ == "__main__":
    main()
