#!/usr/bin/env python3
"""
Закупівлі пального силовими та оборонними замовниками в Prozorro.

Конвеєр (кожен крок можна запускати окремо):
  probe    — перевірка: скільки результатів дає пошук на кожен CPV-код і текстовий запит
  search   — збір кандидатів через пошук prozorro.gov.ua (за датою оголошення, по місяцях)
  details  — повні дані тендерів з public-api.prozorro.gov.ua + договори з модуля contracting
  build    — CSV-таблиці, рейтинги, docs/index.html
  all      — search + details + build

Лише стандартна бібліотека Python 3.10+. Усі сирі відповіді API зберігаються в data/raw/
у вигляді .json.gz, щоб кожне число у звіті можна було перевірити до первинного джерела.
"""
import argparse
import csv
import datetime as dt
import gzip
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "data")
RAW_T = os.path.join(DATA, "raw", "tenders")
RAW_C = os.path.join(DATA, "raw", "contracts")
DOCS = os.path.join(ROOT, "docs")

SEARCH_URL = "https://prozorro.gov.ua/api/search/tenders"
SUMMARY_URL = "https://prozorro.gov.ua/api/tenders/{}/summary"
TENDER_URL = "https://public-api.prozorro.gov.ua/api/2.5/tenders/{}"
CONTRACT_URL = "https://public-api.prozorro.gov.ua/api/2.5/contracts/{}"
PUBLIC_TENDER = "https://prozorro.gov.ua/tender/{}"
SEARCH_CAP = 10000          # пошук віддає не більше 10 000 результатів на запит
HEX_RE = re.compile(r"^[0-9a-f]{32}$")
FINAL_STATUSES = {"complete", "cancelled", "unsuccessful"}

HEADERS = {"User-Agent": "prozorro-fuel-research/1.0 (journalism; github actions)",
           "Accept": "application/json"}
GAP = {"prozorro.gov.ua": 1.0, "public-api.prozorro.gov.ua": 0.15}   # пауза між запитами до хоста, с
_last = {}


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


# ------------------------------------------------------------------ HTTP
def http(url, body=None, tries=6, timeout=60):
    host = urllib.parse.urlparse(url).hostname
    headers = dict(HEADERS)
    data = None
    if body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    for i in range(tries):
        wait = _last.get(host, 0) + GAP.get(host, 0.2) - time.time()
        if wait > 0:
            time.sleep(wait)
        _last[host] = time.time()
        try:
            with urllib.request.urlopen(urllib.request.Request(url, data, headers), timeout=timeout) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            if e.code in (404, 410):
                return None
            if e.code in (400, 401, 403) or i == tries - 1:
                raise RuntimeError(f"HTTP {e.code} від {host}: {url} :: {e.read()[:300]!r}")
            time.sleep((15 if e.code == 429 else 3) * (i + 1))
        except (urllib.error.URLError, TimeoutError, OSError) as e:
            if i == tries - 1:
                raise RuntimeError(f"{host} недоступний: {e}")
            time.sleep(3 * (i + 1))


# ------------------------------------------------------------------ files
def gz_write(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with gzip.open(tmp, "wt", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, path)


def gz_read(path):
    try:
        with gzip.open(path, "rt", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, OSError, ValueError):
        return None


def json_read(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, ValueError):
        return default


def json_write(path, obj):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


def csv_write(path, rows, fields):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:   # BOM — щоб Excel відкрив кирилицю
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow(r)


# ------------------------------------------------------------------ config & classification
def load_config():
    cfg = json_read(os.path.join(ROOT, "config.json"), None)
    if not cfg:
        sys.exit("Немає config.json")
    for g in cfg["groups"]:
        g["_inc"] = [re.compile(p, re.I) for p in g.get("patterns", [])]
        g["_exc"] = [re.compile(p, re.I) for p in g.get("exclude", [])]
    cfg["_kw"] = [k.lower() for k in cfg.get("fuel_keywords", [])]
    return cfg


def load_overrides():
    """overrides/buyers.csv: edrpou,group,note — ручне віднесення замовника до групи (або group=exclude)."""
    out = {}
    path = os.path.join(ROOT, "overrides", "buyers.csv")
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                code = re.sub(r"\D", "", r.get("edrpou") or "")
                if code and (r.get("group") or "").strip():
                    out[code] = r["group"].strip()
    return out


def classify(cfg, name, kind=None, edrpou=None, overrides=None):
    """Повертає (group_id, причина) або (None, None)."""
    if overrides and edrpou and edrpou in overrides:
        g = overrides[edrpou]
        return (None, None) if g == "exclude" else (g, "ручне віднесення (overrides/buyers.csv)")
    n = " ".join((name or "").lower().replace("’", "'").split())
    for g in cfg["groups"]:
        if g["_inc"] and any(p.search(n) for p in g["_inc"]) and not any(p.search(n) for p in g["_exc"]):
            return g["id"], "назва замовника"
        if kind and kind in g.get("kind", []):
            return g["id"], f"kind = {kind}"
    return None, None


def is_fuel_item(cfg, item):
    cpv = ((item.get("classification") or {}).get("id") or "")
    if any(cpv.startswith(p) for p in cfg["cpv_fuel_prefixes"]):
        return True
    for a in item.get("additionalClassifications") or []:
        if any((a.get("id") or "").startswith(p) for p in cfg["cpv_fuel_prefixes"]):
            return True
    d = (item.get("description") or "").lower()
    return any(k in d for k in cfg["_kw"])


def edr(org):
    return re.sub(r"\s", "", str(((org or {}).get("identifier") or {}).get("id") or ""))


def amount(v):
    try:
        return float((v or {}).get("amount"))
    except (TypeError, ValueError):
        return None


# ------------------------------------------------------------------ search
def month_windows(start, end):
    d = dt.date.fromisoformat(start)
    end = dt.date.fromisoformat(end)
    while d <= end:
        nxt = (d.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
        yield d.isoformat(), min(nxt - dt.timedelta(days=1), end).isoformat()
        d = nxt


def search_page(base, start, end, page):
    body = dict(base)
    body["date"] = {"notice": {"start": start, "end": end}}
    return http(f"{SEARCH_URL}?page={page}", body) or {}


def search_window(base, start, end, sink, stats, label):
    """Збирає всі результати за вікно дат; якщо їх >= 10 000 — ділить вікно навпіл."""
    first = search_page(base, start, end, 1)
    total = int(first.get("total") or 0)
    s, e = dt.date.fromisoformat(start), dt.date.fromisoformat(end)
    if total >= SEARCH_CAP and s < e:
        mid = s + (e - s) // 2
        log(f"  {label} {start}..{end}: {total} ≥ ліміту, ділю вікно")
        search_window(base, start, mid.isoformat(), sink, stats, label)
        search_window(base, (mid + dt.timedelta(days=1)).isoformat(), end, sink, stats, label)
        return
    if total >= SEARCH_CAP:
        stats["truncated"].append({"query": label, "date": start, "total": total})
        log(f"  УВАГА: {label} {start}: {total} результатів за один день, отримаю лише перші {SEARCH_CAP}")
    per = int(first.get("per_page") or 20) or 20
    pages = min((total + per - 1) // per, SEARCH_CAP // per)
    got = 0
    for p in range(1, pages + 1):
        r = first if p == 1 else search_page(base, start, end, p)
        rows = r.get("data") or []
        if not rows:
            break
        for x in rows:
            sink(x, label)
        got += len(rows)
    stats["windows"].append({"query": label, "start": start, "end": end, "total": total, "fetched": got})
    log(f"  {label} {start}..{end}: {total} знайдено, отримано {got}")


def queries(cfg):
    yield "cpv", {"cpv": cfg["cpv_search"]}
    for t in cfg.get("text_search", []):
        yield f"text:{t}", {"text": t}


def cmd_probe(cfg, args):
    """Діагностика: показує сирий формат відповіді пошуку і кількість результатів на кожен код за місяць."""
    start = args.probe_from or cfg["date_from"]
    end = (dt.date.fromisoformat(start).replace(day=1) + dt.timedelta(days=32)).replace(day=1) - dt.timedelta(days=1)
    end = end.isoformat()
    r = search_page({"cpv": cfg["cpv_search"][:1]}, start, end, 1)
    sample = (r.get("data") or [None])[0]
    print("== Ключі відповіді пошуку:", sorted(r.keys()))
    print("== Перший запис (обрізано):")
    print(json.dumps(sample, ensure_ascii=False, indent=1)[:3000])
    print(f"\n== Кількість результатів за {start}..{end}")
    for code in cfg["cpv_search"]:
        t = search_page({"cpv": [code]}, start, end, 1).get("total")
        print(f"  CPV {code}: {t}")
    print(f"  усі коди разом: {search_page({'cpv': cfg['cpv_search']}, start, end, 1).get('total')}")
    for t in cfg.get("text_search", []):
        print(f"  текст «{t}»: {search_page({'text': t}, start, end, 1).get('total')}")
    if sample:
        tid = sample.get("tenderID")
        hexid = sample.get("id") if HEX_RE.match(str(sample.get("id") or "")) else None
        print("\n== hex id у пошуку:", hexid or "немає, буде запит до /summary")
        if tid and not hexid:
            print("   summary:", json.dumps(http(SUMMARY_URL.format(tid)), ensure_ascii=False)[:500])


def cmd_search(cfg, args):
    today = dt.date.today().isoformat()
    date_to = cfg.get("date_to") or today
    state = json_read(os.path.join(DATA, "state.json"), {})
    done = set(state.get("searched_months", []))
    refresh_from = (dt.date.today() - dt.timedelta(days=args.refresh_days)).isoformat()
    cand_path = os.path.join(DATA, "candidates.json")
    cands = json_read(cand_path, {})
    overrides = load_overrides()
    stats = {"windows": [], "truncated": [], "seen": 0, "kept": 0}

    def sink(x, label):
        stats["seen"] += 1
        tid = x.get("tenderID")
        if not tid:
            return
        pe = x.get("procuringEntity") or {}
        name, kind, code = pe.get("name"), pe.get("kind"), edr(pe)
        group, _ = classify(cfg, name, kind, code, overrides)
        # якщо пошук не повернув назву замовника — беремо кандидата, перевіримо на етапі details
        if group is None and name:
            return
        stats["kept"] += 1
        c = cands.setdefault(tid, {"tenderID": tid, "found_by": []})
        if HEX_RE.match(str(x.get("id") or "")):
            c["id"] = x["id"]
        c.update({"buyer": name, "buyer_edrpou": code or c.get("buyer_edrpou"), "kind": kind,
                  "dateModified": x.get("dateModified") or c.get("dateModified")})
        if label not in c["found_by"]:
            c["found_by"].append(label)

    for start, end in month_windows(cfg["date_from"], date_to):
        month = start[:7]
        if month in done and end < refresh_from and not args.full:
            continue
        log(f"Місяць {month}")
        for label, base in queries(cfg):
            search_window(base, start, end, sink, stats, label)
        if end < refresh_from:
            done.add(month)
        json_write(cand_path, cands)          # зберігаємо після кожного місяця — можна перезапускати

    state["searched_months"] = sorted(done)
    state["last_search"] = dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")
    json_write(os.path.join(DATA, "state.json"), state)
    log_path = os.path.join(DATA, "search_log.json")
    hist = json_read(log_path, [])
    hist.append({"run": state["last_search"], **stats})
    json_write(log_path, hist[-20:])
    log(f"Пошук завершено: переглянуто {stats['seen']}, кандидатів додано/оновлено {stats['kept']}, всього {len(cands)}")


# ------------------------------------------------------------------ details
def hex_of(c):
    if c.get("id"):
        return c["id"]
    s = http(SUMMARY_URL.format(c["tenderID"]))
    return (s or {}).get("id")


def cmd_details(cfg, args):
    cands = json_read(os.path.join(DATA, "candidates.json"), {})
    n_t = n_c = 0
    for i, (tid, c) in enumerate(sorted(cands.items())):
        path = os.path.join(RAW_T, tid + ".json.gz")
        cached = gz_read(path)
        fresh = cached and (cached.get("status") in FINAL_STATUSES or
                            (c.get("dateModified") and cached.get("dateModified", "") >= c["dateModified"]))
        if cached and fresh and not args.refetch:
            t = cached
        else:
            h = hex_of(c)
            if not h:
                log(f"  {tid}: не знайдено hex id")
                continue
            c["id"] = h
            r = http(TENDER_URL.format(h))
            if not r:
                continue
            t = r["data"]
            gz_write(path, t)
            n_t += 1
        # договори — повна версія з модуля contracting (з додатковими угодами і фактичною оплатою)
        for k in t.get("contracts") or []:
            if k.get("status") not in ("active", "terminated"):
                continue
            cpath = os.path.join(RAW_C, k["id"] + ".json.gz")
            cc = gz_read(cpath)
            if cc and not args.refetch:
                if cc.get("status") == "terminated":
                    continue                      # завершений договір більше не змінюється
                if time.time() - float(cc.get("_fetched") or 0) < args.contract_ttl * 86400:
                    continue                      # активні договори оновлюємо раз на N днів
            r = http(CONTRACT_URL.format(k["id"]))
            if r:
                r["data"]["_fetched"] = time.time()   # службова позначка часу завантаження
                gz_write(cpath, r["data"])
                n_c += 1
        if (i + 1) % 200 == 0:
            log(f"  оброблено {i + 1}/{len(cands)} тендерів")
            json_write(os.path.join(DATA, "candidates.json"), cands)
    json_write(os.path.join(DATA, "candidates.json"), cands)
    log(f"Деталі: завантажено тендерів {n_t}, договорів {n_c}")


# ------------------------------------------------------------------ build
def qty_summary(items):
    agg = defaultdict(float)
    for it in items:
        try:
            q = float(it.get("quantity") or 0)
        except (TypeError, ValueError):
            q = 0
        unit = ((it.get("unit") or {}).get("name") or (it.get("unit") or {}).get("code") or "од.").strip()
        if q:
            agg[unit] += q
    return "; ".join(f"{v:,.0f} {u}".replace(",", " ") for u, v in agg.items())


def lot_items(t, lot_id):
    return [i for i in t.get("items") or [] if not lot_id or i.get("relatedLot") == lot_id]


def cmd_build(cfg, args):
    cands = json_read(os.path.join(DATA, "candidates.json"), {})
    overrides = load_overrides()
    active_from = cfg["active_from"]
    date_from = cfg["date_from"]
    groups = {g["id"]: g["title"] for g in cfg["groups"]}
    tenders_rows, contract_rows, item_rows = [], [], []
    rejected = defaultdict(int)

    for tid in sorted(cands):
        t = gz_read(os.path.join(RAW_T, tid + ".json.gz"))
        if not t:
            rejected["немає даних тендера"] += 1
            continue
        pe = t.get("procuringEntity") or {}
        b_code, b_name = edr(pe), pe.get("name") or ""
        group, why = classify(cfg, b_name, pe.get("kind"), b_code, overrides)
        if not group:
            rejected["замовник поза групами"] += 1
            continue
        fuel_items = [i for i in t.get("items") or [] if is_fuel_item(cfg, i)]
        if not fuel_items:
            rejected["немає позицій пального"] += 1
            continue
        tval = amount(t.get("value"))
        notice = (t.get("date") or t.get("dateCreated") or "")[:10]
        awards = {a.get("id"): a for a in t.get("awards") or []}
        lots = {l.get("id"): l for l in t.get("lots") or []}
        tenders_rows.append({
            "tender_id": tid, "url": PUBLIC_TENDER.format(tid), "title": t.get("title"),
            "status": t.get("status"), "procedure": t.get("procurementMethodType"),
            "notice_date": notice, "tender_period_end": ((t.get("tenderPeriod") or {}).get("endDate") or "")[:10],
            "buyer": b_name, "buyer_edrpou": b_code, "buyer_kind": pe.get("kind"),
            "buyer_region": (pe.get("address") or {}).get("region"),
            "group": group, "group_title": groups.get(group, group), "group_reason": why,
            "expected_value": tval, "currency": (t.get("value") or {}).get("currency"),
            "vat_included": (t.get("value") or {}).get("valueAddedTaxIncluded"),
            "fuel_items": len(fuel_items), "all_items": len(t.get("items") or []),
            "quantity": qty_summary(fuel_items), "bids": len(t.get("bids") or []),
            "n_contracts": sum(1 for k in t.get("contracts") or [] if k.get("status") in ("active", "terminated")),
            "large": bool(tval and tval >= cfg["large_tender_uah"]),
        })

        for k in t.get("contracts") or []:
            if k.get("status") not in ("active", "terminated"):
                continue
            full = gz_read(os.path.join(RAW_C, k["id"] + ".json.gz")) or {}
            award = awards.get(k.get("awardID")) or {}
            lot_id = award.get("lotID")
            sup = ((full.get("suppliers") or k.get("suppliers") or award.get("suppliers") or [{}])[0]) or {}
            items = full.get("items") or k.get("items") or lot_items(t, lot_id)
            f_items = [i for i in items if is_fuel_item(cfg, i)]
            if not f_items:
                continue
            signed = (full.get("dateSigned") or k.get("dateSigned") or "")[:10]
            period = full.get("period") or k.get("period") or {}
            p_end = (period.get("endDate") or "")[:10]
            status = full.get("status") or k.get("status")
            v_init = amount(k.get("value"))
            v_cur = amount(full.get("value")) if full else None
            paid = amount(full.get("amountPaid"))
            v_rank = paid if (status == "terminated" and paid is not None) else (v_cur if v_cur is not None else v_init)
            active_2026 = bool((p_end and p_end >= active_from) or (not p_end and signed >= active_from))
            lot_val = amount((lots.get(lot_id) or {}).get("value")) if lot_id else tval
            row = {
                "tender_id": tid, "url": PUBLIC_TENDER.format(tid), "procedure": t.get("procurementMethodType"),
                "buyer": b_name, "buyer_edrpou": b_code, "group": group, "group_title": groups.get(group, group),
                "contract_id": k["id"], "contract_number": full.get("contractNumber") or k.get("contractNumber"),
                "contract_status": status, "date_signed": signed,
                "period_start": (period.get("startDate") or "")[:10], "period_end": p_end,
                "supplier": sup.get("name") or "не розкрито", "supplier_edrpou": edr(sup) or "",
                "value_initial": v_init, "value_current": v_cur, "amount_paid": paid, "value_rank": v_rank,
                "currency": ((full.get("value") or k.get("value") or {}).get("currency")),
                "vat_included": ((full.get("value") or k.get("value") or {}).get("valueAddedTaxIncluded")),
                "n_changes": len(full.get("changes") or []),
                "expected_value": lot_val, "fuel_items": len(f_items), "all_items": len(items),
                "mixed": len(f_items) < len(items), "quantity": qty_summary(f_items),
                "active_2026": active_2026, "signed_in_period": signed >= date_from,
                "contract_data": "contracting" if full else "лише з тендера",
                "large": bool(v_rank and v_rank >= cfg["large_contract_uah"]),
            }
            contract_rows.append(row)
            for it in f_items:
                u = it.get("unit") or {}
                item_rows.append({
                    "tender_id": tid, "contract_id": k["id"], "date_signed": signed, "group": group,
                    "buyer": b_name, "buyer_edrpou": b_code, "supplier": row["supplier"],
                    "supplier_edrpou": row["supplier_edrpou"], "description": it.get("description"),
                    "cpv": (it.get("classification") or {}).get("id"), "quantity": it.get("quantity"),
                    "unit": u.get("name") or u.get("code"), "unit_price": amount(u.get("value")),
                    "currency": (u.get("value") or {}).get("currency"),
                })

    # рейтинги
    def rank(rows):
        agg = {}
        for r in rows:
            if (r.get("currency") or "UAH") != "UAH" or r["value_rank"] is None:
                continue
            key = r["supplier_edrpou"] or ("? " + r["supplier"])
            a = agg.setdefault(key, {"supplier": r["supplier"], "supplier_edrpou": r["supplier_edrpou"],
                                     "total": 0.0, "contracts": 0, "buyers": set(), "groups": set(),
                                     "max_contract": 0.0, "names": set()})
            a["total"] += r["value_rank"]
            a["contracts"] += 1
            a["buyers"].add(r["buyer_edrpou"])
            a["groups"].add(r["group"])
            a["names"].add(r["supplier"])
            a["max_contract"] = max(a["max_contract"], r["value_rank"])
        grand = sum(a["total"] for a in agg.values()) or 1
        out = []
        for a in sorted(agg.values(), key=lambda x: -x["total"]):
            out.append({"supplier": a["supplier"], "supplier_edrpou": a["supplier_edrpou"],
                        "total": round(a["total"], 2), "share_pct": round(100 * a["total"] / grand, 2),
                        "contracts": a["contracts"], "buyers": len(a["buyers"]),
                        "groups": ", ".join(sorted(groups.get(g, g) for g in a["groups"])),
                        "max_contract": round(a["max_contract"], 2),
                        "other_names": " | ".join(sorted(a["names"] - {a["supplier"]}))})
        for i, r in enumerate(out, 1):
            r["rank"] = i
        return out

    in_period = [r for r in contract_rows if r["signed_in_period"]]
    rank_2026 = rank([r for r in in_period if r["active_2026"]])
    rank_all = rank(in_period)

    buyers = {}
    for r in tenders_rows:
        b = buyers.setdefault(r["buyer_edrpou"] or r["buyer"], {
            "buyer": r["buyer"], "buyer_edrpou": r["buyer_edrpou"], "group": r["group"],
            "group_title": r["group_title"], "group_reason": r["group_reason"], "kind": r["buyer_kind"],
            "tenders": 0, "contracts_value": 0.0})
        b["tenders"] += 1
    for r in in_period:
        if r["value_rank"] and (r.get("currency") or "UAH") == "UAH":
            b = buyers.get(r["buyer_edrpou"] or r["buyer"])
            if b:
                b["contracts_value"] += r["value_rank"]
    buyer_rows = sorted(buyers.values(), key=lambda x: (x["group"], -x["contracts_value"]))

    c_fields = ["tender_id", "url", "procedure", "buyer", "buyer_edrpou", "group", "group_title", "contract_id",
                "contract_number", "contract_status", "date_signed", "period_start", "period_end", "supplier",
                "supplier_edrpou", "value_initial", "value_current", "amount_paid", "value_rank", "currency",
                "vat_included", "n_changes", "expected_value", "fuel_items", "all_items", "mixed", "quantity",
                "active_2026", "signed_in_period", "contract_data", "large"]
    csv_write(os.path.join(DATA, "contracts.csv"), sorted(contract_rows, key=lambda r: -(r["value_rank"] or 0)), c_fields)
    csv_write(os.path.join(DATA, "tenders.csv"), sorted(tenders_rows, key=lambda r: -(r["expected_value"] or 0)),
              list(tenders_rows[0].keys()) if tenders_rows else ["tender_id"])
    csv_write(os.path.join(DATA, "items.csv"), item_rows,
              ["tender_id", "contract_id", "date_signed", "group", "buyer", "buyer_edrpou", "supplier",
               "supplier_edrpou", "description", "cpv", "quantity", "unit", "unit_price", "currency"])
    r_fields = ["rank", "supplier", "supplier_edrpou", "total", "share_pct", "contracts", "buyers", "groups",
                "max_contract", "other_names"]
    csv_write(os.path.join(DATA, "suppliers_active_2026.csv"), rank_2026, r_fields)
    csv_write(os.path.join(DATA, "suppliers_signed_since_start.csv"), rank_all, r_fields)
    csv_write(os.path.join(DATA, "buyers.csv"), buyer_rows,
              ["group", "group_title", "buyer", "buyer_edrpou", "kind", "group_reason", "tenders", "contracts_value"])

    meta = {
        "built": dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "date_from": date_from, "active_from": active_from,
        "large_contract_uah": cfg["large_contract_uah"], "large_tender_uah": cfg["large_tender_uah"],
        "groups": groups, "candidates": len(cands), "tenders": len(tenders_rows),
        "contracts": len(contract_rows), "rejected": dict(rejected),
        "search_log": (json_read(os.path.join(DATA, "search_log.json"), []) or [None])[-1],
    }
    json_write(os.path.join(DATA, "summary.json"), meta)

    payload = {"meta": meta, "contracts": contract_rows,
               "tenders": [r for r in tenders_rows], "buyers": buyer_rows}
    with open(os.path.join(ROOT, "report_template.html"), encoding="utf-8") as f:
        html = f.read()
    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    os.makedirs(DOCS, exist_ok=True)
    with open(os.path.join(DOCS, "index.html"), "w", encoding="utf-8") as f:
        f.write(html.replace("/*__DATA__*/null", blob))
    log(f"Звіт: тендерів {len(tenders_rows)}, договорів {len(contract_rows)}, "
        f"постачальників (діючі 2026) {len(rank_2026)}; відсіяно {dict(rejected)}")


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("command", choices=["probe", "search", "details", "build", "all"])
    ap.add_argument("--full", action="store_true", help="переглянути пошук за всі місяці заново")
    ap.add_argument("--refetch", action="store_true", help="перезавантажити всі тендери і договори")
    ap.add_argument("--refresh-days", type=int, default=60, help="скільки останніх днів шукати повторно")
    ap.add_argument("--contract-ttl", type=int, default=7, help="оновлювати активні договори раз на N днів")
    ap.add_argument("--probe-from", help="місяць для probe, напр. 2026-03-01")
    args = ap.parse_args()
    cfg = load_config()
    if args.command == "probe":
        cmd_probe(cfg, args)
    if args.command in ("search", "all"):
        cmd_search(cfg, args)
    if args.command in ("details", "all"):
        cmd_details(cfg, args)
    if args.command in ("build", "all"):
        cmd_build(cfg, args)


if __name__ == "__main__":
    main()
