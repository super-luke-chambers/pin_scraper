#!/usr/bin/env python3
"""Compare Superbet's published corner O/U prices with the latest Pinnacle corner line
and DM an alert when our vigged price differs from Pinnacle's by more than THRESHOLD.

Pinnacle side : outputs/pinnacle_odds_data.xlsx from github.com/super-luke-chambers/pin_scraper
                (scraped by GitHub Actions at 07/10/13/16 UTC).
Superbet side : public offer API (same endpoints as build_implied_cloud.py), market 704 = Corners,
                outcome 2043 = Over, 2042 = Under.

Comparison    : * Pinnacle line == one of our lines -> compare decimal odds directly (both vigged).
                * otherwise fit a Poisson mean to Pinnacle's line/odds, price Pinnacle at our nearby
                  lines and re-apply Pinnacle's own margin so the comparison stays vigged ("model").
gap = our_odds / pinnacle_odds - 1 ; negative = we are shorter (cheaper to the customer's win) than Pinnacle.

Usage: python pinnacle_vs_superbet_corners.py [--dry-run] [--all] [--days N] [--pin-file X.xlsx] [--no-state]
Run with the repo venv (needs openpyxl): .venv/bin/python pinnacle_vs_superbet_corners.py
"""
import argparse, difflib, json, math, os, re, subprocess, sys, unicodedata, urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

REPO = Path(__file__).parent
STATE_FILE = Path(os.environ.get("CORNERS_ALERT_STATE", REPO / "pinnacle_vs_superbet_corners_state.json"))
LOG_FILE = Path(os.environ.get("CORNERS_ALERT_LOG", REPO / "pinnacle_vs_superbet_corners.log"))

THRESHOLD = 0.05            # alert when |our/pinnacle - 1| > 5%
REALERT_MOVE = 0.03         # re-alert an open alert only if the gap moved by >= 3pp
MAX_PIN_AGE_H = 4           # ignore Pinnacle rows scraped longer ago (scrapes are 3h apart)
MODEL_LINE_WINDOW = 1.0     # model path: only our lines within +-1.0 of Pinnacle's line
ALERT_DM_USER_ID = "U04JV3WMNNP"   # Luke's DM (test target). Swap for C0C2M7NF7FT (#potupoge_alerts) once signed off.

PIN_REPO_FILE = "repos/super-luke-chambers/pin_scraper/contents/outputs/pinnacle_odds_data.xlsx"
# UEFA Nations League A-D (Superbet tournament ids = PROD.DWH.D_EVENT.TOURNAMENT_ID)
TOURNAMENT_IDS = list(range(32374, 32388))
HOST = ("ng-be", "en-BE")   # English team names -> easiest to match against Pinnacle
CORNERS, OVER_ID, UNDER_ID = 704, 2043, 2042

ALIASES = {"turkiye": "turkey", "czechia": "czech republic", "republic of ireland": "ireland",
           "north macedonia": "macedonia", "bosnia and herzegovina": "bosnia", "bosnia herzegovina": "bosnia",
           "faroe islands": "faroe", "kazakhstan": "kazakhstan", "usa": "united states"}


def log(msg):
    line = f"{datetime.now(timezone.utc):%Y-%m-%d %H:%M:%S} {msg}"
    print(line)
    with LOG_FILE.open("a") as f:
        f.write(line + "\n")


def norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\b(fc|cf|sc|the|national team)\b", " ", s)
    s = " ".join(s.split())
    return ALIASES.get(s, s)


def same_team(a, b):
    a, b = norm(a), norm(b)
    return a == b or difflib.SequenceMatcher(None, a, b).ratio() >= 0.82


def http_json(url):
    for _ in range(3):
        try:
            return json.load(urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=60))
        except Exception:
            pass
    return None


# ------------------------------------------------------------------ Pinnacle
def load_pinnacle(pin_file):
    from openpyxl import load_workbook
    if not pin_file:
        out = Path("/tmp/pinnacle_odds_data.xlsx")
        raw = subprocess.run(["gh", "api", PIN_REPO_FILE, "-H", "Accept: application/vnd.github.raw"],
                             capture_output=True, check=True, timeout=120).stdout
        out.write_bytes(raw)
        pin_file = str(out)
    ws = load_workbook(pin_file, read_only=True, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    hdr = {h: i for i, h in enumerate(rows[0])}
    now = datetime.now(timezone.utc)
    best, stale = {}, 0
    for r in rows[1:]:
        g = lambda k: r[hdr[k]]
        if g("Asian Corners Line") is None or g("Asian Corners O Odds") is None or g("Asian Corners U Odds") is None:
            continue
        scraped = g("Time of scraping")
        if not isinstance(scraped, datetime):
            scraped = datetime.fromisoformat(str(scraped))
        scraped = scraped.replace(tzinfo=timezone.utc)
        if now - scraped > timedelta(hours=MAX_PIN_AGE_H):
            stale += 1
            continue
        start = g("Match Start Time")
        start = start if isinstance(start, datetime) else datetime.strptime(str(start), "%Y-%m-%d %H:%M")
        key = g("Game URL")
        if key not in best or scraped > best[key]["scraped"]:
            best[key] = dict(home=g("Home"), away=g("Away"), league=g("League"), start=start.replace(tzinfo=timezone.utc),
                             line=float(g("Asian Corners Line")), o=float(g("Asian Corners O Odds")),
                             u=float(g("Asian Corners U Odds")), scraped=scraped)
    log(f"Pinnacle: {len(best)} matches with a fresh corners line ({stale} stale rows skipped)")
    return list(best.values())


# ------------------------------------------------------------------ Superbet
def load_superbet(days):
    h, loc = HOST
    now = datetime.now(timezone.utc)
    s, e = now.strftime("%Y-%m-%d") + "+00:00:00", (now + timedelta(days=days)).strftime("%Y-%m-%d") + "+00:00:00"
    base = f"https://production-superbet-offer-{h}.freetls.fastly.net/v2/{loc}/events"
    r = http_json(f"{base}/by-date?currentStatus=active&offerState=prematch&tournamentIds={','.join(map(str, TOURNAMENT_IDS))}&startDate={s}&endDate={e}")
    if r is None:
        raise RuntimeError("Superbet events/by-date failed")
    out = []
    for ev in r.get("data", []):
        d = http_json(f"{base}/{ev['eventId']}")
        if not d or not d.get("data"):
            continue
        lines = {}
        for o in d["data"][0].get("odds") or []:
            if o.get("marketId") == CORNERS and o.get("status") == "active":
                lines.setdefault(float(o["specialBetValue"]), {})["o" if o["outcomeId"] == OVER_ID else "u"] = o["price"]
        lines = {k: v for k, v in sorted(lines.items()) if "o" in v and "u" in v}
        home, _, away = ev["matchName"].partition("·")
        out.append(dict(id=ev["eventId"], home=home.strip(), away=away.strip(), tid=ev["tournamentId"],
                        start=datetime.strptime(ev["matchDate"], "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc), lines=lines))
    log(f"Superbet: {len(out)} Nations League events, {sum(1 for x in out if x['lines'])} with corner lines")
    return out


# ------------------------------------------------------------------ model
def pois_cdf(k, lam):
    if k < 0:
        return 0.0
    pm = c = math.exp(-lam)
    for i in range(1, int(k) + 1):
        pm *= lam / i
        c += pm
    return c


def pois_pmf(k, lam):
    return math.exp(-lam + k * math.log(lam) - math.lgamma(k + 1))


def p_over_single(lam, line):
    """Asian-style single line: x.5 clean, integer line pushes (half stake back)."""
    if line % 1 == 0.5:
        return 1 - pois_cdf(line - 0.5, lam)
    return (1 - pois_cdf(line, lam)) + 0.5 * pois_pmf(int(line), lam)


def p_over(lam, line):
    f = line % 1
    if abs(f - 0.25) < 1e-9:
        return 0.5 * (p_over_single(lam, line - 0.25) + p_over_single(lam, line + 0.25))
    if abs(f - 0.75) < 1e-9:
        return 0.5 * (p_over_single(lam, line - 0.25) + p_over_single(lam, line + 0.25))
    return p_over_single(lam, line)


def fit_lambda(line, o, u):
    target = (1 / o) / (1 / o + 1 / u)
    lo, hi = 1.0, 30.0
    for _ in range(60):
        mid = (lo + hi) / 2
        if p_over(mid, line) < target:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


# ------------------------------------------------------------------ compare
def compare(pin_rows, sb_events):
    results, unmatched_pin = [], []
    used = set()
    for p in pin_rows:
        cands = [e for e in sb_events if e["id"] not in used and abs((e["start"] - p["start"]).total_seconds()) <= 36 * 3600
                 and same_team(e["home"], p["home"]) and same_team(e["away"], p["away"])]
        if not cands:
            unmatched_pin.append(p)
            continue
        e = min(cands, key=lambda x: abs((x["start"] - p["start"]).total_seconds()))
        used.add(e["id"])
        if not e["lines"]:
            continue
        margin = 1 / p["o"] + 1 / p["u"] - 1
        if p["line"] in e["lines"]:
            targets = [(p["line"], e["lines"][p["line"]], p["o"], p["u"], "direct")]
        else:
            lam = fit_lambda(p["line"], p["o"], p["u"])
            targets = []
            for ln, v in e["lines"].items():
                if abs(ln - p["line"]) > MODEL_LINE_WINDOW:
                    continue
                po = p_over(lam, ln)
                if not .1 < po < .9:
                    continue
                targets.append((ln, v, 1 / (po * (1 + margin)), 1 / ((1 - po) * (1 + margin)), "model"))
        for ln, v, po, pu, how in targets:
            for side, ours, theirs in (("Over", v["o"], po), ("Under", v["u"], pu)):
                results.append(dict(event=e["id"], match=f"{e['home']} v {e['away']}", start=e["start"], line=ln, side=side,
                                    ours=ours, pin=round(theirs, 3), gap=ours / theirs - 1, how=how,
                                    pin_line=p["line"], pin_o=p["o"], pin_u=p["u"]))
    unmatched_sb = [e for e in sb_events if e["id"] not in used and e["lines"]]
    return results, unmatched_pin, unmatched_sb


# ------------------------------------------------------------------ slack / state
def send_dm(text):
    if os.environ.get("SLACK_BOT_TOKEN"):   # GitHub Actions path: bot token, no claude CLI
        from slack_sdk import WebClient
        WebClient(token=os.environ["SLACK_BOT_TOKEN"]).chat_postMessage(channel=ALERT_DM_USER_ID, text=text)
        return
    prompt = f"Send a Slack message to channel_id {ALERT_DM_USER_ID} with this exact text:\n\n{text}"
    subprocess.run(["claude", "-p", prompt, "--permission-mode", "bypassPermissions",
                    "--allowedTools", "mcp__claude_ai_Slack__slack_send_message"],
                   capture_output=True, text=True, timeout=180)


def fmt(rows):
    lines = [":soccer: *Corners: Superbet vs Pinnacle* (gap > %d%%, our odds vs Pinnacle vigged)" % round(THRESHOLD * 100)]
    for ev, grp in sorted({r["event"]: [] for r in rows}.items(), key=lambda kv: kv[0]):
        g = [r for r in rows if r["event"] == ev]
        lines.append(f"*{g[0]['match']}* - {g[0]['start']:%d/%m %H:%M} UTC (Pinnacle line {g[0]['pin_line']:g}, {g[0]['pin_o']}/{g[0]['pin_u']})")
        for r in sorted(g, key=lambda r: (r["line"], r["side"])):
            lines.append(f"  • {r['side']} {r['line']:g}: ours {r['ours']:.2f} vs Pin {r['pin']:.2f} ({r['gap']*100:+.1f}%)" + (" _[model]_" if r["how"] == "model" else ""))
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--all", action="store_true", help="print the full comparison table")
    ap.add_argument("--days", type=int, default=2)
    ap.add_argument("--pin-file")
    ap.add_argument("--no-state", action="store_true")
    a = ap.parse_args()
    try:
        results, un_pin, un_sb = compare(load_pinnacle(a.pin_file), load_superbet(a.days))
    except Exception as ex:
        log(f"FAILED: {ex!r}")
        if not a.dry_run:
            send_dm(f":warning: Pinnacle-vs-Superbet corners check failed: {ex!r}"[:500])
        raise
    for p in un_pin:
        log(f"UNMATCHED Pinnacle: {p['home']} v {p['away']} {p['start']:%Y-%m-%d %H:%M} ({p['league']})")
    for e in un_sb:
        log(f"UNMATCHED Superbet: {e['home']} v {e['away']} {e['start']:%Y-%m-%d %H:%M} (event {e['id']})")
    if a.all:
        for r in results:
            print(f"{r['match']:40} {r['side']:5} {r['line']:5g} ours {r['ours']:6.2f} pin {r['pin']:6.2f} {r['gap']*100:+6.1f}% [{r['how']}]")
    breaches = [r for r in results if abs(r["gap"]) > THRESHOLD]
    state = {} if a.no_state or not STATE_FILE.exists() else json.loads(STATE_FILE.read_text())
    new_state, to_send = {}, []
    for r in breaches:
        k = f"{r['event']}|{r['line']}|{r['side']}"
        new_state[k] = r["gap"]
        if k not in state or abs(r["gap"] - state[k]) >= REALERT_MOVE:
            to_send.append(r)
    log(f"{len(results)} comparisons, {len(breaches)} breaches, {len(to_send)} new/moved alerts")
    if to_send:
        msg = fmt(to_send)
        print(msg)
        if not a.dry_run:
            send_dm(msg)
    if not a.dry_run and not a.no_state:
        STATE_FILE.write_text(json.dumps(new_state, indent=2))


if __name__ == "__main__":
    main()
