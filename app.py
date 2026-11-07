"""
Car market dashboard (Streamlit) - two datasets, one page each:
  * Chinese-spec cars, UAE (DubiCars)        -> dubicars_chinese.xlsx
  * Luxury cars, Canada + UAE (Autotrader/Dubizzle) -> luxury_cars.xlsx

Runs on Streamlit Community Cloud. Reads the Excel files from Google Drive -
the daily 07:00 run on the VPS replaces those same Drive files, so the
dashboard shows the latest scrape (re-checked every 15 min, or ↻ Refresh).

Secrets (Streamlit Cloud -> app -> Settings -> Secrets):
    drive_link         = "<Drive link of dubicars_chinese.xlsx>"
    drive_link_luxury  = "<Drive link of luxury_cars.xlsx>"
    password           = "..."          # optional
Without the links it looks for the files next to this app.
"""
import hmac
import io
import os
import re

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Car Market Dashboard", page_icon="🚗", layout="wide",
                   initial_sidebar_state="collapsed")

HERE = os.path.dirname(os.path.abspath(__file__))
DATASETS = {
    "Chinese-spec cars (UAE)": dict(secret="drive_link", file="dubicars_chinese.xlsx",
                                    kind="dubicars", currencies=["AED", "USD"],
                                    title="🚗 Chinese-spec cars — UAE market"),
    "Luxury cars (Canada & UAE)": dict(secret="drive_link_luxury", file="luxury_cars.xlsx",
                                       kind="luxury", currencies=["USD", "CAD", "AED"],
                                       title="💎 Luxury cars — Canada & UAE"),
}
TOP_N = 10
PEG = 3.6725                  # 1 USD = 3.6725 AED (UAE fixed peg)
UP, DOWN, FLAT = "#d1495b", "#2a9d8f", "#8d99ae"   # price up = red, down = green

st.markdown("""<style>
.block-container {padding-top: 3rem; padding-bottom: 1rem;}
[data-baseweb="tab"] p {font-size: 1.05rem; font-weight: 600;}
[data-testid="stMetricValue"] {font-size: 2rem;}
</style>""", unsafe_allow_html=True)


def secret(name):
    try:
        return st.secrets.get(name, "")
    except Exception:
        return ""


# ------------------------------------------------------------------ password
def password_ok():
    pw = secret("password")
    if not pw or st.session_state.get("auth"):
        return True
    with st.form("login"):
        st.subheader("🔒 Car market dashboard")
        typed = st.text_input("Password", type="password")
        if st.form_submit_button("Open"):
            if hmac.compare_digest(typed, pw):
                st.session_state["auth"] = True
                st.rerun()
            st.error("Wrong password")
    return False


if not password_ok():
    st.stop()


# ------------------------------------------------------------------ exchange rate
@st.cache_data(ttl=6 * 3600, show_spinner=False)
def live_rate(symbol):
    """today's market rate: 1 USD = ? <symbol>  (None if the lookup fails)"""
    import requests
    for url, pick in [("https://open.er-api.com/v6/latest/USD",
                       lambda j: (j["rates"][symbol], j.get("time_last_update_utc", "")[:16])),
                      (f"https://api.frankfurter.dev/v1/latest?base=USD&symbols={symbol}",
                       lambda j: (j["rates"][symbol], j.get("date", "")))]:
        try:
            r = requests.get(url, timeout=10)
            r.raise_for_status()
            rate, when = pick(r.json())
            return float(rate), when, url.split("/")[2]
        except Exception:
            continue
    return None


# ------------------------------------------------------------------ data
@st.cache_data(show_spinner="Getting the latest data...", ttl=900)
def fetch_drive(link):
    """download the Drive file (shared as 'Anyone with the link')"""
    import requests
    m = re.search(r"/d/([\w-]{20,})", link) or re.search(r"id=([\w-]{20,})", link)
    fid = m.group(1) if m else link.strip()
    s = requests.Session()
    r = s.get("https://drive.google.com/uc", params={"export": "download", "id": fid},
              timeout=120)
    if r.content[:2] != b"PK":                         # Drive's "can't scan" page
        tok = re.search(r'name="confirm" value="([^"]+)"', r.text)
        uuid = re.search(r'name="uuid" value="([^"]+)"', r.text)
        r = s.get("https://drive.usercontent.google.com/download",
                  params={"id": fid, "export": "download", "confirm": tok.group(1) if tok else "t",
                          **({"uuid": uuid.group(1)} if uuid else {})}, timeout=120)
    if r.status_code != 200 or r.content[:2] != b"PK":
        raise RuntimeError("Drive did not return the Excel file - is it shared as "
                           "'Anyone with the link'?")
    when = r.headers.get("Last-Modified") or r.headers.get("Date")
    return r.content, when


def read_sheet(xl, name, cols=None):
    """read one sheet fast (calamine when installed), optionally only some columns"""
    kw = {"usecols": (lambda c: c in cols)} if cols else {}
    return pd.read_excel(xl, name, **kw)


def open_book(src):
    data = io.BytesIO(src) if isinstance(src, bytes) else src
    try:
        return pd.ExcelFile(data, engine="calamine")
    except Exception:
        if hasattr(data, "seek"):
            data.seek(0)
        return pd.ExcelFile(data)


LUX_COLS = {"source", "listing_id", "url", "make", "model", "year", "seller_name", "country",
            "price_usd", "price_aed", "price_cad", "status", "first_seen_at", "last_seen_at",
            "removed_at"}
LUX_HIST_COLS = {"source", "listing_id", "car_listing_id", "price", "recorded_at", "currency"}


def _num(s):
    return pd.to_numeric(s, errors="coerce")


def _ratio(a, b):
    ok = (a > 0) & (b > 0)
    return float((a[ok] / b[ok]).median()) if ok.any() else None


def read_dubicars(xl):
    df = pd.read_excel(xl, "listings")
    hist = pd.read_excel(xl, "price_history") if "price_history" in xl.sheet_names \
        else pd.DataFrame(columns=["car_id", "checked_at", "price_aed"])
    key = "listing_id" if "listing_id" in df.columns and "listing_id" in hist.columns else "car_id"
    col = lambda c, d=None: df[c] if c in df.columns else pd.Series(d, index=df.index)
    df["_key"] = df[key].astype(str)
    aed, usd = _num(col("price_aed")), _num(col("price_usd"))
    aed_rate = _ratio(aed, usd) or PEG
    df["price"] = usd.where(usd > 0, aed / aed_rate)                     # base = USD
    df["ad_status"] = col("ad_status", "").fillna("").astype(str).str.strip().replace("", "Active")
    df["market"] = "UAE"
    h = pd.DataFrame({"_key": hist[key].astype(str),
                      "checked_at": pd.to_datetime(hist["checked_at"], errors="coerce"),
                      "price": _num(hist["price_aed"]) / aed_rate})
    return df, h, {"USD": 1.0, "AED": aed_rate}, "DubiCars"


def read_luxury(xl):
    df = read_sheet(xl, "listings", LUX_COLS)
    hist = read_sheet(xl, "price_history", LUX_HIST_COLS)
    df["_key"] = df["source"].astype(str) + "|" + df["listing_id"].astype(str)
    usd = _num(df["price_usd"])
    rates = {"USD": 1.0, "AED": _ratio(_num(df["price_aed"]), usd) or PEG,
             "CAD": _ratio(_num(df["price_cad"]), usd) or 1.39}
    df["price"] = usd
    df["ad_status"] = df["status"].fillna("Active").astype(str).str.strip()
    df["market"] = df["country"].map({"CA": "Canada", "AE": "UAE"}).fillna(df["country"])
    cur = hist["currency"].fillna("USD").str.upper().map(rates).fillna(1.0)
    # car_listing_id = the car's current ad, so a re-posted car keeps one price line
    hid = hist["car_listing_id"] if "car_listing_id" in hist.columns else hist["listing_id"]
    h = pd.DataFrame({"_key": hist["source"].astype(str) + "|" + hid.astype(str),
                      "checked_at": pd.to_datetime(hist["recorded_at"], errors="coerce"),
                      "price": _num(hist["price"]) / cur})
    src = " + ".join(sorted({"autotrader_ca": "Autotrader.ca", "dubizzle_ae": "Dubizzle"}
                            .get(x, x) for x in df["source"].dropna().unique()))
    return df, h, rates, src


@st.cache_data(show_spinner="Loading...")
def load(src, stamp, kind):                 # stamp in the key -> reloads when the file changes
    xl = open_book(src)
    df, hist, rates, source = (read_luxury if kind == "luxury" else read_dubicars)(xl)
    col = lambda c, d=None: df[c] if c in df.columns else pd.Series(d, index=df.index)
    for c in ["first_seen_at", "last_seen_at"]:
        df[c] = pd.to_datetime(col(c), errors="coerce")
    df["year"] = _num(col("year"))
    for c in ["make", "model", "seller_name"]:
        df[c] = col(c, "").fillna("").astype(str).str.strip()
    df["car"] = (df["make"].str.title() + " " + df["model"].str.title() + " "
                 + df["year"].astype("Int64").astype(str)).str.strip()
    hist = hist.dropna(subset=["checked_at", "price"])
    hist = hist[hist["price"] > 0]

    # one price per ad per day, carried forward until the ad was last seen
    # (price_history only keeps a row when the price moves)
    h = hist.assign(day=hist["checked_at"].dt.normalize()).sort_values("checked_at")
    h = h.groupby(["_key", "day"])["price"].last().reset_index()
    # an ad that is now "price on request" stops at its last real price
    idx = df.set_index("_key")
    end = idx["last_seen_at"].dt.normalize().where(idx["price"].notna())
    end = end[~end.index.duplicated()]
    span = h.groupby("_key")["day"].agg(["min", "max"])
    span["end"] = pd.concat([span["max"], end.reindex(span.index)], axis=1).max(axis=1)
    n = ((span["end"] - span["min"]).dt.days + 1).clip(lower=1).astype(int)
    grid = pd.DataFrame({"_key": span.index.repeat(n), "day": span["min"].repeat(n).values})
    grid["day"] = grid["day"] + pd.to_timedelta(grid.groupby("_key").cumcount(), unit="D")
    daily = grid.merge(h, how="left")
    daily["price"] = daily.groupby("_key")["price"].ffill()
    daily = daily.merge(df[["_key", "car", "make", "seller_name"]], on="_key")
    return df, daily, rates, source


# ------------------------------------------------------------------ one tab per dataset
def render(DS):
    # ------------------------------------------------------------------ dataset
    try:
        link = secret(DS["secret"])
        if link:
            content, when = fetch_drive(link)
            df, daily, RATES, SOURCE = load(content, len(content), DS["kind"])
            updated = pd.to_datetime(when, errors="coerce", utc=True)
        else:
            path = next((p for p in [os.path.join(HERE, DS["file"]),
                                     os.path.join(HERE, "output", DS["file"]),
                                     os.path.join(HERE, "..", "output", DS["file"])]
                         if os.path.exists(p)), None)
            if not path:
                st.error(f"No data for this dataset: add **{DS['secret']}** (the Drive link of "
                         f"{DS['file']}) in the app Secrets.")
                return
            df, daily, RATES, SOURCE = load(path, os.path.getmtime(path), DS["kind"])
            updated = pd.Timestamp(os.path.getmtime(path), unit="s", tz="UTC")
    except Exception as e:
        st.error(f"Couldn't load the data: {e}")
        return
    updated = updated.tz_convert("Asia/Amman") if pd.notna(updated) else None
    last_scrape = df["last_seen_at"].max()

    # ------------------------------------------------------------------ header + filter
    h1, h2 = st.columns([3, 2])
    h1.title(DS["title"])
    h1.caption(f"{SOURCE} · last scrape {last_scrape:%d %b %Y, %H:%M} · data since "
               f"{df['first_seen_at'].min():%d %b %Y}"
               + (f" · file from {updated:%d %b %H:%M}" if updated is not None else ""))
    h2.write("")
    b1, b2 = h2.columns([2, 1])
    CUR = b1.radio("Currency", DS["currencies"], horizontal=True, key=f"{DS['kind']}_cur")
    b2.write("")
    if b2.button("↻ Refresh data", help="Read the latest file from Google Drive now",
                  key=f"{DS['kind']}_refresh"):
        st.cache_data.clear()
        st.rerun()

    # ---- filters: (market) -> make -> model (only models of the chosen makes) -> year
    df["model_t"] = df["model"].str.title()
    markets = sorted(m for m in df["market"].dropna().unique() if m)
    if len(markets) > 1:
        f0, f1, f2, f3 = st.columns([0.8, 1, 1.4, 1])
        pick_market = f0.multiselect("Market", markets, key=f"{DS['kind']}_market", placeholder="All markets")
    else:
        f1, f2, f3 = st.columns([1, 1.4, 1])
        pick_market = []
    sub = df[df["market"].isin(pick_market)] if pick_market else df
    makes = sorted(m for m in sub["make"].unique() if m)
    pick = f1.multiselect("Make", makes, key=f"{DS['kind']}_make", placeholder="All makes")
    if pick:
        sub = sub[sub["make"].isin(pick)]
    models = sorted(m for m in sub["model_t"].unique() if m)
    pick_model = f2.multiselect("Model", models, key=f"{DS['kind']}_model", placeholder="All models")
    if pick_model:
        sub = sub[sub["model_t"].isin(pick_model)]
    years = sorted(int(y) for y in sub["year"].dropna().unique())
    pick_year = f3.multiselect("Year", years, key=f"{DS['kind']}_year", placeholder="All years")

    # prices are stored in USD -> show in the chosen currency
    FX = RATES.get(CUR, 1.0)
    df = df.assign(price=df["price"] * FX)
    daily = daily.assign(price=daily["price"] * FX)
    if pick_market:
        df = df[df["market"].isin(pick_market)]
    if pick:
        df = df[df["make"].isin(pick)]
    if pick_model:
        df = df[df["model_t"].isin(pick_model)]
    if pick_year:
        df = df[df["year"].isin(pick_year)]
    daily = daily[daily["_key"].isin(df["_key"])]
    if df.empty:
        st.warning("No cars match these filters.")
        return

    live = df[df["ad_status"] == "Active"]
    grain = "M"                      # monthly view
    grain_txt = "month" if grain == "M" else "week"
    daily["period"] = daily["day"].dt.to_period(grain).dt.start_time

    # per ad: price in each period = its last price in that period
    per = daily.sort_values("day").groupby(["_key", "period"]).last().reset_index()
    periods = sorted(per["period"].unique())
    cur_p = periods[-1] if periods else None
    prev_p = periods[-2] if len(periods) > 1 else None


    def compare(by):
        """median price per group this period vs last period, only ads online in BOTH."""
        if prev_p is None:
            return pd.DataFrame()
        two = per[per["period"].isin([prev_p, cur_p])]
        both = two.groupby("_key")["period"].nunique()
        two = two[two["_key"].isin(both[both == 2].index)]
        t = two.pivot_table(index=by, columns="period", values="price", aggfunc="median")
        t.columns = ["before", "now"]
        t["change_pct"] = (t["now"] / t["before"] - 1) * 100
        return t.reset_index()


    ads_cmp = compare("_key")
    cut = int((ads_cmp["change_pct"] < -0.5).sum()) if len(ads_cmp) else 0
    rise = int((ads_cmp["change_pct"] > 0.5).sum()) if len(ads_cmp) else 0
    # typos like 1 -> 166,000 would wreck an average, so each ad's change is capped at ±50%
    avg_chg = ads_cmp["change_pct"].clip(-50, 50).mean() if len(ads_cmp) else None

    # ------------------------------------------------------------------ KPIs
    def money(v):
        if pd.isna(v):
            return "—"
        return f"{v / 1e6:,.2f}M {CUR}" if v >= 1e6 else f"{v / 1e3:,.1f}K {CUR}"


    k = st.columns(5)
    k[0].metric("Cars for sale now", f"{len(live):,}")
    med = live["price"].median()
    k[1].metric("Median price", money(med),
                help="The middle price: half the cars cost less, half cost more.")
    if avg_chg is None:
        k[2].metric(f"Prices vs last {grain_txt}", "—")
    else:
        word = "Up" if avg_chg > 0.05 else "Down" if avg_chg < -0.05 else "Flat"
        k[2].metric(f"Prices vs last {grain_txt}", word, f"{avg_chg:+.1f}%",
                    delta_color="inverse" if word != "Flat" else "off")
    k[3].metric("Cars that got cheaper", f"{cut:,}", f"{rise:,} got more expensive",
                delta_color="off")

    rate_cur = CUR if CUR != "USD" else next(c for c in DS["currencies"] if c != "USD")
    file_rate = RATES.get(rate_cur, 1.0)
    lr = live_rate(rate_cur)
    pegged = rate_cur == "AED"
    if lr:
        rate, when, src = lr
        diff = (file_rate / rate - 1) * 100
        k[4].metric("1 USD =", f"{file_rate:.4f} {rate_cur}", f"today {rate:.4f} ({diff:+.2f}%)",
                    delta_color="off",
                    help=f"Rate used in our data: {file_rate:.4f} {rate_cur}"
                         + (" (UAE fixed peg - the AED is pegged to the USD, so the difference is ~0)."
                            if pegged else " (the scraper's rate at the time it saved the prices).")
                         + f" Market rate today from {src} ({when}).")
    else:
        k[4].metric("1 USD =", f"{file_rate:.4f} {rate_cur}",
                    "UAE fixed peg" if pegged else "rate used in our data", delta_color="off",
                    help="Today's market rate couldn't be loaded right now.")

    st.divider()

    # ------------------------------------------------------------------ price over time
    top_cars = live["car"].value_counts().head(TOP_N).index.tolist()
    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        st.subheader(f"Price per {grain_txt} — most listed cars")
        chosen = st.multiselect("Cars", sorted(df["car"].unique()), default=top_cars[:5],
                                key=f"{DS['kind']}_cars_{'|'.join(pick + pick_model + [str(y) for y in pick_year] + pick_market)}",
                                label_visibility="collapsed")
        line = per[per["car"].isin(chosen)].groupby(["period", "car"])["price"]\
            .median().reset_index()
        if line.empty:
            st.info("Pick a car above.")
        else:
            fig = px.line(line, x="period", y="price", color="car", markers=True)
            fig.update_layout(height=420, margin=dict(l=0, r=0, t=10, b=0), legend_title=None,
                              legend=dict(orientation="h", y=-0.15), hovermode="x unified")
            fig.update_xaxes(title=None, tickformat="%d %b" if grain == "W" else "%b %Y",
                             dtick=None if grain == "W" else "M1")
            fig.update_yaxes(title=CUR, tickformat=",.0f")
            st.plotly_chart(fig, width="stretch")
        st.caption(f"Each point = the car's median price that month in {CUR} (each ad's last "
                   "price in the month). The current month is month-to-date.")

    with c2:
        st.subheader(f"Top {TOP_N} cars: median price before vs now ({CUR})")
        t = compare("car")
        t = t[t["car"].isin(top_cars)] if len(t) else t
        if t.empty:
            st.info(f"Needs 2 {grain_txt}s of data.")
        else:
            t["ads"] = t["car"].map(live["car"].value_counts())
            t = t.sort_values("ads", ascending=False)
            t[" "] = t["change_pct"].map(lambda x: "⬆️" if x > 0.5 else "⬇️" if x < -0.5 else "➖")
            bl = f"{prev_p:%b}" if grain == "M" else f"wk {prev_p:%d %b}"
            nl = f"{cur_p:%b}" if grain == "M" else f"wk {cur_p:%d %b}"
            st.dataframe(t[["car", "ads", "before", "now", "change_pct", " "]],
                         hide_index=True, width="stretch", height=420,
                         column_config={
                             "car": "Car", "ads": "Ads",
                             "before": st.column_config.NumberColumn(bl, format="%d"),
                             "now": st.column_config.NumberColumn(nl, format="%d"),
                             "change_pct": st.column_config.NumberColumn("Change", format="%+.1f%%")})

    st.divider()

    # ------------------------------------------------------------------ most listed
    c3, c4 = st.columns(2, gap="large")
    with c3:
        st.subheader("Most listed cars")
        m = live["car"].value_counts().head(TOP_N).rename_axis("car").reset_index(name="ads")
        fig = px.bar(m, x="ads", y="car", orientation="h", text="ads")
        fig.update_layout(height=380, margin=dict(l=0, r=0, t=10, b=0))
        fig.update_yaxes(categoryorder="total ascending", title=None)
        fig.update_xaxes(title=None, showticklabels=False)
        st.plotly_chart(fig, width="stretch")
    with c4:
        st.subheader("Top dealers")
        d = live[live["seller_name"] != ""]["seller_name"].value_counts().head(TOP_N)\
            .rename_axis("dealer").reset_index(name="ads")
        fig = px.bar(d, x="ads", y="dealer", orientation="h", text="ads",
                     color_discrete_sequence=["#5b6c8f"])
        fig.update_layout(height=380, margin=dict(l=0, r=0, t=10, b=0))
        fig.update_yaxes(categoryorder="total ascending", title=None)
        fig.update_xaxes(title=None, showticklabels=False)
        st.plotly_chart(fig, width="stretch")

    # ------------------------------------------------------------------ price changes per ad
    def last_changes(daily, df):
        """per ad: its last price change - old price, since when, changed on, days at each price"""
        d = daily.dropna(subset=["price"]).sort_values(["_key", "day"])[["_key", "day", "price"]]
        if d.empty:
            return pd.DataFrame()
        prev = d.groupby("_key")["price"].shift()
        d["new_run"] = prev.isna() | ((d["price"] - prev).abs() > 0.5)
        d["run"] = d.groupby("_key")["new_run"].cumsum()
        runs = d.groupby(["_key", "run"]).agg(price=("price", "first"),
                                              start=("day", "min"), end=("day", "max")).reset_index()
        n = runs.groupby("_key")["run"].transform("max")
        runs = runs[n >= 2]
        if runs.empty:
            return pd.DataFrame()
        last = runs.groupby("_key").tail(2)
        old = last.groupby("_key").head(1).set_index("_key")
        new = last.groupby("_key").tail(1).set_index("_key")
        first_seen = df.set_index("_key")["first_seen_at"].dt.normalize()
        out = pd.DataFrame({
            "old_price": old["price"],
            "since": pd.concat([old["start"], first_seen.reindex(old.index)], axis=1).min(axis=1)
                     .where(old["run"] == 1, old["start"]),
            "changed_on": new["start"],
            "new_price": new["price"],
            "last_seen": new["end"],
            "changes": runs.groupby("_key")["run"].max() - 1,
        })
        out["days_at_old_price"] = (out["changed_on"] - out["since"]).dt.days
        out["days_at_new_price"] = (out["last_seen"] - out["changed_on"]).dt.days
        out["change_pct"] = (out["new_price"] / out["old_price"] - 1) * 100
        return out.reset_index()


    with st.expander("Price changes per ad: old price, since when, when it changed", expanded=False):
        lc = last_changes(daily, df)
        if lc.empty:
            st.info("No ad has changed its price yet.")
        else:
            cols = ["_key", "car", "seller_name"] + (["url"] if "url" in df.columns else [])
            lc = lc.merge(df[cols], on="_key")
            o1, o2 = st.columns([2, 1])
            order = o1.radio("Sort by", ["Most recent change", "Biggest change %", "Longest at old price"],
                             key=f"{DS['kind']}_sort",
                             horizontal=True)
            direction = o2.radio("Show", ["All", "⬇️ cheaper", "⬆️ pricier"], horizontal=True,
                                 key=f"{DS['kind']}_show")
            if direction == "⬇️ cheaper":
                lc = lc[lc["change_pct"] < 0]
            elif direction == "⬆️ pricier":
                lc = lc[lc["change_pct"] > 0]
            if order == "Most recent change":
                lc = lc.sort_values("changed_on", ascending=False)
            elif order == "Biggest change %":
                lc = lc.reindex(lc["change_pct"].abs().sort_values(ascending=False).index)
            else:
                lc = lc.sort_values("days_at_old_price", ascending=False)
            show = ["car", "seller_name", "old_price", "since", "changed_on", "days_at_old_price",
                    "new_price", "days_at_new_price", "change_pct", "changes"] + \
                   (["url"] if "url" in lc.columns else [])
            st.caption(f"{len(lc):,} ads changed price. Dates are scrape dates: 'changed on' is the first "
                       "scrape that saw the new price. Days at new price = until the ad was last seen.")
            st.dataframe(lc[show], hide_index=True, width="stretch", height=420,
                         column_config={
                             "car": "Car", "seller_name": "Dealer",
                             "old_price": st.column_config.NumberColumn(f"Old price ({CUR})", format="%d"),
                             "since": st.column_config.DateColumn("Since", format="DD MMM YYYY"),
                             "changed_on": st.column_config.DateColumn("Changed on", format="DD MMM YYYY"),
                             "days_at_old_price": st.column_config.NumberColumn("Days at old price"),
                             "new_price": st.column_config.NumberColumn(f"New price ({CUR})", format="%d"),
                             "days_at_new_price": st.column_config.NumberColumn("Days at new price"),
                             "change_pct": st.column_config.NumberColumn("Change", format="%+.1f%%"),
                             "changes": st.column_config.NumberColumn("Changes", help="How many times this ad changed price"),
                             "url": st.column_config.LinkColumn("Ad", display_text="open")})


tabs = st.tabs(list(DATASETS))
for tab, ds in zip(tabs, DATASETS.values()):
    with tab:
        render(ds)
