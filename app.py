"""
Chinese-spec car market (DubiCars) - one-page executive dashboard (Streamlit)

Runs on Streamlit Community Cloud. Reads dubicars_chinese.xlsx from Google
Drive - the daily 07:00 run on the VPS replaces that same Drive file, so the
dashboard always shows the latest scrape (re-checked every hour).

Secrets (Streamlit Cloud -> app -> Settings -> Secrets):
    drive_link = "https://drive.google.com/file/d/<FILE_ID>/view?usp=sharing"
    password   = "..."          # optional
Without drive_link it looks for dubicars_chinese.xlsx next to this file.
"""
import hmac
import io
import os
import re

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Chinese-spec Car Market", page_icon="🚗", layout="wide",
                   initial_sidebar_state="collapsed")

HERE = os.path.dirname(os.path.abspath(__file__))
FILE_NAME = "dubicars_chinese.xlsx"
CANDIDATES = [os.environ.get("DUBICARS_FILE", ""),
              os.path.join(HERE, "..", "output", FILE_NAME),
              os.path.join(HERE, "output", FILE_NAME),
              os.path.join(HERE, FILE_NAME)]
TOP_N = 10
UP, DOWN, FLAT = "#d1495b", "#2a9d8f", "#8d99ae"   # price up = red, down = green

st.markdown("""<style>
.block-container {padding-top: 1.6rem; padding-bottom: 1rem;}
[data-testid="stMetricValue"] {font-size: 2rem;}
</style>""", unsafe_allow_html=True)


# ------------------------------------------------------------------ password
def password_ok():
    try:
        pw = st.secrets.get("password", "")
    except Exception:
        pw = ""
    if not pw or st.session_state.get("auth"):
        return True
    with st.form("login"):
        st.subheader("🔒 Chinese-spec car market")
        typed = st.text_input("Password", type="password")
        if st.form_submit_button("Open"):
            if hmac.compare_digest(typed, pw):
                st.session_state["auth"] = True
                st.rerun()
            st.error("Wrong password")
    return False


if not password_ok():
    st.stop()


# ------------------------------------------------------------------ data
def secret(name):
    try:
        return st.secrets.get(name, "")
    except Exception:
        return ""


@st.cache_data(show_spinner="Getting the latest data...", ttl=3600)
def fetch_drive(link):
    """download the Drive file (shared as 'Anyone with the link')"""
    import requests
    m = re.search(r"/d/([\w-]{20,})", link) or re.search(r"id=([\w-]{20,})", link)
    fid = m.group(1) if m else link.strip()
    s = requests.Session()
    r = s.get("https://drive.google.com/uc", params={"export": "download", "id": fid},
              timeout=60)
    if r.content[:2] != b"PK":                         # Drive's "can't scan" page
        tok = re.search(r'name="confirm" value="([^"]+)"', r.text)
        uuid = re.search(r'name="uuid" value="([^"]+)"', r.text)
        r = s.get("https://drive.usercontent.google.com/download",
                  params={"id": fid, "export": "download", "confirm": tok.group(1) if tok else "t",
                          **({"uuid": uuid.group(1)} if uuid else {})}, timeout=60)
    if r.status_code != 200 or r.content[:2] != b"PK":
        raise RuntimeError("Drive did not return the Excel file - is it shared as "
                           "'Anyone with the link'?")
    when = r.headers.get("Last-Modified") or r.headers.get("Date")
    return r.content, when


@st.cache_data(show_spinner="Loading...")
def load(src, stamp):                       # stamp in the key -> reloads when the file changes
    xl = pd.ExcelFile(io.BytesIO(src) if isinstance(src, bytes) else src)
    df = pd.read_excel(xl, "listings")
    hist = pd.read_excel(xl, "price_history") if "price_history" in xl.sheet_names \
        else pd.DataFrame(columns=["car_id", "checked_at", "price_aed"])
    key = "listing_id" if "listing_id" in df.columns and "listing_id" in hist.columns \
        else "car_id"
    df["_key"] = df[key].astype(str)
    hist["_key"] = hist[key].astype(str)
    def col(c, default=None):                          # works with old files too
        return df[c] if c in df.columns else pd.Series(default, index=df.index)
    for c in ["first_seen_at", "last_seen_at"]:
        df[c] = pd.to_datetime(col(c), errors="coerce")
    for c in ["price_aed", "year"]:
        df[c] = pd.to_numeric(col(c), errors="coerce")
    for c in ["make", "model", "seller_name", "ad_status"]:
        df[c] = col(c, "").fillna("").astype(str).str.strip()
    df["ad_status"] = df["ad_status"].replace("", "Active")
    df["car"] = (df["make"].str.title() + " " + df["model"].str.title() + " "
                 + df["year"].astype("Int64").astype(str)).str.strip()
    hist["checked_at"] = pd.to_datetime(hist["checked_at"], errors="coerce")
    hist["price_aed"] = pd.to_numeric(hist["price_aed"], errors="coerce")
    hist = hist.dropna(subset=["checked_at", "price_aed"])
    hist = hist[hist["price_aed"] > 0]

    # one price per ad per day, carried forward until the ad was last seen
    # (price_history only keeps a row when the price moves)
    h = hist.assign(day=hist["checked_at"].dt.normalize()).sort_values("checked_at")
    h = h.groupby(["_key", "day"])["price_aed"].last().reset_index()
    end = df.set_index("_key")["last_seen_at"].dt.normalize()
    span = h.groupby("_key")["day"].agg(["min", "max"])
    span["end"] = pd.concat([span["max"], end.reindex(span.index)], axis=1).max(axis=1)
    rows = [(k, d) for k, r in span.iterrows() for d in pd.date_range(r["min"], r["end"], freq="D")]
    daily = pd.DataFrame(rows, columns=["_key", "day"]).merge(h, how="left")
    daily["price_aed"] = daily.groupby("_key")["price_aed"].ffill()
    daily = daily.merge(df[["_key", "car", "make", "seller_name"]], on="_key")
    return df, daily


link = secret("drive_link")
try:
    if link:
        content, when = fetch_drive(link)
        df, daily = load(content, len(content))
        updated = pd.to_datetime(when, errors="coerce", utc=True)
    else:
        path = next((os.path.abspath(p) for p in CANDIDATES if p and os.path.exists(p)), None)
        if not path:
            st.error(f"No data: set drive_link in the app secrets, or put {FILE_NAME} "
                     "next to app.py.")
            st.stop()
        df, daily = load(path, os.path.getmtime(path))
        updated = pd.Timestamp(os.path.getmtime(path), unit="s", tz="UTC")
except Exception as e:
    st.error(f"Couldn't load the data: {e}")
    st.stop()
updated = updated.tz_convert("Asia/Amman") if pd.notna(updated) else None
last_scrape = df["last_seen_at"].max()

# ------------------------------------------------------------------ header + filter
h1, h2 = st.columns([3, 2])
h1.title("🚗 Chinese-spec cars — UAE market")
h1.caption(f"DubiCars · last scrape {last_scrape:%d %b %Y, %H:%M} · data since "
           f"{df['first_seen_at'].min():%d %b %Y}")
makes = sorted(m for m in df["make"].unique() if m)
pick = h2.multiselect("Brand", makes, placeholder="All brands")
if pick:
    df = df[df["make"].isin(pick)]
    daily = daily[daily["make"].isin(pick)]

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
    t = two.pivot_table(index=by, columns="period", values="price_aed", aggfunc="median")
    t.columns = ["before", "now"]
    t["change_pct"] = (t["now"] / t["before"] - 1) * 100
    return t.reset_index()


ads_cmp = compare("_key")
cut = int((ads_cmp["change_pct"] < -0.5).sum()) if len(ads_cmp) else 0
rise = int((ads_cmp["change_pct"] > 0.5).sum()) if len(ads_cmp) else 0
avg_chg = ads_cmp["change_pct"].mean() if len(ads_cmp) else None

# ------------------------------------------------------------------ KPIs
k = st.columns(4)
k[0].metric("Cars for sale now", f"{len(live):,}")
med = live["price_aed"].median()
k[1].metric("Typical price", "—" if pd.isna(med) else f"{med / 1000:,.0f}K AED")
if avg_chg is None:
    k[2].metric(f"Prices vs last {grain_txt}", "—")
else:
    word = "up" if avg_chg > 0.05 else "down" if avg_chg < -0.05 else "flat"
    k[2].metric(f"Prices vs last {grain_txt}", f"{avg_chg:+.1f}%", word,
                delta_color="inverse" if word != "flat" else "off")
k[3].metric("Cars that got cheaper", f"{cut:,}", f"{rise:,} got more expensive",
            delta_color="off")

st.divider()

# ------------------------------------------------------------------ price over time
top_cars = live["car"].value_counts().head(TOP_N).index.tolist()
c1, c2 = st.columns([3, 2], gap="large")
with c1:
    st.subheader(f"Price per {grain_txt} — most listed cars")
    chosen = st.multiselect("Cars", sorted(df["car"].unique()), default=top_cars[:5],
                            label_visibility="collapsed")
    line = per[per["car"].isin(chosen)].groupby(["period", "car"])["price_aed"]\
        .median().reset_index()
    if line.empty:
        st.info("Pick a car above.")
    else:
        fig = px.line(line, x="period", y="price_aed", color="car", markers=True)
        fig.update_layout(height=420, margin=dict(l=0, r=0, t=10, b=0), legend_title=None,
                          legend=dict(orientation="h", y=-0.15), hovermode="x unified")
        fig.update_xaxes(title=None, tickformat="%d %b" if grain == "W" else "%b %Y",
                         dtick=None if grain == "W" else "M1")
        fig.update_yaxes(title="AED", tickformat=",.0f")
        st.plotly_chart(fig, width="stretch")
    st.caption("Each point = the car's median price that month (each ad's last price in "
               "the month). The current month is month-to-date.")

with c2:
    st.subheader(f"Top {TOP_N} cars: before vs now")
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

with st.expander("Biggest price changes (single ads)"):
    if len(ads_cmp):
        big = ads_cmp.merge(df[["_key", "car", "seller_name", "url"]] if "url" in df
                            else df[["_key", "car", "seller_name"]], on="_key")
        big = big[big["change_pct"].abs() > 0.5].reindex(
            big["change_pct"].abs().sort_values(ascending=False).index).head(20)
        st.dataframe(big.drop(columns="_key"), hide_index=True, width="stretch",
                     column_config={"before": st.column_config.NumberColumn(format="%d"),
                                    "now": st.column_config.NumberColumn(format="%d"),
                                    "change_pct": st.column_config.NumberColumn(
                                        "change", format="%+.1f%%"),
                                    "url": st.column_config.LinkColumn(display_text="open")})
