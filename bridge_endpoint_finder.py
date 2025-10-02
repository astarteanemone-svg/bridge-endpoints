import streamlit as st
import pandas as pd
import requests
from io import BytesIO

OVERPASS = "https://lz4.overpass-api.de/api/interpreter"

def decimal_to_dms(lat, lon):
    def conv(v, lat=True):
        d = 'N' if lat and v >= 0 else 'S' if lat else ('E' if v >= 0 else 'W')
        v = abs(v)
        deg = int(v)
        m = int((v - deg) * 60)
        s = (v - deg - m/60) * 3600
        return f"{deg}°{m}'{s:.1f}\"{d}"
    return conv(lat, True), conv(lon, False)

def get_way_and_endpoints(name, area_id):
    query = f"""
    [out:json][timeout:60];
    area({area_id})->.a;
    way["bridge"="yes"]["name"~"{name}"](area.a);
    out body;
    >;
    out skel qt;
    """
    r = requests.get(OVERPASS, params={"data": query})
    if r.status_code != 200:
        return None, None, None
    js = r.json()
    nodes = {e["id"]: (e.get("lat"), e.get("lon")) for e in js.get("elements", []) if e["type"] == "node"}
    ways = [e for e in js.get("elements", []) if e["type"] == "way"]
    if not ways:
        return None, None, None
    w = ways[0]
    nids = w["nodes"]
    return w["id"], nodes.get(nids[0]), nodes.get(nids[-1])

st.title("橋リスト → 緯度経度変換ツール")

uploaded = st.file_uploader("橋リスト.xlsx をアップロードしてください", type=["xlsx"])

if uploaded:
    df = pd.read_excel(uploaded, sheet_name="橋リスト")
    st.write("読み込んだ橋リスト", df.head())

    results = []
    for _, row in df.iterrows():
        name = str(row["橋名"]).strip()
        area_id = row["AreaID"]
        if pd.isna(name) or pd.isna(area_id):
            continue
        way_id, s, e = get_way_and_endpoints(name, int(area_id))
        if way_id and s and e:
            slat, slon = s
            elat, elon = e
            sdms_lat, sdms_lon = decimal_to_dms(slat, slon)
            edms_lat, edms_lon = decimal_to_dms(elat, elon)
        else:
            way_id = slat = slon = elat = elon = sdms_lat = sdms_lon = edms_lat = edms_lon = None

        results.append({
            "橋名": name,
            "県名": row["県名"],
            "市町村": row["市町村"],
            "AreaID": area_id,
            "way_id": way_id,
            "起点_緯度(十進)": slat,
            "起点_経度(十進)": slon,
            "終点_緯度(十進)": elat,
            "終点_経度(十進)": elon,
            "起点_緯度(度分秒)": sdms_lat,
            "起点_経度(度分秒)": sdms_lon,
            "終点_緯度(度分秒)": edms_lat,
            "終点_経度(度分秒)": edms_lon
        })

    out_df = pd.DataFrame(results)
    st.write("処理結果", out_df.head())

    # Excel出力
    output = BytesIO()
    out_df.to_excel(output, index=False)
    st.download_button(
        label="📥 結果をダウンロード (Excel)",
        data=output.getvalue(),
        file_name="bridge_endpoints.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
