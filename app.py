import streamlit as st
import requests
import pandas as pd
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from io import BytesIO

# ---------------------------------------
# UI
# ---------------------------------------
st.title("Anwesenheiten Export")

base_url = st.text_input("ChurchTools URL", "https://feg-thayngen.church.tools/")
raw_token = st.text_input("AUTH_TOKEN", type="password")

# Standardwerte: heute bis +4 Monate
today = date.today()
default_end = today + relativedelta(months=4)
start = st.date_input("Von", today)
end = st.date_input("Bis", default_end)

# Gruppen & Tags
GROUPS = [
    {"id": 69, "name": "Programm", "tags": ["Gebet", "Begrüssung", "Deko", "Abendmahl vorbereiten", "Abendmahl abwaschen"]},
    {"id": 7,  "name": "Technik",  "tags": ["Video", "Audio", "Beamer", "Licht"]}
]

# Dienste (vollständig parsen)
SERVICE_MAP_FULL = {
    1:  "Predigt",
    2:  "Lobpreisleitung",
    3:  "Moderation",
    6:  "Audio",
    7:  "Licht",
    8:  "Video",
    9:  "Musik",
    12: "Beamer",
    15: "Gebet",
    18: "Begrüssung",
    30: "Deko",
    82: "Abendmahl vorbereiten",
    36: "Abendmahl abwaschen",
}

# Dienste, die eigene Spalten bekommen
SERVICE_MAP_EXPORT = {
    1: "Predigt",
    2: "Lobpreisleitung",
    3: "Moderation"
}

# Mapping: Tag -> Dienst
TAG_TO_SERVICE = {
    "Gebet": 15,
    "Begrüssung": 18,
    "Deko": 30,
    "Abendmahl abwaschen": 36,
    "Abendmahl vorbereiten": 82,
    "Video": 8,
    "Audio": 6,
    "Beamer": 12,
    "Licht": 7,
}

# ---------------------------------------
# API-Helper (mit Cache)
# ---------------------------------------
@st.cache_data(ttl=3600)
def get_events(from_date, to_date, headers, base_url):
    url = f"{base_url.rstrip('/')}/api/events"
    params = {"from": from_date, "to": to_date, "canceled": "false", "include": "eventServices"}
    r = requests.get(url, params=params, headers=headers)
    r.raise_for_status()
    return [e for e in r.json()["data"] if e["calendar"]["domainIdentifier"] == "2"]

@st.cache_data(ttl=3600)
def get_members(group_id, headers, base_url):
    url = f"{base_url.rstrip('/')}/api/groups/members"
    params = {"ids[]": group_id, "with_deleted": "false"}
    r = requests.get(url, params=params, headers=headers)
    r.raise_for_status()
    return [m["personId"] for m in r.json()["data"]]

@st.cache_data(ttl=3600)
def get_tags(person_id, headers, base_url):
    url = f"{base_url.rstrip('/')}/api/tags/person/{person_id}"
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        return [entry["name"] for entry in r.json()["data"]]
    return []

@st.cache_data(ttl=3600)
def get_group_absences(group_id, from_date, to_date, headers, base_url):
    url = f"{base_url.rstrip('/')}/api/groups/{group_id}/absences"
    params = {"from_date": from_date, "to_date": to_date}
    r = requests.get(url, params=params, headers=headers)
    r.raise_for_status()
    return r.json()["data"]

@st.cache_data(ttl=3600)
def get_full_name(person_id, headers, base_url):
    url = f"{base_url.rstrip('/')}/api/persons/{person_id}"
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        d = r.json()["data"]
        return f"{d['firstName']} {d['lastName']}"
    return f"Person {person_id}"

# ---------------------------------------
# Button
# ---------------------------------------
if st.button("CSV/Excel generieren"):

    if not raw_token:
        st.error("Bitte AUTH_TOKEN eingeben")
        st.stop()
    headers = {"Content-Type": "application/json", "Authorization": "Login " + raw_token}

    progress = st.progress(0, text="Starte Export...")

    # Schritt 1: Events
    events = get_events(str(start), str(end), headers, base_url)
    progress.progress(15, text="Events geladen ✅")

    # Schritt 2: Gruppen & Daten vorbereiten
    absence_cache = {}
    tags_for_persons = {}
    all_person_ids = set()

    for idx, group in enumerate(GROUPS, start=1):
        members = get_members(group["id"], headers, base_url)
        absences = get_group_absences(group["id"], str(start), str(end), headers, base_url)

        # Absenzen speichern
        for absence in absences:
            pid = int(absence["person"]["domainIdentifier"])
            s_date = datetime.strptime(absence["startDate"], "%Y-%m-%d").date()
            e_date = datetime.strptime(absence["endDate"], "%Y-%m-%d").date()
            cur = s_date
            store = absence_cache.setdefault(pid, set())
            while cur <= e_date:
                store.add(cur)
                cur = cur + timedelta(days=1)

        # Tags vorbereiten
        for tag in group["tags"]:
            col_name = f"{group['name']}: {tag}"
            pids_for_tag = []
            for pid in members:
                if tag in get_tags(pid, headers, base_url):
                    pids_for_tag.append(pid)
                    all_person_ids.add(pid)
            tags_for_persons[col_name] = pids_for_tag

        progress.progress(15 + int(45 * idx / len(GROUPS)), text=f"Gruppe {group['name']} geladen ✅")

    # Namen vorab cachen
    names_cache = {pid: get_full_name(pid, headers, base_url) for pid in all_person_ids}

    # Schritt 3: DataFrame bauen
    rows = []
    for i, event in enumerate(events, start=1):
        event_date = date.fromisoformat(event["startDate"][:10])
        event_title = event["name"]
        event_note = event.get("note") or ""
        event_info = f"{event_title} ({event_note})" if event_note else event_title

        row = {"Datum": event_date.strftime("%a, %d.%m.%Y"), "Event": event_info}

        # --- Services parsen (alle IDs) ---
        service_names = {sid: [] for sid in SERVICE_MAP_FULL.keys()}
        assigned_pids = set()

        for svc in event.get("eventServices", []):
            sid = svc.get("serviceId")
            person_obj = svc.get("person")
            if sid in SERVICE_MAP_FULL and person_obj:
                pid = int(person_obj.get("domainIdentifier"))
                person = person_obj.get("domainAttributes", {}) or {}
                full_name = f"{person.get('firstName','')} {person.get('lastName','')}".strip() \
                            or get_full_name(pid, headers, base_url)
                service_names[sid].append(f"*{full_name}")
                assigned_pids.add(pid)

        # Nur die Export-Dienste als eigene Spalten
        for sid, col_name in SERVICE_MAP_EXPORT.items():
            row[col_name] = ", ".join(service_names[sid]) if service_names[sid] else ""

        # --- Tags/Verfügbarkeiten einfügen ---
        for col, persons in tags_for_persons.items():
            tag = col.split(": ", 1)[1] if ": " in col else col
            sid_for_tag = TAG_TO_SERVICE.get(tag)

            if sid_for_tag and service_names.get(sid_for_tag):
                # Dienst besetzt -> nur eingeteilte Person(en) mit *
                row[col] = ", ".join(service_names[sid_for_tag])
            else:
                # sonst alle verfügbaren Personen
                available_normal = []
                available_other = []
                for pid in persons:
                    if event_date not in absence_cache.get(pid, set()):
                        base_name = names_cache.get(pid, f"Person {pid}")
                        if pid in assigned_pids:
                            available_other.append(f"({base_name})")
                        else:
                            available_normal.append(base_name)
                row[col] = "\n".join(available_normal + available_other)

        rows.append(row)
        progress.progress(60 + int(35 * i / max(len(events), 1)), text="Events verarbeitet...")

    df = pd.DataFrame(rows)

    # Schritt 4a: CSV Export
    csv_data = df.to_csv(index=False, encoding="utf-8-sig")

    # Schritt 4b: Excel Export mit schönerer Formatierung
    excel_buffer = BytesIO()
    excel_data = None

    try:
        with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Anwesenheiten")
            workbook  = writer.book
            worksheet = writer.sheets["Anwesenheiten"]

            # Format: Umbruch + vertikal oben
            wrap_format = workbook.add_format({"text_wrap": True, "valign": "top"})
            header_format = workbook.add_format({"bold": True, "bg_color": "#D9E1F2"})

            # Header-Format anwenden
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Spaltenbreite anpassen
            for i, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(i, i, min(max_len, 50), wrap_format)

        excel_data = excel_buffer.getvalue()

    except ModuleNotFoundError:
        st.error("Bitte 'xlsxwriter' installieren für schön formatierten Excel-Export.")


    progress.progress(100, text="Fertig ✅")
    st.success("Dateien erfolgreich erstellt ✅")

    # Download Buttons
    st.download_button(
        "Download CSV",
        csv_data,
        file_name="Anwesenheiten_Gesamt.csv",
        mime="text/csv"
    )

    if excel_data:
        st.download_button(
            "Download Excel",
            excel_data,
            file_name="Anwesenheiten_Gesamt.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Kein Excel-Export möglich – bitte 'openpyxl' oder 'xlsxwriter' installieren.")

    st.dataframe(df, use_container_width=True)
