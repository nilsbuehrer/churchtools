import streamlit as st
import requests
import pandas as pd
from datetime import datetime, date
from io import StringIO

# ---------------------------------------
# UI
# ---------------------------------------
st.title("Anwesenheiten Export")

token = "Login " + st.text_input("AUTH_TOKEN", type="password")
start = st.date_input("Von", date(2025, 11, 2))
end = st.date_input("Bis", date(2026, 3, 1))

# Gruppen & Tags
GROUPS = [
    {"id": 69, "name": "Programm", "tags": ["Gebet", "Begrüssung", "Deko", "Abendmahl vorbereiten", "Abendmahl abwaschen"]},
    {"id": 7,  "name": "Technik",  "tags": ["Video", "Audio", "Beamer", "Licht"]}
]

# Services, die ins CSV sollen
SERVICE_MAP = {
    1: "Predigt",
    2: "Lobpreisleitung",
    3: "Moderation"
}

# ---------------------------------------
# API-Helper (mit Cache)
# ---------------------------------------
@st.cache_data(ttl=3600)
def get_events(from_date, to_date, headers):
    url = "https://feg-thayngen.church.tools/api/events"
    params = {"from": from_date, "to": to_date, "canceled": "false", "include": "eventServices"}
    r = requests.get(url, params=params, headers=headers)
    r.raise_for_status()
    return [e for e in r.json()["data"] if e["calendar"]["domainIdentifier"] == "2"]

@st.cache_data(ttl=3600)
def get_members(group_id, headers):
    url = "https://feg-thayngen.church.tools/api/groups/members"
    params = {"ids[]": group_id, "with_deleted": "false"}
    r = requests.get(url, params=params, headers=headers)
    r.raise_for_status()
    return [m["personId"] for m in r.json()["data"]]

@st.cache_data(ttl=3600)
def get_tags(person_id, headers):
    url = f"https://feg-thayngen.church.tools/api/tags/person/{person_id}"
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        return [entry["name"] for entry in r.json()["data"]]
    return []

@st.cache_data(ttl=3600)
def get_group_absences(group_id, from_date, to_date, headers):
    url = f"https://feg-thayngen.church.tools/api/groups/{group_id}/absences"
    params = {"from_date": from_date, "to_date": to_date}
    r = requests.get(url, params=params, headers=headers)
    r.raise_for_status()
    return r.json()["data"]

@st.cache_data(ttl=3600)
def get_full_name(person_id, headers):
    url = f"https://feg-thayngen.church.tools/api/persons/{person_id}"
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        d = r.json()["data"]
        return f"{d['firstName']} {d['lastName']}"
    return f"Person {person_id}"

# ---------------------------------------
# Button
# ---------------------------------------
if st.button("CSV generieren"):

    if not token:
        st.error("Bitte AUTH_TOKEN eingeben")
        st.stop()

    headers = {"Content-Type": "application/json", "Authorization": token}
    progress = st.progress(0, text="Starte Export...")

    # Schritt 1: Events
    events = get_events(str(start), str(end), headers)
    progress.progress(20, text="Events geladen ✅")

    # Schritt 2: Gruppen & Daten vorbereiten
    absence_cache = {}
    tags_for_persons = {}

    for idx, group in enumerate(GROUPS, start=1):
        members = get_members(group["id"], headers)
        absences = get_group_absences(group["id"], str(start), str(end), headers)

        # Absenzen sammeln
        for absence in absences:
            pid = int(absence["person"]["domainIdentifier"])
            s_date = datetime.strptime(absence["startDate"], "%Y-%m-%d").date()
            e_date = datetime.strptime(absence["endDate"], "%Y-%m-%d").date()
            absence_cache.setdefault(pid, []).append((s_date, e_date))

        # Tags vorbereiten
        for tag in group["tags"]:
            col_name = f"{group['name']}: {tag}"
            tags_for_persons[col_name] = []
            for pid in members:
                if tag in get_tags(pid, headers):
                    tags_for_persons[col_name].append(pid)

        progress.progress(20 + int(60 * idx / len(GROUPS)), text=f"Gruppe {group['name']} geladen ✅")

    # Schritt 3: DataFrame bauen
    rows = []
    for i, event in enumerate(events, start=1):
        event_date = date.fromisoformat(event["startDate"][:10])
        event_title = event["name"]
        event_note = event.get("note") or ""
        event_info = f"{event_title} ({event_note})" if event_note else event_title

        row = {
            "Datum": event_date.strftime("%a, %d.%m.%Y"),
            "Event": event_info
        }

        # --- Services einfügen ---
        assigned_pids = set()
        for svc in event.get("eventServices", []):
            sid = svc.get("serviceId")
            person_obj = svc.get("person")
            pid = None
            person = {}

            if person_obj:
                pid = person_obj.get("domainIdentifier")
                person = person_obj.get("domainAttributes", {})

            if sid in SERVICE_MAP:
                col = SERVICE_MAP[sid]
                if person:
                    full_name = f"{person.get('firstName', '')} {person.get('lastName', '')}".strip()
                    row[col] = full_name
                else:
                    row[col] = ""

            if pid:
                assigned_pids.add(str(pid))

        # --- Tags/Verfügbarkeiten einfügen ---
        for col, persons in tags_for_persons.items():
            available_normal = []
            available_assigned = []
            for pid in persons:
                absent = any(s <= event_date <= e for s, e in absence_cache.get(pid, []))
                if not absent:
                    name = get_full_name(pid, headers)
                    if str(pid) in assigned_pids:
                        available_assigned.append(f"({name})")
                    else:
                        available_normal.append(name)
            # zuerst normale, darunter eingeteilte in Klammern
            row[col] = "\n".join(available_normal + available_assigned)


        rows.append(row)
        progress.progress(80 + int(20 * i / len(events)), text="Events verarbeitet...")

    df = pd.DataFrame(rows)

    # Schritt 4: Export
    csv_data = df.to_csv(index=False)
    progress.progress(100, text="Fertig ✅")
    st.success("CSV erfolgreich erstellt ✅")

    st.download_button(
        "Download CSV",
        csv_data,
        file_name="Anwesenheiten_Gesamt.csv",
        mime="text/csv"
    )

    st.dataframe(df, use_container_width=True)
