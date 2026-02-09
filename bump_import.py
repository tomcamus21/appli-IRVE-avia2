import requests
import sqlite3
from datetime import datetime
import time

# === CONFIGURATION ===
APPLICATION_ID = "2727ac86-4f15-4994-91a5-172e3006ee7b"
APPLICATION_KEY = "6ad38fdc-c5dc-4f7e-af06-3e44b0c59ea5"
BASE_URL = "https://api.bump-charge.com"
DB_PATH = "bump_data.db"

# === AUTHENTIFICATION AVEC RETRY ===
def get_jwt_token():
    url = f"{BASE_URL}/api/applications/{APPLICATION_ID}/authenticate"
    
    # Essayer plusieurs formats de payload
    payloads_to_try = [
        {"applicationKeyId": APPLICATION_KEY},  # Format avec objet
        APPLICATION_KEY,  # Format string directe
    ]
    
    for i, payload in enumerate(payloads_to_try, 1):
        try:
            print(f"  → Tentative {i} d'authentification...")
            response = requests.post(
                url, 
                json=payload,
                headers={"Content-Type": "application/json"},
                timeout=30
            )
            response.raise_for_status()
            token = response.text.strip('"')
            print(f"  ✅ Authentification réussie !")
            return token
        except requests.exceptions.RequestException as e:
            print(f"  ❌ Échec tentative {i}: {e}")
            if i < len(payloads_to_try):
                print(f"  ⏳ Nouvelle tentative dans 2 secondes...")
                time.sleep(2)
            else:
                print(f"\n❌ Toutes les tentatives d'authentification ont échoué")
                raise

# === INITIALISATION DE LA BASE DE DONNÉES ===
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS stations (
            id TEXT PRIMARY KEY,
            name TEXT,
            address TEXT,
            city TEXT
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS chargers (
            id TEXT PRIMARY KEY,
            name TEXT,
            status TEXT,
            location_id TEXT,
            model_id TEXT,
            model_name TEXT,
            vendor_id TEXT,
            vendor_name TEXT
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS charging_sessions (
            id TEXT PRIMARY KEY,
            charger_id TEXT,
            start_time TEXT,
            end_time TEXT,
            energy_kwh REAL,
            total_cost REAL
        );
    """)
    conn.commit()
    conn.close()

# === RÉCUPÉRATION DES STATIONS AVIA VOLT DEPUIS LES SESSIONS ===
def fetch_avia_location_ids(token):
    """Récupère les IDs des stations AVIA VOLT depuis toutes les sessions de recharge"""
    headers = {"Authorization": f"Bearer {token}"}
    location_ids = set()
    location_names = {}
    page = 0
    
    print("  → Parcours de toutes les pages pour identifier les stations AVIA...")
    
    while True:
        url = f"{BASE_URL}/api/v2/charging-sessions?page={page}"
        try:
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            
            sessions = data.get("results", [])
            last_page = data.get("lastPage", 0)
            
            for session in sessions:
                evse = session.get("evse", {})
                location_id = evse.get("locationId")
                location_name = evse.get("locationName", "")
                
                if location_id and "AVIA" in location_name.upper():
                    location_ids.add(location_id)
                    location_names[location_id] = location_name
            
            if page >= last_page:
                break
            
            page += 1
            
        except Exception as e:
            print(f"  ⚠️  Erreur page {page}: {e}")
            break
    
    return list(location_ids), location_names

# === RÉCUPÉRATION D'UNE LOCATION PAR ID ===
def fetch_location_by_id(token, location_id):
    url = f"{BASE_URL}/api/locations/{location_id}"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return response.json()

# === RÉCUPÉRATION DES CHARGERS PAR LOCATION ===
def fetch_chargers_by_location(token, location_id):
    url = f"{BASE_URL}/api/locations/{location_id}/chargers"
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    return response.json()

# === RÉCUPÉRATION DES SESSIONS DE RECHARGE ===
def fetch_charging_sessions(token, avia_charger_ids):
    """Récupère TOUTES les sessions de recharge avec pagination"""
    headers = {"Authorization": f"Bearer {token}"}
    avia_sessions = []
    page = 0
    
    while True:
        url = f"{BASE_URL}/api/v2/charging-sessions?page={page}"
        try:
            response = requests.get(url, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            
            all_sessions = data.get("results", [])
            last_page = data.get("lastPage", 0)
            
            # Filtrer uniquement les sessions des chargers AVIA Volt
            for session in all_sessions:
                authorization = session.get("authorization", {})
                charger = authorization.get("charger") or {}
                charger_id = charger.get("id")
                if charger_id in avia_charger_ids:
                    avia_sessions.append(session)
            
            print(f"  → Page {page}/{last_page}: {len(avia_sessions)} sessions AVIA Volt trouvées")
            
            if page >= last_page:
                break
            
            page += 1
            time.sleep(0.5)  # Petite pause entre les pages pour ne pas surcharger l'API
            
        except requests.exceptions.RequestException as e:
            print(f"  ⚠️  Erreur page {page}: {e}")
            break
    
    return avia_sessions

# === ENREGISTREMENT DES DONNÉES ===
def store_stations(locations):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    for loc in locations:
        address_obj = loc.get("address", {})
        cursor.execute("""
            INSERT OR REPLACE INTO stations (id, name, address, city)
            VALUES (?, ?, ?, ?);
        """, (
            str(loc.get("id")),
            loc.get("name"),
            address_obj.get("address"),
            address_obj.get("city")
        ))
    conn.commit()
    conn.close()

def store_chargers(chargers, location_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    charger_ids = []
    for charger in chargers:
        model = charger.get("model") or {}
        vendor = model.get("vendor") or {}
        charger_id = charger.get("id")
        charger_ids.append(charger_id)
        
        cursor.execute("""
            INSERT OR REPLACE INTO chargers (
                id, name, status, location_id, model_id, model_name, vendor_id, vendor_name
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            charger_id,
            charger.get("name"),
            charger.get("status"),
            location_id,
            model.get("id"),
            model.get("name"),
            vendor.get("id"),
            vendor.get("name")
        ))
    conn.commit()
    conn.close()
    return charger_ids

def store_charging_sessions(sessions):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    sessions_saved = 0
    sessions_skipped = 0
    total_cost_sum = 0
    
    for session in sessions:
        # Extraire l'énergie consommée depuis l'objet consumption
        consumption = session.get("consumption", {})
        energy_wh = consumption.get("consumedEnergy", 0)
        energy_kwh = energy_wh / 1000 if energy_wh else 0
        
        # Extraire le coût total depuis consumption.price.amount (CA TTC)
        price_obj = consumption.get("price", {})
        total_cost = price_obj.get("amount", 0) if price_obj else 0
        
        # Ignorer les sessions avec 0 kWh
        if energy_kwh == 0:
            sessions_skipped += 1
            continue
        
        # Extraire le charger_id depuis authorization.charger.id
        authorization = session.get("authorization", {})
        charger = authorization.get("charger") or {}
        charger_id = charger.get("id")
        
        cursor.execute("""
            INSERT OR REPLACE INTO charging_sessions (
                id, charger_id, start_time, end_time, energy_kwh, total_cost
            ) VALUES (?, ?, ?, ?, ?, ?);
        """, (
            session.get("id"),
            charger_id,
            session.get("startDate"),
            session.get("endDate"),
            energy_kwh,
            total_cost
        ))
        sessions_saved += 1
        total_cost_sum += total_cost
        
    conn.commit()
    conn.close()
    
    print(f"  → {sessions_saved} sessions enregistrées ({sessions_skipped} sessions avec 0 kWh ignorées)")
    print(f"  💰 CA total: {total_cost_sum:.2f}€")
    return sessions_saved

# === SCRIPT PRINCIPAL ===
def main():
    print("Initialisation de la base de données...")
    init_db()
    
    print("Authentification...")
    try:
        token = get_jwt_token()
    except Exception as e:
        print(f"\n❌ Impossible de s'authentifier: {e}")
        print("\n💡 Vérifications à faire:")
        print("   1. Vérifie que les credentials sont corrects")
        print("   2. Vérifie ta connexion internet")
        print("   3. Vérifie que l'API Bump est accessible")
        print("   4. Essaie à nouveau dans quelques minutes")
        return
    
    print("Identification des stations AVIA Volt depuis les sessions...")
    avia_location_ids, location_names = fetch_avia_location_ids(token)
    print(f"  → {len(avia_location_ids)} stations AVIA Volt identifiées:")
    for loc_id in avia_location_ids:
        print(f"     • {location_names[loc_id]}")
    
    if len(avia_location_ids) == 0:
        print("⚠️  Aucune station AVIA Volt trouvée dans les sessions.")
        return
    
    print("\nRécupération des détails des stations AVIA Volt...")
    locations = []
    for loc_id in avia_location_ids:
        location = fetch_location_by_id(token, loc_id)
        locations.append(location)
    
    print("Enregistrement des stations AVIA Volt...")
    store_stations(locations)
    
    print("\nRécupération des chargers pour les stations AVIA Volt...")
    all_charger_ids = []
    for loc_id in avia_location_ids:
        location_name = location_names[loc_id]
        chargers = fetch_chargers_by_location(token, loc_id)
        print(f"  → {location_name}: {len(chargers)} chargers")
        charger_ids = store_chargers(chargers, loc_id)
        all_charger_ids.extend(charger_ids)
    
    print(f"\nTotal de chargers AVIA Volt: {len(all_charger_ids)}")
    
    print("\nRécupération des sessions de recharge AVIA Volt...")
    sessions = fetch_charging_sessions(token, set(all_charger_ids))
    print(f"  → {len(sessions)} sessions AVIA Volt trouvées")
    
    print("\nEnregistrement des sessions...")
    sessions_saved = store_charging_sessions(sessions)
    
    print("\n✅ Import AVIA Volt terminé avec succès !")
    print(f"   • {len(locations)} stations")
    print(f"   • {len(all_charger_ids)} chargers")
    print(f"   • {sessions_saved} sessions de recharge (avec énergie > 0 kWh)")

if __name__ == "__main__":
    main()