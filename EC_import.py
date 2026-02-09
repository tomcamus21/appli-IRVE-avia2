import sqlite3
from datetime import datetime
import pandas as pd


import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')


# Configuration PostgreSQL ETL
PG_HOST = "etl.ocpp.irve.numeos.tech"
PG_PORT = 52012
PG_DATABASE = "etl_tda"
PG_USER = "tda_thevenin_ducrot_view"
PG_PASSWORD = "Fn8ZLFuwLHn7wcPM8e4mZgjGQmGFhoLDgd2d"

# Configuration SQLite
SQLITE_DB = "bump_data.db"
CONFIG_FILE = "config_stations.xlsx"

def load_station_mapping():
    """Charge le mapping id_borne -> station_name depuis config_stations.xlsx"""
    try:
        df = pd.read_excel(CONFIG_FILE)
        
        # Standardiser les séparateurs en virgules
        df['id_borne'] = df['id_borne'].astype(str).str.replace('.', ',')
        
        # Créer un dictionnaire id_borne -> station_name
        mapping = {}
        for _, row in df.iterrows():
            if pd.notna(row['id_borne']) and row['id_borne'] != 'nan':
                bornes = row['id_borne'].split(',')
                for borne in bornes:
                    borne = borne.strip()
                    if borne:
                        mapping[int(borne)] = row['station_name']
        
        print(f"✓ Mapping chargé: {len(mapping)} bornes -> {len(set(mapping.values()))} stations")
        return mapping
    except FileNotFoundError:
        print(f"⚠️  Fichier {CONFIG_FILE} non trouvé")
        return {}
    except Exception as e:
        print(f"⚠️  Erreur lors du chargement du mapping: {e}")
        return {}

def init_etl_database():
    """Initialise la base de données SQLite pour ETL"""
    conn = sqlite3.connect(SQLITE_DB)
    cursor = conn.cursor()
    
    # Supprimer les anciennes tables ETL
    cursor.execute("DROP TABLE IF EXISTS etl_sessions")
    cursor.execute("DROP TABLE IF EXISTS etl_stations")
    
    # Table pour les stations ETL (avec nom de station)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS etl_stations (
            id_borne INTEGER PRIMARY KEY,
            chargebox_id TEXT,
            nom_borne TEXT,
            libelle TEXT,
            alias_borne TEXT,
            nom_zdc TEXT,
            puissance_borne REAL,
            type_pdc TEXT,
            nombre_pdc INTEGER,
            date_installation TEXT,
            perimetre TEXT,
            station_name TEXT,
            last_updated TEXT
        )
    """)
    
    # Table pour les sessions de recharge ETL
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS etl_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_transaction INTEGER UNIQUE,
            id_borne INTEGER,
            pdc_id INTEGER,
            connector_id INTEGER,
            perimetre TEXT,
            station_name TEXT,
            date_debut TEXT,
            date_fin TEXT,
            duree_sec REAL,
            energie_wh REAL,
            prix_ttc REAL,
            rfid TEXT,
            charge_reussie INTEGER,
            last_updated TEXT
        )
    """)
    
    conn.commit()
    conn.close()
    print("✓ Tables ETL créées")

def import_etl_data():
    """Import des données depuis PostgreSQL ETL"""
    print("\n=== IMPORT DONNÉES ETL ===\n")
    
    pg_conn = None
    sqlite_conn = None
    
    # Vérifier psycopg2
    try:
        import psycopg2
        from psycopg2 import Error
        print("✓ Module psycopg2 disponible")
    except ImportError:
        print("✗ Module psycopg2 non installé")
        print("  Installez-le avec: pip install psycopg2-binary")
        return
    
    try:
        # Charger le mapping bornes -> stations
        station_mapping = load_station_mapping()
        
        if not station_mapping:
            print("⚠️  Aucun mapping disponible, import annulé")
            return
        
        # Test de connexion avec diagnostics
        print("\nConnexion à PostgreSQL...")
        print(f"  Host: {PG_HOST}")
        print(f"  Port: {PG_PORT}")
        print(f"  Database: {PG_DATABASE}")
        print(f"  User: {PG_USER}")
        
        try:
            pg_conn = psycopg2.connect(
                host=PG_HOST,
                port=PG_PORT,
                database=PG_DATABASE,
                user=PG_USER,
                password=PG_PASSWORD,
                connect_timeout=90,  # 90 secondes
                keepalives=1,
                keepalives_idle=30,
                keepalives_interval=10,
                keepalives_count=5
            )
            print("✓ Connecté à PostgreSQL")
        except psycopg2.OperationalError as e:
            error_msg = str(e)
            print(f"\n✗ Erreur de connexion PostgreSQL:")
            print(f"  {error_msg}")
            
            if "timeout" in error_msg.lower():
                print("\n💡 Suggestions:")
                print("  1. Vérifiez que le serveur PostgreSQL est accessible depuis votre réseau")
                print("  2. Vérifiez qu'aucun firewall ne bloque le port 52012")
                print("  3. Essayez depuis un autre réseau (VPN, connexion mobile, etc.)")
                print("  4. Contactez l'administrateur du serveur PostgreSQL")
            elif "authentication" in error_msg.lower() or "password" in error_msg.lower():
                print("\n💡 Les identifiants sont peut-être incorrects")
            elif "database" in error_msg.lower():
                print("\n💡 La base de données 'etl_tda' n'existe peut-être pas")
            
            return
        
        pg_cursor = pg_conn.cursor()
        
        # Connexion SQLite
        sqlite_conn = sqlite3.connect(SQLITE_DB)
        sqlite_cursor = sqlite_conn.cursor()
        
        now = datetime.now().isoformat()
        
        # Import uniquement des sessions des bornes mappées
        print("\nImport des sessions des bornes mappées...")
        mapped_ids = list(station_mapping.keys())
        placeholders = ','.join(['%s'] * len(mapped_ids))
        
        print(f"  Recherche de sessions pour {len(mapped_ids)} bornes...")
        
        pg_cursor.execute(f"""
            SELECT 
                id_transaction,
                id_borne,
                pdc_id,
                connector_id,
                perimetre,
                date_debut,
                date_fin,
                duree_charge,
                energie_consommee,
                prix_theorique_ttc,
                rfid,
                charge_reussie
            FROM reporting.historique_recharges
            WHERE id_borne IN ({placeholders})
            ORDER BY date_debut DESC
        """, mapped_ids)
        
        sessions = pg_cursor.fetchall()
        print(f"  {len(sessions)} sessions trouvées")
        
        if len(sessions) == 0:
            print("\n⚠️  Aucune session trouvée pour les bornes mappées")
            print("  Vérifiez que les id_borne dans config_stations.xlsx correspondent bien aux données PostgreSQL")
        
        for session in sessions:
            session_data = list(session)
            id_borne = session_data[1]
            station_name = station_mapping.get(id_borne, 'Inconnu')
            
            # Convertir les datetime en string pour SQLite
            if session_data[5]:  # date_debut
                if isinstance(session_data[5], datetime):
                    session_data[5] = session_data[5].isoformat()
            if session_data[6]:  # date_fin
                if isinstance(session_data[6], datetime):
                    session_data[6] = session_data[6].isoformat()
            
            sqlite_cursor.execute("""
                INSERT OR REPLACE INTO etl_sessions 
                (id_transaction, id_borne, pdc_id, connector_id, perimetre, station_name,
                 date_debut, date_fin, duree_sec, energie_wh, prix_ttc, rfid, charge_reussie, last_updated)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (*session_data[:5], station_name, *session_data[5:], now))
        
        print(f"✓ {len(sessions)} sessions importées")
        
        # Créer les stations à partir des sessions
        print("\nCréation des stations...")
        
        for id_borne, station_name in station_mapping.items():
            # Essayer de récupérer les infos dans la table borne si elles existent
            pg_cursor.execute("""
                SELECT 
                    chargebox_id,
                    nom_borne,
                    libelle,
                    alias_borne,
                    nom_zdc,
                    puissance_borne,
                    type_pdc,
                    nombre_pdc,
                    date_installation,
                    perimetre
                FROM reporting.borne
                WHERE id_borne = %s
            """, [id_borne])
            
            borne_info = pg_cursor.fetchone()
            
            if borne_info:
                # La borne existe dans la table
                station_data = [id_borne] + list(borne_info) + [station_name, now]
            else:
                # La borne n'existe pas, créer une entrée synthétique
                # Récupérer le périmètre depuis les sessions
                sqlite_cursor.execute("""
                    SELECT DISTINCT perimetre 
                    FROM etl_sessions 
                    WHERE id_borne = ?
                    LIMIT 1
                """, [id_borne])
                perimetre_row = sqlite_cursor.fetchone()
                perimetre = perimetre_row[0] if perimetre_row else 'Inconnu'
                
                station_data = [
                    id_borne,
                    None,  # chargebox_id
                    f"Borne {id_borne}",  # nom_borne
                    f"Borne {id_borne}",  # libelle
                    f"{station_name} - Borne {id_borne}",  # alias_borne
                    f"Borne {id_borne}",  # nom_zdc
                    None,  # puissance_borne
                    None,  # type_pdc
                    None,  # nombre_pdc
                    None,  # date_installation
                    perimetre,
                    station_name,
                    now
                ]
            
            # Convertir datetime si nécessaire
            if station_data[9] and isinstance(station_data[9], datetime):
                station_data[9] = station_data[9].isoformat()
            
            sqlite_cursor.execute("""
                INSERT OR REPLACE INTO etl_stations 
                (id_borne, chargebox_id, nom_borne, libelle, alias_borne, nom_zdc, 
                 puissance_borne, type_pdc, nombre_pdc, date_installation, perimetre, station_name, last_updated)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, station_data)
        
        print(f"✓ {len(station_mapping)} bornes créées")
        
        # Commit SQLite
        sqlite_conn.commit()
        
        # Résumé détaillé
        print("\n=== RÉSUMÉ PAR STATION ===")
        sqlite_cursor.execute("""
            SELECT station_name, COUNT(DISTINCT id_borne) as nb_bornes, COUNT(*) as nb_sessions
            FROM etl_sessions
            GROUP BY station_name
            ORDER BY nb_sessions DESC
        """)
        
        for row in sqlite_cursor.fetchall():
            print(f"{row[0]}: {row[1]} bornes, {row[2]} sessions")
        
        print("\n=== RÉSUMÉ GLOBAL ===")
        sqlite_cursor.execute("SELECT COUNT(*) FROM etl_stations")
        print(f"Total bornes: {sqlite_cursor.fetchone()[0]}")
        
        sqlite_cursor.execute("SELECT COUNT(*) FROM etl_sessions")
        print(f"Total sessions: {sqlite_cursor.fetchone()[0]}")
        
        # Date de la première et dernière session
        sqlite_cursor.execute("SELECT MIN(date_debut), MAX(date_debut) FROM etl_sessions WHERE date_debut IS NOT NULL")
        dates = sqlite_cursor.fetchone()
        if dates[0]:
            print(f"Période: de {dates[0][:10]} à {dates[1][:10]}")
        
    except Exception as e:
        print(f"✗ Erreur: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if pg_conn:
            pg_cursor.close()
            pg_conn.close()
            print("\n✓ Connexion PostgreSQL fermée")
        if sqlite_conn:
            sqlite_cursor.close()
            sqlite_conn.close()
            print("✓ Connexion SQLite fermée")

if __name__ == "__main__":
    print("=== IMPORT ETL POSTGRESQL ===")
    init_etl_database()
    import_etl_data()
    print("\n✓ Import terminé")