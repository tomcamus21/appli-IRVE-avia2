#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configuration centralisée avec variables d'environnement
"""
import os

# === BUMP API ===
BUMP_APP_ID = os.environ.get('BUMP_APP_ID', '2727ac86-4f15-4994-91a5-172e3006ee7b')
BUMP_APP_KEY = os.environ.get('BUMP_APP_KEY', '6ad38fdc-c5dc-4f7e-af06-3e44b0c59ea5')

# === PostgreSQL (pour EC_import) ===
PG_CONFIG = {
    'host': os.environ.get('PG_HOST', 'localhost'),
    'port': int(os.environ.get('PG_PORT', 5432)),
    'database': os.environ.get('PG_DATABASE', 'recharge_db'),
    'user': os.environ.get('PG_USER', 'postgres'),
    'password': os.environ.get('PG_PASSWORD', '')
}

# === Base de données SQLite ===
DB_PATH = os.environ.get('DB_PATH', 'bump_data.db')

# === Flask ===
SECRET_KEY = os.environ.get('SECRET_KEY', 'dev-key-change-in-production')
DEBUG = os.environ.get('DEBUG', 'False').lower() == 'true'
