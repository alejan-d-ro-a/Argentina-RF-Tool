import json
import os
import sys
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ConfigManager:
    APP_NAME = "ArgentinaRFTool"
    
    if sys.platform == "win32":
        BASE_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), APP_NAME)
    else:
        BASE_DIR = os.path.join(os.path.expanduser("~"), f".{APP_NAME}")
        
    os.makedirs(BASE_DIR, exist_ok=True)
    CONFIG_FILE = os.path.join(BASE_DIR, "rf_config.json")

    @staticmethod
    def load_config():
        if not os.path.exists(ConfigManager.CONFIG_FILE):
            return {}
        try:
            with open(ConfigManager.CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                valid_config = {}
                for key, value in config.items():
                    if isinstance(value, str) and os.path.exists(value):
                        valid_config[key] = value
                    else:
                        valid_config[key] = ""
                return valid_config
        except json.JSONDecodeError as e:
            logging.error(f"El archivo de configuración está corrupto: {e}. Iniciando vacío.")
            return {}
        except Exception as e:
            logging.error(f"Error al leer la configuración: {e}")
            return {}

    @staticmethod
    def save_config(config_data):
        try:
            with open(ConfigManager.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            logging.error(f"Error al guardar la configuración: {e}")
