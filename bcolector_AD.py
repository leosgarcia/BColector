import os
import sqlite3
import pandas as pd
from collections import Counter
from datetime import datetime, timedelta
import logging
import shutil
import tempfile
import psutil
import time
import tldextract

# Configuração de logging
logging.basicConfig(level=logging.DEBUG, filename='debug.log', filemode='w',
                    format='%(asctime)s - %(levelname)s - %(message)s')

def copy_db_file(original_path):
    try:
        temp_dir = tempfile.gettempdir()
        temp_file = os.path.join(temp_dir, 'temp_history.db')
        shutil.copy2(original_path, temp_file)
        return temp_file
    except Exception as e:
        logging.error(f"Erro ao copiar arquivo de histórico: {str(e)}")
        return None

def is_browser_running(browser_name):
    for proc in psutil.process_iter(['name']):
        if browser_name.lower() in proc.info['name'].lower():
            return True
    return False

def detect_installed_browsers():
    browsers = {}
    current_user = os.getenv('USERNAME')
    user_path = os.environ['USERPROFILE']
    logging.debug(f"Usuário atual: {current_user}")
    logging.debug(f"Caminho do usuário: {user_path}")
    
    browser_paths = {
        'chrome': [
            os.path.join(user_path, r"AppData\Local\Google\Chrome\User Data\Default"),
            os.path.join(r"C:\Users", current_user, r"AppData\Local\Google\Chrome\User Data\Default")
        ],
        'firefox': [
            os.path.join(user_path, r"AppData\Roaming\Mozilla\Firefox\Profiles"),
            os.path.join(r"C:\Users", current_user, r"AppData\Roaming\Mozilla\Firefox\Profiles")
        ],
        'edge': [
            os.path.join(user_path, r"AppData\Local\Microsoft\Edge\User Data\Default"),
            os.path.join(r"C:\Users", current_user, r"AppData\Local\Microsoft\Edge\User Data\Default")
        ],
        'opera': [
            os.path.join(user_path, r"AppData\Roaming\Opera Software\Opera Stable"),
            os.path.join(r"C:\Users", current_user, r"AppData\Roaming\Opera Software\Opera Stable")
        ],
        'brave': [
            os.path.join(user_path, r"AppData\Local\BraveSoftware\Brave-Browser\User Data\Default"),
            os.path.join(r"C:\Users", current_user, r"AppData\Local\BraveSoftware\Brave-Browser\User Data\Default")
        ]
    }
    
    for browser, paths in browser_paths.items():
        for path in paths:
            if os.path.exists(path):
                if browser == 'firefox':
                    profiles = [os.path.join(path, p) for p in os.listdir(path) if os.path.isdir(os.path.join(path, p))]
                    browsers[browser] = profiles
                else:
                    browsers[browser] = path
                logging.debug(f"{browser.capitalize()} encontrado em: {path}")
                break
    
    return browsers

def extract_domain(url):
    ext = tldextract.extract(url)
    return f"{ext.domain}.{ext.suffix}"

def get_history_from_chromium_based(path):
    history_db = os.path.join(path, 'History')
    temp_db = copy_db_file(history_db)
    if not temp_db:
        return Counter()
    
    c = Counter()
    try:
        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        ninety_days_ago = datetime.now() - timedelta(days=90)
        cursor.execute("SELECT url, visit_count FROM urls WHERE last_visit_time > ?", (ninety_days_ago.timestamp() * 1000000,))
        for row in cursor.fetchall():
            domain = extract_domain(row[0])
            c[domain] += row[1]
        conn.close()
        logging.debug(f"Histórico coletado com sucesso de: {path}")
    except sqlite3.OperationalError as e:
        logging.error(f"Erro ao acessar o histórico em {path}: {str(e)}")
    except Exception as e:
        logging.error(f"Erro inesperado ao coletar histórico de {path}: {str(e)}")
    finally:
        try:
            os.remove(temp_db)
        except:
            pass
    return c

def get_history_from_firefox(paths):
    c = Counter()
    for path in paths:
        history_db = os.path.join(path, 'places.sqlite')
        temp_db = copy_db_file(history_db)
        if not temp_db:
            continue
        
        try:
            conn = sqlite3.connect(temp_db)
            cursor = conn.cursor()
            ninety_days_ago = datetime.now() - timedelta(days=90)
            cursor.execute("SELECT url, visit_count FROM moz_places WHERE last_visit_date > ?", (ninety_days_ago.timestamp() * 1000000,))
            for row in cursor.fetchall():
                domain = extract_domain(row[0])
                c[domain] += row[1]
            conn.close()
            logging.debug(f"Histórico coletado com sucesso de: {path}")
        except sqlite3.OperationalError as e:
            logging.error(f"Erro ao acessar o histórico em {path}: {str(e)}")
        except Exception as e:
            logging.error(f"Erro inesperado ao coletar histórico de {path}: {str(e)}")
        finally:
            try:
                os.remove(temp_db)
            except:
                pass
    return c

def get_history_with_retry(func, path, max_retries=3):
    for attempt in range(max_retries):
        try:
            return func(path)
        except sqlite3.OperationalError as e:
            if "database is locked" in str(e) and attempt < max_retries - 1:
                logging.warning(f"Database locked, retrying in 5 seconds... (Attempt {attempt + 1})")
                time.sleep(5)
            else:
                raise

def get_running_browsers():
    browser_processes = {
        'chrome': ['chrome.exe', 'googlecrashhandler.exe', 'googlecrashhandler64.exe'],
        'firefox': ['firefox.exe'],
        'edge': ['msedge.exe'],
        'opera': ['opera.exe'],
        'brave': ['brave.exe']
    }
    
    running_browsers = []
    for browser, processes in browser_processes.items():
        if any(p.info['name'].lower() in processes for p in psutil.process_iter(['name'])):
            running_browsers.append(browser)
    
    return running_browsers

def close_browser(browser_name):
    browser_processes = {
        'chrome': ['chrome.exe', 'googlecrashhandler.exe', 'googlecrashhandler64.exe'],
        'firefox': ['firefox.exe'],
        'edge': ['msedge.exe'],
        'opera': ['opera.exe'],
        'brave': ['brave.exe']
    }
    
    for process_name in browser_processes.get(browser_name, []):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'].lower() == process_name.lower():
                try:
                    proc.terminate()
                    logging.info(f"Processo {process_name} terminado.")
                except psutil.NoSuchProcess:
                    logging.warning(f"Processo {process_name} não encontrado.")
                except psutil.AccessDenied:
                    logging.error(f"Acesso negado ao tentar terminar {process_name}.")
    
    time.sleep(2)  # Aguarda um pouco para os processos serem encerrados

def generate_report():
    running_browsers = get_running_browsers()
    if running_browsers:
        print(f"Os seguintes navegadores estão em execução: {', '.join(running_browsers)}")
        choice = input("Deseja fechar estes navegadores automaticamente? (s/n): ").lower()
        if choice == 's':
            for browser in running_browsers:
                close_browser(browser)
        else:
            print("Por favor, feche os navegadores manualmente antes de continuar.")
            return

    browsers = detect_installed_browsers()
    overall_counter = Counter()

    for browser, path in browsers.items():
        if browser == 'firefox':
            overall_counter += get_history_with_retry(get_history_from_firefox, path)
        else:
            overall_counter += get_history_with_retry(get_history_from_chromium_based, path)
    
    # Gerar relatório ordenado por frequência de acesso
    report = sorted(overall_counter.items(), key=lambda x: x[1], reverse=True)
    
    # Gerar planilha Excel
    df = pd.DataFrame(report, columns=['Domínio', 'Visitas'])
    output_file = 'historico_navegacao_por_dominio.xlsx'
    df.to_excel(output_file, index=False)
    logging.info(f"Relatório salvo em '{output_file}'")
    print(f"Relatório salvo em '{output_file}'")

if __name__ == "__main__":
    try:
        generate_report()
    except Exception as e:
        logging.critical(f"Erro crítico durante a execução do script: {str(e)}", exc_info=True)
        print(f"Ocorreu um erro. Por favor, verifique o arquivo de log 'debug.log' para mais detalhes.")
