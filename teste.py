import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox, ttk # <--- ADICIONADO O TTK
import sys
import threading
import time
import queue
import os
import pandas as pd
from datetime import datetime, date # <--- GARANTINDO O IMPORT DO DATETIME
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ==============================================================================
# CONFIGURAÇÕES E FUNÇÕES AUXILIARES
# ==============================================================================

ARQUIVO_LOG = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'log_robo_iss.txt')

def registrar_log(mensagem):
    """Registra no arquivo e exibe na tela (via print)"""
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    texto = f"[{timestamp}] {mensagem}\n"
    print(texto.strip())
    try:
        os.makedirs(os.path.dirname(ARQUIVO_LOG), exist_ok=True)
        with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
            f.write(texto)
    except:
        pass

def clicar_seguro(driver, xpath, timeout=5):
    try:
        wait = WebDriverWait(driver, timeout)
        elemento = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            elemento.click()
            return True
        except:
            driver.execute_script("arguments[0].click();", elemento)
            return True
    except:
        return False

def tratar_mensagens_alerta(driver, wait):
    print("   [Check] Verificando mensagens/alertas...")
    while True:
        try:
            if not clicar_seguro(driver, '//*[@id="mensagensForm:mensagemDataTable:0:linkVisualizar"]/i', timeout=2):
                break
            time.sleep(1)
            clicar_seguro(driver, '//*[@id="mensagensForm:botaoDarCiencia"]', timeout=2)
            time.sleep(2)
        except:
            break

# ==============================================================================
# LÓGICA DA AUTOMAÇÃO (THREAD)
# ==============================================================================
# Agora recebemos mes e ano como argumentos
def minha_automacao(caminho_planilha, diretorio_notas, modo_headless, mes_selecionado, ano_selecionado):
    
    # --- 1. CONVERSÃO E DEFINIÇÃO DA COMPETÊNCIA MANUAL ---
    try:
        mes_alvo = int(mes_selecionado)
        ano_alvo = int(ano_selecionado)
        str_competencia = f"{mes_alvo:02d}.{ano_alvo}"
        print(f">>> Competência definida manualmente: {str_competencia}")
    except ValueError:
        registrar_log("ERRO CRÍTICO: Mês ou Ano inválidos selecionados.")
        return
    # ------------------------------------------------------

    print(">>> INICIANDO AUTOMAÇÃO...")
    print(f">>> Modo Invisível: {'ATIVADO' if modo_headless else 'DESATIVADO'}")
    
    if not os.path.exists(caminho_planilha):
        registrar_log(f"ERRO: A planilha não foi encontrada.")
        return

    if not os.path.exists(diretorio_notas):
        try:
            os.makedirs(diretorio_notas)
        except Exception as e:
            registrar_log(f"ERRO AO CRIAR PASTA DE DESTINO: {e}")
            return

    try:
        df_empresas = pd.read_excel(caminho_planilha, dtype=str)
        df_empresas.columns = df_empresas.columns.str.strip().str.upper()
    except Exception as e:
        registrar_log(f"ERRO CRÍTICO AO LER PLANILHA: {e}")
        return

    for index, row in df_empresas.iterrows():
        try:
            NOME_EMPRESA = str(row['NOME']).strip()
            CNPJ_RAW = str(row['CNPJ']).replace(".", "").replace("/", "").replace("-", "").strip().zfill(14)
            LOGIN_CPF = str(row['CPF']).replace(".", "").replace("-", "").strip()
            SENHA_ACESSO = str(row['SENHA']).strip()
        except KeyError as e:
            registrar_log(f"ERRO: Coluna {e} não encontrada na planilha.")
            return

        print(f"\n{'='*60}")
        print(f"EMPRESA {index + 1}/{len(df_empresas)}: {NOME_EMPRESA}")
        print(f"Competência: {str_competencia}")
        print(f"{'='*60}")
        
        # Garante caminho absoluto e cria pasta
        pasta_final = os.path.abspath(os.path.join(diretorio_notas, f"{NOME_EMPRESA} {CNPJ_RAW}", str_competencia))
        if not os.path.exists(pasta_final):
            os.makedirs(pasta_final)

        # --- CONFIGURAÇÕES DO CHROME ---
        chrome_prefs = {
            "download.default_directory": pasta_final,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "safebrowsing.disable_download_protection": True,
            "profile.default_content_setting_values.automatic_downloads": 1,
            "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
            "profile.default_content_settings.popups": 0,
            "profile.password_manager_enabled": False,
        }
        
        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", chrome_prefs)
        options.add_argument("--disable-safebrowsing-disable-download-protection")
        options.add_argument("--safebrowsing-disable-download-protection")
        options.add_argument("--disable-features=InsecureDownloadWarnings")
        options.add_argument("--log-level=3")
        options.add_argument("--window-size=1920,1080")

        if modo_headless:
            options.add_argument("--headless=new")

        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            wait = WebDriverWait(driver, 15)
        except Exception as e:
            registrar_log(f"ERRO AO ABRIR CHROME: {e}")
            return

        try:
            # 1. LOGIN
            driver.get("https://idp2.sefin.fortaleza.ce.gov.br/realms/sefin/protocol/openid-connect/auth?nonce=705900bf-b300-47cb-b193-414e60359c46&response_type=code&client_id=iss.sefin.fortaleza.ce.gov.br&redirect_uri=https%3A%2F%2Fiss.fortaleza.ce.gov.br%2Fgrpfor%2Foauth2%2Fcallback&scope=openid+profile&state=secret-c677fae4-b35c-4fbb-a30d-11befc8b05fb")
            driver.find_element(By.XPATH, '//*[@id="username"]').send_keys(LOGIN_CPF)
            driver.find_element(By.XPATH, '//*[@id="password"]').send_keys(SENHA_ACESSO)
            
            if not clicar_seguro(driver, '//*[@id="botao-entrar"]'):
                registrar_log(f"FALHA: Não foi possível clicar no botão de entrar.")
                driver.quit()
                continue

            time.sleep(2)
            try:
                driver.find_element(By.XPATH, "//*[contains(text(), 'Nome de usuário ou senha inválida')]")
                registrar_log(f"FALHA LOGIN: {NOME_EMPRESA} - Senha Incorreta.")
                driver.quit()
                continue
            except NoSuchElementException:
                pass

            clicar_seguro(driver, '//*[@id="login"]/div[1]/div[2]/a[1]', timeout=5)
            tratar_mensagens_alerta(driver, wait)

            # 2. SELEÇÃO DA EMPRESA
            print("2. Verificando empresa...")
            time.sleep(2)
            try:
                wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="alteraInscricaoForm:empresaDataTable:0:linkImagem"]/i')))
                if not clicar_seguro(driver, '//*[@id="alteraInscricaoForm:tipoPesquisa:1"]', timeout=3):
                     clicar_seguro(driver, "//label[contains(text(), 'CNPJ')]", timeout=3)
                
                time.sleep(1)
                driver.find_element(By.XPATH, "//form[@id='alteraInscricaoForm']//input[@type='text']").clear()
                driver.find_element(By.XPATH, "//form[@id='alteraInscricaoForm']//input[@type='text']").send_keys(CNPJ_RAW)
                clicar_seguro(driver, '//*[@id="alteraInscricaoForm:btnPesquisar"]')
                time.sleep(3)
                
                if not clicar_seguro(driver, '//*[@id="alteraInscricaoForm:empresaDataTable:0:linkImagem"]/i'):
                    print("   -> Empresa não encontrada na busca.")
                time.sleep(3)
            except TimeoutException:
                try:
                    driver.find_element(By.XPATH, f"//a[contains(text(), '{CNPJ_RAW[:8]}')]").click()
                    time.sleep(3)
                except: pass

            try:
                wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="navbar"]')))
            except:
                driver.refresh()
                wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="navbar"]')))

            tratar_mensagens_alerta(driver, wait)

            # 3. DOWNLOAD XMLs
            print("3. Baixando Notas (XMLs)...")
            
            if "consultar_nfse" not in driver.current_url:
                clicar_seguro(driver, '//*[@id="navbar"]/ul/li[5]/a')
                time.sleep(1)
                clicar_seguro(driver, '//*[@id="formMenuTopo:menuNfse:j_id61"]')
            
            time.sleep(2)

            def baixar_xmls_paginado():
                pag = 1
                ultimo_conteudo_tabela = ""

                while True:
                    try:
                        # Seleciona "Todos"
                        if not clicar_seguro(driver, '//*[@id="consultarnfseForm:j_id324"]', timeout=3):
                             try:
                                 chk = driver.find_element(By.XPATH, "//th//input[@type='checkbox']")
                                 driver.execute_script("arguments[0].click();", chk)
                             except:
                                 print("   -> Nenhuma nota encontrada nesta página.")
                                 break 

                        time.sleep(1)
                        # Clica em Exportar
                        if not clicar_seguro(driver, "//input[contains(@title, 'Exportar XML')]", timeout=2):
                            clicar_seguro(driver, '//*[@id="consultarnfseForm:j_id321"]/div[1]/input[3]')

                        time.sleep(5) 

                        # --- TRAVA DE LOOP ---
                        try:
                            elemento_tabela = driver.find_element(By.XPATH, "//tbody")
                            conteudo_atual = elemento_tabela.text[:100]
                        except:
                            conteudo_atual = "vazio"

                        if pag > 1 and conteudo_atual == ultimo_conteudo_tabela:
                            print(f"   -> Página {pag} é igual à anterior. Fim da paginação.")
                            break
                        
                        ultimo_conteudo_tabela = conteudo_atual
                        # ---------------------
                        
                        prox = pag + 1
                        xpath_prox = f"//td[contains(@class, 'rich-datascr-inact') and normalize-space()='{prox}']"
                        
                        try:
                            btn = driver.find_element(By.XPATH, xpath_prox)
                            driver.execute_script("arguments[0].scrollIntoView();", btn)
                            driver.execute_script("arguments[0].click();", btn)
                            print(f"   -> Indo para página {prox}...")
                            time.sleep(4)
                            pag = prox
                        except:
                            print(f"   -> Fim das páginas.")
                            break
                    except Exception as e:
                        break

            # [A] PRESTADOS
            print("   -> Processando Prestados...")
            clicar_seguro(driver, '//*[@id="consultarnfseForm:competencia_prestador_tab_lbl"]')
            time.sleep(1)
            if clicar_seguro(driver, '//*[@id="consultarnfseForm:competenciaHeader"]/label/div'):
                time.sleep(1)
                # Seleção de ano e mês baseada na escolha manual
                clicar_seguro(driver, f'//*[@id="consultarnfseForm:competenciaDateEditorLayoutY{ano_alvo - 2022}"]')
                time.sleep(0.5)
                clicar_seguro(driver, f'//*[@id="consultarnfseForm:competenciaDateEditorLayoutM{mes_alvo - 1}"]')
                time.sleep(0.5)
                clicar_seguro(driver, '//*[@id="consultarnfseForm:competenciaDateEditorButtonOk"]/span')
                time.sleep(1)
                
                if not clicar_seguro(driver, '//*[@id="consultarnfseForm:j_id237"]'):
                    clicar_seguro(driver, "//input[@value='Consultar']")
                
                time.sleep(3)
                baixar_xmls_paginado()

            # [B] TOMADOS
            print("   -> Processando Tomados...")
            clicar_seguro(driver, '//*[@id="consultarnfseForm:opTipoRelatorio:1"]') 
            time.sleep(3)
            clicar_seguro(driver, '//*[@id="consultarnfseForm:abaPorCompetenciaTomador_tab_lbl"]')
            time.sleep(1)
            if clicar_seguro(driver, '//*[@id="consultarnfseForm:competenciaTomadorHeader"]/label/div'):
                time.sleep(1)
                clicar_seguro(driver, f'//*[@id="consultarnfseForm:competenciaTomadorDateEditorLayoutY{ano_alvo - 2022}"]')
                time.sleep(0.5)
                clicar_seguro(driver, f'//*[@id="consultarnfseForm:competenciaTomadorDateEditorLayoutM{mes_alvo - 1}"]')
                time.sleep(0.5)
                clicar_seguro(driver, '//*[@id="consultarnfseForm:competenciaTomadorDateEditorButtonOk"]/span')
                time.sleep(1)
                
                if not clicar_seguro(driver, '//*[@id="consultarnfseForm:j_id317"]'):
                    try:
                         btns = driver.find_elements(By.XPATH, "//input[@value='Consultar']")
                         for btn in btns:
                             if btn.is_displayed():
                                 driver.execute_script("arguments[0].click();", btn)
                                 break
                    except: pass
                
                time.sleep(3)
                baixar_xmls_paginado()

            # 4. ENCERRAMENTO
            print("4. Escrituração (Encerramento)...")
            clicar_seguro(driver, '//*[@id="navbar"]/ul/li[6]/a')
            clicar_seguro(driver, '//*[@id="formMenuTopo:menuEscrituracao:j_id80"]')
            
            if clicar_seguro(driver, '//*[@id="manterEscrituracaoForm:dataInicialHeader"]/label/div'):
                time.sleep(1)
                clicar_seguro(driver, f'//*[@id="manterEscrituracaoForm:dataInicialDateEditorLayoutY{ano_alvo - 2022}"]')
                time.sleep(0.5)
                clicar_seguro(driver, f'//*[@id="manterEscrituracaoForm:dataInicialDateEditorLayoutM{mes_alvo - 1}"]')
                time.sleep(0.5)
                clicar_seguro(driver, '//*[@id="manterEscrituracaoForm:dataInicialDateEditorButtonOk"]/span')
                time.sleep(1)

            if clicar_seguro(driver, '//*[@id="manterEscrituracaoForm:dataFinalHeader"]/label/div'):
                time.sleep(1)
                clicar_seguro(driver, f'//*[@id="manterEscrituracaoForm:dataFinalDateEditorLayoutY{ano_alvo - 2022}"]')
                time.sleep(0.5)
                clicar_seguro(driver, f'//*[@id="manterEscrituracaoForm:dataFinalDateEditorLayoutM{mes_alvo - 1}"]')
                time.sleep(0.5)
                clicar_seguro(driver, '//*[@id="manterEscrituracaoForm:dataFinalDateEditorButtonOk"]/span')
                time.sleep(1)
            
            clicar_seguro(driver, '//*[@id="manterEscrituracaoForm:btnConsultar"]')
            time.sleep(3)
            
            if clicar_seguro(driver, '//*[@id="manterEscrituracaoForm:dataTablePendentes:0:linkEscriturarPendente"]/span', timeout=5):
                print("   -> Aberta. Realizando Encerramento...")
                time.sleep(3)
                
                if clicar_seguro(driver, '//*[@id="aba_servicos_pendentes_lbl"]', timeout=3):
                    time.sleep(2)
                    clicar_seguro(driver, '//*[@id="servicos_pendentes_form:idLinkaceitarDocTomados"]')
                    clicar_seguro(driver, '//*[@id="aceite_todos_doc_tomados_modal_panel_form:btnSim"]')
                    time.sleep(5)
                
                clicar_seguro(driver, '//*[@id="abaEncerramento_lbl"]')
                time.sleep(2)
                clicar_seguro(driver, '//*[@id="abaEncerramentoForm:btnEncerrarEscrituracao"]')
                clicar_seguro(driver, '//*[@id="formEncerramento:btnSim"]')
                time.sleep(5)
                print("   -> Escrituração ENCERRADA com sucesso.")
            else:
                print("   -> Nenhuma pendência encontrada ou já encerrada.")

            registrar_log(f"SUCESSO TOTAL: {NOME_EMPRESA}")

        except Exception as e:
            registrar_log(f"ERRO GERAL: {NOME_EMPRESA} - {e}")
        
        finally:
            driver.quit()

    print("\n>>> PROCESSO CONCLUÍDO! Verifique a pasta escolhida.")

# ==============================================================================
# INTERFACE GRÁFICA (GUI)
# ==============================================================================
class LogQueue:
    def __init__(self, log_queue):
        self.log_queue = log_queue
    def write(self, message):
        self.log_queue.put(message)
    def flush(self):
        pass

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Robô ISS Fortaleza - Seletor de Data")
        self.root.geometry("750x650")
        self.root.configure(bg="#f0f0f0")

        self.log_queue = queue.Queue()

        frame_config = tk.LabelFrame(root, text="Configurações", padx=10, pady=10, bg="#f0f0f0", font=("Arial", 10, "bold"))
        frame_config.pack(padx=10, pady=10, fill=tk.X)

        tk.Label(frame_config, text="Planilha de Acessos:", bg="#f0f0f0").pack(anchor="w")
        frame_p = tk.Frame(frame_config, bg="#f0f0f0")
        frame_p.pack(fill=tk.X, pady=(0, 5))
        self.var_planilha = tk.StringVar()
        tk.Entry(frame_p, textvariable=self.var_planilha, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        tk.Button(frame_p, text="Selecionar...", command=self.selecionar_planilha).pack(side=tk.RIGHT)

        tk.Label(frame_config, text="Página de download:", bg="#f0f0f0").pack(anchor="w")
        frame_d = tk.Frame(frame_config, bg="#f0f0f0")
        frame_d.pack(fill=tk.X, pady=(0, 5))
        self.var_destino = tk.StringVar(value="") 
        tk.Entry(frame_d, textvariable=self.var_destino, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        tk.Button(frame_d, text="Selecionar...", command=self.selecionar_pasta_destino).pack(side=tk.RIGHT)

        self.var_headless = tk.BooleanVar(value=False)
        chk_headless = tk.Checkbutton(frame_config, text="Executar em modo invisível", variable=self.var_headless, bg="#f0f0f0", font=("Arial", 9))
        chk_headless.pack(anchor="w", pady=5)

        # --- SELEÇÃO DE DATA ---
        frame_data = tk.LabelFrame(frame_config, text="Competência Alvo", bg="#f0f0f0", font=("Arial", 9, "bold"))
        frame_data.pack(fill=tk.X, pady=5)

        tk.Label(frame_data, text="Mês:", bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        self.combo_mes = ttk.Combobox(frame_data, values=[f"{i:02d}" for i in range(1, 13)], width=5, state="readonly")
        self.combo_mes.pack(side=tk.LEFT, padx=5)
        self.combo_mes.set(datetime.now().strftime("%m"))

        tk.Label(frame_data, text="Ano:", bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        anos_disponiveis = [str(a) for a in range(2023, 2031)]
        self.combo_ano = ttk.Combobox(frame_data, values=anos_disponiveis, width=8, state="readonly")
        self.combo_ano.pack(side=tk.LEFT, padx=5)
        self.combo_ano.set(datetime.now().strftime("%Y"))
        # ------------------------

        self.btn_iniciar = tk.Button(root, text="INICIAR AUTOMAÇÃO", command=self.iniciar_thread, height=2, bg="#0056b3", fg="white", font=("Arial", 12, "bold"))
        self.btn_iniciar.pack(pady=5, padx=10, fill=tk.X)

        tk.Label(root, text="Logs de Execução:", bg="#f0f0f0", font=("Arial", 10, "bold")).pack(anchor="w", padx=10)
        self.txt_log = scrolledtext.ScrolledText(root, state='disabled', height=15, font=("Consolas", 9))
        self.txt_log.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        sys.stdout = LogQueue(self.log_queue)
        self.atualizar_logs()

    def selecionar_planilha(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f: self.var_planilha.set(f)

    def selecionar_pasta_destino(self):
        d = filedialog.askdirectory()
        if d: self.var_destino.set(d)

    def iniciar_thread(self):
        p = self.var_planilha.get()
        d = self.var_destino.get()
        h = self.var_headless.get()
        
        # Obtém a escolha do usuário
        m = self.combo_mes.get()
        a = self.combo_ano.get()

        if not p or not os.path.exists(p):
            messagebox.showwarning("Erro", "Selecione uma planilha válida.")
            return
        
        if not d:
            messagebox.showwarning("Erro", "Selecione a Página de download.")
            return

        # Passamos Mês e Ano para a função
        self.btn_iniciar.config(state=tk.DISABLED, text="Rodando...", bg="#6c757d")
        
        thread = threading.Thread(target=minha_automacao, args=(p, d, h, m, a))
        thread.daemon = True
        thread.start()

    def atualizar_logs(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.txt_log.config(state='normal')
            self.txt_log.insert(tk.END, msg)
            self.txt_log.see(tk.END)
            self.txt_log.config(state='disabled')
        self.root.after(100, self.atualizar_logs)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()