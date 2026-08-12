# -*- coding: utf-8 -*-

import customtkinter as ctk
import pandas as pd
from datetime import datetime, timedelta
import os
import warnings
import sys
import requests
import io
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import traceback
import threading
import time
import logging

# =========================================
# CONFIGURAÇÃO DE LOGGING
# =========================================

if getattr(sys, 'frozen', False):
    PASTA = os.path.dirname(sys.executable)
else:
    PASTA = os.path.dirname(os.path.abspath(__file__))

log_file = os.path.join(PASTA, 'operacional.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)  # também exibe no console
    ]
)

logger = logging.getLogger('OperacionalDB')

def handle_exception(exc_type, exc_value, exc_traceback):
    """Captura exceções não tratadas e registra no log."""
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logger.critical("Exceção não tratada:", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = handle_exception

# =========================================
# IGNORAR AVISOS DO SISTEMA (OPENPYXL/PANDAS)
# =========================================
warnings.filterwarnings('ignore', category=UserWarning)
warnings.filterwarnings('ignore')

# =========================================
# CONFIGURAÇÕES DE DIRETÓRIO E VARIÁVEIS
# =========================================

LITORAL_SP = ['CGT', 'BASE', 'SST', 'SSTBO', 'IBL']

RJ = [
    'RSD', 'MPE', 'VAS', 'VRD',
    'PNDO', 'VLC', 'IZA', 'TRS', 'BMA',
    'PORE', 'COLG', 'BPI', 'PFS', 'PDS',
    'PNHE'
]

TODAS_CIDADES = LITORAL_SP + RJ

DICIONARIO_CIDADES = {
    'CGT': 'Caraguatatuba',
    'BASE': 'Caraguatatuba1',
    'SST': 'São Sebastião',
    'SSTBO': 'Boiçucanga',
    'IBL': 'Ilhabela',
    'RSD': 'Resende',
    'MPE': 'Miguel Pereira',
    'VAS': 'Vassouras',
    'VRD': 'Volta Redonda',
    'PNDO': 'Penedo',
    'VLC': 'Valença',
    'IZA': 'Itatiaia',
    'TRS': 'Três Rios',
    'BMA': 'Barra Mansa',
    'PORE': 'Porto Real',
    'COLG': 'Comendador Levy Gasparian',
    'BPI': 'Barra do Piraí',
    'PFS': 'Paty do Alferes',
    'PDS': 'Paraíba do Sul',
    'PNHE': 'Pinheiral'
}

# =========================================
# CLASSE DE CACHE
# =========================================
class DataCache:
    def __init__(self):
        self._cache = {}
        self.last_activity = time.time()

    def is_valid(self, key, ttl=300):
        if key not in self._cache:
            return False
        return (time.time() - self._cache[key]['timestamp']) < ttl

    def get(self, key):
        return self._cache.get(key, {}).get('data')

    def set(self, key, data):
        self._cache[key] = {'data': data, 'timestamp': time.time()}

    def clear(self):
        self._cache.clear()
        logger.debug("Cache limpo manualmente.")

    def update_activity(self):
        self.last_activity = time.time()

    def is_idle(self, seconds=300):
        return (time.time() - self.last_activity) > seconds

# =========================================
# FUNÇÕES AUXILIARES DE DADOS (PANDAS)
# =========================================

def carregar_excel(nome):
    caminho = f"{PASTA}\\{nome}"
    try:
        if nome.lower().endswith('.csv'):
            try:
                df = pd.read_csv(caminho, sep=';', encoding='utf-8', on_bad_lines='skip')
            except:
                try:
                    df = pd.read_csv(caminho, sep=';', encoding='latin1', on_bad_lines='skip')
                except:
                    df = pd.read_csv(caminho, sep=';', on_bad_lines='skip')
        else:
            df = pd.read_excel(caminho)

        df.columns = df.columns.str.strip()
        return df

    except Exception:
        logger.exception(f"Erro ao carregar arquivo {nome}")
        return pd.DataFrame()

def carregar_google_sheet(sheet_id, aba="Sheet1"):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            os.path.join(PASTA, "google_credentials.json"), scope
        )
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id)
        worksheet = sheet.worksheet(aba)
        dados = worksheet.get_all_records()
        df = pd.DataFrame(dados)
        df.columns = df.columns.str.strip()
        return df
    except Exception:
        logger.exception(f"Falha ao carregar Google Sheet {sheet_id}/{aba}")
        return pd.DataFrame()

def carregar_csv_robusto(nome_arquivo):
    caminho = os.path.join(PASTA, nome_arquivo)
    for sep in [';', ',']:
        for enc in ['utf-8', 'latin1']:
            try:
                df = pd.read_csv(caminho, sep=sep, encoding=enc, on_bad_lines='skip')
                if len(df.columns) > 1:
                    df.columns = df.columns.str.strip()
                    return df
            except:
                continue
    logger.warning(f"Não foi possível ler {nome_arquivo} como CSV, tentando Excel.")
    return carregar_excel(nome_arquivo)

def normalizar_data(coluna):
    return pd.to_datetime(coluna, errors='coerce', dayfirst=True).dt.normalize()

def tratar_contrato(coluna):
    return coluna.astype(str).str.split('.').str[0].str.strip()

def filtrar_cidades(df):
    return df

def encontrar_coluna(df, possibilidades):
    for p in possibilidades:
        for col in df.columns:
            if col.strip().lower() == p.lower():
                return col
    return None

def parse_data_flexivel(data_str):
    """
    Interpreta string de data nos formatos:
    'dd/mm/aaaa', 'dd/mm/aa', 'dd/mm' (assume ano que torna a data futura).
    Retorna pd.Timestamp normalizado ou NaT se não conseguir.
    """
    if pd.isna(data_str) or str(data_str).strip() == '':
        return pd.NaT

    texto = str(data_str).strip()

    # Tenta formatos com dia/mês/ano completos ou de dois dígitos
    for fmt in ['%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y', '%d-%m-%y', '%d.%m.%Y', '%d.%m.%y']:
        try:
            return pd.to_datetime(texto, format=fmt, dayfirst=True, errors='raise').normalize()
        except (ValueError, TypeError):
            continue

    # Apenas dia/mês (ex: "30/06") ou "30/06/ano?" já tratado acima, mas aqui apenas dois números
    if '/' in texto and texto.count('/') == 1:
        partes = texto.split('/')
        if len(partes) == 2 and partes[0].isdigit() and partes[1].isdigit():
            dia, mes = int(partes[0]), int(partes[1])
            if not (1 <= mes <= 12 and 1 <= dia <= 31):
                return pd.NaT
            hoje = pd.Timestamp.now().normalize()
            try:
                # Tenta ano atual
                tentativa = pd.Timestamp(year=hoje.year, month=mes, day=dia)
                if tentativa > hoje:
                    return tentativa
                # Tenta ano seguinte
                tentativa = pd.Timestamp(year=hoje.year + 1, month=mes, day=dia)
                return tentativa if tentativa > hoje else pd.NaT
            except (ValueError, TypeError):
                return pd.NaT

    # Similar para "dd-mm"
    if '-' in texto and texto.count('-') == 1:
        partes = texto.split('-')
        if len(partes) == 2 and partes[0].isdigit() and partes[1].isdigit():
            dia, mes = int(partes[0]), int(partes[1])
            if not (1 <= mes <= 12 and 1 <= dia <= 31):
                return pd.NaT
            hoje = pd.Timestamp.now().normalize()
            try:
                tentativa = pd.Timestamp(year=hoje.year, month=mes, day=dia)
                if tentativa > hoje:
                    return tentativa
                tentativa = pd.Timestamp(year=hoje.year + 1, month=mes, day=dia)
                return tentativa if tentativa > hoje else pd.NaT
            except (ValueError, TypeError):
                return pd.NaT

    return pd.NaT

def carregar_conveniencias():
    SHEET_ID = "SEU_ID_DA_PLANILHA_GOOGLE"
    ABA = "CONVENIENCIA"
    
    try:
        logger.info("Carregando planilha de conveniência do Google...")
        df = carregar_google_sheet(SHEET_ID, ABA)
        logger.info("Planilha de conveniência carregada online.")
    except Exception:
        logger.exception("Erro ao carregar conveniência do Google Sheets.")
        df = carregar_excel("PLANILHA DE CONVENIENCIA.xlsx")
        if df.empty:
            logger.warning("Arquivo local de conveniência não encontrado ou vazio.")
    
    if df.empty:
        return []
    try:
        # Procura colunas por nomes (mais robusto), caso não encontre usa índices fixos
        col_contrato = encontrar_coluna(df, ['CONTRATO', 'Contrato', 'CÓDIGO CONTRATO'])
        col_data = encontrar_coluna(df, ['DATA', 'Data', 'Data de Vencimento', 'Vencimento'])

        if not col_contrato or not col_data:
            logger.warning("Colunas não encontradas por nome, usando índices fixos (A e C).")
            col_contrato = df.columns[0] if not df.empty else None
            col_data = df.columns[2] if len(df.columns) > 2 else None

        if not col_contrato or not col_data:
            logger.error("Não foi possível identificar colunas de contrato e data na planilha de conveniência.")
            return []

        contratos = tratar_contrato(df[col_contrato])
        # Aplica o parser flexível de datas
        datas = df[col_data].apply(parse_data_flexivel)

        hoje = pd.Timestamp.now().normalize()
        valid_mask = datas > hoje
        invalidas = datas.isna() | (datas <= hoje)
        logger.info(f"Conveniências: {valid_mask.sum()} válidas, {invalidas.sum()} ignoradas (datas inválidas/passadas).")

        validos = contratos[valid_mask]
        return validos.tolist()
    except Exception:
        logger.exception("Erro ao processar dados da conveniência.")
        return []

def carregar_forms():
    SHEET_FORMS_ID = "SEU_ID_DA_PLANILHA_GOOGLE"
    ABA_FORMS = "Respostas ao formulário 1"
    try:
        logger.info("Carregando respostas do Forms...")
        df = carregar_google_sheet(SHEET_FORMS_ID, ABA_FORMS)
        logger.info("Forms carregado online.")
    except Exception:
        logger.exception("Erro ao carregar forms do Google Sheets.")
        df = pd.DataFrame()

    if df.empty:
        df = carregar_excel("forms confirmação.xlsx")
        if df.empty:
            df = carregar_csv_robusto("forms confirmação.csv")
            if df.empty:
                df = carregar_csv_robusto("forms confirmação.xlsx - Respostas ao formulário 1.csv")
    return df

def carregar_chamados():
    ch = carregar_excel("chamados_abertos_field_service.xlsx")
    if ch.empty:
        logger.warning("Arquivo de chamados vazio ou não encontrado.")
        return ch

    ch.columns = ch.columns.str.strip()
    if 'CÓDIGO CONTRATO' in ch.columns:
        ch['CÓDIGO CONTRATO'] = tratar_contrato(ch['CÓDIGO CONTRATO'])
    ch = filtrar_cidades(ch)
    if 'DATA AGENDAMENTO' in ch.columns:
        ch['DATA AGENDAMENTO'] = normalizar_data(ch['DATA AGENDAMENTO'])
    if 'DATA DE INGRESSO' in ch.columns:
        ch['DATA DE INGRESSO'] = pd.to_datetime(ch['DATA DE INGRESSO'], errors='coerce', dayfirst=True)
    return ch

# =========================================
# FUNÇÕES DO AUTENTICADOR
# =========================================

def consultar_autenticador_status(lista_contratos):
    try:
        import urllib3
        urllib3.disable_warnings()

        sessao = requests.Session()
        payload = {"contratos": "\n".join(lista_contratos)}
        
        # Conexão estabelece em 5s e espera resposta por até 35s para evitar ReadTimeout
        sessao.post("https://provedor.example/status.php?action=save", data=payload, verify=False, timeout=(5, 35))
        sessao.get("https://provedor.example/processa.php?bg=1", verify=False, timeout=(5, 35))
        res = sessao.get("https://provedor.example/ler_csv.php", verify=False, timeout=(5, 35))
        html = res.text

        if '<table>' not in html:
            return pd.DataFrame(), "Resposta do servidor não contém tabela."

        tabelas = pd.read_html(io.StringIO(html))
        if not tabelas:
            return pd.DataFrame(), "Nenhuma tabela encontrada na resposta."

        df = tabelas[0]
        df.columns = [c.lower().strip() for c in df.columns]

        col_map = {
            'contrato': 'CONTRATO', 'username': 'USERNAME', 'acctstarttime': 'INÍCIO',
            'acctstoptime': 'FIM', 'circuitid': 'CIRCUITO', 'callingstationid': 'MAC',
            'trafego': 'TRÁFEGO', 'servidor': 'SERVIDOR'
        }
        df.rename(columns=col_map, inplace=True)
        for col in ['CONTRATO', 'USERNAME', 'INÍCIO', 'FIM', 'CIRCUITO', 'MAC', 'TRÁFEGO', 'SERVIDOR']:
            if col not in df.columns:
                df[col] = ''

        df['CONTRATO'] = df['CONTRATO'].apply(lambda x: str(int(x)) if pd.notna(x) else '')

        status_contratos = {}
        for contrato in lista_contratos:
            contrato_str = str(contrato).strip()
            df_contrato = df[df['CONTRATO'] == contrato_str]
            if df_contrato.empty:
                status_contratos[contrato_str] = 'NÃO LOCALIZADO'
            else:
                tem_ativo = any(pd.isna(val) or str(val).strip() == '' for val in df_contrato['FIM'])
                status_contratos[contrato_str] = 'ONLINE' if tem_ativo else 'OFFLINE'

        linhas_resumo = []
        for contrato in lista_contratos:
            c_str = str(contrato).strip()
            status = status_contratos.get(c_str, 'ERRO')
            df_c = df[df['CONTRATO'] == c_str]
            if not df_c.empty:
                ativo = df_c[pd.isna(df_c['FIM']) | (df_c['FIM'].astype(str).str.strip() == '')]
                info = ativo.iloc[0] if not ativo.empty else df_c.iloc[0]
                linhas_resumo.append({
                    'CONTRATO': c_str, 'STATUS': status,
                    'USERNAME': info.get('USERNAME', ''),
                    'INÍCIO': info.get('INÍCIO', '') if status == 'ONLINE' else '',
                    'FIM': info.get('FIM', '') if status != 'ONLINE' else '',
                    'CIRCUITO': info.get('CIRCUITO', ''),
                    'MAC': info.get('MAC', ''),
                    'TRÁFEGO': info.get('TRÁFEGO', '') if status == 'ONLINE' else '',
                    'SERVIDOR': info.get('SERVIDOR', '')
                })
            else:
                linhas_resumo.append({
                    'CONTRATO': c_str, 'STATUS': status, 'USERNAME': '', 'INÍCIO': '', 'FIM': '',
                    'CIRCUITO': '', 'MAC': '', 'TRÁFEGO': '', 'SERVIDOR': ''
                })

        return pd.DataFrame(linhas_resumo), None

    except Exception as e:
        logger.exception("Erro na consulta ao Autenticador.")
        return pd.DataFrame(), f"ERRO NA CONSULTA: {str(e)}"

def adicionar_status_autenticador(df_garantia):
    if df_garantia is None or df_garantia.empty:
        return df_garantia
    contratos_unicos = df_garantia['CÓDIGO CONTRATO'].unique().tolist()
    if not contratos_unicos:
        df_garantia['STATUS'] = 'ERRO'
        return df_garantia
    df_status, erro = consultar_autenticador_status(contratos_unicos)
    if erro or df_status.empty:
        df_garantia['STATUS'] = 'ERRO'
        logger.error(f"Falha ao adicionar status Autenticador: {erro}")
        return df_garantia
    df_status = df_status.rename(columns={'CONTRATO': 'CÓDIGO CONTRATO'})
    df_status = df_status[['CÓDIGO CONTRATO', 'STATUS']]
    df_garantia = df_garantia.merge(df_status, on='CÓDIGO CONTRATO', how='left')
    df_garantia['STATUS'] = df_garantia['STATUS'].fillna('DESCONHECIDO')
    return df_garantia

# =========================================
# FUNÇÃO PARA OBTER INFORMAÇÕES DE STATUS DOS ARQUIVOS
# =========================================

def obter_info_rodape():
    arquivos = [
        "chamados_abertos_field_service.xlsx",
        "base OFS ok.xlsx",
        "OFS GERAL.csv"
    ]
    linhas = []
    for arquivo in arquivos:
        caminho = os.path.join(PASTA, arquivo)
        if not os.path.exists(caminho):
            linhas.append(f"❌ {arquivo} - Arquivo não encontrado")
            continue
        if arquivo == "base OFS ok.xlsx":
            try:
                df_ofs = pd.read_excel(caminho)
                col_data = encontrar_coluna(df_ofs, ['Data'])
                if col_data:
                    df_ofs[col_data] = normalizar_data(df_ofs[col_data])
                    ultima_data = df_ofs[col_data].max()
                    if pd.notna(ultima_data):
                        data_esperada = datetime.now().date() - timedelta(days=1)
                        if ultima_data.date() == data_esperada:
                            linhas.append(f"✅ base OFS ok.xlsx - Base OFS atualizada")
                        else:
                            linhas.append(f"⚠️ base OFS ok.xlsx - Desatualizada (última data: {ultima_data.strftime('%d/%m/%Y')})")
                    else:
                        linhas.append("⚠️ base OFS ok.xlsx - Datas não encontradas")
                else:
                    linhas.append("⚠️ base OFS ok.xlsx - Coluna 'Data' não encontrada")
            except:
                linhas.append("⚠️ base OFS ok.xlsx - Erro ao ler arquivo")
        else:
            mod_time = datetime.fromtimestamp(os.path.getmtime(caminho))
            data_mod = mod_time.strftime('%d/%m/%Y %H:%M')
            linhas.append(f"📅 {arquivo} - Atualizado em: {data_mod}")
    return "\n".join(linhas)

# =========================================
# INTERFACE GRÁFICA (CUSTOMTKINTER)
# =========================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class OperacionalApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        logger.info("Iniciando OperacionalApp...")

        self.title("OPERACIONALDATABASE2026 V1")
        self.geometry("1400x850")
        self.configure(fg_color="#081B3A")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Cache
        self.cache = DataCache()
        self.bind("<Button-1>", lambda e: self.cache.update_activity())
        self.bind("<Motion>", lambda e: self.cache.update_activity())
        self.iniciar_pre_carregamento()

        # Menu lateral
        self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_autenticador=0, fg_color="#06152E")
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(7, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="OPERACIONAL\nDATABASE", text_color="#00BFFF", font=ctk.CTkFont(size=30, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 30))

        self.btn_garantias = ctk.CTkButton(self.sidebar_frame, text="Lista de Garantias", command=self.acao_garantias)
        self.btn_garantias.grid(row=1, column=0, padx=20, pady=10)

        self.btn_reparos = ctk.CTkButton(self.sidebar_frame, text="Reparos", command=self.acao_reparos)
        self.btn_reparos.grid(row=2, column=0, padx=20, pady=10)

        self.btn_upgrade = ctk.CTkButton(self.sidebar_frame, text="Upgrade e M.C", command=self.acao_upgrade)
        self.btn_upgrade.grid(row=3, column=0, padx=20, pady=10)

        self.btn_atv = ctk.CTkButton(self.sidebar_frame, text="ATV e MDE", command=self.acao_atv)
        self.btn_atv.grid(row=4, column=0, padx=20, pady=10)

        self.btn_autenticador = ctk.CTkButton(self.sidebar_frame, text="Consultar Autenticador", command=self.acao_autenticador)
        self.btn_autenticador.grid(row=5, column=0, padx=20, pady=10)

        self.btn_conf_agenda = ctk.CTkButton(self.sidebar_frame, text="Conf. Agenda", command=self.acao_confirmacao_agenda)
        self.btn_conf_agenda.grid(row=6, column=0, padx=20, pady=10)

        self.assinatura = ctk.CTkLabel(self.sidebar_frame, text="Desenvolvido por\nGuilherme Santos", font=ctk.CTkFont(size=12))
        self.assinatura.grid(row=8, column=0, padx=20, pady=20)

        # Área principal
        self.main_frame = ctk.CTkFrame(self, corner_autenticador=15, fg_color="#102B57")
        self.main_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=3)
        self.main_frame.grid_columnconfigure(1, weight=1)

        self.text_mode_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.text_mode_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        self.text_mode_frame.grid_rowconfigure(0, weight=1)
        self.text_mode_frame.grid_columnconfigure(0, weight=1)

        self.textbox_full = ctk.CTkTextbox(
            self.text_mode_frame, font=ctk.CTkFont(family="Consolas", size=12),
            fg_color="#0A1931", text_color="#FFFFFF", wrap="none"
        )
        self.textbox_full.grid(row=0, column=0, sticky="nsew")

        self.file_status_label = ctk.CTkLabel(
            self.text_mode_frame, text="", font=ctk.CTkFont(size=10),
            text_color="#AAAAAA", justify="left"
        )
        self.file_status_label.grid(row=1, column=0, sticky="sw", padx=10, pady=(0, 5))
        self.file_status_label.configure(text=obter_info_rodape())

        # Frames das visualizações
        self.scroll_frame = ctk.CTkScrollableFrame(self.main_frame, fg_color="#102B57")
        self.scroll_frame.grid_columnconfigure(0, weight=1)

        self.details_frame = ctk.CTkFrame(self.main_frame, fg_color="#06152E", corner_autenticador=10)
        self.details_frame.grid_rowconfigure(1, weight=1)
        self.details_frame.grid_columnconfigure(0, weight=1)

        self.details_header = ctk.CTkFrame(self.details_frame, fg_color="transparent")
        self.details_header.grid(row=0, column=0, sticky="ew", padx=10, pady=5)

        ctk.CTkLabel(self.details_header, text="CONTRATOS SOMADOS", text_color="#00BFFF", font=ctk.CTkFont(weight="bold")).pack(side="left")
        ctk.CTkButton(self.details_header, text="Limpar", width=60, height=24, fg_color="#C10037", hover_color="#900028", font=ctk.CTkFont(size=11), command=self.limpar_detalhes).pack(side="right")

        self.textbox_details = ctk.CTkTextbox(
            self.details_frame, font=ctk.CTkFont(family="Consolas", size=13),
            fg_color="#0A1931", text_color="#FFFFFF", wrap="none"
        )
        self.textbox_details.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.textbox_details.configure(state="disabled")

        self.modo_texto()

        texto_inicial = (
            "================ BEM-VINDO =================\n\n"
            "Bem-vindo ao Operacional Database.\n\n"
            "- Salve os chamados do CAMPO sobre o arquivo 'chamados_abertos_field_service.xlsx'.\n\n"
            "- Mantenha a base OFS (base OFS ok.xlsx) atualizada com os concluídos do dia anterior.\n\n"
            "- Exporte diariamente o OFS D0 + D1 e salve como 'OFS GERAL.csv'.\n\n"
            "✅ O Forms de Confirmação e a Planilha de Conveniência agora são carregados automaticamente pela API do Google,\n"
            "não sendo mais necessário atualizar esses arquivos manualmente.\n\n"
            "- Para consultar o Autenticador, a VPN deve estar conectada.\n\n"
            "Novidades:\n"
            " • Efetividade Geral.\n"
            " • Efetividade por Operador.\n"
            " • Ranking automático.\n"
            " • Integração automática com Google Sheets.\n"
            " • Cache inteligente: dados pré‑carregados, atualização após 5 min de inatividade.\n"
            " • Log de erros salvo em 'operacional.log'.\n"
        )
        self.escrever_na_tela_full(texto_inicial)

    # =========================================
    # PRÉ‑CARREGAMENTO DE DADOS (CACHE)
    # =========================================
    def iniciar_pre_carregamento(self):
        def carregar():
            self._preload_data()
        threading.Thread(target=carregar, daemon=True).start()

    def _preload_data(self):
        logger.info("Pré‑carregamento de dados iniciado.")
        try:
            ch = carregar_chamados()
            if not ch.empty:
                self.cache.set('chamados', ch)
            conv = carregar_conveniencias()
            if conv:
                self.cache.set('conveniencias', conv)
            forms = carregar_forms()
            if not forms.empty:
                self.cache.set('forms', forms)
            ofs = carregar_csv_robusto("OFS GERAL.csv")
            if not ofs.empty:
                self.cache.set('ofs_geral', ofs)
            base_ofs = carregar_excel("base OFS ok.xlsx")
            if not base_ofs.empty:
                self.cache.set('base_ofs', base_ofs)
            logger.info("Pré‑carregamento concluído.")
        except Exception:
            logger.exception("Erro durante o pré‑carregamento dos dados.")

    # =========================================
    # MÉTODOS DE ACESSO AOS DADOS COM CACHE
    # =========================================

    def get_chamados(self):
        if self.cache.is_valid('chamados'):
            return self.cache.get('chamados')
        ch = carregar_chamados()
        if not ch.empty:
            self.cache.set('chamados', ch)
        return ch

    def get_conveniencias(self):
        if self.cache.is_valid('conveniencias'):
            return self.cache.get('conveniencias')
        conv = carregar_conveniencias()
        if conv:
            self.cache.set('conveniencias', conv)
        return conv

    def get_forms_data(self):
        if self.cache.is_valid('forms'):
            return self.cache.get('forms')
        df = carregar_forms()
        if not df.empty:
            self.cache.set('forms', df)
        return df

    def get_ofs_geral(self):
        if self.cache.is_valid('ofs_geral'):
            return self.cache.get('ofs_geral')
        df = carregar_csv_robusto("OFS GERAL.csv")
        if not df.empty:
            self.cache.set('ofs_geral', df)
        return df

    def get_base_ofs(self):
        if self.cache.is_valid('base_ofs'):
            return self.cache.get('base_ofs')
        df = carregar_excel("base OFS ok.xlsx")
        if not df.empty:
            self.cache.set('base_ofs', df)
        return df

    # =========================================
    # MÉTODO PARA REMOVER STATUS DOS ARQUIVOS
    # =========================================
    def remover_status_label(self):
        if self.file_status_label.winfo_exists():
            self.file_status_label.destroy()

    # =========================================
    # CONTROLE DE EXIBIÇÃO (MODOS DE TELA)
    # =========================================

    def modo_texto(self):
        self.scroll_frame.grid_forget()
        self.details_frame.grid_forget()
        self.text_mode_frame.grid(row=0, column=0, columnspan=2, sticky="nsew")
        if self.file_status_label.winfo_exists():
            self.file_status_label.configure(text=obter_info_rodape())

    def modo_lista_detalhada(self):
        self.remover_status_label()
        self.text_mode_frame.grid_forget()
        self.scroll_frame.grid(row=0, column=0, columnspan=1, padx=(10, 5), pady=10, sticky="nsew")
        self.details_frame.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="nsew")

    def modo_lista_cheia(self):
        self.remover_status_label()
        self.text_mode_frame.grid_forget()
        self.details_frame.grid_forget()
        self.scroll_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

    def limpar_tela_cards(self):
        self.remover_status_label()
        self.modo_lista_detalhada()
        self.limpar_detalhes()
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

    def limpar_tela_lista_cheia(self):
        self.remover_status_label()
        self.modo_lista_cheia()
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

    def escrever_na_tela_full(self, texto):
        self.modo_texto()
        self.textbox_full.configure(state="normal")
        self.textbox_full.delete("0.0", "end")
        self.textbox_full.insert("0.0", texto)
        self.textbox_full.configure(state="disabled")

    def exibir_mensagem_resultado(self, mensagem, cor="#FFFFFF"):
        for widget in self.result_container.winfo_children():
            widget.destroy()
        lbl = ctk.CTkLabel(
            self.result_container, text=mensagem, font=ctk.CTkFont(size=15),
            text_color=cor, justify="left", wraplength=700
        )
        lbl.pack(pady=30, padx=20, anchor="w")

    # =========================================
    # GRÁFICO DE BARRAS NATIVO (MINI BARRAS)
    # =========================================

    def criar_grafico_barra(self, parent, titulo, valor, total, cor):
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.pack(fill="x", pady=6, padx=20)
        lbl_titulo = ctk.CTkLabel(container, text=titulo, width=150, anchor="w", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_titulo.pack(side="left", padx=(0, 10))
        pct = valor / total if total > 0 else 0
        bar_bg = ctk.CTkFrame(container, fg_color="#1A3461", height=18, corner_autenticador=9)
        bar_bg.pack(side="left", fill="x", expand=True, padx=10)
        if pct > 0:
            bar_fg = ctk.CTkFrame(bar_bg, fg_color=cor, height=18, corner_autenticador=9)
            bar_fg.place(relx=0, rely=0, relwidth=pct, relheight=1)
        pct_str = f"{pct*100:.1f}%"
        lbl_valor = ctk.CTkLabel(container, text=f"{valor}/{total} ({pct_str})", width=110, anchor="e", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_valor.pack(side="right", padx=(10, 0))

    def criar_linha_estilo_autenticador(self, parent, titulo, qtd, pct, cor_barra, texto_extra=""):
        row = ctk.CTkFrame(parent, fg_color="#16376B", height=45, corner_autenticador=0)
        row.pack(fill="x", pady=2)
        ctk.CTkLabel(row, text=titulo, width=200, anchor="w", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=20)
        ctk.CTkLabel(row, text=str(qtd), width=60, anchor="center").pack(side="left")
        bar_bg = ctk.CTkFrame(row, fg_color="#06152E", height=14, corner_autenticador=7)
        bar_bg.pack(side="left", fill="x", expand=True, padx=20)
        if pct > 0:
            ctk.CTkFrame(bar_bg, fg_color=cor_barra, height=14, corner_autenticador=7).place(relwidth=min(pct, 1.0), relheight=1)
        ctk.CTkLabel(row, text=f"{pct*100:.1f}%", width=60, anchor="e").pack(side="right", padx=20)

    # =========================================
    # FUNÇÕES DO PAINEL LATERAL DE DETALHES
    # =========================================

    def limpar_detalhes(self):
        self.textbox_details.configure(state="normal")
        self.textbox_details.delete("0.0", "end")
        self.textbox_details.configure(state="disabled")

    def adicionar_detalhamento(self, sigla, secoes):
        self.textbox_details.configure(state="normal")
        conteudo = f"\n=== {sigla.upper()} ===\n"
        for titulo_secao, lista_contratos in secoes:
            conteudo += f">> {titulo_secao}:\n"
            if not lista_contratos:
                conteudo += "   -\n"
            else:
                for c in lista_contratos:
                    conteudo += f"   {c}\n"
        conteudo += "-" * 35 + "\n"
        if self.textbox_details.get("0.0", "end").strip() == "":
            conteudo = conteudo.lstrip()
        self.textbox_details.insert("end", conteudo)
        self.textbox_details.see("end")
        self.textbox_details.configure(state="disabled")

    # =========================================
    # CONSTRUTORES VISUAIS DA TABELA DE GARANTIAS
    # =========================================

    def gerar_tabela_garantias(self, df_garantia):
        self.limpar_tela_lista_cheia()
        ctk.CTkLabel(
            self.scroll_frame, text="LISTA DE GARANTIAS", text_color="#00BFFF", font=ctk.CTkFont(size=24, weight="bold")
        ).pack(pady=(5, 15))
        if df_garantia is None or df_garantia.empty:
            ctk.CTkLabel(self.scroll_frame, text="Sem garantias pendentes no momento.", text_color="#FFFFFF", font=ctk.CTkFont(size=15)).pack(pady=20)
            return

        def desenhar_tabela_regiao(regiao_nome, lista_cidades):
            df_regiao = df_garantia[df_garantia['UNIDADE'].isin(lista_cidades)]
            if df_regiao.empty:
                ctk.CTkLabel(
                    self.scroll_frame,
                    text=f"{regiao_nome} – Sem garantias no momento.",
                    text_color="#FFFFFF",
                    font=ctk.CTkFont(size=15)
                ).pack(anchor="w", padx=10, pady=10)
                return

            ctk.CTkLabel(
                self.scroll_frame, text=regiao_nome,
                font=ctk.CTkFont(size=20, weight="bold"), text_color="#007BFF"
            ).pack(anchor="w", padx=10, pady=(20, 10))

            for sigla in lista_cidades:
                df_cidade = df_garantia[df_garantia['UNIDADE'] == sigla]
                if df_cidade.empty: continue

                nome_cidade = DICIONARIO_CIDADES.get(sigla, sigla).upper()

                ctk.CTkLabel(
                    self.scroll_frame, text=f"📍 {nome_cidade}",
                    font=ctk.CTkFont(size=16, weight="bold"), text_color="#00BFFF"
                ).pack(anchor="w", padx=15, pady=(15, 5))

                header_frame = ctk.CTkFrame(self.scroll_frame, fg_color="#06152E", corner_autenticador=6, height=35)
                header_frame.pack(fill="x", padx=15, pady=(0, 5))
                header_frame.pack_propagate(False)

                col_widths = [100, 70, 60, 140, 250]
                headers = ["CONTRATO", "STATUS", "AGING", "SERVIÇO", "TÉCNICO (OFS)"]

                for i, text in enumerate(headers):
                    ctk.CTkLabel(header_frame, text=text, width=col_widths[i], anchor="w", text_color="#00FF88", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", padx=10)

                ctk.CTkLabel(header_frame, text="ENDEREÇO", anchor="w", text_color="#00FF88", font=ctk.CTkFont(size=13, weight="bold")).pack(side="left", fill="x", expand=True, padx=10)

                for idx, (_, row) in enumerate(df_cidade.iterrows()):
                    contrato = str(row['CÓDIGO CONTRATO']).strip()[:12]
                    aging = str(row['AGING']).split('.')[0][:6]

                    tipo_bruto = str(row['TIPO_ANTERIOR_RAW']).lower()
                    if 'reparo' in tipo_bruto:
                        tipo = 'IRR'
                    elif 'mudança' in tipo_bruto or 'mde' in tipo_bruto:
                        tipo = 'IFI de MDE'
                    elif 'ativação' in tipo_bruto or 'ativacao' in tipo_bruto:
                        tipo = 'IFI'
                    else:
                        tipo = str(row['TIPO_ANTERIOR_RAW'])[:16]

                    tecnico = str(row['TECNICO']).upper()[:35]
                    endereco = str(row['ENDERECO']).upper().replace("\n", " ")

                    status_val = str(row.get('STATUS', ''))
                    if status_val == 'ONLINE':
                        cor_status = "#00FF88"
                    elif status_val == 'OFFLINE':
                        cor_status = "#FF4D4D"
                    else:
                        cor_status = "#FFA500"

                    cor_fundo = "#16376B" if idx % 2 == 0 else "#102B57"

                    row_frame = ctk.CTkFrame(self.scroll_frame, fg_color=cor_fundo, corner_autenticador=4, height=35)
                    row_frame.pack(fill="x", padx=15, pady=2)
                    row_frame.pack_propagate(False)

                    valores = [
                        (contrato, "#FFFFFF", "normal"),
                        (status_val, cor_status, "bold"),
                        (aging, "#FFFFFF", "normal"),
                        (tipo, "#FFFFFF", "normal"),
                        (tecnico, "#FFFFFF", "normal")
                    ]

                    for i, (val, cor, peso) in enumerate(valores):
                        celula = ctk.CTkEntry(
                            row_frame, width=col_widths[i], fg_color="transparent",
                            border_width=0, text_color=cor, font=ctk.CTkFont(size=13, weight=peso)
                        )
                        celula.pack(side="left", padx=10)
                        celula.insert(0, val)
                        celula.icursor(0)
                        celula.configure(state="readonly")

                    celula_end = ctk.CTkEntry(
                        row_frame, fg_color="transparent",
                        border_width=0, text_color="#FFFFFF", font=ctk.CTkFont(size=13)
                    )
                    celula_end.pack(side="left", fill="x", expand=True, padx=10)
                    celula_end.insert(0, endereco)
                    celula_end.icursor(0)
                    celula_end.configure(state="readonly")

        desenhar_tabela_regiao("LITORAL NORTE SP", LITORAL_SP)
        desenhar_tabela_regiao("SUL RJ", RJ)

        invalidos = df_garantia[~df_garantia['UNIDADE'].isin(TODAS_CIDADES)]
        if not invalidos.empty:
            ctk.CTkLabel(
                self.scroll_frame,
                text="⚠️ SIGLA INVÁLIDA / VERIFICAR",
                font=ctk.CTkFont(size=20, weight="bold"),
                text_color="#FFAA00"
            ).pack(anchor="w", padx=10, pady=(30,10))

            for _, row in invalidos.iterrows():
                texto = f"{row['CÓDIGO CONTRATO']} | {row['UNIDADE']} | {row['TECNICO']}"
                ctk.CTkLabel(
                    self.scroll_frame,
                    text=texto,
                    anchor="w"
                ).pack(fill="x", padx=20)

    # =========================================
    # CONSTRUTORES VISUAIS DA LISTA COMPACTA
    # =========================================

    def mostrar_titulo_lista(self, texto):
        ctk.CTkLabel(
            self.scroll_frame, text=texto, text_color="#00BFFF", font=ctk.CTkFont(size=24, weight="bold")
        ).pack(pady=(5, 15))

    def criar_linha_lista(self, frame_pai, sigla, dados_texto, command_detalhar):
        linha = ctk.CTkFrame(frame_pai, fg_color="#16376B", corner_autenticador=6, height=40)
        linha.pack(fill="x", padx=5, pady=3)
        linha.pack_propagate(False)

        ctk.CTkLabel(
            linha, text=sigla, width=60, anchor="w",
            text_color="#00BFFF", font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left", padx=10)

        ctk.CTkFrame(linha, width=2, fg_color="#30568C").pack(side="left", fill="y", pady=5)

        container_metricas = ctk.CTkFrame(linha, fg_color="transparent")
        container_metricas.pack(side="left", fill="x", expand=True, padx=10)

        for rotulo, valor in dados_texto:
            if rotulo == "ONLINE":
                cor = "#00FF88"
            elif rotulo == "OFFLINE":
                cor = "#FF4D4D"
            else:
                cor = "#FFFFFF"
            lbl = ctk.CTkLabel(container_metricas, text=f"{rotulo}: {valor}", text_color=cor, font=ctk.CTkFont(size=13))
            lbl.pack(side="left", padx=10)

        ctk.CTkButton(
            linha, text="Detalhar", width=80, height=24, fg_color="#007BFF", text_color="white", hover_color="#0056B3",
            font=ctk.CTkFont(size=12, weight="bold"), command=command_detalhar
        ).pack(side="right", padx=10)

    def gerar_visao_padrao(self, df_filtrado, titulo_principal, usar_status_autenticador=False):
        self.limpar_tela_cards()
        self.mostrar_titulo_lista(titulo_principal)
        hoje = pd.Timestamp.now().normalize()

        def renderizar_regiao(cidades_lista, nome_regiao):
            ctk.CTkLabel(
                self.scroll_frame, text=nome_regiao,
                font=ctk.CTkFont(size=18, weight="bold"), text_color="#007BFF"
            ).pack(anchor="w", padx=5, pady=(15, 5))

            total_online = 0
            total_offline = 0
            total_d0 = 0
            total_fut = 0
            total_geral = 0

            for sigla in cidades_lista:
                temp = df_filtrado[df_filtrado['UNIDADE'] == sigla]
                if temp.empty:
                    continue

                if usar_status_autenticador and 'STATUS' in temp.columns:
                    online = len(temp[temp['STATUS'] == 'ONLINE'])
                    offline = len(temp[temp['STATUS'] == 'OFFLINE'])
                    d0_df = temp[temp['DATA AGENDAMENTO'] == hoje]
                    futuro_df = temp[temp['DATA AGENDAMENTO'] > hoje]
                    qtd_d0 = len(d0_df)
                    qtd_futuro = len(futuro_df)

                    qtd_total = len(temp)

                    dados = [
                        ("ONLINE", online),
                        ("OFFLINE", offline),
                        ("D0", qtd_d0),
                        ("Futuro", qtd_futuro),
                        ("Total", qtd_total)
                    ]

                    secoes_detalhe = [
                        ("ONLINE", temp[temp['STATUS'] == 'ONLINE']['CÓDIGO CONTRATO'].tolist()),
                        ("OFFLINE", temp[temp['STATUS'] == 'OFFLINE']['CÓDIGO CONTRATO'].tolist()),
                        ("D0 (Hoje)", d0_df['CÓDIGO CONTRATO'].tolist()),
                        ("FUTURO", futuro_df['CÓDIGO CONTRATO'].tolist())
                    ]

                    total_online += online
                    total_offline += offline
                    total_d0 += qtd_d0
                    total_fut += qtd_futuro
                    total_geral += qtd_total
                else:
                    d0_df = temp[temp['DATA AGENDAMENTO'] == hoje]
                    futuro_df = temp[temp['DATA AGENDAMENTO'] > hoje]

                    qtd_d0 = len(d0_df)
                    qtd_futuro = len(futuro_df)

                    qtd_total = len(temp)

                    dados = [("D0", qtd_d0), ("Futuro", qtd_futuro), ("Total", qtd_total)]
                    secoes_detalhe = [
                        ("D0 (Hoje)", d0_df['CÓDIGO CONTRATO'].tolist()),
                        ("FUTURO", futuro_df['CÓDIGO CONTRATO'].tolist())
                    ]
                    total_d0 += qtd_d0
                    total_fut += qtd_futuro
                    total_geral += qtd_total

                self.criar_linha_lista(
                    self.scroll_frame, sigla, dados,
                    lambda c=sigla, s=secoes_detalhe: self.adicionar_detalhamento(c, s)
                )

            if usar_status_autenticador:
                geral = total_geral
                ctk.CTkLabel(
                    self.scroll_frame,
                    text=f"TOTAL DA REGIÃO  |  ONLINE: {total_online}  |  OFFLINE: {total_offline}  |  D0: {total_d0}  |  Futuro: {total_fut}  |  Geral: {geral}",
                    text_color="#FFFFFF", font=ctk.CTkFont(size=14, weight="bold")
                ).pack(anchor="e", padx=30, pady=(5, 20))
            else:
                geral = total_geral
                ctk.CTkLabel(
                    self.scroll_frame,
                    text=f"TOTAL DA REGIÃO  |  D0: {total_d0}  |  Futuro: {total_fut}  |  Geral: {geral}",
                    text_color="#FFFFFF", font=ctk.CTkFont(size=14, weight="bold")
                ).pack(anchor="e", padx=30, pady=(5, 20))

        renderizar_regiao(LITORAL_SP, "LITORAL NORTE SP")
        renderizar_regiao(RJ, "SUL RJ")

    # =========================================
    # AÇÕES DO MENU (com verificação de inatividade)
    # =========================================

    def buscar_chamados(self):
        if self.cache.is_idle(300):
            self.cache.clear()
        ch = self.get_chamados()
        if ch.empty:
            logger.error("Arquivo de chamados não encontrado ou vazio.")
            self.remover_status_label()
            self.escrever_na_tela_full("Erro: Arquivo de chamados não encontrado ou vazio na pasta raiz.")
            return ch
        return ch

    def obter_dados_garantias(self, ch):
        hist = self.get_base_ofs()
        if hist.empty:
            logger.error("Base OFS não encontrada.")
            return None, "Arquivo 'base OFS ok.xlsx' não encontrado."

        hist.columns = hist.columns.str.strip()
        contrato_hist = encontrar_coluna(hist, ['Número do contrato', 'Contrato'])
        data_hist = encontrar_coluna(hist, ['Data'])
        status_hist = encontrar_coluna(hist, ['Status da Atividade'])
        tipo_hist = encontrar_coluna(hist, ['Tipo de Atividade.1', 'Tipo de Atividade 2', 'Tipo de Atividade'])

        tecnico_hist = encontrar_coluna(hist, ['Recurso', 'Técnico', 'Nome do Técnico', 'Nome', 'Resource', 'Técnico Executante'])

        if not contrato_hist or not tipo_hist:
            logger.error("Colunas de contrato ou atividade não encontradas no OFS.")
            return None, "Colunas de contrato ou atividade não encontradas no OFS."

        if not tecnico_hist:
            hist['TÉCNICO'] = "Não Identificado"
            tecnico_hist = 'TÉCNICO'

        hist[contrato_hist] = tratar_contrato(hist[contrato_hist])
        hist[data_hist] = normalizar_data(hist[data_hist])
        hist = hist[hist[status_hist].str.contains('conclu', case=False, na=False)]

        reparos = ch[ch['FILA'].str.contains('REPARO', case=False, na=False)]
        reparos = reparos[~reparos['FILA'].str.contains('PREVENTIVO', case=False, na=False)]

        df = pd.merge(
            reparos,
            hist,
            left_on='CÓDIGO CONTRATO',
            right_on=contrato_hist,
            how='inner'
        )

        col_ingresso = encontrar_coluna(df, ['DATA DE INGRESSO'])
        if col_ingresso:
            df['AGING'] = (df[col_ingresso].dt.normalize() - df[data_hist].dt.normalize()).dt.days
        else:
            df['AGING'] = (pd.to_datetime(df.iloc[:, 16], errors='coerce', dayfirst=True).dt.normalize() - df[data_hist].dt.normalize()).dt.days

        # Exclui apenas atividades finalizadas em data POSTERIOR à abertura do reparo
        # (o que indicaria inconsistência de datas). Reparos abertos no MESMO DIA em que
        # o serviço anterior foi concluído (AGING == 0) agora são considerados garantia
        # normalmente — antes eram descartados por engano, mesmo sendo um caso válido
        # (ex.: cliente reclama horas depois da OS de Ativação/Mudança/Reparo ter sido concluída).
        df = df[df['AGING'] >= 0]

        garantia = df[
            ((df[tipo_hist].str.contains('REPARO', case=False, na=False)) & (df['AGING'] <= 30)) |
            ((df[tipo_hist].str.contains('ATIVAÇÃO|MUDANÇA', case=False, na=False)) & (df['AGING'] <= 15))
        ]

        endereco_col = encontrar_coluna(garantia, ['Endereço', 'Endereco', 'Logradouro', 'Rua', 'Endereço do Cliente', 'ENDEREÇO'])
        if not endereco_col:
            garantia['ENDEREÇO'] = "Endereço Não Encontrado"
            endereco_col = 'ENDEREÇO'

        garantia['TIPO_ANTERIOR_RAW'] = garantia[tipo_hist]
        garantia['TECNICO'] = garantia[tecnico_hist]
        garantia['ENDERECO'] = garantia[endereco_col]
        garantia['SIGLA_INVALIDA'] = ~garantia['UNIDADE'].isin(TODAS_CIDADES)

        return garantia, ""

    def acao_garantias(self):
        if self.cache.is_idle(300):
            self.cache.clear()
        ch = self.buscar_chamados()
        if ch.empty: return

        df_garantia, erro = self.obter_dados_garantias(ch)
        if erro:
            self.escrever_na_tela_full(erro)
            return

        if df_garantia is not None and not df_garantia.empty:
            self.escrever_na_tela_full("Consultando status online no Autenticador...")
            self.update()
            df_garantia = adicionar_status_autenticador(df_garantia)

        self.gerar_tabela_garantias(df_garantia)

    def acao_reparos(self):
        if self.cache.is_idle(300):
            self.cache.clear()
        ch = self.buscar_chamados()
        if ch.empty: return
        df_reparo = ch[ch['FILA'].str.contains('REPARO', case=False, na=False)]
        if not df_reparo.empty:
            self.escrever_na_tela_full("Consultando status online dos reparos...")
            self.update()
            df_reparo = adicionar_status_autenticador(df_reparo)
        self.gerar_visao_padrao(df_reparo, "REPAROS PENDENTES", usar_status_autenticador=True)

    def acao_upgrade(self):
        if self.cache.is_idle(300):
            self.cache.clear()
        ch = self.buscar_chamados()
        if ch.empty: return
        df_upgrade = ch[ch['FILA'].str.contains('UPGRADE|MUDANÇA DE CÔMODO', case=False, na=False)]
        self.gerar_visao_padrao(df_upgrade, "UPGRADE E M.C", usar_status_autenticador=False)

    def acao_atv(self):
        if self.cache.is_idle(300):
            self.cache.clear()
        ch = self.buscar_chamados()
        if ch.empty: return

        self.limpar_tela_cards()
        self.mostrar_titulo_lista("ATIVAÇÃO E M.D.E")

        hoje = pd.Timestamp.now().normalize()
        conveniencias = self.get_conveniencias()

        df_ativacao = ch[ch['FILA'].str.contains('ATIVAÇÃO', case=False, na=False)].copy()
        df_mudanca = ch[
            ch['FILA'].str.contains('MUDANÇA', case=False, na=False) &
            ~ch['FILA'].str.contains('CÔMODO|COMODO', case=False, na=False)
        ].copy()

        tipos = [("ATIVAÇÕES", df_ativacao), ("MUDANÇAS DE ENDEREÇO", df_mudanca)]

        for nome_visao, df_filtro in tipos:
            bloco_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
            bloco_frame.pack(fill="x", pady=(15, 5))

            ctk.CTkLabel(
                bloco_frame, text=nome_visao, text_color="#007BFF",
                font=ctk.CTkFont(size=20, weight="bold")
            ).pack(anchor="center")

            def renderizar_regiao_atv(cidades_lista, nome_regiao):
                ctk.CTkLabel(
                    self.scroll_frame, text=nome_regiao, text_color="#00BFFF",
                    font=ctk.CTkFont(size=16, weight="bold")
                ).pack(anchor="w", padx=5, pady=(5, 2))

                t_d0, t_fut, t_conv, t_inj = 0, 0, 0, 0

                for sigla in cidades_lista:
                    temp = df_filtro[df_filtro['UNIDADE'] == sigla]
                    if temp.empty:
                        continue

                    d0_df = temp[temp['DATA AGENDAMENTO'] == hoje]
                    futuro_df = temp[temp['DATA AGENDAMENTO'] > hoje]
                    conv_df = futuro_df[futuro_df['CÓDIGO CONTRATO'].isin(conveniencias)]
                    inj_df = futuro_df[~futuro_df['CÓDIGO CONTRATO'].isin(conveniencias)]

                    q_d0 = len(d0_df)
                    q_fut = len(futuro_df)
                    q_conv = len(conv_df)
                    q_inj = len(inj_df)

                    t_d0 += q_d0
                    t_fut += q_fut
                    t_conv += q_conv
                    t_inj += q_inj

                    dados = [
                        ("D0", q_d0),
                        ("Futuro", q_fut),
                        ("Conv.", q_conv),
                        ("Inj.", q_inj),
                        ("Total", q_d0 + q_fut)
                    ]

                    secoes_detalhe = [
                        ("D0 (Hoje)", d0_df['CÓDIGO CONTRATO'].tolist()),
                        ("Conveniência", conv_df['CÓDIGO CONTRATO'].tolist()),
                        ("Injeção", inj_df['CÓDIGO CONTRATO'].tolist())
                    ]

                    self.criar_linha_lista(
                        self.scroll_frame, sigla, dados,
                        lambda c=sigla, s=secoes_detalhe: self.adicionar_detalhamento(c, s)
                    )

                geral = t_d0 + t_fut

                ctk.CTkLabel(
                    self.scroll_frame,
                    text=f"TOTAL DA REGIÃO  |  D0: {t_d0}  |  Futuro: {t_fut}  |  Conv: {t_conv}  |  Inj: {t_inj}  |  Geral: {geral}",
                    text_color="#FFFFFF", font=ctk.CTkFont(size=14, weight="bold")
                ).pack(anchor="e", padx=30, pady=(5, 20))

            renderizar_regiao_atv(LITORAL_SP, "LITORAL NORTE SP")
            renderizar_regiao_atv(RJ, "SUL RJ")

    def acao_autenticador(self):
        if self.cache.is_idle(300):
            self.cache.clear()
        self.limpar_tela_lista_cheia()

        ctk.CTkLabel(
            self.scroll_frame,
            text="CONSULTAR SESSÕES NO AUTENTICADOR",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="#00BFFF"
        ).pack(pady=10)

        ctk.CTkLabel(
            self.scroll_frame,
            text="Cole os contratos (um por linha):",
            font=ctk.CTkFont(size=13)
        ).pack(anchor="w", padx=20)

        self.entry_contratos = ctk.CTkTextbox(self.scroll_frame, height=150)
        self.entry_contratos.pack(fill="x", padx=20, pady=(5, 10))

        ctk.CTkButton(
            self.scroll_frame,
            text="Consultar",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#007BFF",
            command=self.executar_consulta_autenticador
        ).pack(pady=10)

        # Usando CTkFrame para evitar o colapso de componentes associados a frames roláveis aninhados.
        self.result_frame = ctk.CTkFrame(self.scroll_frame, fg_color="#102B57")
        self.result_frame.pack(fill="both", expand=True, padx=10, pady=10)

    def executar_consulta_autenticador(self):
        contratos = self.entry_contratos.get("0.0", "end").splitlines()
        contratos = [c.strip() for c in contratos if c.strip()]

        for widget in self.result_frame.winfo_children():
            widget.destroy()

        if not contratos:
            ctk.CTkLabel(self.result_frame, text="Nenhum contrato informado.", text_color="#FF4D4D").pack(pady=20)
            return

        # Indicador visual de carregamento
        loading_lbl = ctk.CTkLabel(
            self.result_frame, 
            text="Consultando Autenticador... Por favor, aguarde.", 
            text_color="#00BFFF", 
            font=ctk.CTkFont(size=14, weight="bold")
        )
        loading_lbl.pack(pady=20)

        def worker():
            try:
                df, erro = consultar_autenticador_status(contratos)
                self.after(0, lambda: self._atualizar_resultados_autenticador(df, erro, loading_lbl))
            except Exception as e:
                logger.exception("Erro na thread do Autenticador")
                self.after(0, lambda: self._atualizar_resultados_autenticador(pd.DataFrame(), f"Erro inesperado: {e}", loading_lbl))

        threading.Thread(target=worker, daemon=True).start()

    def _atualizar_resultados_autenticador(self, df, erro, loading_lbl):
        try:
            if loading_lbl.winfo_exists():
                loading_lbl.destroy()
        except Exception:
            pass

        if erro:
            ctk.CTkLabel(self.result_frame, text=erro, text_color="#FF4D4D", font=ctk.CTkFont(size=14)).pack(pady=20)
            return

        if df.empty:
            ctk.CTkLabel(self.result_frame, text="Sem dados retornados.", text_color="#FFFFFF").pack(pady=20)
            return

        header_frame = ctk.CTkFrame(self.result_frame, fg_color="#06152E", corner_autenticador=6, height=35)
        header_frame.pack(fill="x", padx=5, pady=(0, 5))
        header_frame.pack_propagate(False)

        colunas_exibir = ['CONTRATO', 'STATUS', 'USERNAME', 'INÍCIO', 'FIM', 'CIRCUITO', 'MAC', 'TRÁFEGO', 'SERVIDOR']
        col_widths = [80, 70, 150, 140, 140, 160, 120, 80, 180]

        for i, col in enumerate(colunas_exibir):
            ctk.CTkLabel(
                header_frame, text=col, width=col_widths[i], anchor="w",
                text_color="#00FF88", font=ctk.CTkFont(size=12, weight="bold")
            ).pack(side="left", padx=5)

        for idx, (_, row) in enumerate(df.iterrows()):
            cor_fundo = "#16376B" if idx % 2 == 0 else "#102B57"
            row_frame = ctk.CTkFrame(self.result_frame, fg_color=cor_fundo, corner_autenticador=4, height=30)
            row_frame.pack(fill="x", padx=5, pady=1)
            row_frame.pack_propagate(False)

            for i, col in enumerate(colunas_exibir):
                valor = str(row.get(col, ''))
                if col == 'STATUS':
                    cor = "#00FF88" if valor == 'ONLINE' else "#FF4D4D" if valor == 'OFFLINE' else "#FFA500"
                    peso = "bold"
                else:
                    cor = "#FFFFFF"
                    peso = "normal"

                celula = ctk.CTkEntry(
                    row_frame, width=col_widths[i], fg_color="transparent",
                    border_width=0, text_color=cor, font=ctk.CTkFont(size=12, weight=peso)
                )
                celula.pack(side="left", padx=5)
                celula.insert(0, valor)
                celula.configure(state="readonly")

    # =========================================
    # FUNCIONALIDADE: CONFIRMAÇÃO DE AGENDA
    # =========================================

    def acao_confirmacao_agenda(self):
        if self.cache.is_idle(300):
            self.cache.clear()
        self.limpar_tela_lista_cheia()

        cabecalho = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        cabecalho.pack(fill="x", padx=24, pady=(18, 10))
        ctk.CTkLabel(cabecalho, text="CONFIRMAÇÃO DE AGENDA",
                     font=ctk.CTkFont(size=24, weight="bold"),
                     text_color="#00BFFF").pack(anchor="w")
        ctk.CTkLabel(cabecalho, text="Selecione a visão que deseja consultar.",
                     font=ctk.CTkFont(size=14), text_color="#AAAAAA").pack(anchor="w", pady=(3, 0))

        self.view_frame = ctk.CTkFrame(self.scroll_frame, fg_color="#0A1931", corner_autenticador=10)
        self.view_frame.pack(fill="x", padx=24, pady=18)

        botoes = [
            ("Visão Confirmado", "Acompanhamento de confirmações por região.", "confirmado"),
            ("Efetividade Geral", "Resultado de produtividade do D0.", "efetividade"),
            ("Efetividade por Operador", "Resultado de produtividade do D0 por operador.", "operador"),
        ]
        for texto, descricao, visao in botoes:
            linha = ctk.CTkFrame(self.view_frame, fg_color="transparent")
            linha.pack(fill="x", padx=16, pady=9)
            ctk.CTkButton(linha, text=texto, width=220, height=38,
                          command=lambda v=visao: self.preparar_tela_confirmacao_date(v)).pack(side="left")
            ctk.CTkLabel(linha, text=descricao, text_color="#AAAAAA",
                         font=ctk.CTkFont(size=13)).pack(side="left", padx=14)

    def preparar_tela_confirmacao_date(self, visao):
        self.view_mode = visao
        self.limpar_tela_lista_cheia()

        nomes = {"confirmado": "CONFIRMADO", "efetividade": "EFETIVIDADE GERAL", "operador": "EFETIVIDADE POR OPERADOR"}
        topo = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        topo.pack(fill="x", padx=24, pady=(18, 8))
        ctk.CTkLabel(topo, text=nomes[visao], font=ctk.CTkFont(size=22, weight="bold"),
                     text_color="#00BFFF").pack(side="left")
        ctk.CTkButton(topo, text="← Voltar", width=90, height=30, fg_color="#555555",
                      command=self.acao_confirmacao_agenda).pack(side="right")

        if visao in ("efetividade", "operador"):
            ctk.CTkLabel(self.scroll_frame, text="D0 • resultado de hoje",
                         text_color="#AAAAAA", font=ctk.CTkFont(size=13)).pack(anchor="w", padx=24, pady=(0, 10))
            self.result_container = ctk.CTkFrame(self.scroll_frame, fg_color="#0A1931", corner_autenticador=10)
            self.result_container.pack(fill="x", padx=24, pady=(0, 16))
            self.detail_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
            self.detail_frame.pack(fill="x", padx=24)
            self.processar_confirmacao_novo(0)
            return

        controles = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        controles.pack(fill="x", padx=24, pady=(0, 10))
        ctk.CTkButton(controles, text="Hoje (D0)", command=lambda: self.processar_confirmacao_novo(0)).pack(side="left", padx=(0, 8))
        ctk.CTkButton(controles, text="Amanhã (D+1)", command=lambda: self.processar_confirmacao_novo(1)).pack(side="left")

        self.result_container = ctk.CTkFrame(self.scroll_frame, fg_color="#0A1931", corner_autenticador=10)
        self.result_container.pack(fill="x", padx=24, pady=(0, 12))
        self.detail_frame = ctk.CTkFrame(self.scroll_frame, fg_color="transparent")
        self.detail_frame.pack(fill="x", padx=24, pady=(0, 10))
        self.exibir_mensagem_resultado("Escolha o período para consultar.")

    # =========================================
    # MÉTODOS ATUALIZADOS DE PROCESSAMENTO
    # =========================================

    def processar_confirmacao_novo(self, delta_days):
        if delta_days == 0:
            data_alvo = pd.Timestamp.now().normalize()
            texto_data = "Hoje (D0)"
        else:
            data_alvo = pd.Timestamp.now().normalize() + timedelta(days=1)
            texto_data = "Amanhã (D+1)"

        self.data_alvo = data_alvo

        ofs = self.get_ofs_geral()
        if ofs.empty:
            self.exibir_mensagem_resultado("Arquivo 'OFS GERAL.csv' não encontrado ou vazio.")
            return

        forms = self.get_forms_data()
        if forms.empty:
            self.exibir_mensagem_resultado("Arquivos de forms não encontrados ou vazios.")
            return

        ofs.columns = ofs.columns.str.strip()
        forms.columns = forms.columns.str.strip()

        col_os_ofs = encontrar_coluna(ofs, ['Ordem de Serviço', 'ID da Ordem de Serviço', 'PROTOCOLO', 'OS', 'ID'])
        col_data_ofs = encontrar_coluna(ofs, ['Data', 'Data Agendamento', 'DATA DE AGENDAMENTO', 'Data da Atividade'])
        col_area_ofs = encontrar_coluna(ofs, ['Cidade', 'AREA', 'UNIDADE', 'Área de Trabalho', 'Chave Workzone'])
        col_contrato_ofs = encontrar_coluna(ofs, ['Número do contrato', 'CÓDIGO CONTRATO', 'Contrato'])
        col_status_ofs = encontrar_coluna(ofs, ['Status da Atividade', 'STATUS'])
        col_motivo_ofs = encontrar_coluna(ofs, ['Motivo de Encerramento das atividades', 'Motivo'])
        col_tipo_ofs = encontrar_coluna(ofs, ['Tipo de Atividade.1', 'Tipo de Atividade 2', 'Tipo de Atividade', 'Tipo', 'TIPO'])

        self.col_contrato_ofs = col_contrato_ofs
        self.col_tipo_ofs = col_tipo_ofs
        self.col_status_ofs = col_status_ofs  # guarda para uso no detalhamento

        col_os_forms = encontrar_coluna(forms, ['ORDEM DE SERVIÇO', 'OS', 'Ordem de Serviço'])
        col_status_forms = encontrar_coluna(forms, ['STATUS', 'Status'])

        col_data_forms = encontrar_coluna(forms, [
            'DATA DE AGENDAMENTO',
            'Data de Agendamento',
            'Data Agendamento',
            'Data do Agendamento'
        ])

        if col_data_forms is not None:
            if 'AGENDAMENTO' not in col_data_forms.upper():
                col_data_forms = None
                logger.warning("Coluna de data encontrada não parece ser DATA DE AGENDAMENTO, ignorando.")

        if col_data_forms is None and len(forms.columns) >= 5:
            col_data_forms = forms.columns[4]
            logger.info(f"Usando a coluna E (índice 4) como DATA DE AGENDAMENTO: {col_data_forms}")

        if not col_os_ofs or not col_data_ofs or not col_os_forms:
            diag = (f"[DIAGNÓSTICO] Não foi possível mapear colunas essenciais.\n"
                    f"Colunas do OFS GERAL: {list(ofs.columns)}\n"
                    f"Colunas do FORMS: {list(forms.columns)}")
            self.exibir_mensagem_resultado(diag)
            return

        # Remove apenas canceladas, NÃO remove concluídas para a visão confirmado
        if col_status_ofs:
            ofs = ofs[~ofs[col_status_ofs].astype(str).str.contains('cancelad', case=False, na=False)]

        ofs[col_data_ofs] = normalizar_data(ofs[col_data_ofs])
        df = ofs[ofs[col_data_ofs] == data_alvo].copy()

        if df.empty:
            self.exibir_mensagem_resultado(f"Sem ordens agendadas para {texto_data}.")
            return

        # NÃO aplicamos o filtro de concluídas para a visão confirmado; mantemos todos os agendamentos
        df['OS_CLEAN'] = df[col_os_ofs].astype(str).str.split('.').str[0].str.strip()

        self.forms_raw = forms.copy()

        if col_data_forms:
            forms[col_data_forms] = normalizar_data(forms[col_data_forms])
            forms = forms[forms[col_data_forms] == data_alvo].copy()
            if forms.empty:
                self.exibir_mensagem_resultado(f"Nenhuma confirmação no forms para {texto_data}.")
                return

        MOTIVO_PRODUTIVO = {
            'Troca de ONU': True, 'Troca de conector interno': True, 'Cliente ausente': False,
            'Troca de conector externo': True, 'Drop refeito': True, 'Concluída': True,
            'Área de risco': False, 'Entrada não autorizada': False, 'Problema de infraestrutura': False,
            'Não cumprimento de agenda': False, 'Interna cliente': False, 'Normalizado sem intervenção técnica': True,
            'Chuva': False, 'Provisionamento ONU': True, 'Problema CTO': False, 'Falha massiva': False,
            'Solicitação de reagendamento': False, 'Reconexão externa - CTO': True, 'Limpeza de conector interno': True,
            'Desistiu do serviço': False, 'Reposição ONU': True, 'Tubulação obstruída': False,
            'Abertura indevida': False, 'Endereço não localizado': False, 'Reconexão interna': True,
            'Troca de fonte': True, 'Endereço incorreto': False, 'Troca de cabo telefônico': True,
            'Reconfiguração ONU': True, 'Situação de risco': False,
        }

        if self.view_mode in ('efetividade', 'operador'):
            if not col_motivo_ofs:
                self.exibir_mensagem_resultado("Coluna de motivo não encontrada no OFS. Impossível calcular efetividade.")
                return
            df['MOTIVO_CLEAN'] = df[col_motivo_ofs].astype(str).str.strip()
            df = df[df['MOTIVO_CLEAN'].isin(MOTIVO_PRODUTIVO.keys())].copy()
            if df.empty:
                self.exibir_mensagem_resultado("Nenhuma OS com motivo de conclusão mapeado.")
                return
            df['PRODUTIVO'] = df['MOTIVO_CLEAN'].map(MOTIVO_PRODUTIVO)
        else:
            df['PRODUTIVO'] = False

        def classificar_regiao(val):
            v = str(val).upper().strip()
            litoral_nomes = ['CARAGUATATUBA', 'CARAGUATATUBA1', 'SÃO SEBASTIÃO', 'SAO SEBASTIAO', 'BOIÇUCANGA', 'ILHABELA', 'CGT', 'BASE', 'SST', 'SSTBO', 'IBL']
            rj_nomes = ['RESENDE', 'MIGUEL PEREIRA', 'VASSOURAS', 'VOLTA REDONDA', 'PENEDO', 'VALENÇA', 'VALENCA', 'ITATIAIA', 'TRÊS RIOS', 'TRES RIOS', 'BARRA MANSA', 'PORTO REAL', 'COMENDADOR LEVY GASPARIAN', 'BARRA DO PIRAÍ', 'BARRA DO PIRAI', 'PATY DO ALFERES', 'PARAÍBA DO SUL', 'PARAIBA DO SUL', 'PINHEIRAL', 'RSD', 'MPE', 'VAS', 'VRD', 'PNDO', 'VLC', 'IZA', 'TRS', 'BMA', 'PORE', 'COLG', 'BPI', 'PFS', 'PDS', 'PNHE', 'RIO', 'RJ', 'SUL RJ']

            if any(x in v for x in litoral_nomes):
                return 'LITORAL NORTE SP'
            elif any(x in v for x in rj_nomes):
                return 'SUL RJ'
            return 'OUTROS'

        df['REGIAO'] = df[col_area_ofs].apply(classificar_regiao) if col_area_ofs else 'DESCONHECIDO'
        df = df[df['REGIAO'].isin(['LITORAL NORTE SP', 'SUL RJ'])]

        try:
            if self.view_mode == 'confirmado':
                self.processar_visao_confirmado(df, forms, col_os_forms, col_status_forms, texto_data)
            elif self.view_mode == 'efetividade':
                self.processar_visao_efetividade(df, col_motivo_ofs, texto_data)
            elif self.view_mode == 'operador':
                self.renderizar_visao_operador(df, forms, col_os_forms, col_status_forms)
            else:
                self.exibir_mensagem_resultado("Visão desconhecida.")
        except Exception as e:
            logger.exception(f"Erro ao processar a visão {self.view_mode}.")
            self.exibir_mensagem_resultado(f"Erro ao processar a visão: {e}")
            for widget in self.detail_frame.winfo_children():
                widget.destroy()

    def renderizar_visao_operador(self, df, forms, col_os_forms, col_status_forms):
        MAPEAMENTO_MOTIVOS = {
            'CLIENTE - AUSENTE': 'CLIENTE', 'REDE - FALHA MASSIVA': 'TÉCNICA',
            'CLIENTE - CAIXA CHEIA': 'TÉCNICA', 'REDE - PROBLEMA DE INFRAESTRUTURA': 'TÉCNICA',
            'CLIENTE - REAGENDOU': 'CLIENTE', 'CLIENTE - ENDERECO NAO LOCALIZADO': 'COMERCIAL',
            'CLIENTE - ENTRADA / AREA NAO LIBERADA': 'TÉCNICA', 'REDE - PREDIO SEM MDU': 'TÉCNICA',
            'REDE - PROBLEMA DE REDE': 'TÉCNICA', 'CLIENTE - ABERTURA INDEVIDA': 'COMERCIAL',
            'CLIENTE - DESISTIU DO SERVICO': 'CLIENTE', 'CLIENTE - TUBULAÇÃO OBSTRUÍDA': 'TÉCNICA',
            'CLIENTE - INTERNA CLIENTE': 'TÉCNICA', 'CAMPO - NÃO CUMPRIMENTO DE AGENDA': 'TÉCNICA',
            'CLIENTE - AREA/SITUAÇÃO DE RISCO': 'TÉCNICA', 'CLIENTE - ABERTURA INDEVIDA REPARO': 'CLIENTE',
            'CAMPO - CHUVA': 'TÉCNICA', 'CLIENTE - SUSPEITA DE FRAUDE': 'CLIENTE',
            'CAMPO - FALTA MATERIAL': 'TÉCNICA', 'CLIENTE - MUDOU DE ENDERECO': 'CLIENTE',
            'CLIENTE - AREA/SITUAÇÃO DE RISCO RETIRADA': 'TÉCNICA', 'CLIENTE - ENDEREÇO INCORRETO': 'COMERCIAL',
            'CLIENTE - SOLICITAÇÃO DE REAGENDAMENTO': 'CLIENTE',
            'Cliente ausente': 'CLIENTE', 'Área de risco': 'TÉCNICA',
            'Entrada não autorizada': 'COMERCIAL', 'Problema de infraestrutura': 'TÉCNICA',
            'Não cumprimento de agenda': 'TÉCNICA', 'Interna cliente': 'TÉCNICA',
            'Chuva': 'TÉCNICA', 'Problema CTO': 'TÉCNICA', 'Falha massiva': 'TÉCNICA',
            'Solicitação de reagendamento': 'CLIENTE', 'Desistiu do serviço': 'COMERCIAL',
            'Tubulação obstruída': 'TÉCNICA', 'Abertura indevida': 'COMERCIAL',
            'Endereço não localizado': 'COMERCIAL', 'Endereço incorreto': 'COMERCIAL',
            'Situação de risco': 'TÉCNICA',
        }

        def categorizar_motivo(motivo):
            if pd.isna(motivo): return 'OUTROS'
            m = str(motivo).strip()
            for chave, categoria in MAPEAMENTO_MOTIVOS.items():
                if chave.upper() == m.upper(): return categoria
            return 'OUTROS'

        forms_clean = forms.copy()
        forms_clean['OS_CLEAN'] = forms_clean[col_os_forms].astype(str).str.split('.').str[0].str.strip()
        status_map = dict(zip(forms_clean['OS_CLEAN'], forms_clean[col_status_forms].str.upper()))
        col_op = encontrar_coluna(forms_clean, ['OPERADOR', 'Operador', 'Nome do Operador', 'Técnico', 'Recurso'])
        if not col_op:
            self.exibir_mensagem_resultado("Coluna de OPERADOR não encontrada no forms de confirmação.")
            return
        op_map = dict(zip(forms_clean['OS_CLEAN'], forms_clean[col_op].str.upper()))
        df['STATUS_CONF'] = df['OS_CLEAN'].map(status_map).fillna('NÃO TRATADO')
        df['OPERADOR'] = df['OS_CLEAN'].map(op_map).fillna('NÃO TRATADO')
        if 'PRODUTIVO' not in df.columns:
            self.exibir_mensagem_resultado("Informação de produtividade não disponível.")
            return
        if 'MOTIVO_MACRO' not in df.columns:
            motivo_col = encontrar_coluna(df, ['Motivo de Encerramento das atividades', 'Motivo', 'MOTIVO'])
            if motivo_col:
                df['MOTIVO_MACRO'] = df[motivo_col].apply(categorizar_motivo)
            else:
                df['MOTIVO_MACRO'] = 'OUTROS'

        confirmados = df[df['STATUS_CONF'] == 'CONFIRMADO'].copy()
        if confirmados.empty:
            self.exibir_mensagem_resultado("Nenhuma OS confirmada encontrada para os operadores.")
            return

        metricas = {}
        for op in confirmados['OPERADOR'].unique():
            sub = confirmados[confirmados['OPERADOR'] == op]
            total = len(sub)
            produtivos = sub['PRODUTIVO'].sum()
            tecnicos = (sub['MOTIVO_MACRO'] == 'TÉCNICA').sum()
            denominador = total - tecnicos
            pct = produtivos / denominador if denominador > 0 else 0
            metricas[op] = {'produtivos': produtivos, 'pct': pct}

        ordem = sorted(metricas, key=lambda op: metricas[op]['pct'], reverse=True)

        for widget in self.result_container.winfo_children():
            widget.destroy()

        header = ctk.CTkFrame(self.result_container, fg_color="#06152E", height=40)
        header.pack(fill="x", pady=5)
        ctk.CTkLabel(header, text="OPERADOR", width=200, anchor="w", text_color="#00BFFF", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=20)
        ctk.CTkLabel(header, text="EFETIVIDADE", text_color="#00BFFF", font=ctk.CTkFont(weight="bold")).pack(side="left", expand=True)

        for op in ordem:
            m = metricas[op]
            self.criar_linha_estilo_autenticador(self.result_container, op, m['produtivos'], m['pct'], "#00FF88")

        sep = ctk.CTkFrame(self.result_container, height=2, fg_color="#30568C")
        sep.pack(fill="x", padx=20, pady=(15, 10))
        ctk.CTkLabel(self.result_container, text="PREENCHIMENTO DE FORMS", text_color="#00BFFF", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(5,10))
        ctk.CTkLabel(self.result_container, text="(total de preenchimentos no forms)", text_color="#AAAAAA", font=ctk.CTkFont(size=11)).pack(pady=(0,5))

        if hasattr(self, 'forms_raw'):
            forms_total = self.forms_raw.copy()
        else:
            forms_total = forms_clean.copy()
        forms_total.columns = forms_total.columns.str.strip()
        col_os_forms_total = encontrar_coluna(forms_total, ['ORDEM DE SERVIÇO', 'OS', 'Ordem de Serviço'])
        col_status_forms_total = encontrar_coluna(forms_total, ['STATUS', 'Status'])
        col_op_total = encontrar_coluna(forms_total, ['OPERADOR', 'Operador', 'Nome do Operador', 'Técnico', 'Recurso'])
        if not col_op_total:
            forms_total['OPERADOR'] = forms_total[col_os_forms_total].astype(str).str.split('.').str[0].str.strip().map(op_map)
        else:
            forms_total['OPERADOR'] = forms_total[col_op_total].str.upper()
        forms_total['STATUS_FORM'] = forms_total[col_status_forms_total].str.upper().replace('CANCELAR', 'CANCELADO')
        status_interesse = ['CONFIRMADO', 'CANCELADO', 'REAGENDAMENTO']
        forms_total = forms_total[forms_total['STATUS_FORM'].isin(status_interesse)]
        contagem_total = forms_total.groupby(['OPERADOR', 'STATUS_FORM']).size().unstack(fill_value=0)
        totais_gerais = contagem_total.sum()

        for op in ordem:
            if op in contagem_total.index:
                counts = contagem_total.loc[op]
            else:
                counts = pd.Series(0, index=status_interesse)
            total_forms = counts.sum()
            if total_forms == 0:
                continue
            op_frame = ctk.CTkFrame(self.result_container, fg_color="#102B57", corner_autenticador=6)
            op_frame.pack(fill="x", padx=20, pady=5)
            ctk.CTkLabel(op_frame, text=op, width=150, anchor="w", text_color="#00BFFF", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)
            bar_frame = ctk.CTkFrame(op_frame, fg_color="transparent")
            bar_frame.pack(side="left", fill="x", expand=True, padx=10)
            cores = {'CONFIRMADO': '#00FF88', 'CANCELADO': '#6495ED', 'REAGENDAMENTO': '#4682B4'}
            for st in status_interesse:
                qtd = counts.get(st, 0)
                total_geral = totais_gerais.get(st, 1)
                self.criar_grafico_barra(bar_frame, st.capitalize(), qtd, total_geral, cores[st])

        for widget in self.detail_frame.winfo_children():
            widget.destroy()

    def processar_visao_efetividade(self, df, col_motivo_ofs, texto_data):
        offset_seg = 300
        MAPEAMENTO_MOTIVOS = {
            'CLIENTE - AUSENTE': 'CLIENTE', 'REDE - FALHA MASSIVA': 'TÉCNICA',
            'CLIENTE - CAIXA CHEIA': 'TÉCNICA', 'REDE - PROBLEMA DE INFRAESTRUTURA': 'TÉCNICA',
            'CLIENTE - REAGENDOU': 'CLIENTE', 'CLIENTE - ENDERECO NAO LOCALIZADO': 'COMERCIAL',
            'CLIENTE - ENTRADA / AREA NAO LIBERADA': 'TÉCNICA', 'REDE - PREDIO SEM MDU': 'TÉCNICA',
            'REDE - PROBLEMA DE REDE': 'TÉCNICA', 'CLIENTE - ABERTURA INDEVIDA': 'COMERCIAL',
            'CLIENTE - DESISTIU DO SERVICO': 'CLIENTE', 'CLIENTE - TUBULAÇÃO OBSTRUÍDA': 'TÉCNICA',
            'CLIENTE - INTERNA CLIENTE': 'TÉCNICA', 'CAMPO - NÃO CUMPRIMENTO DE AGENDA': 'TÉCNICA',
            'CLIENTE - AREA/SITUAÇÃO DE RISCO': 'TÉCNICA', 'CLIENTE - ABERTURA INDEVIDA REPARO': 'CLIENTE',
            'CAMPO - CHUVA': 'TÉCNICA', 'CLIENTE - SUSPEITA DE FRAUDE': 'CLIENTE',
            'CAMPO - FALTA MATERIAL': 'TÉCNICA', 'CLIENTE - MUDOU DE ENDERECO': 'CLIENTE',
            'CLIENTE - AREA/SITUAÇÃO DE RISCO RETIRADA': 'TÉCNICA', 'CLIENTE - ENDEREÇO INCORRETO': 'COMERCIAL',
            'CLIENTE - SOLICITAÇÃO DE REAGENDAMENTO': 'CLIENTE',
            'Cliente ausente': 'CLIENTE', 'Área de risco': 'TÉCNICA',
            'Entrada não autorizada': 'COMERCIAL', 'Problema de infraestrutura': 'TÉCNICA',
            'Não cumprimento de agenda': 'TÉCNICA', 'Interna cliente': 'TÉCNICA',
            'Chuva': 'TÉCNICA', 'Problema CTO': 'TÉCNICA', 'Falha massiva': 'TÉCNICA',
            'Solicitação de reagendamento': 'CLIENTE', 'Desistiu do serviço': 'COMERCIAL',
            'Tubulação obstruída': 'TÉCNICA', 'Abertura indevida': 'COMERCIAL',
            'Endereço não localizado': 'COMERCIAL', 'Endereço incorreto': 'COMERCIAL',
            'Situação de risco': 'TÉCNICA',
        }

        def categorizar_motivo(motivo):
            if pd.isna(motivo): return 'OUTROS'
            m = str(motivo).strip()
            for chave, categoria in MAPEAMENTO_MOTIVOS.items():
                if chave.upper() == m.upper(): return categoria
            return 'OUTROS'

        if col_motivo_ofs:
            df['MOTIVO_MACRO'] = df[col_motivo_ofs].apply(categorizar_motivo)
        else:
            df['MOTIVO_MACRO'] = 'OUTROS'

        for widget in self.result_container.winfo_children():
            widget.destroy()

        ctk.CTkLabel(
            self.result_container, text=f"VISÃO EFETIVIDADE • {texto_data}",
            font=ctk.CTkFont(size=18, weight="bold"), text_color="#00BFFF"
        ).pack(pady=(10, 20))

        for reg, nome_reg in [('LITORAL NORTE SP', 'LITORAL NORTE SP'), ('SUL RJ', 'SUL RJ')]:
            df_reg = df[df['REGIAO'] == reg]
            total = len(df_reg)
            if total == 0: continue

            produtivas = df_reg['PRODUTIVO'].sum()
            nao_produtivas = total - produtivas
            efetividade = (produtivas / total * 100) if total > 0 else 0.0

            reg_frame = ctk.CTkFrame(self.result_container, fg_color="#102B57", corner_autenticador=10)
            reg_frame.pack(fill="x", padx=10, pady=10)

            ctk.CTkLabel(
                reg_frame, text=nome_reg, font=ctk.CTkFont(size=16, weight="bold"), text_color="#FFFFFF"
            ).pack(anchor="w", padx=15, pady=(10, 5))

            self.criar_grafico_barra(reg_frame, "Produtivas", produtivas, total, "#00FF88")
            self.criar_grafico_barra(reg_frame, "Não Produtivas", nao_produtivas, total, "#607B8B")

            ctk.CTkLabel(
                reg_frame, text=f"Total: {total} | Efetividade: {efetividade:.1f}%",
                font=ctk.CTkFont(size=13, weight="bold"), text_color="#AAAAAA"
            ).pack(anchor="e", padx=20, pady=(5, 10))

            if nao_produtivas > 0:
                cond_nao_prod = ~df_reg['PRODUTIVO']
                f_cliente = df_reg[(df_reg['MOTIVO_MACRO'] == 'CLIENTE') & cond_nao_prod]
                f_tecnica = df_reg[(df_reg['MOTIVO_MACRO'] == 'TÉCNICA') & cond_nao_prod]
                f_comercial = df_reg[(df_reg['MOTIVO_MACRO'] == 'COMERCIAL') & cond_nao_prod]

                ofensores_frame = ctk.CTkFrame(reg_frame, fg_color="#0A1931", corner_autenticador=6)
                ofensores_frame.pack(fill="x", padx=15, pady=(0, 15))
                ctk.CTkLabel(
                    ofensores_frame, text="Ofensores da Não Produtividade",
                    font=ctk.CTkFont(size=13, weight="bold"), text_color="#AAAAAA"
                ).pack(anchor="w", padx=15, pady=(5, 5))
                self.criar_grafico_barra(ofensores_frame, "Cliente", len(f_cliente), nao_produtivas, "#6495ED")
                self.criar_grafico_barra(ofensores_frame, "Técnica", len(f_tecnica), nao_produtivas, "#4682B4")
                self.criar_grafico_barra(ofensores_frame, "Comercial", len(f_comercial), nao_produtivas, "#00BFFF")

        self.df_efetividade = df
        for widget in self.detail_frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(self.detail_frame, text="Detalhar Não Produtivas:").pack(side="left", padx=(0, 10))
        for reg in ['LITORAL NORTE SP', 'SUL RJ']:
            ctk.CTkButton(
                self.detail_frame, text=reg.replace("NORTE SP", "").strip(), width=120,
                command=lambda r=reg: self.detalhar_efetividade(r)
            ).pack(side="left", padx=5)

    def processar_visao_confirmado(self, df, forms, col_os_forms, col_status_forms, texto_data):
        forms_clean = forms.copy()
        forms_clean['OS_CLEAN'] = forms_clean[col_os_forms].astype(str).str.split('.').str[0].str.strip()

        # Dicionário de status consolidado: dá prioridade a CONFIRMADO se existir em qualquer linha
        status_dict = {}
        if col_status_forms:
            for _, r in forms_clean.iterrows():
                os_clean = r['OS_CLEAN']
                st = str(r[col_status_forms]).upper().strip()
                if 'CONFIRMADO' in st:
                    status_dict[os_clean] = st
                elif os_clean not in status_dict:
                    status_dict[os_clean] = st

        def verificar_status(os_val):
            st_f = status_dict.get(os_val, 'NÃO ENCONTRADO')
            return 'CONFIRMADO' if 'CONFIRMADO' in st_f else 'NÃO CONFIRMADO'

        df['STATUS_CONFIRMACAO'] = df['OS_CLEAN'].apply(verificar_status)
        self.df_confirmado = df

        for widget in self.result_container.winfo_children():
            widget.destroy()

        ctk.CTkLabel(
            self.result_container, text=f"VISÃO CONFIRMADO • {texto_data}",
            font=ctk.CTkFont(size=18, weight="bold"), text_color="#00BFFF"
        ).pack(pady=(10, 10))

        tabview = ctk.CTkTabview(
            self.result_container,
            fg_color="#0A1931",
            segmented_button_fg_color="#06152E",
            segmented_button_selected_color="#007BFF",
            segmented_button_selected_hover_color="#0056B3",
            segmented_button_unselected_hover_color="#102B57"
        )
        tabview.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        tabview.add('LITORAL NORTE SP')
        tabview.add('SUL RJ')

        for reg, nome_reg in [('LITORAL NORTE SP', 'LITORAL NORTE SP'), ('SUL RJ', 'SUL RJ')]:
            aba = tabview.tab(nome_reg)
            df_reg = df[df['REGIAO'] == reg]
            total = len(df_reg)
            if total == 0:
                ctk.CTkLabel(aba, text=f"Sem agendamentos mapeados para {nome_reg}.", text_color="#AAAAAA").pack(pady=40)
                continue

            df_conf = df_reg[df_reg['STATUS_CONFIRMACAO'] == 'CONFIRMADO']
            conf = len(df_conf)
            n_conf = total - conf

            total_ativacao = total_mudanca = total_outros = 0
            qtd_ativacao = qtd_mudanca = qtd_outros = 0

            if hasattr(self, 'col_tipo_ofs') and self.col_tipo_ofs:
                for _, row in df_reg.iterrows():
                    tipo_str = str(row[self.col_tipo_ofs]).upper()
                    is_conf = (row['STATUS_CONFIRMACAO'] == 'CONFIRMADO')

                    if 'ATIVAÇÃO' in tipo_str or 'ATIVACAO' in tipo_str:
                        total_ativacao += 1
                        if is_conf: qtd_ativacao += 1
                    elif 'MUDANÇA' in tipo_str or 'MUDANCA' in tipo_str or 'MDE' in tipo_str:
                        total_mudanca += 1
                        if is_conf: qtd_mudanca += 1
                    else:
                        total_outros += 1
                        if is_conf: qtd_outros += 1
            else:
                total_outros = total
                qtd_outros = conf

            reg_frame = ctk.CTkFrame(aba, fg_color="#102B57", corner_autenticador=10)
            reg_frame.pack(fill="x", padx=10, pady=5)
            ctk.CTkLabel(
                reg_frame, text="Indicadores Principais", font=ctk.CTkFont(size=14, weight="bold"), text_color="#FFFFFF"
            ).pack(anchor="w", padx=15, pady=(10, 5))

            self.criar_grafico_barra(reg_frame, "Confirmados", conf, total, "#00FF88")
            self.criar_grafico_barra(reg_frame, "Não Confirmados", n_conf, total, "#AAAAAA")
            ctk.CTkFrame(reg_frame, height=2, fg_color="#16376B").pack(fill="x", padx=15, pady=10)

            ctk.CTkLabel(
                reg_frame, text="Detalhamento dos Confirmados", font=ctk.CTkFont(size=14, weight="bold"), text_color="#FFFFFF"
            ).pack(anchor="w", padx=15, pady=(5, 5))

            if conf > 0:
                self.criar_grafico_barra(reg_frame, "Ativações", qtd_ativacao, total_ativacao, "#00BFFF")
                self.criar_grafico_barra(reg_frame, "Mudanças de Endereço", qtd_mudanca, total_mudanca, "#6495ED")
                self.criar_grafico_barra(reg_frame, "Outros Serviços", qtd_outros, total_outros, "#B84DFF")
            else:
                ctk.CTkLabel(reg_frame, text="Nenhum agendamento confirmado ainda.", font=ctk.CTkFont(size=13), text_color="#AAAAAA").pack(pady=10)

            ctk.CTkLabel(
                reg_frame, text=f"Total Agendado: {total}", font=ctk.CTkFont(size=13, weight="bold"), text_color="#AAAAAA"
            ).pack(anchor="e", padx=20, pady=(15, 10))

        for widget in self.detail_frame.winfo_children():
            widget.destroy()
        ctk.CTkLabel(self.detail_frame, text="Detalhar Não Confirmados:").pack(side="left", padx=(0, 10))
        ctk.CTkButton(
            self.detail_frame, text="Litoral Norte", width=120,
            command=lambda: self.detalhar_confirmado('LITORAL NORTE SP', 'NÃO CONFIRMADO')
        ).pack(side="left", padx=5)
        ctk.CTkButton(
            self.detail_frame, text="Sul RJ", width=120,
            command=lambda: self.detalhar_confirmado('SUL RJ', 'NÃO CONFIRMADO')
        ).pack(side="left", padx=5)

    def detalhar_confirmado(self, regiao, status):
        if not hasattr(self, 'df_confirmado'): return
        df_reg = self.df_confirmado[(self.df_confirmado['REGIAO'] == regiao) & (self.df_confirmado['STATUS_CONFIRMACAO'] == status)]

        # *** CORREÇÃO: Remover OS que já foram concluídas (baixadas) na lista detalhada ***
        if self.col_status_ofs and self.col_status_ofs in df_reg.columns:
            df_reg = df_reg[~df_reg[self.col_status_ofs].astype(str).str.contains('conclu', case=False, na=False)]

        for widget in self.result_container.winfo_children():
            widget.destroy()
        detalhe_textbox = ctk.CTkTextbox(
            self.result_container, font=ctk.CTkFont(family="Consolas", size=13),
            fg_color="transparent", text_color="#FFFFFF", wrap="none"
        )
        detalhe_textbox.pack(fill="both", expand=True, padx=5, pady=5)
        if df_reg.empty:
            detalhe_textbox.insert("end", f"Nenhuma OS pendente com status '{status}' em {regiao}.\n")
        else:
            detalhe_textbox.insert("end", f"=== {status} PENDENTES em {regiao} ===\n\n")
            for _, row in df_reg.iterrows():
                contrato_val = row.get(self.col_contrato_ofs, '?') if hasattr(self, 'col_contrato_ofs') and self.col_contrato_ofs else '?'
                detalhe_textbox.insert("end", f"{row['OS_CLEAN']} | Contrato: {contrato_val}\n")
        detalhe_textbox.configure(state="disabled")

    def detalhar_efetividade(self, regiao):
        if not hasattr(self, 'df_efetividade'): return
        cond_nao_prod = ~self.df_efetividade['PRODUTIVO']
        df_reg = self.df_efetividade[(self.df_efetividade['REGIAO'] == regiao) & cond_nao_prod]
        for widget in self.result_container.winfo_children():
            widget.destroy()
        detalhe_textbox = ctk.CTkTextbox(
            self.result_container, font=ctk.CTkFont(family="Consolas", size=13),
            fg_color="transparent", text_color="#FFFFFF", wrap="none"
        )
        detalhe_textbox.pack(fill="both", expand=True, padx=5, pady=5)
        if df_reg.empty:
            detalhe_textbox.insert("end", f"Nenhuma não produtiva em {regiao}.\n")
        else:
            detalhe_textbox.insert("end", f"=== NÃO PRODUTIVAS - {regiao} ===\n\n")
            for _, row in df_reg.iterrows():
                motivo = row.get('MOTIVO_MACRO', '?')
                detalhe_textbox.insert("end", f"{row['OS_CLEAN']} | Motivo Macro: {motivo}\n")
        detalhe_textbox.configure(state="disabled")

if __name__ == "__main__":
    app = OperacionalApp()
    app.mainloop()