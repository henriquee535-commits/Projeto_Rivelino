import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import sqlite3
import uuid
from datetime import datetime, timedelta
import io
import base64
import smtplib
from email.mime.text import MIMEText
import random

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Inventário José Rivelino", layout="wide", page_icon="📦")

# --- CONFIGURAÇÕES ---
ARQUIVO_PLANILHA = 'Almoxarifado.xlsm'
SENHA_ACESSO = st.secrets["SENHA_ACESSO"]
SENHA_ZERAR_ESTOQUE = st.secrets["SENHA_ZERAR_ESTOQUE"]
LIMITE_PESSOAS = 40
TEMPO_INATIVIDADE = 1
DB_NAME = "almoxarifado.db"

# --- CSS GLOBAL ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Sora', sans-serif; }
.stButton > button {
    background: linear-gradient(135deg, #1a3a4a, #0d5c8a);
    color: white; border: none; border-radius: 8px;
    font-family: 'Sora', sans-serif; font-weight: 600;
    padding: 0.5rem 1.5rem; transition: opacity 0.2s;
}
.stButton > button:hover { opacity: 0.88; color: white; }
.stAlert { border-radius: 10px; }
</style>
""", unsafe_allow_html=True)

# --- CONEXÃO COM SQLITE ---
def get_conn():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

# --- BANCO DE DADOS ---
def init_db():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS estoque (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Codigo TEXT,
                Descricao TEXT,
                Quantidade INTEGER,
                CC TEXT,
                Preco_Custo REAL DEFAULT 0.00,
                Preco_Venda REAL DEFAULT 0.00
            )
        ''')
        c.execute('''
            CREATE TABLE IF NOT EXISTS acessos (
                sessao_id TEXT PRIMARY KEY,
                ultimo_clique DATETIME
            )
        ''')
        c.execute('''
            CREATE TABLE IF NOT EXISTS centros_custo (
                nome TEXT PRIMARY KEY
            )
        ''')
        c.execute('''
            CREATE TABLE IF NOT EXISTS financeiro (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Codigo TEXT,
                Descricao TEXT,
                tipo TEXT,
                quantidade INTEGER,
                valor_total REAL,
                data DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        conn.commit()

init_db()

@st.cache_data(ttl=300)
def carregar_estoque():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('SELECT Codigo, Descricao, Quantidade, CC, Preco_Custo, Preco_Venda FROM estoque')
        rows = c.fetchall()
    df = pd.DataFrame([dict(row) for row in rows], columns=['Codigo', 'Descricao', 'Quantidade', 'CC', 'Preco_Custo', 'Preco_Venda'])
    if not df.empty:
        df['Quantidade'] = df['Quantidade'].astype(int)
        df['Preco_Custo'] = df['Preco_Custo'].astype(float)
        df['Preco_Venda'] = df['Preco_Venda'].astype(float)
    return df

@st.cache_data
def carregar_ccs():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("SELECT nome FROM centros_custo ORDER BY nome")
        rows = c.fetchall()
    lista_cc = [r['nome'] for r in rows]
    if not lista_cc:
        lista_cc = ["Centro de Custo Geral"]
        with get_conn() as conn:
            c = conn.cursor()
            c.execute("INSERT OR IGNORE INTO centros_custo (nome) VALUES (?)", ("Centro de Custo Geral",))
            conn.commit()
    return lista_cc

def buscar_descricao_por_codigo(cod):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('SELECT DISTINCT Descricao FROM estoque WHERE Codigo = ?', (cod,))
        result = c.fetchone()
    return result['Descricao'] if result else None

def gerar_template_xlsx():
    template_df = pd.DataFrame({
        'Codigo': ['ABC001', 'ABC002'],
        'Descricao': ['Parafuso M8', 'Cabo Elétrico 2,5mm'],
        'Quantidade': [100, 50],
        'CC': ['01/0001 - LIVRE DEMANDA', '01/0001 - LIVRE DEMANDA'],
        'Preco_Custo': [0.50, 2.50],
        'Preco_Venda': [1.00, 5.00]
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False, sheet_name='Inventario')
    return buf.getvalue()

def gerar_template_depara():
    df = pd.DataFrame({
        'De': ['Centro de Custo Antigo 1', 'Centro de Custo Antigo 2'],
        'Para': ['Centro de Custo Novo 1', 'Centro de Custo Novo 2']
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='DePara')
    return buf.getvalue()

def logo_para_base64(path):
    for tentativa in [path, path.replace('.png', '.jpg'), path.replace('.png', '.jpeg')]:
        try:
            with open(tentativa, "rb") as f:
                data = base64.b64encode(f.read()).decode()
            ext = tentativa.rsplit('.', 1)[-1].lower()
            mime = 'image/png' if ext == 'png' else 'image/jpeg'
            return f"data:{mime};base64,{data}"
        except FileNotFoundError:
            continue
    return None

def formatar_moeda(valor):
    return f"R$ {valor:_.2f}".replace('.', ',').replace('_', '.')

# --- SISTEMA DE APROVAÇÃO POR E-MAIL ---
def aprovar_acao_master(chave, descricao_acao):
    if f"token_{chave}" not in st.session_state:
        st.session_state[f"token_{chave}"] = None

    email_solicitante = st.text_input("📧 Seu e-mail (para identificação):", key=f"email_{chave}", placeholder="seunome@gmail.com")

    if st.button(f"📩 Solicitar Liberação: {descricao_acao}", key=f"req_{chave}"):
        if not email_solicitante:
            st.error("⛔ Informe seu e-mail antes de solicitar.")
            return False

        codigo = str(random.randint(100000, 999999))
        st.session_state[f"token_{chave}"] = codigo

        remetente    = st.secrets["email"]["remetente"]
        senha_email  = st.secrets["email"]["senha"]
        destinatario = st.secrets["email"]["destinatario"]

        msg = MIMEText(
            f"Solicitação de ação no sistema de Almoxarifado:\n\n"
            f"SOLICITANTE: {email_solicitante}\n"
            f"AÇÃO: {descricao_acao}\n\n"
            f"Para autorizar, informe o código abaixo:\n"
            f"CÓDIGO: {codigo}"
        )
        msg['Subject'] = 'Aprovação de Sistema - Almoxarifado'
        msg['From'] = remetente
        msg['To'] = destinatario

        try:
            with smtplib.SMTP('smtp.office365.com', 587) as server:
                server.starttls()
                server.login(remetente, senha_email)
                server.sendmail(remetente, [destinatario], msg.as_string())
            st.info("✅ Solicitação enviada!")
        except Exception as e:
            st.error(f"Erro ao enviar e-mail: {e}")

    if st.session_state[f"token_{chave}"]:
        token_input = st.text_input("🔑 Código enviado para Eduardo Sousa - Controladoria. Código:", key=f"inp_{chave}")
        if st.button("✅ Confirmar Execução", key=f"exec_{chave}"):
            if token_input == st.session_state[f"token_{chave}"]:
                st.session_state[f"token_{chave}"] = None
                return True
            else:
                st.error("⛔ Código incorreto!")
    return False

# --- CONTROLE DE ACESSO E LIMITE DE USUÁRIOS ---
if 'sessao_id' not in st.session_state:
    st.session_state.sessao_id = str(uuid.uuid4())

with get_conn() as conn:
    c = conn.cursor()
    tempo_limite = (datetime.now() - timedelta(minutes=TEMPO_INATIVIDADE)).strftime('%Y-%m-%d %H:%M:%S')
    c.execute("DELETE FROM acessos WHERE ultimo_clique < ?", (tempo_limite,))
    
    agora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("""
        INSERT INTO acessos (sessao_id, ultimo_clique) VALUES (?, ?)
        ON CONFLICT(sessao_id) DO UPDATE SET ultimo_clique=excluded.ultimo_clique
    """, (st.session_state.sessao_id, agora))
    
    c.execute("SELECT COUNT(*) as total FROM acessos")
    total_ativos = c.fetchone()['total']
    conn.commit()

if total_ativos > LIMITE_PESSOAS:
    st.error(f"⚠️ O sistema está lotado ({total_ativos}/{LIMITE_PESSOAS} usuários). Tente novamente em 1 minuto.")
    st.stop()

# --- NAVEGAÇÃO ---
df = carregar_estoque()
lista_cc = carregar_ccs()

st.sidebar.title("Navegação")
menu = st.sidebar.radio("Ir para:", ["📊 Consulta", "💰 Financeiro", "🔒 Almoxarifado"])
st.sidebar.divider()
st.sidebar.markdown(f"🟢 **{total_ativos}/{LIMITE_PESSOAS}** pessoas online")

# ==========================================
# TELA 1: CONSULTA
# ==========================================
if menu == "📊 Consulta":
    src1 = logo_para_base64("logo1.png")
    src2 = logo_para_base64("logo2.png")
    img1 = f'<img class="img-logo1" src="{src1}">' if src1 else '<span style="color:#102a43;font-weight:700;">LOGO 1</span>'
    img2 = f'<img class="img-logo2" src="{src2}">' if src2 else '<span style="color:#102a43;font-weight:700;">LOGO 2</span>'

    df_ativos = df[df['Quantidade'] > 0]
    total_pecas = f"{df_ativos['Quantidade'].sum():.0f}" if not df_ativos.empty else "0"
    total_itens = str(df_ativos['Codigo'].nunique())    if not df_ativos.empty else "0"
    total_cc    = str(df_ativos['CC'].nunique())        if not df_ativos.empty else "0"

    components.html(f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link href="https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
      * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Sora', sans-serif; }}
      .header-container {{ display: grid; grid-template-columns: 1fr auto 1fr; align-items: center; padding: 20px 32px; border-radius: 16px; margin-bottom: 16px; background: linear-gradient(135deg, #f0f4f8 0%, #d9e2ec 100%); box-shadow: 0 4px 12px rgba(0,0,0,0.05); border: 1px solid #e2e8f0; }}
      .left-logo {{ justify-self: start; display: flex; align-items: center; }}
      .title-box {{ text-align: center; padding: 0 20px; }}
      .right-logo {{ justify-self: end; display: flex; align-items: center; }}
      .img-logo1 {{ height: 85px; width: auto; max-width: 240px; object-fit: contain; mix-blend-mode: darken; }}
      .img-logo2 {{ height: 35px; width: auto; max-width: 120px; object-fit: contain; mix-blend-mode: darken; }}
      .title-box h1 {{ font-size: 1.8rem; font-weight: 700; color: #102a43; letter-spacing: 0.02em; line-height: 1.15; }}
      .title-box p {{ font-size: 0.75rem; color: #334e68; margin-top: 5px; font-weight: 600; letter-spacing: 0.22em; text-transform: uppercase; }}
      .metrics-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-top: 4px; }}
      .metric-card {{ background: #fff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px 20px; box-shadow: 0 1px 4px rgba(0,0,0,0.05); }}
      .metric-label {{ font-size: 0.78rem; color: #718096; font-weight: 600; margin-bottom: 4px; }}
      .metric-value {{ font-size: 1.9rem; font-weight: 700; color: #1a202c; line-height: 1.1; }}
      @media (max-width: 768px) {{
        .header-container {{ grid-template-columns: 1fr; gap: 15px; padding: 15px; text-align: center; }}
        .left-logo, .right-logo {{ justify-self: center; }}
        .img-logo1 {{ height: 60px; }}
        .img-logo2 {{ height: 30px; }}
        .title-box h1 {{ font-size: 1.4rem; }}
        .metrics-grid {{ grid-template-columns: 1fr; gap: 8px; }}
      }}
    </style>
    </head>
    <body>
    <div class="header-container"><div class="left-logo">{img1}</div><div class="title-box"><h1>INVENTÁRIO JOSÉ RIVELINO</h1><p>ALMOXARIFADO</p></div><div class="right-logo">{img2}</div></div>
    <div class="metrics-grid">
      <div class="metric-card"><div class="metric-label">📦 Total de Peças</div><div class="metric-value">{total_pecas}</div></div>
      <div class="metric-card"><div class="metric-label">🏷️ Itens Únicos</div><div class="metric-value">{total_itens}</div></div>
      <div class="metric-card"><div class="metric-label">🏢 Centros de Custo</div><div class="metric-value">{total_cc}</div></div>
    </div>
    </body>
    </html>
    """, height=350, scrolling=True)

    st.divider()

    c_busca, c_filtro = st.columns([2, 1])
    busca     = c_busca.text_input("🔍 Pesquisar Código ou Descrição:")
    cc_filtro = c_filtro.selectbox("🏢 Filtrar por Centro de Custo:", ["Todos"] + lista_cc)

    df_filt = df_ativos.copy()
    if cc_filtro != "Todos":
        df_filt = df_filt[df_filt['CC'] == cc_filtro]
    if busca:
        df_filt = df_filt[
            df_filt['Codigo'].astype(str).str.contains(busca, case=False) |
            df_filt['Descricao'].str.contains(busca, case=False, na=False)
        ]

    df_exibicao = df_filt.copy()
    if not df_exibicao.empty:
        df_exibicao['Preco_Custo'] = df_exibicao['Preco_Custo'].apply(formatar_moeda)
        df_exibicao['Preco_Venda'] = df_exibicao['Preco_Venda'].apply(formatar_moeda)

    st.dataframe(df_exibicao, use_container_width=True, hide_index=True)

# ==========================================
# TELA 1.5: FINANCEIRO
# ==========================================
elif menu == "💰 Financeiro":
    st.title("💰 Dashboard Financeiro")
    
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute("SELECT COALESCE(SUM(valor_total), 0) as total FROM financeiro WHERE tipo='Venda'")
        total_vendas = cur.fetchone()['total']
        
        cur.execute("SELECT COALESCE(SUM(valor_total), 0) as total FROM financeiro WHERE tipo='Perda'")
        total_perdas = cur.fetchone()['total']
        
        cur.execute('SELECT COALESCE(SUM(Quantidade * Preco_Custo), 0) as total FROM estoque WHERE Quantidade > 0')
        valor_estoque_atual = cur.fetchone()['total']

    c1, c2, c3 = st.columns(3)
    c1.metric("Entrada em Caixa (Vendas)", formatar_moeda(float(total_vendas)))
    c2.metric("Prejuízo Registrado (Perdas)", formatar_moeda(float(total_perdas)))
    c3.metric("Valor Atual do Estoque (Custo)", formatar_moeda(float(valor_estoque_atual)))

    st.divider()
    st.subheader("Últimas Movimentações")
    with get_conn() as conn:
        cur = conn.cursor()
        cur.execute('SELECT data, tipo, Codigo, Descricao, quantidade, valor_total FROM financeiro ORDER BY data DESC LIMIT 50')
        mov_rows = cur.fetchall()
            
    if mov_rows:
        df_mov = pd.DataFrame([dict(r) for r in mov_rows])
        df_mov['data'] = pd.to_datetime(df_mov['data']).dt.strftime('%d/%m/%Y %H:%M')
        df_mov['valor_total'] = df_mov['valor_total'].astype(float).apply(formatar_moeda)
        df_mov.columns = ['Data', 'Operação', 'Código', 'Descrição', 'Qtd', 'Valor Total']
        st.dataframe(df_mov, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma movimentação financeira registrada.")

# ==========================================
# TELA 2: ALMOXARIFADO
# ==========================================
elif menu == "🔒 Almoxarifado":
    st.title("🔒 Área Restrita — Almoxarifado")
    senha = st.text_input("Senha:", type="password")

    if senha == SENHA_ACESSO or senha == SENHA_ZERAR_ESTOQUE:
        abas_nomes = ["📝 Registro Individual", "📤 Carga em Massa"]
        if senha == SENHA_ZERAR_ESTOQUE:
            abas_nomes.extend(["🗑️ Excluir Item (Master)", "🏢 Gerenciar CCs (Master)", "⚠️ Limpar Dados (Master)"])

        abas = st.tabs(abas_nomes)

        with abas[0]:
            with st.form("registro", clear_on_submit=True):
                st.markdown("**1. Informações do Produto**")
                c1, c2 = st.columns(2)
                cod        = c1.text_input("Código:")
                desc_input = c2.text_input("Descrição (somente para itens novos):")
                
                st.markdown("**2. Preços (Preenchimento obrigatório para NOVAS Entradas)**")
                c3, c4 = st.columns(2)
                preco_custo = c3.number_input("Preço de Custo Unitário (R$)", min_value=0.0, format="%.2f", step=0.50)
                preco_venda = c4.number_input("Preço de Venda Unitário (R$)", min_value=0.0, format="%.2f", step=0.50)

                st.markdown("**3. Movimentação**")
                c5, c6, c7 = st.columns([2, 2, 1])
                cc_sel = c5.selectbox("Centro de Custo:", lista_cc)
                op     = c6.selectbox("Operação:", ["Entrada", "Venda", "Perda"])
                qtd    = c7.number_input("Qtd:", min_value=1, step=1, format="%d")

                if st.form_submit_button("✅ Confirmar"):
                    if not cod:
                        st.error("⛔ Informe o Código do item.")
                    else:
                        desc_existente = buscar_descricao_por_codigo(cod)
                        if not desc_existente and not desc_input:
                            st.error("⛔ A Descrição é OBRIGATÓRIA para cadastrar um novo item.")
                        elif desc_existente and desc_input and desc_input.strip() != desc_existente.strip():
                            st.error(f"⛔ Conflito! O código **{cod}** já está cadastrado como: **\"{desc_existente}\"**")
                        else:
                            desc_final = desc_existente if desc_existente else desc_input
                            with get_conn() as conn:
                                cur = conn.cursor()
                                cur.execute('SELECT Quantidade, Preco_Custo, Preco_Venda FROM estoque WHERE Codigo=? AND CC=?', (cod, cc_sel))
                                res = cur.fetchone()
                                
                                if res:
                                    p_custo_db = res['Preco_Custo']
                                    p_venda_db = res['Preco_Venda']

                                    if op == "Venda":
                                        if res['Quantidade'] < qtd:
                                            st.error(f"⛔ FALTA DE ESTOQUE! Saldo atual: {res['Quantidade']} unidades.")
                                        else:
                                            cur.execute('UPDATE estoque SET Quantidade = Quantidade - ? WHERE Codigo=? AND CC=?', (qtd, cod, cc_sel))
                                            cur.execute('INSERT INTO financeiro (Codigo, Descricao, tipo, quantidade, valor_total) VALUES (?, ?, ?, ?, ?)', 
                                                        (cod, desc_final, 'Venda', qtd, qtd * p_venda_db))
                                            st.success(f"✅ Venda registrada! Valor total: {formatar_moeda(qtd * p_venda_db)}")
                                            st.cache_data.clear()
                                            
                                    elif op == "Perda":
                                        if res['Quantidade'] < qtd:
                                            st.error(f"⛔ FALTA DE ESTOQUE! Saldo atual: {res['Quantidade']} unidades.")
                                        else:
                                            cur.execute('UPDATE estoque SET Quantidade = Quantidade - ? WHERE Codigo=? AND CC=?', (qtd, cod, cc_sel))
                                            cur.execute('INSERT INTO financeiro (Codigo, Descricao, tipo, quantidade, valor_total) VALUES (?, ?, ?, ?, ?)', 
                                                        (cod, desc_final, 'Perda', qtd, qtd * p_custo_db))
                                            st.success(f"✅ Perda registrada! Prejuízo somado: {formatar_moeda(qtd * p_custo_db)}")
                                            st.cache_data.clear()
                                            
                                    elif op == "Entrada":
                                        novo_custo = preco_custo if preco_custo > 0 else p_custo_db
                                        novo_venda = preco_venda if preco_venda > 0 else p_venda_db
                                        cur.execute('UPDATE estoque SET Quantidade = Quantidade + ?, Preco_Custo = ?, Preco_Venda = ? WHERE Codigo=? AND CC=?', 
                                                    (qtd, novo_custo, novo_venda, cod, cc_sel))
                                        st.success(f"✅ Entrada registrada. Saldo atualizado: {res['Quantidade'] + qtd}")
                                        st.cache_data.clear()
                                else:
                                    if op in ["Venda", "Perda"]:
                                        st.error("⛔ ITEM NÃO ENCONTRADO para realizar Venda ou Perda.")
                                    else:
                                        cur.execute('INSERT INTO estoque (Codigo, Descricao, Quantidade, CC, Preco_Custo, Preco_Venda) VALUES (?, ?, ?, ?, ?, ?)', 
                                                    (cod, desc_final, qtd, cc_sel, preco_custo, preco_venda))
                                        st.success("✅ Item novo cadastrado com sucesso.")
                                        st.cache_data.clear()
                            conn.commit()

        with abas[1]:
            st.info("Upload de arquivo Excel (.xlsx) com colunas: `Codigo` | `Descricao` | `Quantidade` | `CC` | `Preco_Custo` | `Preco_Venda`")
            st.download_button("⬇️ Template Inventário", gerar_template_xlsx(), "template_inventario.xlsx")
            arquivo = st.file_uploader("Arquivo de Inventário (.xlsx):", type=["xlsx"], key="upload_massa")

            if arquivo:
                try:
                    df_upload = pd.read_excel(arquivo, engine='openpyxl')
                    faltando = {'Codigo', 'Descricao', 'Quantidade', 'CC'} - set(df_upload.columns)
                    if faltando:
                        st.error(f"⛔ Colunas obrigatórias ausentes: {', '.join(faltando)}")
                    else:
                        if st.button("🚀 Processar Importação"):
                            df_upload['Codigo']     = df_upload['Codigo'].astype(str).str.strip()
                            df_upload['Descricao']  = df_upload['Descricao'].astype(str).str.strip()
                            df_upload['CC']         = df_upload['CC'].astype(str).str.strip()
                            df_upload['Quantidade'] = pd.to_numeric(df_upload['Quantidade'], errors='coerce')
                            
                            if 'Preco_Custo' not in df_upload.columns: df_upload['Preco_Custo'] = 0.0
                            if 'Preco_Venda' not in df_upload.columns: df_upload['Preco_Venda'] = 0.0
                            
                            df_upload['Preco_Custo'] = pd.to_numeric(df_upload['Preco_Custo'], errors='coerce').fillna(0.0)
                            df_upload['Preco_Venda'] = pd.to_numeric(df_upload['Preco_Venda'], errors='coerce').fillna(0.0)

                            df_upload = df_upload.dropna(subset=['Quantidade'])
                            df_upload = df_upload[df_upload['Quantidade'] > 0]
                            df_upload['Quantidade'] = df_upload['Quantidade'].astype(int)
                            df_upload = df_upload[(df_upload['Codigo'] != 'nan') & (df_upload['Codigo'] != '')]

                            ccs_invalidos = set(df_upload['CC'].unique()) - set(lista_cc)
                            if ccs_invalidos:
                                st.error(f"⛔ IMPORTAÇÃO BLOQUEADA! CCs não encontrados: **{', '.join(ccs_invalidos)}**.")
                            else:
                                with get_conn() as conn:
                                    cur = conn.cursor()
                                    cur.execute('SELECT Codigo, CC FROM estoque')
                                    db_set = set((r['Codigo'], r['CC']) for r in cur.fetchall())

                                    inserts, updates = [], []
                                    for _, row in df_upload.iterrows():
                                        cod_r, desc_r, cc_r, qtd_r = row['Codigo'], row['Descricao'], row['CC'], row['Quantidade']
                                        cust_r, vend_r = row['Preco_Custo'], row['Preco_Venda']

                                        if (cod_r, cc_r) not in db_set and (not desc_r or desc_r.lower() == 'nan'):
                                            continue
                                        if (cod_r, cc_r) in db_set:
                                            updates.append((qtd_r, cust_r, vend_r, cod_r, cc_r))
                                        else:
                                            inserts.append((cod_r, desc_r, qtd_r, cc_r, cust_r, vend_r))
                                            db_set.add((cod_r, cc_r))

                                    if inserts:
                                        cur.executemany('INSERT INTO estoque (Codigo,Descricao,Quantidade,CC,Preco_Custo,Preco_Venda) VALUES (?,?,?,?,?,?)', inserts)
                                    if updates:
                                        cur.executemany('UPDATE estoque SET Quantidade = Quantidade + ?, Preco_Custo = MAX(Preco_Custo, ?), Preco_Venda = MAX(Preco_Venda, ?) WHERE Codigo=? AND CC=?', updates)
                                conn.commit()

                                st.success(f"✅ Importação concluída! {len(inserts)} novos, {len(updates)} atualizados.")
                                st.cache_data.clear()
                                st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")

        if senha == SENHA_ZERAR_ESTOQUE:
            with abas[2]:
                st.subheader("🗑️ Excluir Item do Banco")
                cod_excluir = st.text_input("Digite o Código do item que deseja apagar:")
                if cod_excluir and aprovar_acao_master("del_item", f"Excluir código {cod_excluir}"):
                    with get_conn() as conn:
                        cur = conn.cursor()
                        cur.execute('SELECT * FROM estoque WHERE Codigo=?', (cod_excluir,))
                        if cur.fetchone():
                            cur.execute('DELETE FROM estoque WHERE Codigo=?', (cod_excluir,))
                            st.success(f"✅ Código **{cod_excluir}** apagado!")
                        else:
                            st.error("⛔ Código não encontrado.")
                        conn.commit()
                    st.cache_data.clear()

            with abas[3]:
                c_sec1, c_sec2 = st.columns(2)
                with c_sec1:
                    st.subheader("➕ Novo Centro de Custo")
                    novo_cc = st.text_input("Nome:")
                    if novo_cc and aprovar_acao_master("new_cc", f"Criar CC: {novo_cc}"):
                        with get_conn() as conn:
                            cur = conn.cursor()
                            cur.execute("INSERT OR IGNORE INTO centros_custo (nome) VALUES (?)", (novo_cc,))
                            conn.commit()
                        st.success("Centro de Custo cadastrado!")
                        st.cache_data.clear()
                        st.rerun()

                with c_sec2:
                    st.subheader("🔄 De/Para (Individual)")
                    cc_antigo = st.selectbox("De:", lista_cc)
                    cc_novo   = st.text_input("Para (Novo Nome):")
                    if cc_novo and cc_antigo and aprovar_acao_master("rename_cc", f"Renomear {cc_antigo} → {cc_novo}"):
                        with get_conn() as conn:
                            cur = conn.cursor()
                            cur.execute("INSERT OR IGNORE INTO centros_custo (nome) VALUES (?)", (cc_novo,))
                            cur.execute('UPDATE estoque SET CC = ? WHERE CC = ?', (cc_novo, cc_antigo))
                            cur.execute("DELETE FROM centros_custo WHERE nome = ?", (cc_antigo,))
                            conn.commit()
                        st.success("Centro de Custo renomeado!")
                        st.cache_data.clear()
                        st.rerun()

            with abas[4]:
                st.subheader("⚠️ Área de Risco - Acesso Master")
                opcao = st.radio("Selecione a ação desejada:", [
                    "1️⃣ Apenas zerar o estoque (Mantém os códigos salvos)",
                    "2️⃣ Excluir tudo (Limpa o banco de estoque e códigos)",
                    "3️⃣ Limpar Histórico Financeiro"
                ])

                if aprovar_acao_master("limpeza", f"Limpeza de Banco: {opcao}"):
                    with get_conn() as conn:
                        cur = conn.cursor()
                        if "1️⃣" in opcao:
                            cur.execute('UPDATE estoque SET Quantidade = 0')
                            st.success("Quantidades zeradas!")
                        elif "2️⃣" in opcao:
                            cur.execute("DELETE FROM estoque")
                            st.success("Todos os itens de estoque apagados!")
                        elif "3️⃣" in opcao:
                            cur.execute("DELETE FROM financeiro")
                            st.success("Histórico financeiro limpo!")
                        conn.commit()
                    st.cache_data.clear()
                    st.rerun()
