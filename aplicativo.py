import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import sqlite3
import uuid
from datetime import datetime, timedelta, date
import io
import base64
import smtplib
from email.mime.text import MIMEText
import random

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Inventário José Rivelino", layout="wide", page_icon="📦")

# --- CONFIGURAÇÕES ---
SENHA_ACESSO = st.secrets["SENHA_ACESSO"]
SENHA_ZERAR_ESTOQUE = st.secrets["SENHA_ZERAR_ESTOQUE"]
LIMITE_PESSOAS = 40
TEMPO_INATIVIDADE = 1
DB_NAME = "estoque_v2.db"

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
                Vencimento DATE,
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
        c.execute('SELECT id, Codigo, Descricao, Quantidade, Vencimento, Preco_Custo, Preco_Venda FROM estoque')
        rows = c.fetchall()
    df = pd.DataFrame([dict(row) for row in rows])
    if not df.empty:
        df['Quantidade'] = df['Quantidade'].astype(int)
        df['Preco_Custo'] = df['Preco_Custo'].astype(float)
        df['Preco_Venda'] = df['Preco_Venda'].astype(float)
    else:
        df = pd.DataFrame(columns=['id', 'Codigo', 'Descricao', 'Quantidade', 'Vencimento', 'Preco_Custo', 'Preco_Venda'])
    return df

def buscar_descricao_por_codigo(cod):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('SELECT DISTINCT Descricao FROM estoque WHERE Codigo = ?', (cod,))
        result = c.fetchone()
    return result['Descricao'] if result else None

def gerar_template_xlsx():
    template_df = pd.DataFrame({
        'Codigo': ['ABC001', 'ABC002'],
        'Descricao': ['Arroz 5kg', 'Feijão 1kg'],
        'Quantidade': [100, 50],
        'Vencimento': ['2026-12-31', '2026-10-15'],
        'Preco_Custo': [20.00, 6.50],
        'Preco_Venda': [28.00, 9.00]
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        template_df.to_excel(writer, index=False, sheet_name='Inventario')
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
        msg = MIMEText(f"Solicitação:\nSOLICITANTE: {email_solicitante}\nAÇÃO: {descricao_acao}\nCÓDIGO: {codigo}")
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
        token_input = st.text_input("🔑 Código recebido:", key=f"inp_{chave}")
        if st.button("✅ Confirmar Execução", key=f"exec_{chave}"):
            if token_input == st.session_state[f"token_{chave}"]:
                st.session_state[f"token_{chave}"] = None
                return True
            else:
                st.error("⛔ Código incorreto!")
    return False

# --- CONTROLE DE ACESSO ---
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

st.sidebar.title("Navegação")
menu = st.sidebar.radio("Ir para:", ["📊 Consulta Estoque", "💰 Financeiro", "🔒 Almoxarifado (Entrada/Saída)"])
st.sidebar.divider()
st.sidebar.markdown(f"🟢 **{total_ativos}/{LIMITE_PESSOAS}** pessoas online")

hoje = date.today()

# ==========================================
# TELA 1: CONSULTA ESTOQUE
# ==========================================
if menu == "📊 Consulta Estoque":
    src1 = logo_para_base64("logo1.png")
    img1 = f'<img class="img-logo1" src="{src1}">' if src1 else '<span style="color:#102a43;font-weight:700;">LOGO 1</span>'

    df_ativos = df[df['Quantidade'] > 0].copy()
    total_pecas = f"{df_ativos['Quantidade'].sum():.0f}" if not df_ativos.empty else "0"
    total_itens = str(df_ativos['Codigo'].nunique())    if not df_ativos.empty else "0"

    components.html(f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link href="https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
      * {{ box-sizing: border-box; margin: 0; padding: 0; font-family: 'Sora', sans-serif; }}
      .header-container {{ display: flex; justify-content: space-between; align-items: center; padding: 20px 32px; border-radius: 16px; margin-bottom: 16px; background: linear-gradient(135deg, #f0f4f8 0%, #d9e2ec 100%); border: 1px solid #e2e8f0; }}
      .img-logo1 {{ height: 60px; object-fit: contain; }}
      .title-box h1 {{ font-size: 1.6rem; color: #102a43; }}
      .metrics-grid {{ display: flex; gap: 12px; margin-top: 4px; }}
      .metric-card {{ background: #fff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 16px 20px; width: 100%; }}
      .metric-label {{ font-size: 0.8rem; color: #718096; font-weight: 600; }}
      .metric-value {{ font-size: 1.8rem; font-weight: 700; color: #1a202c; }}
    </style>
    </head>
    <body>
    <div class="header-container">
        <div>{img1}</div>
        <div class="title-box"><h1>ESTOQUE GERAL</h1></div>
    </div>
    <div class="metrics-grid">
      <div class="metric-card"><div class="metric-label">📦 Total de Unidades</div><div class="metric-value">{total_pecas}</div></div>
      <div class="metric-card"><div class="metric-label">🏷️ Produtos Únicos</div><div class="metric-value">{total_itens}</div></div>
    </div>
    </body>
    </html>
    """, height=220)

    st.divider()
    busca = st.text_input("🔍 Pesquisar Código ou Descrição:")

    if not df_ativos.empty:
        # Lógica de Vencimento
        df_ativos['Vencimento_Date'] = pd.to_datetime(df_ativos['Vencimento'], errors='coerce').dt.date
        
        def classificar_vencimento(data_venc):
            if pd.isnull(data_venc): return '⚪ Sem data'
            dias_restantes = (data_venc - hoje).days
            if dias_restantes < 0: return '🔴 Vencido (Descartar)'
            elif dias_restantes <= 30: return f'🟡 Vence em {dias_restantes} dias'
            else: return '🟢 No Prazo'

        df_ativos['Status Vencimento'] = df_ativos['Vencimento_Date'].apply(classificar_vencimento)
        
        if busca:
            df_ativos = df_ativos[
                df_ativos['Codigo'].astype(str).str.contains(busca, case=False) |
                df_ativos['Descricao'].str.contains(busca, case=False, na=False)
            ]

        # Ordenar: Os mais próximos do vencimento primeiro (ou vencidos)
        df_ativos = df_ativos.sort_values(by='Vencimento_Date', ascending=True, na_position='last')

        df_exibicao = df_ativos[['Status Vencimento', 'Codigo', 'Descricao', 'Quantidade', 'Vencimento', 'Preco_Custo', 'Preco_Venda']].copy()
        df_exibicao['Vencimento'] = pd.to_datetime(df_exibicao['Vencimento']).dt.strftime('%d/%m/%Y')
        df_exibicao['Preco_Custo'] = df_exibicao['Preco_Custo'].apply(formatar_moeda)
        df_exibicao['Preco_Venda'] = df_exibicao['Preco_Venda'].apply(formatar_moeda)

        st.dataframe(df_exibicao, use_container_width=True, hide_index=True)
    else:
        st.info("O estoque está vazio.")

# ==========================================
# TELA 2: FINANCEIRO
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
        
        # Colorir operação para facilitar leitura
        def color_op(val):
            color = 'green' if val == 'Venda' else 'red'
            return f'color: {color}; font-weight: bold'
        
        st.dataframe(df_mov.style.map(color_op, subset=['Operação']), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma movimentação financeira registrada.")

# ==========================================
# TELA 3: ALMOXARIFADO (Operação)
# ==========================================
elif menu == "🔒 Almoxarifado (Entrada/Saída)":
    st.title("🔒 Área Operacional")
    senha = st.text_input("Senha de Acesso:", type="password")

    if senha == SENHA_ACESSO or senha == SENHA_ZERAR_ESTOQUE:
        abas_nomes = ["📥 Entrada de Itens", "📤 Saída de Itens", "🗑️ Descarte de Vencidos", "📦 Carga em Massa"]
        if senha == SENHA_ZERAR_ESTOQUE:
            abas_nomes.extend(["⚠️ Master"])

        abas = st.tabs(abas_nomes)

        # ---------------------------------------------------------
        # ABA 1: ENTRADA DE ITENS
        # ---------------------------------------------------------
        with abas[0]:
            st.subheader("Registrar Nova Entrada")
            with st.form("form_entrada", clear_on_submit=True):
                c1, c2 = st.columns([1, 2])
                cod = c1.text_input("Código do Produto:")
                desc_input = c2.text_input("Descrição (Apenas se for item novo):")
                
                c3, c4, c5 = st.columns([1, 1, 1])
                data_vencimento = c3.date_input("Data de Vencimento:", format="DD/MM/YYYY")
                qtd = c4.number_input("Quantidade (Entrada):", min_value=1, step=1, format="%d")
                
                st.markdown("**Valores Unitários (R$)**")
                c6, c7 = st.columns(2)
                preco_custo = c6.number_input("Preço de Custo", min_value=0.0, format="%.2f", step=0.50)
                preco_venda = c7.number_input("Preço de Venda", min_value=0.0, format="%.2f", step=0.50)

                if st.form_submit_button("✅ Registrar Entrada"):
                    if not cod:
                        st.error("⛔ Informe o Código.")
                    else:
                        venc_str = data_vencimento.strftime('%Y-%m-%d')
                        desc_existente = buscar_descricao_por_codigo(cod)
                        
                        if not desc_existente and not desc_input:
                            st.error("⛔ Produto novo! A Descrição é obrigatória.")
                        elif desc_existente and desc_input and desc_input.strip().lower() != desc_existente.strip().lower():
                            st.error(f"⛔ Conflito! O código **{cod}** já está registrado como: **{desc_existente}**")
                        else:
                            desc_final = desc_existente if desc_existente else desc_input
                            with get_conn() as conn:
                                cur = conn.cursor()
                                # Verifica se já existe o mesmo produto COM A MESMA DATA DE VENCIMENTO
                                cur.execute('SELECT id, Preco_Custo, Preco_Venda FROM estoque WHERE Codigo=? AND Vencimento=?', (cod, venc_str))
                                res = cur.fetchone()
                                
                                if res: # Lote já existe, atualiza qtd e sobrepõe preços se informados
                                    novo_custo = preco_custo if preco_custo > 0 else res['Preco_Custo']
                                    novo_venda = preco_venda if preco_venda > 0 else res['Preco_Venda']
                                    cur.execute('UPDATE estoque SET Quantidade = Quantidade + ?, Preco_Custo = ?, Preco_Venda = ? WHERE id=?', 
                                                (qtd, novo_custo, novo_venda, res['id']))
                                    st.success(f"✅ Quantidade adicionada ao lote existente (Vencimento: {data_vencimento.strftime('%d/%m/%Y')}).")
                                else: # Lote novo
                                    cur.execute('INSERT INTO estoque (Codigo, Descricao, Quantidade, Vencimento, Preco_Custo, Preco_Venda) VALUES (?, ?, ?, ?, ?, ?)', 
                                                (cod, desc_final, qtd, venc_str, preco_custo, preco_venda))
                                    st.success("✅ Novo lote/produto registrado com sucesso.")
                            st.cache_data.clear()

        # ---------------------------------------------------------
        # ABA 2: SAÍDA DE ITENS
        # ---------------------------------------------------------
        with abas[1]:
            st.subheader("Registrar Saída / Venda")
            df_disponivel = df[df['Quantidade'] > 0].copy()
            
            if df_disponivel.empty:
                st.warning("O estoque está zerado.")
            else:
                # Criar um dicionário de opções para facilitar a escolha do usuário
                df_disponivel['Vencimento_BR'] = pd.to_datetime(df_disponivel['Vencimento']).dt.strftime('%d/%m/%Y')
                opcoes = {}
                for _, row in df_disponivel.iterrows():
                    label = f"{row['Codigo']} - {row['Descricao']} | Lote: {row['Vencimento_BR']} | Disp: {row['Quantidade']} un"
                    opcoes[label] = row['id']
                
                with st.form("form_saida", clear_on_submit=True):
                    item_selecionado = st.selectbox("Selecione o Produto/Lote:", list(opcoes.keys()))
                    c1, c2 = st.columns(2)
                    op = c1.radio("Motivo da Saída:", ["Venda", "Perda (Avaria/Roubo)"])
                    qtd_saida = c2.number_input("Quantidade:", min_value=1, step=1, format="%d")
                    
                    if st.form_submit_button("📤 Confirmar Saída"):
                        item_id = opcoes[item_selecionado]
                        
                        with get_conn() as conn:
                            cur = conn.cursor()
                            cur.execute('SELECT Codigo, Descricao, Quantidade, Preco_Custo, Preco_Venda FROM estoque WHERE id=?', (item_id,))
                            row_db = cur.fetchone()
                            
                            if row_db['Quantidade'] < qtd_saida:
                                st.error("⛔ Quantidade solicitada é maior que o saldo em estoque deste lote.")
                            else:
                                tipo_operacao = "Venda" if op == "Venda" else "Perda"
                                valor_base = row_db['Preco_Venda'] if tipo_operacao == "Venda" else row_db['Preco_Custo']
                                valor_total = qtd_saida * valor_base
                                
                                cur.execute('UPDATE estoque SET Quantidade = Quantidade - ? WHERE id=?', (qtd_saida, item_id))
                                cur.execute('INSERT INTO financeiro (Codigo, Descricao, tipo, quantidade, valor_total) VALUES (?, ?, ?, ?, ?)', 
                                            (row_db['Codigo'], row_db['Descricao'], tipo_operacao, qtd_saida, valor_total))
                                
                                conn.commit()
                                st.success(f"✅ {tipo_operacao} registrada com sucesso!")
                                st.cache_data.clear()
                                st.rerun()

        # ---------------------------------------------------------
        # ABA 3: DESCARTE DE VENCIDOS
        # ---------------------------------------------------------
        with abas[2]:
            st.subheader("🚨 Orientação de Descarte (Produtos Vencidos)")
            st.write("Aqui estão listados todos os itens cujo prazo de validade expirou. O descarte lançará automaticamente o valor de custo no caixa de perdas.")
            
            df_vencidos = df[df['Quantidade'] > 0].copy()
            df_vencidos['Vencimento_Date'] = pd.to_datetime(df_vencidos['Vencimento'], errors='coerce').dt.date
            df_vencidos = df_vencidos[df_vencidos['Vencimento_Date'] < hoje]

            if df_vencidos.empty:
                st.success("Tudo certo! Não há nenhum produto vencido no estoque.")
            else:
                custo_total_perda = (df_vencidos['Quantidade'] * df_vencidos['Preco_Custo']).sum()
                st.error(f"⚠️ Atenção! Você possui {len(df_vencidos)} lote(s) vencido(s). Prejuízo estimado: {formatar_moeda(custo_total_perda)}.")
                
                df_show_vencidos = df_vencidos[['Codigo', 'Descricao', 'Quantidade', 'Vencimento', 'Preco_Custo']].copy()
                df_show_vencidos['Vencimento'] = pd.to_datetime(df_show_vencidos['Vencimento']).dt.strftime('%d/%m/%Y')
                df_show_vencidos['Preco_Custo'] = df_show_vencidos['Preco_Custo'].apply(formatar_moeda)
                st.dataframe(df_show_vencidos, use_container_width=True, hide_index=True)

                st.warning("A ação abaixo removerá do estoque TODOS os itens listados acima e lançará seus valores como PERDA.")
                if st.button("🗑️ Processar Descarte de Vencidos"):
                    with get_conn() as conn:
                        cur = conn.cursor()
                        for _, row_venc in df_vencidos.iterrows():
                            # Remove do estoque
                            cur.execute('UPDATE estoque SET Quantidade = 0 WHERE id=?', (row_venc['id'],))
                            # Lança a perda
                            valor_perda = row_venc['Quantidade'] * row_venc['Preco_Custo']
                            cur.execute('INSERT INTO financeiro (Codigo, Descricao, tipo, quantidade, valor_total) VALUES (?, ?, ?, ?, ?)', 
                                        (row_venc['Codigo'], f"{row_venc['Descricao']} (VENCIDO)", 'Perda', row_venc['Quantidade'], valor_perda))
                        conn.commit()
                    st.success("✅ Descarte processado. Valores lançados como perda.")
                    st.cache_data.clear()
                    st.rerun()

        # ---------------------------------------------------------
        # ABA 4: CARGA EM MASSA
        # ---------------------------------------------------------
        with abas[3]:
            st.info("Upload de arquivo Excel (.xlsx) com colunas: `Codigo` | `Descricao` | `Quantidade` | `Vencimento` (AAAA-MM-DD) | `Preco_Custo` | `Preco_Venda`")
            st.download_button("⬇️ Template Inventário", gerar_template_xlsx(), "template_inventario.xlsx")
            arquivo = st.file_uploader("Arquivo de Inventário (.xlsx):", type=["xlsx"], key="upload_massa")

            if arquivo:
                try:
                    df_upload = pd.read_excel(arquivo, engine='openpyxl')
                    faltando = {'Codigo', 'Descricao', 'Quantidade', 'Vencimento'} - set(df_upload.columns)
                    if faltando:
                        st.error(f"⛔ Colunas obrigatórias ausentes: {', '.join(faltando)}")
                    else:
                        if st.button("🚀 Processar Importação"):
                            df_upload['Codigo']     = df_upload['Codigo'].astype(str).str.strip()
                            df_upload['Descricao']  = df_upload['Descricao'].astype(str).str.strip()
                            df_upload['Quantidade'] = pd.to_numeric(df_upload['Quantidade'], errors='coerce')
                            
                            # Tratamento de Data
                            df_upload['Vencimento'] = pd.to_datetime(df_upload['Vencimento'], errors='coerce').dt.strftime('%Y-%m-%d')
                            
                            if 'Preco_Custo' not in df_upload.columns: df_upload['Preco_Custo'] = 0.0
                            if 'Preco_Venda' not in df_upload.columns: df_upload['Preco_Venda'] = 0.0
                            
                            df_upload['Preco_Custo'] = pd.to_numeric(df_upload['Preco_Custo'], errors='coerce').fillna(0.0)
                            df_upload['Preco_Venda'] = pd.to_numeric(df_upload['Preco_Venda'], errors='coerce').fillna(0.0)

                            df_upload = df_upload.dropna(subset=['Quantidade', 'Vencimento'])
                            df_upload = df_upload[df_upload['Quantidade'] > 0]
                            df_upload['Quantidade'] = df_upload['Quantidade'].astype(int)

                            with get_conn() as conn:
                                cur = conn.cursor()
                                cur.execute('SELECT id, Codigo, Vencimento FROM estoque')
                                db_list = cur.fetchall()
                                db_dict = {(r['Codigo'], r['Vencimento']): r['id'] for r in db_list}

                                inserts, updates = [], []
                                for _, row in df_upload.iterrows():
                                    cod_r, desc_r, venc_r, qtd_r = row['Codigo'], row['Descricao'], row['Vencimento'], row['Quantidade']
                                    cust_r, vend_r = row['Preco_Custo'], row['Preco_Venda']
                                    key = (cod_r, venc_r)

                                    if key in db_dict:
                                        updates.append((qtd_r, cust_r, vend_r, db_dict[key]))
                                    else:
                                        inserts.append((cod_r, desc_r, qtd_r, venc_r, cust_r, vend_r))
                                        # Atualiza dicionário para evitar duplicadas na mesma planilha
                                        db_dict[key] = -1 

                                if inserts:
                                    cur.executemany('INSERT INTO estoque (Codigo,Descricao,Quantidade,Vencimento,Preco_Custo,Preco_Venda) VALUES (?,?,?,?,?,?)', inserts)
                                if updates:
                                    cur.executemany('UPDATE estoque SET Quantidade = Quantidade + ?, Preco_Custo = MAX(Preco_Custo, ?), Preco_Venda = MAX(Preco_Venda, ?) WHERE id=?', updates)
                            conn.commit()
                            st.success(f"✅ Importação concluída! {len(inserts)} novos lotes, {len(updates)} atualizados.")
                            st.cache_data.clear()
                            st.rerun()
                except Exception as e:
                    st.error(f"Erro ao processar planilha: {e}")

        # ---------------------------------------------------------
        # ABA 5: MASTER
        # ---------------------------------------------------------
        if senha == SENHA_ZERAR_ESTOQUE:
            with abas[4]:
                st.subheader("⚠️ Área de Risco - Acesso Master")
                opcao = st.radio("Selecione a ação desejada:", [
                    "1️⃣ Apagar apenas as quantidades (Mantém cadastro de produtos)",
                    "2️⃣ Excluir Banco de Estoque (Limpa tudo, incluindo códigos)",
                    "3️⃣ Limpar Histórico Financeiro"
                ])

                if aprovar_acao_master("limpeza", f"Limpeza Master: {opcao}"):
                    with get_conn() as conn:
                        cur = conn.cursor()
                        if "1️⃣" in opcao:
                            cur.execute('UPDATE estoque SET Quantidade = 0')
                            st.success("Quantidades zeradas!")
                        elif "2️⃣" in opcao:
                            cur.execute("DELETE FROM estoque")
                            st.success("Estoque apagado do sistema!")
                        elif "3️⃣" in opcao:
                            cur.execute("DELETE FROM financeiro")
                            st.success("Histórico financeiro apagado!")
                        conn.commit()
                    st.cache_data.clear()
                    st.rerun()
