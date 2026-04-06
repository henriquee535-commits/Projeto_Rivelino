import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import psycopg2
from psycopg2.extras import RealDictCursor
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
DATABASE_URL = st.secrets["DATABASE_URL"]
LIMITE_PESSOAS = 40
TEMPO_INATIVIDADE = 1

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

# --- CONEXÃO COM SUPABASE ---
def get_conn():
    return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)

# --- BANCO DE DADOS ---
def init_db():
    with get_conn() as conn:
        with conn.cursor() as c:
            c.execute('''
                CREATE TABLE IF NOT EXISTS estoque (
                    id SERIAL PRIMARY KEY,
                    "Codigo" TEXT,
                    "Descricao" TEXT,
                    "Quantidade" INTEGER,
                    "CC" TEXT
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS acessos (
                    sessao_id TEXT PRIMARY KEY,
                    ultimo_clique TIMESTAMP
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS centros_custo (
                    nome TEXT PRIMARY KEY
                )
            ''')
        conn.commit()

init_db()

@st.cache_data(ttl=300)
def carregar_estoque():
    with get_conn() as conn:
        with conn.cursor() as c:
            c.execute('SELECT "Codigo", "Descricao", "Quantidade", "CC" FROM estoque')
            rows = c.fetchall()
    df = pd.DataFrame(rows, columns=['Codigo', 'Descricao', 'Quantidade', 'CC'])
    if not df.empty:
        df['Quantidade'] = df['Quantidade'].astype(int)
    return df

@st.cache_data
def carregar_ccs():
    with get_conn() as conn:
        with conn.cursor() as c:
            c.execute("SELECT nome FROM centros_custo ORDER BY nome")
            rows = c.fetchall()
    lista_cc = [r['nome'] for r in rows]
    if not lista_cc:
        lista_cc = ["Centro de Custo Geral"]
        with get_conn() as conn:
            with conn.cursor() as c:
                c.execute("INSERT INTO centros_custo (nome) VALUES (%s) ON CONFLICT DO NOTHING", ("Centro de Custo Geral",))
            conn.commit()
    return lista_cc

def buscar_descricao_por_codigo(cod):
    with get_conn() as conn:
        with conn.cursor() as c:
            c.execute('SELECT DISTINCT "Descricao" FROM estoque WHERE "Codigo" = %s', (cod,))
            result = c.fetchone()
    return result['Descricao'] if result else None

def gerar_template_xlsx():
    template_df = pd.DataFrame({
        'Codigo': ['ABC001', 'ABC002'],
        'Descricao': ['Parafuso M8', 'Cabo Elétrico 2,5mm'],
        'Quantidade': [100, 50],
        'CC': ['01/0001 - LIVRE DEMANDA', '01/0001 - LIVRE DEMANDA']
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

# --- SISTEMA DE APROVAÇÃO POR E-MAIL ---
def aprovar_acao_master(chave, descricao_acao):
    if f"token_{chave}" not in st.session_state:
        st.session_state[f"token_{chave}"] = None

    email_solicitante = st.text_input(
        "📧 Seu e-mail (para identificação):",
        key=f"email_{chave}",
        placeholder="seunome@gmail.com"
    )

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
        token_input = st.text_input("🔑 Código enviado para Eduardo Sousa - Controladoria, solicite a ele. Código:", key=f"inp_{chave}")
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
    with conn.cursor() as c:
        tempo_limite = datetime.now() - timedelta(minutes=TEMPO_INATIVIDADE)
        c.execute("DELETE FROM acessos WHERE ultimo_clique < %s", (tempo_limite,))
        c.execute("""
            INSERT INTO acessos (sessao_id, ultimo_clique) VALUES (%s, %s)
            ON CONFLICT (sessao_id) DO UPDATE SET ultimo_clique = EXCLUDED.ultimo_clique
        """, (st.session_state.sessao_id, datetime.now()))
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
menu = st.sidebar.radio("Ir para:", ["📊 Consulta", "🔒 Almoxarifado"])
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
      
      /* AJUSTE PARA MOBILE */
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

    st.dataframe(df_filt, use_container_width=True, hide_index=True)

# ==========================================
# TELA 2: ALMOXARIFADO
# ==========================================
else:
    st.title("🔒 Área Restrita — Almoxarifado")
    senha = st.text_input("Senha:", type="password")

    if senha == SENHA_ACESSO or senha == SENHA_ZERAR_ESTOQUE:
        abas_nomes = ["📝 Registro Individual", "📤 Carga em Massa"]
        if senha == SENHA_ZERAR_ESTOQUE:
            abas_nomes.extend(["🗑️ Excluir Item (Master)", "🏢 Gerenciar CCs (Master)", "⚠️ Limpar Dados (Master)"])

        abas = st.tabs(abas_nomes)

        # TAB 1: REGISTRO INDIVIDUAL
        with abas[0]:
            with st.form("registro", clear_on_submit=True):
                c1, c2 = st.columns(2)
                cod        = c1.text_input("Código:")
                desc_input = c2.text_input("Descrição (somente para itens novos):")
                c3, c4, c5 = st.columns([2, 2, 1])
                cc_sel = c3.selectbox("Centro de Custo:", lista_cc)
                op     = c4.selectbox("Operação:", ["Entrada", "Saída"])
                qtd    = c5.number_input("Qtd:", min_value=1, step=1, format="%d")

                if st.form_submit_button("✅ Confirmar"):
                    if not cod:
                        st.error("⛔ Informe o Código do item.")
                    else:
                        desc_existente = buscar_descricao_por_codigo(cod)
                        if not desc_existente and not desc_input:
                            st.error("⛔ A Descrição é OBRIGATÓRIA para cadastrar um novo item.")
                        elif desc_existente and desc_input and desc_input.strip() != desc_existente.strip():
                            st.error(f"⛔ Conflito! O código **{cod}** já está cadastrado como:\n\n**\"{desc_existente}\"**")
                        else:
                            desc_final = desc_existente if desc_existente else desc_input
                            with get_conn() as conn:
                                with conn.cursor() as cur:
                                    cur.execute('SELECT "Quantidade" FROM estoque WHERE "Codigo"=%s AND "CC"=%s', (cod, cc_sel))
                                    res = cur.fetchone()
                                    if res:
                                        if op == "Saída":
                                            if res['Quantidade'] < qtd:
                                                st.error(f"⛔ FALTA DE ESTOQUE! Saldo atual: {res['Quantidade']} unidades.")
                                            else:
                                                cur.execute('UPDATE estoque SET "Quantidade" = "Quantidade" - %s WHERE "Codigo"=%s AND "CC"=%s', (qtd, cod, cc_sel))
                                                st.success(f"✅ Saída registrada. Saldo: {res['Quantidade'] - qtd}")
                                                st.cache_data.clear()
                                        else:
                                            cur.execute('UPDATE estoque SET "Quantidade" = "Quantidade" + %s WHERE "Codigo"=%s AND "CC"=%s', (qtd, cod, cc_sel))
                                            st.success(f"✅ Entrada registrada. Saldo: {res['Quantidade'] + qtd}")
                                            st.cache_data.clear()
                                    else:
                                        if op == "Saída":
                                            st.error("⛔ ITEM NÃO ENCONTRADO neste Centro de Custo.")
                                        else:
                                            cur.execute('INSERT INTO estoque ("Codigo", "Descricao", "Quantidade", "CC") VALUES (%s, %s, %s, %s)', (cod, desc_final, qtd, cc_sel))
                                            st.success("✅ Item novo cadastrado com sucesso.")
                                            st.cache_data.clear()
                                conn.commit()

        # TAB 2: CARGA EM MASSA
        with abas[1]:
            st.info("Upload de arquivo Excel (.xlsx) com colunas: `Codigo` | `Descricao` | `Quantidade` | `CC`")
            st.download_button("⬇️ Template Inventário", gerar_template_xlsx(), "template_inventario.xlsx")
            arquivo = st.file_uploader("Arquivo de Inventário (.xlsx):", type=["xlsx"], key="upload_massa")

            if arquivo:
                try:
                    df_upload = pd.read_excel(arquivo, engine='openpyxl')
                    faltando = {'Codigo', 'Descricao', 'Quantidade', 'CC'} - set(df_upload.columns)
                    if faltando:
                        st.error(f"⛔ Colunas ausentes: {', '.join(faltando)}")
                    else:
                        if st.button("🚀 Processar Importação"):
                            df_upload['Codigo']     = df_upload['Codigo'].astype(str).str.strip()
                            df_upload['Descricao']  = df_upload['Descricao'].astype(str).str.strip()
                            df_upload['CC']         = df_upload['CC'].astype(str).str.strip()
                            df_upload['Quantidade'] = pd.to_numeric(df_upload['Quantidade'], errors='coerce')
                            df_upload = df_upload.dropna(subset=['Quantidade'])
                            df_upload = df_upload[df_upload['Quantidade'] > 0]
                            df_upload['Quantidade'] = df_upload['Quantidade'].astype(int)
                            df_upload = df_upload[(df_upload['Codigo'] != 'nan') & (df_upload['Codigo'] != '')]

                            ccs_invalidos = set(df_upload['CC'].unique()) - set(lista_cc)
                            if ccs_invalidos:
                                st.error(f"⛔ IMPORTAÇÃO BLOQUEADA! CCs não encontrados: **{', '.join(ccs_invalidos)}**.")
                            else:
                                with get_conn() as conn:
                                    with conn.cursor() as cur:
                                        cur.execute('SELECT "Codigo", "CC" FROM estoque')
                                        db_set = set((r['Codigo'], r['CC']) for r in cur.fetchall())

                                        inserts, updates = [], []
                                        for _, row in df_upload.iterrows():
                                            cod_r  = row['Codigo']
                                            desc_r = row['Descricao']
                                            cc_r   = row['CC']
                                            qtd_r  = row['Quantidade']

                                            if (cod_r, cc_r) not in db_set and (not desc_r or desc_r.lower() == 'nan'):
                                                continue
                                            if (cod_r, cc_r) in db_set:
                                                updates.append((qtd_r, cod_r, cc_r))
                                            else:
                                                inserts.append((cod_r, desc_r, qtd_r, cc_r))
                                                db_set.add((cod_r, cc_r))

                                        if inserts:
                                            cur.executemany('INSERT INTO estoque ("Codigo","Descricao","Quantidade","CC") VALUES (%s,%s,%s,%s)', inserts)
                                        if updates:
                                            cur.executemany('UPDATE estoque SET "Quantidade" = "Quantidade" + %s WHERE "Codigo"=%s AND "CC"=%s', updates)
                                conn.commit()

                                st.success(f"✅ Importação concluída! {len(inserts)} novos, {len(updates)} atualizados.")
                                st.cache_data.clear()
                                st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")

        # ÁREA MASTER
        if senha == SENHA_ZERAR_ESTOQUE:
            with abas[2]:
                st.subheader("🗑️ Excluir Item do Banco")
                st.warning("Esta ação apagará o código de todos os CCs.")
                cod_excluir = st.text_input("Digite o Código do item que deseja apagar:")
                if cod_excluir and aprovar_acao_master("del_item", f"Excluir código {cod_excluir}"):
                    with get_conn() as conn:
                        with conn.cursor() as cur:
                            cur.execute('SELECT * FROM estoque WHERE "Codigo"=%s', (cod_excluir,))
                            if cur.fetchone():
                                cur.execute('DELETE FROM estoque WHERE "Codigo"=%s', (cod_excluir,))
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
                            with conn.cursor() as cur:
                                cur.execute("INSERT INTO centros_custo (nome) VALUES (%s) ON CONFLICT DO NOTHING", (novo_cc,))
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
                            with conn.cursor() as cur:
                                cur.execute("INSERT INTO centros_custo (nome) VALUES (%s) ON CONFLICT DO NOTHING", (cc_novo,))
                                cur.execute('UPDATE estoque SET "CC" = %s WHERE "CC" = %s', (cc_novo, cc_antigo))
                                cur.execute("DELETE FROM centros_custo WHERE nome = %s", (cc_antigo,))
                            conn.commit()
                        st.success("Centro de Custo renomeado!")
                        st.cache_data.clear()
                        st.rerun()

                st.divider()
                st.subheader("📂 De/Para em Massa")
                st.download_button("⬇️ Template De/Para", gerar_template_depara(), "template_depara.xlsx")
                arq_depara = st.file_uploader("Arquivo De/Para (.xlsx):", type=["xlsx"])

                if arq_depara and aprovar_acao_master("depara_massa", "Processar De/Para em massa"):
                    df_dp = pd.read_excel(arq_depara)
                    if 'De' in df_dp.columns and 'Para' in df_dp.columns:
                        with get_conn() as conn:
                            with conn.cursor() as cur:
                                for _, row in df_dp.iterrows():
                                    de   = str(row['De']).strip()
                                    para = str(row['Para']).strip()
                                    if de != 'nan' and para != 'nan':
                                        cur.execute("INSERT INTO centros_custo (nome) VALUES (%s) ON CONFLICT DO NOTHING", (para,))
                                        cur.execute('UPDATE estoque SET "CC" = %s WHERE "CC" = %s', (para, de))
                                        cur.execute("DELETE FROM centros_custo WHERE nome = %s", (de,))
                            conn.commit()
                        st.success("De/Para em massa concluído!")
                        st.cache_data.clear()
                        st.rerun()
                
                # --- NOVA SESSÃO: INCLUSÃO EM MASSA DE CCs ---
                st.divider()
                st.subheader("➕ Inclusão em Massa de Centros de Custo")
                ccs_massa = st.text_area("Cole a lista de Centros de Custo (um por linha):")
                if ccs_massa and aprovar_acao_master("add_cc_massa", "Adicionar CCs em Massa"):
                    novos_ccs = [c.strip() for c in ccs_massa.split('\n') if c.strip()]
                    with get_conn() as conn:
                        with conn.cursor() as cur:
                            for cc in novos_ccs:
                                cur.execute("INSERT INTO centros_custo (nome) VALUES (%s) ON CONFLICT DO NOTHING", (cc,))
                        conn.commit()
                    st.success(f"✅ {len(novos_ccs)} Centros de Custo processados com sucesso!")
                    st.cache_data.clear()
                    st.rerun()

            with abas[4]:
                st.subheader("⚠️ Área de Risco - Acesso Master")
                opcao = st.radio("Selecione a ação desejada:", [
                    "1️⃣ Apenas zerar o estoque (Mantém os códigos salvos)",
                    "2️⃣ Excluir tudo (Limpa o banco de estoque e códigos)"
                ])

                if aprovar_acao_master("limpeza", f"Limpeza de Banco: {opcao}"):
                    with get_conn() as conn:
                        with conn.cursor() as cur:
                            if "1️⃣" in opcao:
                                cur.execute('UPDATE estoque SET "Quantidade" = 0')
                                st.success("Quantidades zeradas com sucesso!")
                            else:
                                cur.execute("DELETE FROM estoque")
                                st.success("Todos os itens apagados!")
                        conn.commit()
                    st.cache_data.clear()
                    st.rerun()
