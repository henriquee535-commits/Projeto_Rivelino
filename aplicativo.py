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

# ==========================================
# CONFIGURAÇÃO DA PÁGINA
# ==========================================
st.set_page_config(page_title="Inventário José Rivelino", layout="wide", page_icon="📦")

SENHA_ACESSO        = st.secrets["SENHA_ACESSO"]
SENHA_ZERAR_ESTOQUE = st.secrets["SENHA_ZERAR_ESTOQUE"]
LIMITE_PESSOAS      = 20
TEMPO_INATIVIDADE   = 15
DB_NAME             = "almoxarifado.db"

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

# ==========================================
# BANCO DE DADOS
# ==========================================
def get_conn():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS estoque (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                Codigo       TEXT,
                Descricao    TEXT,
                Quantidade   INTEGER,
                Preco_Custo  REAL DEFAULT 0.00,
                Preco_Venda  REAL DEFAULT 0.00,
                Vencimento   TEXT
            )
        ''')
        c.execute('''
            CREATE TABLE IF NOT EXISTS acessos (
                sessao_id     TEXT PRIMARY KEY,
                ultimo_clique DATETIME
            )
        ''')
        c.execute('''
            CREATE TABLE IF NOT EXISTS financeiro (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                Codigo           TEXT,
                Descricao        TEXT,
                tipo             TEXT,
                quantidade       INTEGER,
                preco_custo_unit REAL DEFAULT 0,
                valor_total      REAL,
                data             DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        conn.commit()

init_db()

# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def formatar_moeda(valor):
    return f"R$ {valor:_.2f}".replace('.', ',').replace('_', '.')

def status_vencimento(dias):
    if dias is None:
        return "sem_vencimento", "—"
    if dias < 0:
        return "vencido", f"Vencido há {abs(dias)}d"
    if dias == 0:
        return "vencido", "Vence hoje!"
    if dias <= 30:
        return "critico", f"{dias}d restantes"
    if dias <= 90:
        return "atencao", f"{dias}d restantes"
    return "ok", f"{dias}d restantes"

def logo_para_base64(path):
    for t in [path, path.replace('.png','.jpg'), path.replace('.png','.jpeg')]:
        try:
            with open(t, "rb") as f:
                data = base64.b64encode(f.read()).decode()
            mime = 'image/png' if t.endswith('.png') else 'image/jpeg'
            return f"data:{mime};base64,{data}"
        except FileNotFoundError:
            continue
    return None

def buscar_item_por_codigo(cod):
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('SELECT DISTINCT Descricao FROM estoque WHERE Codigo = ?', (cod,))
        r = c.fetchone()
    return r['Descricao'] if r else None

def gerar_excel_download(dataframe, nome_aba="Estoque"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        dataframe.to_excel(w, index=False, sheet_name=nome_aba)
    return buf.getvalue()

def gerar_template_xlsx():
    df_temp = pd.DataFrame({
        'Codigo':['ABC001','ABC002'],'Descricao':['Arroz 5kg','Feijão 1kg'],
        'Quantidade':[100,50],'Preco_Custo':[20.50,7.20],'Preco_Venda':[25.00,9.50],
        'Vencimento':['31/12/2026','30/06/2026'],
    })
    return gerar_excel_download(df_temp, "Template Importação")

# ==========================================
# PROCESSAR VENCIDOS AUTOMÁTICO
# ==========================================
def processar_vencidos_automatico():
    hoje_iso = date.today().isoformat()
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('''
            SELECT id, Codigo, Descricao, Quantidade, Preco_Custo FROM estoque
            WHERE Vencimento IS NOT NULL AND Vencimento != ''
              AND Vencimento < ? AND Quantidade > 0
        ''', (hoje_iso,))
        vencidos = c.fetchall()
        for item in vencidos:
            vp = item['Quantidade'] * item['Preco_Custo']
            c.execute('''
                INSERT INTO financeiro (Codigo,Descricao,tipo,quantidade,preco_custo_unit,valor_total)
                VALUES (?,?,'Perda',?,?,?)
            ''', (item['Codigo'], item['Descricao'], item['Quantidade'], item['Preco_Custo'], vp))
            c.execute('UPDATE estoque SET Quantidade=0 WHERE id=?', (item['id'],))
        conn.commit()
    return vencidos

# ==========================================
# CACHE DE DADOS
# ==========================================
@st.cache_data(ttl=300)
def carregar_estoque():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('SELECT Codigo,Descricao,Quantidade,Preco_Custo,Preco_Venda,Vencimento FROM estoque')
        rows = c.fetchall()
    df_e = pd.DataFrame([dict(r) for r in rows],
                      columns=['Codigo','Descricao','Quantidade','Preco_Custo','Preco_Venda','Vencimento'])
    if not df_e.empty:
        df_e['Quantidade']  = df_e['Quantidade'].astype(int)
        df_e['Preco_Custo'] = df_e['Preco_Custo'].astype(float)
        df_e['Preco_Venda'] = df_e['Preco_Venda'].astype(float)
        df_e['Vencimento']  = pd.to_datetime(df_e['Vencimento'], errors='coerce')
    return df_e

@st.cache_data(ttl=60)
def carregar_financeiro():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('SELECT id,Codigo,Descricao,tipo,quantidade,preco_custo_unit,valor_total,data FROM financeiro ORDER BY data DESC')
        rows = c.fetchall()
    df_f = pd.DataFrame([dict(r) for r in rows])
    if not df_f.empty:
        df_f['data']             = pd.to_datetime(df_f['data'])
        df_f['valor_total']      = df_f['valor_total'].astype(float)
        df_f['preco_custo_unit'] = df_f['preco_custo_unit'].astype(float)
        df_f['quantidade']       = df_f['quantidade'].astype(int)
    return df_f

# ==========================================
# APROVAÇÃO MASTER POR GMAIL
# ==========================================
def aprovar_acao_master(chave, descricao_acao):
    if f"token_{chave}" not in st.session_state:
        st.session_state[f"token_{chave}"] = None
    email_sol = st.text_input("📧 Seu e-mail:", key=f"email_{chave}", placeholder="exemplo@gmail.com")
    if st.button(f"📩 Solicitar Liberação: {descricao_acao}", key=f"req_{chave}"):
        if not email_sol:
            st.error("⛔ Informe seu e-mail."); return False
        codigo = str(random.randint(100000, 999999))
        st.session_state[f"token_{chave}"] = codigo
        try:
            # Configurado para Gmail
            rem  = st.secrets["email"]["remetente"]
            pwd  = st.secrets["email"]["senha"] # Deve ser uma 'Senha de App' do Google
            dest = st.secrets["email"]["destinatario"]
            
            msg  = MIMEText(f"Solicitante: {email_sol}\nAção: {descricao_acao}\nCódigo: {codigo}")
            msg['Subject'] = 'Aprovação - Sistema Supermercado'
            msg['From'] = rem; msg['To'] = dest
            
            with smtplib.SMTP('smtp.gmail.com', 587) as s:
                s.starttls()
                s.login(rem, pwd)
                s.sendmail(rem, [dest], msg.as_string())
            st.info("✅ Solicitação enviada via Gmail!")
        except Exception as e:
            st.error(f"Erro ao enviar e-mail: {e}. Verifique se a 'Senha de App' está correta nos Secrets.")
    
    if st.session_state[f"token_{chave}"]:
        tok = st.text_input("🔑 Código de Liberação:", key=f"inp_{chave}")
        if st.button("✅ Confirmar Execução", key=f"exec_{chave}"):
            if tok == st.session_state[f"token_{chave}"]:
                st.session_state[f"token_{chave}"] = None; return True
            else:
                st.error("⛔ Código incorreto!")
    return False

# ==========================================
# CONTROLE DE SESSÃO
# ==========================================
if 'sessao_id' not in st.session_state:
    st.session_state.sessao_id = str(uuid.uuid4())

with get_conn() as conn:
    c = conn.cursor()
    lim = (datetime.now() - timedelta(minutes=TEMPO_INATIVIDADE)).strftime('%Y-%m-%d %H:%M:%S')
    c.execute("DELETE FROM acessos WHERE ultimo_clique < ?", (lim,))
    agora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("INSERT INTO acessos (sessao_id,ultimo_clique) VALUES (?,?) ON CONFLICT(sessao_id) DO UPDATE SET ultimo_clique=excluded.ultimo_clique",
              (st.session_state.sessao_id, agora))
    c.execute("SELECT COUNT(*) as total FROM acessos")
    total_ativos = c.fetchone()['total']
    conn.commit()

if total_ativos > LIMITE_PESSOAS:
    st.error(f"⚠️ Sistema lotado ({total_ativos}/{LIMITE_PESSOAS}). Tente em alguns minutos.")
    st.stop()

# Auto-processamento
vencidos_proc = processar_vencidos_automatico()
if vencidos_proc:
    st.cache_data.clear()

df_global = carregar_estoque()

# ==========================================
# NAVEGAÇÃO
# ==========================================
st.sidebar.title("Navegação")
menu = st.sidebar.radio("Ir para:", ["📊 Consulta", "📥 Entrada", "📤 Saída", "🔒 Administrativo", "🔒 Financeiro"])
st.sidebar.divider()
st.sidebar.markdown(f"🟢 **{total_ativos}/{LIMITE_PESSOAS}** pessoas online")

# ==========================================
# TELA 1 — CONSULTA
# ==========================================
if menu == "📊 Consulta":
    src1 = logo_para_base64("logo1.png")
    img1 = f'<img class="il1" src="{src1}">' if src1 else '<span style="color:#102a43;font-weight:700;">LOGO 1</span>'

    hoje = date.today()
    df_ativos = df_global[df_global['Quantidade'] > 0].copy()
    
    tp = f"{df_ativos['Quantidade'].sum():.0f}" if not df_ativos.empty else "0"
    ti = str(df_ativos['Codigo'].nunique())     if not df_ativos.empty else "0"
    vc = cr = 0
    if not df_ativos.empty:
        df_ativos['_d'] = df_ativos['Vencimento'].apply(lambda v: (v.date()-hoje).days if pd.notna(v) else None)
        vc = int((df_ativos['_d'].dropna() < 0).sum())
        cr = int(((df_ativos['_d'].dropna() >= 0) & (df_ativos['_d'].dropna() <= 30)).sum())

    components.html(f"""<!DOCTYPE html><html><head>
    <link href="https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
      *{{box-sizing:border-box;margin:0;padding:0;font-family:'Sora',sans-serif;}}
      .hdr{{display:grid;grid-template-columns:auto 1fr;align-items:center;padding:20px 32px;border-radius:16px;margin-bottom:16px;background:linear-gradient(135deg,#f0f4f8,#d9e2ec);box-shadow:0 4px 12px rgba(0,0,0,.05);border:1px solid #e2e8f0;}}
      .il1{{height:85px;width:auto;max-width:240px;object-fit:contain;mix-blend-mode:darken;}}
      .tb{{text-align:left;padding-left:20px;}}
      .tb h1{{font-size:1.8rem;font-weight:700;color:#102a43;}}
      .mg{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:4px;}}
      .mc{{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:16px 20px;}}
      .ml{{font-size:.78rem;color:#718096;font-weight:600;}}
      .mv{{font-size:1.9rem;font-weight:700;color:#1a202c;}}
      @media(max-width:768px){{.hdr{{grid-template-columns:1fr;text-align:center;}}.mg{{grid-template-columns:repeat(2,1fr);}}}}
    </style></head><body>
    <div class="hdr"><div>{img1}</div><div class="tb"><h1>ESTOQUE ATUAL</h1><p>GERENCIAMENTO DE VALIDADE</p></div></div>
    <div class="mg">
      <div class="mc"><div class="ml">📦 Total Unidades</div><div class="mv">{tp}</div></div>
      <div class="mc"><div class="ml">🏷️ Itens Únicos</div><div class="mv">{ti}</div></div>
      <div class="mc" style="border-left:5px solid orange;"><div class="ml">⚠️ Vence em 30 dias</div><div class="mv" style="color:orange;">{cr}</div></div>
      <div class="mc" style="border-left:5px solid red;"><div class="ml">🔴 Vencidos</div><div class="mv" style="color:red;">{vc}</div></div>
    </div></body></html>""", height=300)

    st.divider()
    
    col_s1, col_s2 = st.columns([3, 1])
    busca = col_s1.text_input("🔍 Pesquisar Produto:")
    
    df_f = df_ativos.copy()
    if busca:
        df_f = df_f[df_f['Codigo'].astype(str).str.contains(busca, case=False) |
                    df_f['Descricao'].str.contains(busca, case=False, na=False)]

    if not df_f.empty:
        # Preparação para download ANTES de formatar para o HTML
        df_export = df_f.copy()
        df_export['Vencimento'] = df_export['Vencimento'].dt.strftime('%d/%m/%Y')
        excel_data = gerar_excel_download(df_export, "Estoque_Consultado")
        
        col_s2.markdown("<br>", unsafe_allow_html=True)
        col_s2.download_button("📥 Baixar Excel", excel_data, "Estoque_Filtrado.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Exibição Visual (HTML)
        df_f['_dias'] = df_f['Vencimento'].apply(lambda v: (v.date()-hoje).days if pd.notna(v) else 99999)
        df_f = df_f.sort_values('_dias')
        
        def linha_html(row):
            d = None if row['_dias']==99999 else row['_dias']
            st_class, lbl = status_vencimento(d)
            v_formatada = row['Vencimento'].strftime('%d/%m/%Y') if pd.notna(row['Vencimento']) else '—'
            
            # Cores dinâmicas
            bg = "#fff"
            if st_class == 'vencido': bg = "#fff5f5"
            elif st_class == 'critico': bg = "#fffaf0"
            
            return (f'<tr style="background:{bg};"><td>{row["Codigo"]}</td><td>{row["Descricao"]}</td>'
                    f'<td style="text-align:center;">{row["Quantidade"]}</td>'
                    f'<td style="text-align:center;">{formatar_moeda(row["Preco_Venda"])}</td>'
                    f'<td style="text-align:center; font-weight:700;">{v_formatada}</td>'
                    f'<td style="text-align:center; font-size:0.8rem;">{lbl}</td></tr>')

        linhas = "\n".join(df_f.apply(linha_html, axis=1))
        components.html(f"""
        <style>
          .t{{width:100%; border-collapse:collapse; font-family:'Sora',sans-serif; font-size:0.9rem;}}
          .t th{{background:#1a3a4a; color:#fff; padding:12px; text-align:left;}}
          .t td{{padding:10px; border-bottom:1px solid #eee;}}
        </style>
        <table class="t"><thead><tr>
          <th>Código</th><th>Descrição</th><th>Qtd</th><th>Venda</th><th>Vencimento</th><th>Status</th>
        </tr></thead><tbody>{linhas}</tbody></table>""", height=500, scrolling=True)
    else:
        st.info("Nenhum item em estoque para exibir.")

# ==========================================
# TELA 2 — ENTRADA
# ==========================================
elif menu == "📥 Entrada":
    st.title("📥 Entrada de Estoque")
    abas = st.tabs(["Individual", "Carga em Massa"])

    with abas[0]:
        with st.form("f_ent", clear_on_submit=True):
            c1, c2 = st.columns(2)
            cod = c1.text_input("Código:")
            desc = c2.text_input("Descrição:")
            venc = st.date_input("Vencimento:", value=date.today()+timedelta(days=365))
            
            c3, c4, c5 = st.columns(3)
            pc = c3.number_input("Custo Unitário", min_value=0.0)
            pv = c4.number_input("Venda Unitário", min_value=0.0)
            qt = c5.number_input("Quantidade", min_value=1)

            if st.form_submit_button("Confirmar Entrada"):
                # Salvamos no banco no formato ISO (AAAA-MM-DD) para lógica, mas no Excel aceitamos DD/MM/AAAA
                v_iso = venc.isoformat()
                with get_conn() as conn:
                    cur = conn.cursor()
                    cur.execute('SELECT id FROM estoque WHERE Codigo=? AND Vencimento=?', (cod, v_iso))
                    res = cur.fetchone()
                    if res:
                        cur.execute('UPDATE estoque SET Quantidade=Quantidade+?, Preco_Custo=?, Preco_Venda=? WHERE id=?', (qt, pc, pv, res['id']))
                    else:
                        cur.execute('INSERT INTO estoque (Codigo,Descricao,Quantidade,Preco_Custo,Preco_Venda,Vencimento) VALUES (?,?,?,?,?,?)', (cod, desc, qt, pc, pv, v_iso))
                    conn.commit()
                st.success("Registrado!")
                st.cache_data.clear()

    with abas[1]:
        st.download_button("⬇️ Baixar Template Excel (Modelo)", gerar_template_xlsx(), "modelo_importacao.xlsx")
        arq = st.file_uploader("Subir planilha preenchida:", type="xlsx")
        if arq and st.button("🚀 Importar Agora"):
            try:
                du = pd.read_excel(arq)
                # Converte datas brasileiras para ISO antes de salvar
                du['Vencimento'] = pd.to_datetime(du['Vencimento'], dayfirst=True).dt.strftime('%Y-%m-%d')
                with get_conn() as conn:
                    cur = conn.cursor()
                    for _, r in du.iterrows():
                        cur.execute('INSERT INTO estoque (Codigo,Descricao,Quantidade,Preco_Custo,Preco_Venda,Vencimento) VALUES (?,?,?,?,?,?)',
                                   (str(r['Codigo']), r['Descricao'], r['Quantidade'], r['Preco_Custo'], r['Preco_Venda'], r['Vencimento']))
                    conn.commit()
                st.success("Importação concluída!")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Erro na data ou formato: {e}")

# ==========================================
# TELA 3 — SAÍDA
# ==========================================
elif menu == "📤 Saída":
    st.title("📤 Saída (Venda ou Perda)")
    with st.form("f_sai", clear_on_submit=True):
        c1, c2 = st.columns(2)
        cod = c1.text_input("Código do Produto:")
        qtd = c2.number_input("Quantidade:", min_value=1)
        tipo = st.selectbox("Tipo de Saída:", ["Venda", "Perda"])
        
        if st.form_submit_button("Confirmar Saída"):
            with get_conn() as conn:
                cur = conn.cursor()
                # Pega o lote que vence primeiro (FIFO)
                cur.execute('SELECT id, Preco_Custo, Preco_Venda, Descricao FROM estoque WHERE Codigo=? AND Quantidade >= ? ORDER BY Vencimento ASC LIMIT 1', (cod, qtd))
                item = cur.fetchone()
                if item:
                    valor = item['Preco_Venda'] if tipo == "Venda" else item['Preco_Custo']
                    cur.execute('UPDATE estoque SET Quantidade = Quantidade - ? WHERE id = ?', (qtd, item['id']))
                    cur.execute('INSERT INTO financeiro (Codigo, Descricao, tipo, quantidade, preco_custo_unit, valor_total) VALUES (?,?,?,?,?,?)',
                               (cod, item['Descricao'], tipo, qtd, item['Preco_Custo'], qtd * valor))
                    conn.commit()
                    st.success("Saída efetuada!")
                    st.cache_data.clear()
                else:
                    st.error("Produto não encontrado ou estoque insuficiente no lote mais antigo.")

# ==========================================
# ADMINISTRATIVO & FINANCEIRO (Mantidos com Lógica de Senha)
# ==========================================
elif menu == "🔒 Administrativo":
    st.title("🔒 Administrativo")
    if st.text_input("Senha Master:", type="password") == SENHA_ZERAR_ESTOQUE:
        if st.button("🗑️ Zerar Todo o Estoque"):
            if aprovar_acao_master("limpar", "Zerar tudo"):
                with get_conn() as conn:
                    conn.cursor().execute("DELETE FROM estoque")
                    conn.commit()
                st.success("Banco limpo.")
                st.cache_data.clear()

elif menu == "🔒 Financeiro":
    st.title("🔒 Financeiro")
    if st.text_input("Senha Financeira:", type="password") == SENHA_ACESSO:
        df_fin = carregar_financeiro()
        if not df_fin.empty:
            # Formatação de data na tabela financeira para o usuário
            df_display = df_fin.copy()
            df_display['data'] = df_display['data'].dt.strftime('%d/%m/%Y %H:%M')
            st.dataframe(df_display, use_container_width=True)
            
            total_vendas = df_fin[df_fin['tipo']=='Venda']['valor_total'].sum()
            st.metric("Total de Vendas no Período", formatar_moeda(total_vendas))
