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
LIMITE_PESSOAS      = 40
TEMPO_INATIVIDADE   = 1
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
        for col, typedef in [("Vencimento","TEXT"),("Preco_Custo","REAL DEFAULT 0"),("Preco_Venda","REAL DEFAULT 0")]:
            try:
                c.execute(f"ALTER TABLE estoque ADD COLUMN {col} {typedef}")
            except sqlite3.OperationalError:
                pass

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
        try:
            c.execute("ALTER TABLE financeiro ADD COLUMN preco_custo_unit REAL DEFAULT 0")
        except sqlite3.OperationalError:
            pass
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

def gerar_template_xlsx():
    df = pd.DataFrame({
        'Codigo':['ABC001','ABC002'],'Descricao':['Parafuso M8','Cabo Elétrico 2,5mm'],
        'Quantidade':[100,50],'Preco_Custo':[0.50,2.50],'Preco_Venda':[1.00,5.00],
        'Vencimento':['2025-12-31','2026-06-30'],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='Inventario')
    return buf.getvalue()

# ==========================================
# PROCESSAR VENCIDOS AUTOMÁTICO
# ==========================================
def processar_vencidos_automatico():
    hoje = date.today().isoformat()
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('''
            SELECT id, Codigo, Descricao, Quantidade, Preco_Custo FROM estoque
            WHERE Vencimento IS NOT NULL AND Vencimento != ''
              AND Vencimento < ? AND Quantidade > 0
        ''', (hoje,))
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
    df = pd.DataFrame([dict(r) for r in rows],
                      columns=['Codigo','Descricao','Quantidade','Preco_Custo','Preco_Venda','Vencimento'])
    if not df.empty:
        df['Quantidade']  = df['Quantidade'].astype(int)
        df['Preco_Custo'] = df['Preco_Custo'].astype(float)
        df['Preco_Venda'] = df['Preco_Venda'].astype(float)
        df['Vencimento']  = pd.to_datetime(df['Vencimento'], errors='coerce')
    return df

@st.cache_data(ttl=60)
def carregar_financeiro():
    with get_conn() as conn:
        c = conn.cursor()
        c.execute('SELECT id,Codigo,Descricao,tipo,quantidade,preco_custo_unit,valor_total,data FROM financeiro ORDER BY data DESC')
        rows = c.fetchall()
    df = pd.DataFrame([dict(r) for r in rows])
    if not df.empty:
        df['data']             = pd.to_datetime(df['data'])
        df['valor_total']      = df['valor_total'].astype(float)
        df['preco_custo_unit'] = df['preco_custo_unit'].astype(float)
        df['quantidade']       = df['quantidade'].astype(int)
    return df

# ==========================================
# APROVAÇÃO MASTER POR E-MAIL
# ==========================================
def aprovar_acao_master(chave, descricao_acao):
    if f"token_{chave}" not in st.session_state:
        st.session_state[f"token_{chave}"] = None
    email_sol = st.text_input("📧 Seu e-mail:", key=f"email_{chave}", placeholder="seunome@empresa.com")
    if st.button(f"📩 Solicitar Liberação: {descricao_acao}", key=f"req_{chave}"):
        if not email_sol:
            st.error("⛔ Informe seu e-mail."); return False
        codigo = str(random.randint(100000, 999999))
        st.session_state[f"token_{chave}"] = codigo
        try:
            rem  = st.secrets["email"]["remetente"]
            pwd  = st.secrets["email"]["senha"]
            dest = st.secrets["email"]["destinatario"]
            msg  = MIMEText(f"Solicitante: {email_sol}\nAção: {descricao_acao}\nCódigo: {codigo}")
            msg['Subject'] = 'Aprovação - Almoxarifado'
            msg['From'] = rem; msg['To'] = dest
            with smtplib.SMTP('smtp.office365.com', 587) as s:
                s.starttls(); s.login(rem, pwd); s.sendmail(rem, [dest], msg.as_string())
            st.info("✅ Solicitação enviada!")
        except Exception as e:
            st.error(f"Erro ao enviar e-mail: {e}")
    if st.session_state[f"token_{chave}"]:
        tok = st.text_input("🔑 Código (enviado a Eduardo Sousa - Controladoria):", key=f"inp_{chave}")
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
    st.error(f"⚠️ Sistema lotado ({total_ativos}/{LIMITE_PESSOAS}). Tente em 1 minuto.")
    st.stop()

vencidos_proc = processar_vencidos_automatico()
if vencidos_proc:
    st.cache_data.clear()
    nomes = ", ".join([f"{v['Descricao']} ({v['Quantidade']} un.)" for v in vencidos_proc])
    st.warning(f"⚠️ **Descarte automático!** Itens vencidos baixados como Perda: {nomes}")

df = carregar_estoque()

# ==========================================
# NAVEGAÇÃO
# ==========================================
st.sidebar.title("Navegação")
menu = st.sidebar.radio("Ir para:", ["📊 Consulta","💰 Financeiro","📥 Entrada","📤 Saída","🔒 Administrativo"])
st.sidebar.divider()
st.sidebar.markdown(f"🟢 **{total_ativos}/{LIMITE_PESSOAS}** pessoas online")

# ==========================================
# TELA 1 — CONSULTA
# ==========================================
if menu == "📊 Consulta":
    src1 = logo_para_base64("logo1.png")
    src2 = logo_para_base64("logo2.png")
    img1 = f'<img class="il1" src="{src1}">' if src1 else '<span style="color:#102a43;font-weight:700;">LOGO 1</span>'
    img2 = f'<img class="il2" src="{src2}">' if src2 else '<span style="color:#102a43;font-weight:700;">LOGO 2</span>'

    hoje      = date.today()
    df_ativos = df[df['Quantidade'] > 0].copy()
    tp = f"{df_ativos['Quantidade'].sum():.0f}" if not df_ativos.empty else "0"
    ti = str(df_ativos['Codigo'].nunique())     if not df_ativos.empty else "0"
    vc = cr = 0
    if not df_ativos.empty:
        df_ativos['_d'] = df_ativos['Vencimento'].apply(lambda v: (v.date()-hoje).days if pd.notna(v) else None)
        vc = int((df_ativos['_d'].dropna() < 0).sum())
        cr = int(((df_ativos['_d'].dropna() >= 0) & (df_ativos['_d'].dropna() <= 30)).sum())

    ca = "alerta" if cr > 0 else ""
    cp = "perigo" if vc > 0 else ""

    components.html(f"""<!DOCTYPE html><html><head>
    <link href="https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
      *{{box-sizing:border-box;margin:0;padding:0;font-family:'Sora',sans-serif;}}
      .hdr{{display:grid;grid-template-columns:1fr auto 1fr;align-items:center;padding:20px 32px;border-radius:16px;margin-bottom:16px;background:linear-gradient(135deg,#f0f4f8,#d9e2ec);box-shadow:0 4px 12px rgba(0,0,0,.05);border:1px solid #e2e8f0;}}
      .ll{{justify-self:start;}}.tb{{text-align:center;padding:0 20px;}}.rl{{justify-self:end;}}
      .il1{{height:85px;width:auto;max-width:240px;object-fit:contain;mix-blend-mode:darken;}}
      .il2{{height:35px;width:auto;max-width:120px;object-fit:contain;mix-blend-mode:darken;}}
      .tb h1{{font-size:1.8rem;font-weight:700;color:#102a43;letter-spacing:.02em;line-height:1.15;}}
      .tb p{{font-size:.75rem;color:#334e68;margin-top:5px;font-weight:600;letter-spacing:.22em;text-transform:uppercase;}}
      .mg{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-top:4px;}}
      .mc{{background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:16px 20px;box-shadow:0 1px 4px rgba(0,0,0,.05);}}
      .mc.alerta{{border-color:#f6ad55;background:#fffbf0;}}.mc.perigo{{border-color:#fc8181;background:#fff5f5;}}
      .ml{{font-size:.78rem;color:#718096;font-weight:600;margin-bottom:4px;}}
      .mv{{font-size:1.9rem;font-weight:700;color:#1a202c;line-height:1.1;}}
      .mv.alerta{{color:#c05621;}}.mv.perigo{{color:#c53030;}}
      @media(max-width:768px){{.hdr{{grid-template-columns:1fr;gap:15px;padding:15px;text-align:center;}}.ll,.rl{{justify-self:center;}}.mg{{grid-template-columns:repeat(2,1fr);gap:8px;}}}}
    </style></head><body>
    <div class="hdr"><div class="ll">{img1}</div><div class="tb"><h1>INVENTÁRIO JOSÉ RIVELINO</h1><p>ALMOXARIFADO</p></div><div class="rl">{img2}</div></div>
    <div class="mg">
      <div class="mc"><div class="ml">📦 Total de Peças</div><div class="mv">{tp}</div></div>
      <div class="mc"><div class="ml">🏷️ Itens Únicos</div><div class="mv">{ti}</div></div>
      <div class="mc {ca}"><div class="ml">⚠️ Vencem em 30 dias</div><div class="mv {ca}">{cr}</div></div>
      <div class="mc {cp}"><div class="ml">🔴 Itens Vencidos</div><div class="mv {cp}">{vc}</div></div>
    </div></body></html>""", height=330, scrolling=False)

    st.divider()
    busca = st.text_input("🔍 Pesquisar Código ou Descrição:")
    df_f = df_ativos.copy()
    if busca:
        df_f = df_f[df_f['Codigo'].astype(str).str.contains(busca, case=False) |
                    df_f['Descricao'].str.contains(busca, case=False, na=False)]

    if not df_f.empty:
        df_f['_dias'] = df_f['Vencimento'].apply(lambda v: (v.date()-hoje).days if pd.notna(v) else 99999)
        df_f = df_f.sort_values('_dias')
        CORES = {
            'vencido':        ('#fff0f0','#c53030','🔴'),
            'critico':        ('#fff8e1','#92400e','🟠'),
            'atencao':        ('#fffde7','#78350f','🟡'),
            'ok':             ('#f0fff4','#276749','🟢'),
            'sem_vencimento': ('#f7fafc','#718096','⚪'),
        }
        def linha(row):
            d = None if row['_dias']==99999 else row['_dias']
            st_, lbl = status_vencimento(d)
            bg,fg,ic = CORES[st_]
            vs = row['Vencimento'].strftime('%d/%m/%Y') if pd.notna(row['Vencimento']) else '—'
            return (f'<tr style="background:{bg};"><td>{row["Codigo"]}</td><td>{row["Descricao"]}</td>'
                    f'<td style="text-align:center;">{row["Quantidade"]}</td>'
                    f'<td style="text-align:center;">{formatar_moeda(row["Preco_Custo"])}</td>'
                    f'<td style="text-align:center;">{formatar_moeda(row["Preco_Venda"])}</td>'
                    f'<td style="text-align:center;color:{fg};font-weight:600;">{ic} {vs}</td>'
                    f'<td style="text-align:center;color:{fg};font-weight:600;font-size:.85rem;">{lbl}</td></tr>')
        linhas = "\n".join(df_f.apply(linha, axis=1))
        components.html(f"""
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&display=swap');
          .t{{width:100%;border-collapse:collapse;font-family:'Sora',sans-serif;font-size:.88rem;}}
          .t th{{background:#1a3a4a;color:#fff;padding:10px 12px;text-align:left;font-weight:600;}}
          .t td{{padding:9px 12px;border-bottom:1px solid #e2e8f0;}}
          .leg{{display:flex;gap:18px;margin-top:12px;font-size:.8rem;font-family:'Sora',sans-serif;flex-wrap:wrap;}}
          .li{{display:flex;align-items:center;gap:5px;}}
        </style>
        <table class="t"><thead><tr>
          <th>Código</th><th>Descrição</th><th>Qtd</th><th>Preço Custo</th><th>Preço Venda</th><th>Vencimento</th><th>Status</th>
        </tr></thead><tbody>{linhas}</tbody></table>
        <div class="leg">
          <div class="li">🔴 Vencido</div><div class="li">🟠 ≤30 dias</div>
          <div class="li">🟡 ≤90 dias</div><div class="li">🟢 OK</div><div class="li">⚪ Sem vencimento</div>
        </div>""", height=min(600, 120+len(df_f)*42), scrolling=True)
    else:
        st.info("Nenhum item encontrado.")

# ==========================================
# TELA 2 — FINANCEIRO
# ==========================================
elif menu == "💰 Financeiro":
    st.title("💰 Dashboard Financeiro")

    col_fi, col_ff = st.columns(2)
    data_ini = col_fi.date_input("📅 De:", value=date.today().replace(day=1))
    data_fim = col_ff.date_input("📅 Até:", value=date.today())

    df_fin = carregar_financeiro()

    if df_fin.empty:
        st.info("Nenhuma movimentação registrada ainda.")
    else:
        mask  = (df_fin['data'].dt.date >= data_ini) & (df_fin['data'].dt.date <= data_fim)
        df_p  = df_fin[mask].copy()
        df_v  = df_p[df_p['tipo'] == 'Venda']
        df_pr = df_p[df_p['tipo'] == 'Perda']

        receita     = df_v['valor_total'].sum()
        custo_vend  = (df_v['preco_custo_unit'] * df_v['quantidade']).sum()
        lucro       = receita - custo_vend
        margem      = (lucro / receita * 100) if receita > 0 else 0
        tot_perdas  = df_pr['valor_total'].sum()
        resultado   = lucro - tot_perdas

        with get_conn() as conn:
            cur = conn.cursor()
            cur.execute('SELECT COALESCE(SUM(Quantidade*Preco_Custo),0) as v FROM estoque WHERE Quantidade>0')
            val_estoque = float(cur.fetchone()['v'])

        st.markdown("### 📊 Resumo do Período")
        k1,k2,k3,k4,k5,k6 = st.columns(6)
        k1.metric("💵 Receita Bruta",      formatar_moeda(receita))
        k2.metric("🏷️ Custo das Vendas",  formatar_moeda(custo_vend))
        k3.metric("📈 Lucro Bruto",        formatar_moeda(lucro), delta=f"{margem:.1f}% margem")
        k4.metric("🗑️ Perdas",            formatar_moeda(tot_perdas))
        k5.metric("✅ Resultado Líquido",  formatar_moeda(resultado))
        k6.metric("📦 Valor em Estoque",   formatar_moeda(val_estoque))

        st.divider()
        tab_mov, tab_ins, tab_mg, tab_giro = st.tabs(
            ["📋 Movimentações", "💡 Insights", "📊 Margem por Item", "🔄 Giro de Estoque"]
        )

        # ---- Movimentações ----
        with tab_mov:
            if df_p.empty:
                st.info("Sem movimentações no período.")
            else:
                ex = df_p.copy()
                ex['Custo Total'] = (ex['preco_custo_unit'] * ex['quantidade']).apply(formatar_moeda)
                ex['Ganho']       = ex.apply(
                    lambda r: formatar_moeda(r['valor_total'] - r['preco_custo_unit']*r['quantidade'])
                    if r['tipo']=='Venda' else '—', axis=1)
                ex['data']        = ex['data'].dt.strftime('%d/%m/%Y %H:%M')
                ex['valor_total'] = ex['valor_total'].apply(formatar_moeda)
                ex = ex.rename(columns={'data':'Data','tipo':'Operação','Codigo':'Código',
                                        'Descricao':'Descrição','quantidade':'Qtd','valor_total':'Valor Total'})
                st.dataframe(ex[['Data','Operação','Código','Descrição','Qtd','Custo Total','Valor Total','Ganho']],
                             use_container_width=True, hide_index=True)

        # ---- Insights ----
        with tab_ins:
            st.subheader("💡 Insights do Período")
            if df_v.empty:
                st.info("Sem vendas no período para gerar insights.")
            else:
                top_rec = df_v.groupby('Descricao')['valor_total'].sum().sort_values(ascending=False).head(5).reset_index()
                top_rec.columns = ['Produto','Receita Total']

                top_vol = df_v.groupby('Descricao')['quantidade'].sum().sort_values(ascending=False).head(5).reset_index()
                top_vol.columns = ['Produto','Qtd Vendida']

                dv2 = df_v.copy()
                dv2['ganho_u'] = dv2['valor_total'] - dv2['preco_custo_unit']*dv2['quantidade']
                top_mg = dv2.groupby('Descricao').apply(
                    lambda g: g['ganho_u'].sum()/g['valor_total'].sum()*100 if g['valor_total'].sum()>0 else 0
                ).sort_values(ascending=False).head(5).reset_index()
                top_mg.columns = ['Produto','Margem (%)']

                c1,c2 = st.columns(2)
                with c1:
                    st.markdown("#### 🥇 Maior Receita")
                    df_r = top_rec.copy(); df_r['Receita Total'] = df_r['Receita Total'].apply(formatar_moeda)
                    st.dataframe(df_r, use_container_width=True, hide_index=True)
                    st.markdown("#### 🔄 Maior Volume de Vendas")
                    st.dataframe(top_vol, use_container_width=True, hide_index=True)
                with c2:
                    st.markdown("#### 📈 Melhor Margem de Lucro")
                    df_m = top_mg.copy(); df_m['Margem (%)'] = df_m['Margem (%)'].apply(lambda x: f"{x:.1f}%")
                    st.dataframe(df_m, use_container_width=True, hide_index=True)
                    if not df_pr.empty:
                        st.markdown("#### 🗑️ Maiores Perdas")
                        tp2 = df_pr.groupby('Descricao')['valor_total'].sum().sort_values(ascending=False).head(5).reset_index()
                        tp2.columns=['Produto','Valor Perdido']; tp2['Valor Perdido']=tp2['Valor Perdido'].apply(formatar_moeda)
                        st.dataframe(tp2, use_container_width=True, hide_index=True)

                st.divider()
                st.markdown("#### ⚡ Alertas Automáticos")
                alertas = []

                # Vendas abaixo do custo
                dv2['abaixo'] = dv2['ganho_u'] < 0
                for p in dv2[dv2['abaixo']]['Descricao'].unique():
                    alertas.append(("🔴", f"**{p}** vendido abaixo do custo em alguma transação."))

                # Perdas > 20% da receita
                if receita > 0 and tot_perdas/receita > 0.2:
                    alertas.append(("🟠", f"Perdas = **{tot_perdas/receita*100:.1f}%** da receita — acima de 20%."))

                # Margem bruta baixa
                if 0 < margem < 15:
                    alertas.append(("🟡", f"Margem bruta baixa: **{margem:.1f}%**. Revisar preços de venda."))

                # Alto giro com baixa margem
                vol_prod = dv2.groupby('Descricao')['quantidade'].sum()
                mg_prod  = dv2.groupby('Descricao').apply(
                    lambda g: g['ganho_u'].sum()/g['valor_total'].sum()*100 if g['valor_total'].sum()>0 else 0)
                for prod in vol_prod.index:
                    if vol_prod[prod] >= 10 and mg_prod.get(prod, 100) < 10:
                        alertas.append(("🟡", f"**{prod}** — alto giro mas margem < 10%. Revisar precificação."))

                if not alertas:
                    st.success("✅ Nenhum alerta identificado. Operação saudável!")
                else:
                    for em, msg in alertas:
                        if em=="🔴": st.error(f"{em} {msg}")
                        elif em=="🟠": st.warning(f"{em} {msg}")
                        else: st.info(f"{em} {msg}")

        # ---- Margem por Item ----
        with tab_mg:
            st.subheader("📊 Análise de Margem por Produto")
            if df_v.empty:
                st.info("Sem vendas no período.")
            else:
                dm = df_v.copy()
                dm['custo_t'] = dm['preco_custo_unit'] * dm['quantidade']
                dm['ganho']   = dm['valor_total'] - dm['custo_t']
                res = dm.groupby('Descricao').agg(
                    Qtd=('quantidade','sum'), Receita=('valor_total','sum'),
                    Custo=('custo_t','sum'), Lucro=('ganho','sum')
                ).reset_index()
                res['Margem (%)'] = res.apply(lambda r: r['Lucro']/r['Receita']*100 if r['Receita']>0 else 0, axis=1)
                res = res.sort_values('Margem (%)', ascending=False)
                fmt = res.copy()
                fmt['Receita']    = fmt['Receita'].apply(formatar_moeda)
                fmt['Custo']      = fmt['Custo'].apply(formatar_moeda)
                fmt['Lucro']      = fmt['Lucro'].apply(formatar_moeda)
                fmt['Margem (%)'] = fmt['Margem (%)'].apply(lambda x: f"{x:.1f}%")
                fmt.columns = ['Produto','Qtd Vendida','Receita','Custo das Vendas','Lucro Bruto','Margem (%)']
                st.dataframe(fmt, use_container_width=True, hide_index=True)

        # ---- Giro de Estoque ----
        with tab_giro:
            st.subheader("🔄 Giro de Estoque")
            st.caption("Giro = Qtd vendida no período ÷ Qtd atual em estoque.")
            if df_v.empty:
                st.info("Sem vendas no período.")
            else:
                vc2 = df_v.groupby('Codigo')['quantidade'].sum().reset_index()
                vc2.columns = ['Codigo','Qtd_Vendida']
                est = df[df['Quantidade']>0][['Codigo','Descricao','Quantidade']].copy()
                est = est.groupby(['Codigo','Descricao'])['Quantidade'].sum().reset_index()
                giro = est.merge(vc2, on='Codigo', how='left').fillna(0)
                giro['Giro'] = giro.apply(lambda r: r['Qtd_Vendida']/r['Quantidade'] if r['Quantidade']>0 else 0, axis=1)
                giro = giro.sort_values('Giro', ascending=False)
                def tag(g):
                    if g==0: return "⚪ Parado"
                    if g<0.3: return "🔴 Baixo"
                    if g<0.7: return "🟡 Médio"
                    if g<1.5: return "🟢 Bom"
                    return "🔵 Alto"
                giro['Status'] = giro['Giro'].apply(tag)
                giro['Giro']   = giro['Giro'].apply(lambda x: f"{x:.2f}x")
                giro.columns   = ['Código','Produto','Qtd em Estoque','Qtd Vendida','Giro','Status']
                st.dataframe(giro, use_container_width=True, hide_index=True)
                st.caption("🔵 Alto (>1.5x) · 🟢 Bom (0.7–1.5x) · 🟡 Médio (0.3–0.7x) · 🔴 Baixo (<0.3x) · ⚪ Sem venda")

# ==========================================
# TELA 3 — ENTRADA
# ==========================================
elif menu == "📥 Entrada":
    st.title("📥 Entrada de Itens")
    senha = st.text_input("Senha:", type="password", key="senha_entrada")

    if senha in (SENHA_ACESSO, SENHA_ZERAR_ESTOQUE):
        abas = st.tabs(["📝 Registro Individual", "📤 Carga em Massa"])

        with abas[0]:
            with st.form("entrada_individual", clear_on_submit=True):
                st.markdown("**1. Identificação do Produto**")
                c1,c2 = st.columns(2)
                cod        = c1.text_input("Código:")
                desc_input = c2.text_input("Descrição (obrigatória para itens novos):")

                st.markdown("**2. Vencimento** *(obrigatório)*")
                venc_in = st.date_input("Data de Vencimento:",
                    min_value=date.today(), value=date.today()+timedelta(days=365), key="venc_entrada")

                st.markdown("**3. Preços** *(obrigatório para novas entradas)*")
                c3,c4 = st.columns(2)
                preco_custo = c3.number_input("Preço de Custo Unitário (R$)", min_value=0.0, format="%.2f", step=0.50)
                preco_venda = c4.number_input("Preço de Venda Unitário (R$)", min_value=0.0, format="%.2f", step=0.50)

                st.markdown("**4. Quantidade**")
                qtd = st.number_input("Qtd:", min_value=1, step=1, format="%d")

                if st.form_submit_button("✅ Registrar Entrada"):
                    if not cod:
                        st.error("⛔ Informe o Código do item.")
                    else:
                        de = buscar_item_por_codigo(cod)
                        if not de and not desc_input:
                            st.error("⛔ Descrição obrigatória para novos itens.")
                        elif de and desc_input and desc_input.strip() != de.strip():
                            st.error(f"⛔ Conflito! Código **{cod}** já cadastrado como: **\"{de}\"**")
                        else:
                            df2 = de if de else desc_input
                            vs  = venc_in.isoformat()
                            with get_conn() as conn:
                                cur = conn.cursor()
                                cur.execute('SELECT id,Quantidade,Preco_Custo,Preco_Venda FROM estoque WHERE Codigo=? AND Vencimento=?', (cod,vs))
                                res = cur.fetchone()
                                if res:
                                    nc = preco_custo if preco_custo>0 else res['Preco_Custo']
                                    nv = preco_venda if preco_venda>0 else res['Preco_Venda']
                                    cur.execute('UPDATE estoque SET Quantidade=Quantidade+?,Preco_Custo=?,Preco_Venda=?,Vencimento=? WHERE id=?',
                                                (qtd,nc,nv,vs,res['id']))
                                    st.success(f"✅ Entrada registrada. Novo saldo: {res['Quantidade']+qtd}")
                                else:
                                    cur.execute('INSERT INTO estoque (Codigo,Descricao,Quantidade,Preco_Custo,Preco_Venda,Vencimento) VALUES (?,?,?,?,?,?)',
                                                (cod,df2,qtd,preco_custo,preco_venda,vs))
                                    st.success("✅ Novo item cadastrado.")
                                conn.commit()
                            st.cache_data.clear()

        with abas[1]:
            st.info("Colunas: `Codigo` | `Descricao` | `Quantidade` | `Preco_Custo` | `Preco_Venda` | `Vencimento` (AAAA-MM-DD)")
            st.download_button("⬇️ Baixar Template", gerar_template_xlsx(), "template_inventario.xlsx")
            arq = st.file_uploader("Arquivo (.xlsx):", type=["xlsx"], key="upload_massa")
            if arq:
                try:
                    du = pd.read_excel(arq, engine='openpyxl')
                    falt = {'Codigo','Descricao','Quantidade','Vencimento'} - set(du.columns)
                    if falt:
                        st.error(f"⛔ Colunas ausentes: {', '.join(falt)}")
                    else:
                        if st.button("🚀 Processar Importação"):
                            du['Codigo']    = du['Codigo'].astype(str).str.strip()
                            du['Descricao'] = du['Descricao'].astype(str).str.strip()
                            du['Quantidade']= pd.to_numeric(du['Quantidade'], errors='coerce')
                            du['Vencimento']= pd.to_datetime(du['Vencimento'], errors='coerce').dt.strftime('%Y-%m-%d')
                            for col in ['Preco_Custo','Preco_Venda']:
                                if col not in du.columns: du[col]=0.0
                                du[col] = pd.to_numeric(du[col], errors='coerce').fillna(0.0)
                            du = du.dropna(subset=['Quantidade','Vencimento'])
                            du = du[du['Quantidade']>0]; du['Quantidade']=du['Quantidade'].astype(int)
                            du = du[~du['Codigo'].isin(['nan',''])]
                            with get_conn() as conn:
                                cur = conn.cursor()
                                cur.execute('SELECT Codigo,Vencimento FROM estoque')
                                dbs = set((r['Codigo'],r['Vencimento']) for r in cur.fetchall())
                                ins,upd=[],[]
                                for _,row in du.iterrows():
                                    k=(row['Codigo'],row['Vencimento'])
                                    if k in dbs: upd.append((row['Quantidade'],row['Preco_Custo'],row['Preco_Venda'],row['Codigo'],row['Vencimento']))
                                    else: ins.append((row['Codigo'],row['Descricao'],row['Quantidade'],row['Preco_Custo'],row['Preco_Venda'],row['Vencimento'])); dbs.add(k)
                                if ins: cur.executemany('INSERT INTO estoque (Codigo,Descricao,Quantidade,Preco_Custo,Preco_Venda,Vencimento) VALUES (?,?,?,?,?,?)',ins)
                                if upd: cur.executemany('UPDATE estoque SET Quantidade=Quantidade+?,Preco_Custo=MAX(Preco_Custo,?),Preco_Venda=MAX(Preco_Venda,?) WHERE Codigo=? AND Vencimento=?',upd)
                                conn.commit()
                            st.success(f"✅ {len(ins)} inseridos, {len(upd)} atualizados.")
                            st.cache_data.clear(); st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")
    elif senha:
        st.error("⛔ Senha incorreta.")

# ==========================================
# TELA 4 — SAÍDA
# ==========================================
elif menu == "📤 Saída":
    st.title("📤 Saída de Itens")
    senha = st.text_input("Senha:", type="password", key="senha_saida")

    if senha in (SENHA_ACESSO, SENHA_ZERAR_ESTOQUE):

        st.markdown("**1. Identificação do Item**")
        c1,c2 = st.columns(2)
        cod_s = c1.text_input("Código:", key="cod_s")
        op_s  = c2.selectbox("Operação:", ["Venda","Perda"], key="op_s")

        st.markdown("**2. Lote / Vencimento**")
        st.caption("Deixe em branco para usar o lote mais antigo (FIFO).")
        venc_s = st.date_input("Vencimento do Lote (opcional):", value=None, key="venc_s")

        st.markdown("**3. Quantidade**")
        qtd_s = st.number_input("Qtd:", min_value=1, step=1, format="%d", key="qtd_s")

        # Pré-visualização do item
        item_pv = None
        preco_conf = None

        if cod_s:
            with get_conn() as conn:
                cur = conn.cursor()
                if venc_s:
                    cur.execute('''SELECT id,Descricao,Quantidade,Preco_Custo,Preco_Venda,Vencimento
                                   FROM estoque WHERE Codigo=? AND Vencimento=? AND Quantidade>0''',
                                (cod_s, venc_s.isoformat()))
                else:
                    cur.execute('''SELECT id,Descricao,Quantidade,Preco_Custo,Preco_Venda,Vencimento
                                   FROM estoque WHERE Codigo=? AND Quantidade>0
                                   ORDER BY Vencimento ASC LIMIT 1''', (cod_s,))
                item_pv = cur.fetchone()

            if item_pv:
                vf = item_pv['Vencimento'] or '—'
                st.info(
                    f"📦 **{item_pv['Descricao']}** | Lote: `{vf}` | "
                    f"Saldo: **{item_pv['Quantidade']}** un. | "
                    f"Custo unit.: {formatar_moeda(item_pv['Preco_Custo'])} | "
                    f"Venda padrão: {formatar_moeda(item_pv['Preco_Venda'])}"
                )
                if op_s == "Venda":
                    st.markdown("**4. Confirmar / Ajustar Preço de Venda**")
                    st.caption("Preço padrão do cadastro já preenchido. Altere para aplicar promoção ou desconto.")
                    preco_conf = st.number_input(
                        "💲 Preço de Venda Unitário (R$):",
                        min_value=0.01,
                        value=float(item_pv['Preco_Venda']),
                        format="%.2f", step=0.50, key="preco_conf"
                    )
                    valor_total_prev = qtd_s * preco_conf
                    lucro_prev       = valor_total_prev - qtd_s * item_pv['Preco_Custo']
                    margem_prev      = lucro_prev / valor_total_prev * 100 if valor_total_prev > 0 else 0

                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("Receita prevista",    formatar_moeda(valor_total_prev))
                    col_b.metric("Lucro bruto previsto", formatar_moeda(lucro_prev))
                    col_c.metric("Margem prevista",      f"{margem_prev:.1f}%")

                    if preco_conf < item_pv['Preco_Custo']:
                        st.warning(
                            f"⚠️ Preço ({formatar_moeda(preco_conf)}) abaixo do custo "
                            f"({formatar_moeda(item_pv['Preco_Custo'])}). Venda com prejuízo."
                        )
                else:
                    preco_conf = item_pv['Preco_Custo']
            else:
                st.error("⛔ Item não encontrado ou sem estoque para este lote.")

        st.divider()
        if st.button("✅ Confirmar Saída", key="btn_saida"):
            if not cod_s:
                st.error("⛔ Informe o Código.")
            elif item_pv is None:
                st.error("⛔ Item não encontrado.")
            elif item_pv['Quantidade'] < qtd_s:
                st.error(f"⛔ Estoque insuficiente! Disponível: {item_pv['Quantidade']} un.")
            else:
                with get_conn() as conn:
                    cur = conn.cursor()
                    cur.execute('UPDATE estoque SET Quantidade=Quantidade-? WHERE id=?', (qtd_s, item_pv['id']))
                    if op_s == "Venda":
                        val  = qtd_s * preco_conf
                        ganho = val - qtd_s * item_pv['Preco_Custo']
                        cur.execute('INSERT INTO financeiro (Codigo,Descricao,tipo,quantidade,preco_custo_unit,valor_total) VALUES (?,?,?,?,?,?)',
                                    (cod_s, item_pv['Descricao'], 'Venda', qtd_s, item_pv['Preco_Custo'], val))
                        st.success(f"✅ Venda registrada! Receita: {formatar_moeda(val)} | Lucro bruto: {formatar_moeda(ganho)}")
                    else:
                        val = qtd_s * item_pv['Preco_Custo']
                        cur.execute('INSERT INTO financeiro (Codigo,Descricao,tipo,quantidade,preco_custo_unit,valor_total) VALUES (?,?,?,?,?,?)',
                                    (cod_s, item_pv['Descricao'], 'Perda', qtd_s, item_pv['Preco_Custo'], val))
                        st.success(f"✅ Perda registrada! Prejuízo: {formatar_moeda(val)}")
                    conn.commit()
                st.cache_data.clear()
    elif senha:
        st.error("⛔ Senha incorreta.")

# ==========================================
# TELA 5 — ADMINISTRATIVO
# ==========================================
elif menu == "🔒 Administrativo":
    st.title("🔒 Área Administrativa")
    senha = st.text_input("Senha:", type="password", key="senha_admin")

    if senha == SENHA_ZERAR_ESTOQUE:
        abas = st.tabs(["🗑️ Excluir Item","⚠️ Limpar Dados"])

        with abas[0]:
            st.subheader("🗑️ Excluir Item do Banco")
            cod_del = st.text_input("Código do item:")
            if cod_del and aprovar_acao_master("del_item", f"Excluir código {cod_del}"):
                with get_conn() as conn:
                    cur = conn.cursor()
                    cur.execute('SELECT * FROM estoque WHERE Codigo=?', (cod_del,))
                    if cur.fetchone():
                        cur.execute('DELETE FROM estoque WHERE Codigo=?', (cod_del,))
                        st.success(f"✅ Código **{cod_del}** apagado!")
                    else:
                        st.error("⛔ Código não encontrado.")
                    conn.commit()
                st.cache_data.clear()

        with abas[1]:
            st.subheader("⚠️ Área de Risco")
            opcao = st.radio("Ação:", [
                "1️⃣ Zerar quantidades (mantém cadastros)",
                "2️⃣ Excluir todo o estoque",
                "3️⃣ Limpar histórico financeiro",
            ])
            if aprovar_acao_master("limpeza", f"Limpeza: {opcao}"):
                with get_conn() as conn:
                    cur = conn.cursor()
                    if "1️⃣" in opcao: cur.execute('UPDATE estoque SET Quantidade=0'); st.success("Quantidades zeradas!")
                    elif "2️⃣" in opcao: cur.execute("DELETE FROM estoque"); st.success("Estoque apagado!")
                    elif "3️⃣" in opcao: cur.execute("DELETE FROM financeiro"); st.success("Histórico financeiro limpo!")
                    conn.commit()
                st.cache_data.clear(); st.rerun()
    elif senha:
        st.error("⛔ Senha incorreta.")
