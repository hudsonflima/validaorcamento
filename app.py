import streamlit as st
import pandas as pd
import locale
import io
import json
import unicodedata
import re

# Tenta definir locale para pt_BR, com fallback
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_ALL, '')

# Configuração da página
st.set_page_config(
    page_title="Validador de Operações Orçamentárias",  # ← título da aba
    layout="wide",
    page_icon="🧾"  # ← opcional: ícone na aba
)
st.title("Validador de Operações Orçamentárias")

@st.cache_data
def carregar_lista_mestre(path="dados/operacoes.csv"):
    df = pd.read_csv(path, sep=";", dtype={"Cod": str})
    df["Cod"] = df["Cod"].str.zfill(7)
    df["Descr"] = df["Descr"].str.lower().str.strip()
    return dict(zip(df["Cod"], df["Descr"]))

def para_float(valor):
    try:
        return float(str(valor).replace(',', '.'))
    except:
        return 0.0

def formatar_moeda_br(valor):
    return f"{valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def limpar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = texto.lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    texto = re.sub(r'[^a-z0-9 ]', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()

lista_mestre = carregar_lista_mestre()
with open('dados/termos_orcamentarios.json', encoding='utf-8') as f:
    termos_orcamentarios = json.load(f)

def sugerir_operacao(row):
    descricao = str(row["Descr. Op. Orc."]).lower().strip()
    operacao = str(row["Op. Orc."]).strip().lower()
    historico = limpar_texto(row["Histórico"])

    termos_limpeza = ["papel higienico", "papel toalha", "desinfetante", "alcool em gel", "sabonete", "peroxy"]
    if any(t in historico for t in termos_limpeza) and descricao == "material de expediente":
        return "0800108 - Material de copa / limpeza"

    if "cafe" in historico and operacao == "0800108" and "copa" in descricao:
        return "0800102 - Maquinas de Café - Insumos"

    if operacao in lista_mestre and lista_mestre[operacao].lower() == descricao:
        return "Conferido"

    cond_sem_op = descricao == 'sem op. orc.' or operacao in ['', '-', '0000001', 'nan', 'null']
    cond_descricao_errada = operacao in lista_mestre and lista_mestre[operacao] != descricao
    cond_palavra_proibida = descricao == 'material de copa / limpeza' and "cafe" in historico

    if cond_sem_op or cond_descricao_errada or cond_palavra_proibida:
        termos_ordenados = sorted(
            termos_orcamentarios.items(),
            key=lambda item: -len(item[0])
        )
        for termo, (codigo, desc) in termos_ordenados:
            if limpar_texto(termo) in historico:
                return f"{codigo} - {desc}"
        return "Sem sugestão"

    return "-"

uploaded_file = st.file_uploader("Faça upload do arquivo Excel (Razão)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype={
        "Centro de Custo": str,
        "Descr. Op. Orc.": str,
        "Op. Orc.": str,
        "Mês": str,
        "Conta": str,
        "Histórico": str
    })

    df["Op. Orc."] = df["Op. Orc."].astype(str).str.replace(r'\.0$', '', regex=True).str.zfill(7)
    df["Descr. Op. Orc."] = df["Descr. Op. Orc."].str.lower().str.strip()
    df["Data Contábil"] = pd.to_datetime(df["Data Contábil"], errors="coerce")
    df["Valor_Numérico"] = df["Valor Realizado"].apply(para_float).round(2)
    df["Valor Realizado"] = df["Valor_Numérico"].apply(formatar_moeda_br)

    df["Sugestão Op. Orc."] = df.apply(sugerir_operacao, axis=1)

    def validar(cod, descr):
        if cod not in lista_mestre:
            return "Código não encontrado"
        elif lista_mestre[cod].lower() != descr.lower():
            return "Descrição divergente"
        return "OK"

    df["Validação"] = df.apply(lambda row: validar(row["Op. Orc."], row["Descr. Op. Orc."]), axis=1)

    df["Descr. Op. Orc."] = df["Descr. Op. Orc."].str.capitalize()
    df["Data Contábil"] = df["Data Contábil"].dt.strftime('%d/%m/%Y').fillna("")

    def texto_colorido(valor):
        if valor == "Conferido":
            return "✅ Conferido"
        elif valor == "Sem sugestão":
            return "❌ Sem sugestão"
        elif "-" in str(valor):
            return f"🟠 {valor}"
        return valor

    def status_validacao(valor):
        return "✅ OK" if valor == "OK" else f"❌ {valor}"

    df["Sugestão Op. Orc."] = df["Sugestão Op. Orc."].apply(texto_colorido)
    df["Validação"] = df["Validação"].apply(status_validacao)

    st.subheader("KPIs")
    sem_op_orc = df[df["Descr. Op. Orc."] == "Sem op. orc."]
    soma_sem_op_orc = sem_op_orc["Valor_Numérico"].sum()
    kpis_conta = df.groupby("Conta")["Valor_Numérico"].sum().reset_index()
    kpi_data = [("Impacto Financeiro - Sem Op. Orc.", soma_sem_op_orc)]
    for _, row in kpis_conta.iterrows():
        kpi_data.append((f"KPI - Conta {row['Conta']}", row["Valor_Numérico"]))

    for linha in [kpi_data[i:i+3] for i in range(0, len(kpi_data), 3)]:
        cols = st.columns(3)
        for col, (titulo, valor) in zip(cols, linha):
            col.metric(titulo, formatar_moeda_br(valor))

    st.badge("Arquivo processado com sucesso!", icon=":material/check:", color="green")

    termo_busca = st.text_input("🔍 Buscar por termo (Histórico, Op. Orc., Sugestão):", "").strip().lower()

    if termo_busca:
        df_filtrado = df[
            df["Histórico"].str.lower().str.contains(termo_busca, na=False) |
            df["Op. Orc."].str.lower().str.contains(termo_busca, na=False) |
            df["Sugestão Op. Orc."].str.lower().str.contains(termo_busca, na=False)
        ]
    else:
        df_filtrado = df.copy()


    status_opcao = st.selectbox("Filtrar por status de validação:", ["Todos", "✅ OK", "❌ Descrição divergente", "❌ Código não encontrado"])

    if status_opcao != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Validação"] == status_opcao]


    
    colunas_visiveis = [col for col in df.columns if col != "Valor_Numérico"]
    with st.container():
        # st.data_editor(df[colunas_visiveis], hide_index=True, width=1755, height=740)
        st.data_editor(df_filtrado[colunas_visiveis], hide_index=True, width=1755, height=740)


    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Validação')
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Baixar resultado em Excel",
        data=processed_data,
        file_name="resultado_validacao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
