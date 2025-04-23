import streamlit as st
import pandas as pd
import locale
import io

# Tenta definir locale para pt_BR, com fallback
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_ALL, '')

# Configuração da página
st.set_page_config(layout="wide")
st.title("Validador de Operações Orçamentárias")

# Cache para carregar lista mestre
@st.cache_data
def carregar_lista_mestre(path="dados/operacoes.csv"):
    df = pd.read_csv(path, sep=";", dtype={"Cod": str})
    df["Cod"] = df["Cod"].str.zfill(8)
    df["Descr"] = df["Descr"].str.lower().str.strip()
    return dict(zip(df["Cod"], df["Descr"]))

# Função para converter valores para float (de forma robusta)
def para_float(valor):
    try:
        if isinstance(valor, str):
            return float(valor.replace('R$', '').replace('.', '').replace(',', '.'))
        return float(valor)
    except:
        return 0.0

# Função para formatar moeda brasileira
def formatar_moeda_br(valor):
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

# Carrega a lista mestre
lista_mestre = carregar_lista_mestre()

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Faça upload do arquivo Excel (Razão)", type=["xlsx"])

# Se um arquivo for carregado
if uploaded_file:
    df = pd.read_excel(uploaded_file, dtype={
        "Centro de Custo": str,
        "Descr. Op. Orc.": str,
        "Op. Orc.": str,
        "Mês": str,
        "Conta": str,
        "Histórico": str
    })

    # Tratamento de colunas
    df["Op. Orc."] = df["Op. Orc."].astype(str).str.zfill(7)
    df["Descr. Op. Orc."] = df["Descr. Op. Orc."].str.lower().str.strip()
    df["Data Contábil"] = pd.to_datetime(df["Data Contábil"], errors="coerce")
    df["Valor_Numérico"] = df["Valor Realizado"].apply(para_float)
    df["Valor Realizado"] = df["Valor_Numérico"].apply(formatar_moeda_br)

    # Validação
    def validar(cod, descr):
        if cod not in lista_mestre:
            return "Código não encontrado"
        elif lista_mestre[cod] != descr:
            return "Descrição divergente"
        return "OK"

    df["Validação"] = df.apply(lambda row: validar(row["Op. Orc."], row["Descr. Op. Orc."]), axis=1)

    # Exibições finais
    df["Descr. Op. Orc."] = df["Descr. Op. Orc."].str.capitalize()
    df["Data Contábil"] = df["Data Contábil"].dt.strftime('%d/%m/%Y').fillna("")

    # KPIs
    sem_op_orc = df[df["Descr. Op. Orc."] == "Sem op. orc."]
    soma_sem_op_orc = sem_op_orc["Valor_Numérico"].sum()
    kpis_conta = df.groupby("Conta")["Valor_Numérico"].sum().reset_index()

    st.subheader("KPIs")
    kpi_data = [("Impacto Financeiro - Sem Op. Orc.", soma_sem_op_orc)]
    for _, row in kpis_conta.iterrows():
        kpi_data.append((f"KPI - Conta {row['Conta']}", row["Valor_Numérico"]))

    # Organiza KPIs em colunas
    num_cols = 3
    rows = [kpi_data[i:i + num_cols] for i in range(0, len(kpi_data), num_cols)]
    for row in rows:
        cols = st.columns(num_cols)
        for col, (label, value) in zip(cols, row):
            col.metric(label, formatar_moeda_br(value))

    st.badge("Arquivo processado com sucesso!", icon=":material/check:", color="green")

    # Estilo para DataFrame
    st.markdown("""
        <style>
            .element-container:has(.stDataFrame) {
                width: 100vw !important;
                max-width: 100vw !important;
            }
        </style>
    """, unsafe_allow_html=True)

    # Exibe a tabela ocultando coluna interna
    colunas_visiveis = [col for col in df.columns if col != "Valor_Numérico"]
    with st.container():
        st.data_editor(df[colunas_visiveis], hide_index=True, width=1755, height=740)

    # Prepara download do resultado
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Validação')
    processed_data = output.getvalue()
    
    # Botão para download do resultado
    st.download_button(
        label="📥 Baixar resultado em Excel",
        data=processed_data,
        file_name="resultado_validacao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
