import streamlit as st
import pandas as pd
import locale
import io

# Definir o locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Configuração da página
st.set_page_config(layout="wide")
st.title("Validador de Operações Orçamentárias")

# Função para carregar a lista mestre
@st.cache_data
def carregar_lista_mestre(path="dados/operacoes.csv"):
    df = pd.read_csv(path, sep=";", dtype={"Cod": str})
    df["Cod"] = df["Cod"].str.zfill(8)
    df["Descr"] = df["Descr"].str.lower().str.strip()
    return dict(zip(df["Cod"], df["Descr"]))


# Carrega DataFrame da lista mestre
operacoes = pd.read_csv("dados/operacoes.csv", sep=";")
operacoes["Cod"] = operacoes["Cod"].astype(str).str.zfill(7)
operacoes["Descr"] = operacoes["Descr"].str.lower().str.strip()
lista_mestre = dict(zip(operacoes["Cod"], operacoes["Descr"]))

# Upload do arquivo Excel
uploaded_file = st.file_uploader("Faça upload do arquivo Excel (Razão)", type=["xlsx"])

# Função para parse de moeda
def parse_moeda(valor):
    try:
        valor_float = float(valor)
        return locale.currency(valor_float, grouping=True)
    except:
        return None

# Se um arquivo for carregado
if uploaded_file:

    # Inicializa o DataFrame de resultado no estado da sessão, se necessário
    if "df_resultado" not in st.session_state:
        st.session_state.df_resultado = None

    # Carrega o arquivo Excel com os dados
    df = pd.read_excel(uploaded_file, dtype={
        "Centro de Custo": str,
        "Descr. Op. Orc.": str,
        "Op. Orc.": str,
        "Mês": str,
        "Conta": str,
        "Histórico": str
    })

    # Converter colunas específicas
    df["Op. Orc."] = df["Op. Orc."].astype(str).str.zfill(7)
    df["Descr. Op. Orc."] = df["Descr. Op. Orc."].str.lower().str.strip()

    # Converter data contábil
    df["Data Contábil"] = pd.to_datetime(df["Data Contábil"], format="%d/%m/%Y", errors="coerce")

    # Converter valor realizado para moeda
    df["Valor Realizado"] = df["Valor Realizado"].apply(parse_moeda)

    # Função de validação
    def validar(cod, descr):
        if cod not in lista_mestre:
            return "Código não encontrado"
        elif lista_mestre[cod] != descr:
            return "Descrição divergente"
        else:
            return "OK"

    # Aplicar a validação em cada linha
    df["Validação"] = df.apply(lambda row: validar(row["Op. Orc."], row["Descr. Op. Orc."]), axis=1)

    # Exibir melhorias na tabela
    df["Descr. Op. Orc."] = df["Descr. Op. Orc."].str.capitalize()
    df["Data Contábil"] = df["Data Contábil"].dt.strftime('%d/%m/%Y')
    df["Valor Realizado"] = df["Valor Realizado"]

    # Mostrar o KPI de "Sem Op. Orc."
    sem_op_orc = df[df["Descr. Op. Orc."] == "Sem op. orc."]
    soma_sem_op_orc = sem_op_orc["Valor Realizado"].apply(lambda x: float(x.replace('R$', '').replace('.', '').replace(',', '.'))).sum()

    # Mostrar KPIs por tipo de conta
    kpis_conta = df.groupby("Conta")["Valor Realizado"].apply(
        lambda x: x.apply(lambda val: float(val.replace('R$', '').replace('.', '').replace(',', '.'))).sum()
    ).reset_index()

    # # Exibir indicadores
    # st.subheader("Indicadores")
    
    # # KPI de "Sem Op. Orc."
    # st.metric("Impacto Financeiro - Sem Op. Orc.", f"R$ {soma_sem_op_orc:,.2f}")

    # # KPIs por tipo de conta
    # for index, row in kpis_conta.iterrows():
    #     st.metric(f"KPI - Conta {row['Conta']}", f"R$ {row['Valor Realizado']:,.2f}")
    st.subheader("KPIs")

    # KPI 1 - Impacto Financeiro - Sem Op. Orc.
    kpi_data = [("Impacto Financeiro - Sem Op. Orc.", soma_sem_op_orc)]

    # KPIs das contas
    for index, row in kpis_conta.iterrows():
        kpi_data.append((f"KPI - Conta {row['Conta']}", row["Valor Realizado"]))

    # Organiza os KPIs em 3 colunas e 2 linhas
    num_cols = 3
    rows = [kpi_data[i:i + num_cols] for i in range(0, len(kpi_data), num_cols)]

    for row in rows:
        cols = st.columns(num_cols)
        for col, (label, value) in zip(cols, row):
            col.metric(label, f"R$ {value:,.2f}")
    # Exibir sucesso no processamento
    st.badge("Arquivo processado com sucesso!", icon=":material/check:", color="green")

    # Estilos personalizados para a exibição do DataFrame
    st.markdown("""
        <style>
            .element-container:has(.stDataFrame) {
                width: 100vw !important;
                max-width: 100vw !important;
            }
        </style>
    """, unsafe_allow_html=True)
    
    # Exibir a tabela de validação
    with st.container():
        st.data_editor(df, hide_index=True, width=1755, height=740)

    # Salva o DataFrame de resultado no estado da sessão
    st.session_state.df_resultado = df

    # Exibe o DataFrame
    # st.dataframe(st.session_state.df_resultado)

    # Gerar o arquivo Excel com os dados processados
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
