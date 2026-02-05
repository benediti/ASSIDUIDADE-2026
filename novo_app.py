import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="Cálculo de Prêmio - Nova Lógica", layout="wide")
st.title("Sistema de Cálculo de Prêmio - Nova Lógica")

st.sidebar.header("Importação de Dados")

# Uploads
func_file = st.sidebar.file_uploader("Base de Funcionários", type=["xlsx"])
aus_file = st.sidebar.file_uploader("Base de Ausências", type=["xlsx"])
tipo_file = st.sidebar.file_uploader("Tipos de Afastamento (opcional)", type=["xlsx"])

data_limite = st.sidebar.date_input("Data Limite de Admissão", value=datetime.now())

# Função para cálculo do prêmio considerando atestados
def calcular_premio(row, ausencias):
    VALOR_BASE = 315.00
    SALARIO_LIMITE = 2720.86
    horas = row['Qtd_Horas_Mensais']
    salario = row['Salario_Mes_Atual']
    status = "Tem direito"
    valor = VALOR_BASE
    detalhes = []
    # Filtra ausências do funcionário
    aus = ausencias[ausencias['Matricula'] == row['Matricula']]
    # Conta atestados
    dias_atestado = aus['Afastamentos'].str.lower().str.contains('atestado').sum()
    # Regras de cálculo
    if salario > SALARIO_LIMITE:
        status = "Não tem direito"
        valor = 0
        detalhes.append("Salário acima do limite")
    elif dias_atestado >= 3:
        status = "Não tem direito"
        valor = 0
        detalhes.append(f"{dias_atestado} dias de atestado (perde o direito)")
    elif dias_atestado == 2:
        valor = VALOR_BASE * 0.25
        detalhes.append("2 dias de atestado (25% do valor)")
    elif dias_atestado == 1:
        valor = VALOR_BASE * 0.5
        detalhes.append("1 dia de atestado (50% do valor)")
    # Jornada 4h
    if horas <= 120 and valor > 0:
        valor = round(valor * 0.5, 2)
        detalhes.append("Jornada 4h (50%)")
    return pd.Series({
        'Valor_Premio': valor,
        'Status': status,
        'Detalhes': "; ".join(detalhes),
        'Qtd_Atestados': dias_atestado
    })

# Processamento principal
def processar():
    if not (func_file and aus_file):
        st.warning("Carregue as bases de funcionários e ausências.")
        return
    df_func = pd.read_excel(func_file)
    df_aus = pd.read_excel(aus_file)
    # Padroniza nomes de colunas
    df_func.columns = [c.strip() for c in df_func.columns]
    df_aus.columns = [c.strip() for c in df_aus.columns]
    # Garante colunas essenciais
    if 'Matricula' not in df_func.columns:
        st.error("Coluna 'Matricula' não encontrada na base de funcionários.")
        return
    if 'Afastamentos' not in df_aus.columns:
        st.error("Coluna 'Afastamentos' não encontrada na base de ausências.")
        return
    # Filtra por data de admissão
    df_func['Data_Admissao'] = pd.to_datetime(df_func['Data_Admissao'], errors='coerce', dayfirst=True)
    df_func = df_func[df_func['Data_Admissao'] <= pd.to_datetime(data_limite)]
    # Aplica cálculo
    resultado = df_func.apply(lambda row: calcular_premio(row, df_aus), axis=1)
    df_final = pd.concat([df_func, resultado], axis=1)
    st.subheader("Relatório de Prêmios Calculados")
    st.dataframe(df_final)
    # Exportação Excel
    if st.button("Exportar Relatório Executivo Excel"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Relatório Detalhado')
            # Aba de atestados
            df_atest = df_final[df_final['Qtd_Atestados'] > 0][['Matricula','Nome_Funcionario','Status','Valor_Premio','Qtd_Atestados','Detalhes']]
            if not df_atest.empty:
                df_atest.to_excel(writer, index=False, sheet_name='Atestados')
        st.download_button("Baixar Excel Executivo", output.getvalue(), "relatorio_executivo.xlsx")

processar()
