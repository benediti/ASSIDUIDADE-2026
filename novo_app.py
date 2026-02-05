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
    horas = row['horas']
    salario = row['salario']
    status = "Tem direito"
    valor = VALOR_BASE
    detalhes = []
    # Filtra ausências do funcionário
    aus = ausencias[ausencias['Matricula'] == row['Matricula']]
    # Normaliza nomes e status
    aus['Afastamento_Lower'] = aus['Afastamentos'].str.lower().str.strip()
    aus['Status_Lower'] = aus.iloc[:,1].astype(str).str.lower().str.strip() if aus.shape[1] > 1 else ''


    # 1. Se houver "Atraso" ou "Férias" na coluna de afastamentos, aplicar lógica especial
    if aus['Afastamento_Lower'].str.contains('atraso').any():
        return pd.Series({
            'Valor_Premio': 0,
            'Status': 'Aguardando decisão',
            'Detalhes': 'Afastamento: Atraso',
            'Qtd_Atestados': aus['Afastamento_Lower'].str.contains('atestado').sum()
        })

    # Conta atestados
    dias_atestado = aus['Afastamento_Lower'].str.contains('atestado').sum()

    # 2. Férias: descontar dias proporcionalmente se houver "Férias" na coluna
    dias_ferias = 0
    mask_ferias = aus['Afastamento_Lower'].str.contains('ferias')
    if mask_ferias.any():
        # Tenta extrair quantidade de dias da coluna de status, se for número
        try:
            dias_ferias = pd.to_numeric(aus.loc[mask_ferias, 'Status_Lower'], errors='coerce').sum()
        except Exception:
            dias_ferias = 0

    # Regras de cálculo padrão
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
    # Desconto proporcional de férias
    if dias_ferias > 0 and valor > 0:
        desconto = min(dias_ferias / 30, 1)
        valor = round(valor * (1 - desconto), 2)
        detalhes.append(f"Desconto férias: {dias_ferias} dias")
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
    # Função para encontrar colunas por possíveis nomes
    def encontrar_coluna(possibilidades):
        for nome in df_func.columns:
            nome_limpo = nome.lower().replace(' ', '').replace('ç','c').replace('ã','a').replace('é','e').replace('í','i').replace('ê','e').replace('ó','o').replace('á','a').replace('ú','u')
            if nome_limpo in possibilidades:
                return nome
        return None

    col_data_adm = encontrar_coluna(['datadeadmissao','dataadmissao','admissao'])
    col_horas = encontrar_coluna(['qtdhorasmensais','horasmensais','horas','qtdhoras'])
    col_salario = encontrar_coluna(['salariomesatual','salariomesatu','salariomes','salario','saláriomesatual','saláriomesatu','saláriomes'])

    if not col_data_adm:
        st.error("Coluna de data de admissão não encontrada na base de funcionários.")
        return
    if not col_horas:
        st.error("Coluna de horas mensais não encontrada na base de funcionários.")
        return
    if not col_salario:
        st.error("Coluna de salário não encontrada na base de funcionários.")
        return

    # Filtra por data de admissão
    df_func[col_data_adm] = pd.to_datetime(df_func[col_data_adm], errors='coerce', dayfirst=True)
    df_func = df_func[df_func[col_data_adm] <= pd.to_datetime(data_limite)]

    # Aplica cálculo, passando horas e salário explicitamente
    resultado = df_func.apply(
        lambda row: calcular_premio(
            pd.Series({
                **row,
                'horas': row[col_horas],
                'salario': row[col_salario]
            }),
            df_aus
        ),
        axis=1
    )
    df_final = pd.concat([df_func, resultado], axis=1)
    st.subheader("Relatório de Prêmios Calculados")
    st.dataframe(df_final)
    # Exportação Excel
    if st.button("Exportar Relatório Executivo Excel"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Relatório Detalhado')
            # Aba de atestados (detecta nome da coluna de nome do funcionário)
            col_nome = encontrar_coluna(['nome','nomefuncionario','nome_funcionario']) or 'Nome'
            cols_atest = ['Matricula', col_nome, 'Status', 'Valor_Premio', 'Qtd_Atestados', 'Detalhes']
            df_atest = df_final[df_final['Qtd_Atestados'] > 0][cols_atest if all(c in df_final.columns for c in cols_atest) else df_final.columns]
            if not df_atest.empty:
                df_atest.to_excel(writer, index=False, sheet_name='Atestados')
        st.download_button("Baixar Excel Executivo", output.getvalue(), "relatorio_executivo.xlsx")

processar()
