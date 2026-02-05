from datetime import datetime
import pandas as pd
import streamlit as st
import os
import logging
import io
from utils import editar_valores_status, exportar_novo_excel  # Importar funções do utils.py

# Configuração do logging
logging.basicConfig(
    filename='sistema_premios.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def carregar_tipos_afastamento():
    # Verificar se o diretório 'data' existe e criá-lo se não existir
    if not os.path.exists("data"):
        os.makedirs("data")
        
    if os.path.exists("data/tipos_afastamento.pkl"):
        return pd.read_pickle("data/tipos_afastamento.pkl")
    return pd.DataFrame({"tipo": [], "categoria": []})

def salvar_tipos_afastamento(df):
    # Verificar se o diretório 'data' existe e criá-lo se não existir
    if not os.path.exists("data"):
        os.makedirs("data")
    df.to_pickle("data/tipos_afastamento.pkl")
    
def processar_ausencias(df):
    # Renomear colunas e configurar dados iniciais
    df = df.rename(columns={
        "Matrícula": "Matricula",
        "Centro de Custo": "Centro_de_Custo",
        "Ausência Integral": "Ausencia_Integral",
        "Ausência Parcial": "Ausencia_Parcial",
        "Data de Demissão": "Data_de_Demissao"
    })
    
    df['Matricula'] = pd.to_numeric(df['Matricula'], errors='coerce')
    df = df.dropna(subset=['Matricula'])
    df['Matricula'] = df['Matricula'].astype(int)
    
    # Processar faltas marcadas com X na coluna Falta
    df['Faltas'] = df['Falta'].fillna('')
    df['Faltas'] = df['Faltas'].apply(lambda x: 1 if str(x).upper().strip() == 'X' else 0)
    
    # Detectar faltas não justificadas na coluna Ausência Parcial
    df['Tem_Falta_Nao_Justificada'] = df['Ausencia_Parcial'].fillna('').astype(str).str.contains('Falta não justificada', case=False)
    
    def converter_para_horas(tempo):
        if pd.isna(tempo) or tempo == '' or tempo == '00:00':
            return 0
        try:
            if ':' in str(tempo):
                horas, minutos = map(int, str(tempo).split(':'))
                return horas + minutos / 60
            return 0
        except:
            return 0
    
    df['Horas_Atraso'] = df['Ausencia_Parcial'].apply(converter_para_horas)
    
    # Processar informações de atraso na coluna Ausência Parcial
    df['Tem_Atraso'] = df['Ausencia_Parcial'].fillna('').astype(str).str.contains('Atraso', case=False)
    
    # Adicionar tipos de afastamento à coluna Afastamentos quando encontrados na coluna Ausência Parcial
    df['Afastamentos'] = df.apply(
        lambda row: row['Afastamentos'] + '; Atraso' if row['Tem_Atraso'] and 'Atraso' not in str(row['Afastamentos']) 
        else row['Afastamentos'],
        axis=1
    )
    
    # Adicionar Falta não justificada aos afastamentos quando encontrado na coluna Ausência Parcial ou Falta é X
    df['Afastamentos'] = df.apply(
        lambda row: row['Afastamentos'] + '; Falta não justificada' 
        if (row['Tem_Falta_Nao_Justificada'] or row['Faltas'] == 1) and 'Falta não justificada' not in str(row['Afastamentos']) 
        else row['Afastamentos'],
        axis=1
    )
    
    df['Afastamentos'] = df['Afastamentos'].fillna('').astype(str)
    
    # Armazenar os valores de atraso para uso posterior
    df['Atrasos'] = df.apply(
        lambda row: row['Ausencia_Parcial'] if row['Tem_Atraso'] else '',
        axis=1
    )
    
    # Carregar tipos de afastamento
    df_tipos = carregar_tipos_afastamento()
    tipos_conhecidos = df_tipos['tipo'].unique() if not df_tipos.empty else []

    # Identificar afastamentos desconhecidos
    df['Afastamentos_Desconhecidos'] = df['Afastamentos'].apply(
        lambda x: '; '.join([a for a in x.split(';') if a.strip() not in tipos_conhecidos])
    )
    
    # Classificar status
    def classificar_status(afastamentos):
        afastamentos_list = afastamentos.split(';')
        if any(a.strip() in afastamentos_impeditivos for a in afastamentos_list):
            return "Não Tem Direito"
        elif any(a.strip() in afastamentos_decisao for a in afastamentos_list):
            return "Aguardando Decisão"
        return "Tem Direito"
    
    afastamentos_impeditivos = [
        "Licença Maternidade", "Atestado Médico", "Férias", "Feriado", "Falta não justificada"
    ]
    afastamentos_decisao = ["Abono", "Atraso"]
    
    df['Status'] = df['Afastamentos'].apply(classificar_status)
    
    # Retornar DataFrame atualizado
    return df

def calcular_cesta_basica(df_funcionarios, df_ausencias, data_limite_admissao):
    VALOR_BASE = 315.00
    SALARIO_LIMITE = 2720.86
    resultados = []
    df_funcionarios['Data_Admissao'] = pd.to_datetime(df_funcionarios['Data_Admissao'], format='%d/%m/%Y')
    df_funcionarios = df_funcionarios[df_funcionarios['Data_Admissao'] <= pd.to_datetime(data_limite_admissao)]
    for idx, func in df_funcionarios.iterrows():
        matricula = func['Matricula']
        ausencias = df_ausencias[df_ausencias['Matricula'] == matricula]
        salario = func['Salario_Mes_Atual']
        horas = func['Qtd_Horas_Mensais']
        status = "Tem direito"
        valor = VALOR_BASE
        detalhes = []
        dias_atestado = 0
        falta_injustificada = False
        # Verifica salário
        if salario > SALARIO_LIMITE:
            status = "Não tem direito"
            valor = 0
            detalhes.append("Salário acima do limite")
        # Verifica ausências
        else:
            if not ausencias.empty:
                # Falta injustificada
                if 'Tem_Falta_Nao_Justificada' in ausencias.columns and ausencias['Tem_Falta_Nao_Justificada'].any():
                    status = "Não tem direito"
                    valor = 0
                    detalhes.append("Falta injustificada")
                    falta_injustificada = True
                # Falta marcada com X
                elif 'Faltas' in ausencias.columns and ausencias['Faltas'].sum() > 0:
                    status = "Não tem direito"
                    valor = 0
                    detalhes.append("Falta injustificada (X)")
                    falta_injustificada = True
                # Dias de atestado
                else:
                    # Considera cada linha com "Atestado" na ausência integral/parcial
                    for _, row in ausencias.iterrows():
                        texto = str(row.get('Ausencia_Integral', '')) + ' ' + str(row.get('Ausencia_Parcial', ''))
                        if 'atestado' in texto.lower():
                            dias_atestado += 1
                    if dias_atestado == 1:
                        valor = 240.00
                        detalhes.append("1 dia de atestado")
                    elif dias_atestado == 2:
                        valor = 140.00
                        detalhes.append("2 dias de atestado")
                    elif dias_atestado >= 3:
                        status = "Não tem direito"
                        valor = 0
                        detalhes.append(f"{dias_atestado} dias de atestado")
            # Proporcionalidade férias/afastamento previdenciário
            if status == "Tem direito":
                dias_trabalhados = 30
                if 'Férias' in str(ausencias.get('Afastamentos', '')).title() or 'INSS' in str(ausencias.get('Afastamentos', '')).upper():
                    # Aqui, para simplificação, considera 30 dias no mês, descontando dias de férias/afastamento
                    dias_faltantes = 0
                    for _, row in ausencias.iterrows():
                        if 'férias' in str(row.get('Afastamentos', '')).lower() or 'inss' in str(row.get('Afastamentos', '')).lower():
                            dias_faltantes += 1
                    dias_trabalhados = max(0, 30 - dias_faltantes)
                    valor = round(valor * (dias_trabalhados / 30), 2)
                    detalhes.append(f"Proporcional: {dias_trabalhados} dias trabalhados")
        # Jornada 4h: 50%
        if horas <= 120 and valor > 0:
            valor = round(valor * 0.5, 2)
            detalhes.append("Jornada 4h (50%)")
        resultado = {
            'Matricula': func['Matricula'],
            'Nome': func['Nome_Funcionario'],
            'Cargo': func['Cargo'],
            'Local': func['Nome_Local'],
            'Horas_Mensais': func['Qtd_Horas_Mensais'],
            'Data_Admissao': func['Data_Admissao'],
            'Valor_Premio': valor,
            'Status': status,
            'Detalhes_Afastamentos': "; ".join(detalhes),
            'Observações': ''
        }
        resultados.append(resultado)
    return pd.DataFrame(resultados)

def exportar_excel(df_mostrar, df_funcionarios):
    output = io.BytesIO()
    df_export = df_mostrar.copy()
    df_export['Salario'] = df_funcionarios.set_index('Matricula').loc[df_export['Matricula'], 'Salario_Mes_Atual'].values
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Resultados Detalhados')
        
        relatorio_diretoria = pd.DataFrame([
            ["RELATÓRIO DE PRÊMIOS - VISÃO EXECUTIVA", ""],
            [f"Data do relatório: {datetime.now().strftime('%d/%m/%Y')}", ""],
            ["", ""],
            ["RESUMO GERAL", ""],
            [f"Total de Funcionários Analisados: {len(df_export)}", ""],
            [f"Funcionários com Direito: {len(df_export[df_export['Status'] == 'Tem direito'])}", ""],
            [f"Funcionários Aguardando Decisão: {len(df_export[df_export['Status'].str.contains('Aguardando decisão', na=False)])}", ""],
            [f"Valor Total dos Prêmios: R$ {df_export['Valor_Premio'].sum():,.2f}", ""],
            ["", ""],
            ["DETALHAMENTO POR STATUS", ""],
        ])
        
        for status in df_export['Status'].unique():
            df_status = df_export[df_export['Status'] == status]
            relatorio_diretoria = pd.concat([relatorio_diretoria, pd.DataFrame([
                [f"\nStatus: {status}", ""],
                [f"Quantidade de Funcionários: {len(df_status)}", ""],
                [f"Valor Total: R$ {df_status['Valor_Premio'].sum():,.2f}", ""],
                ["Locais Afetados:", ""],
                [", ".join(df_status['Local'].unique()), ""],
                ["", ""]
            ])])
        
        relatorio_diretoria.to_excel(writer, index=False, header=False, sheet_name='Relatório Executivo')
    
    return output.getvalue()

def main():
    st.set_page_config(page_title="Sistema de Verificação da CESTA BÁSICA II", page_icon="🛒", layout="wide")
    st.title("Sistema de Verificação da CESTA BÁSICA II")
    
    with st.sidebar:
        st.header("Configurações")
        
        data_limite = st.date_input(
            "Data Limite de Admissão",
             help="Funcionários admitidos após esta data não terão direito ao prêmio",
            value=datetime.now(),
            format="DD/MM/YYYY"
        )
        
        st.subheader("Base de Funcionários")
        uploaded_func = st.file_uploader("Carregar base de funcionários", type=['xlsx'])
        
        st.subheader("Base de Ausências")
        uploaded_ausencias = st.file_uploader("Carregar base de ausências", type=['xlsx'])
        
        st.subheader("Tipos de Afastamento")
        uploaded_tipos = st.file_uploader("Atualizar tipos de afastamento", type=['xlsx'])
        
        if uploaded_tipos is not None:
            try:
                df_tipos_novo = pd.read_excel(uploaded_tipos)
                # Verificar se as colunas do arquivo carregado estão corretas
                if 'tipo de afastamento' in df_tipos_novo.columns and 'Direito Pagamento' in df_tipos_novo.columns:
                    # Renomear as colunas para os nomes esperados pelo sistema
                    df_tipos = df_tipos_novo.rename(columns={'tipo de afastamento': 'tipo', 'Direito Pagamento': 'categoria'})
                    salvar_tipos_afastamento(df_tipos)
                    st.success("Tipos de afastamento atualizados!")
                else:
                    st.error("Arquivo deve conter colunas 'tipo de afastamento' e 'Direito Pagamento'")
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {str(e)}")
    
    if uploaded_func is not None and uploaded_ausencias is not None and data_limite is not None:
        try:
            df_funcionarios = pd.read_excel(uploaded_func)
            colunas_esperadas = [
                "Matricula", "Nome_Funcionario", "Cargo", 
                "Codigo_Local", "Nome_Local", "Qtd_Horas_Mensais",
                "Tipo_Contrato", "Data_Termino_Contrato", 
                "Dias_Experiencia", "Salario_Mes_Atual", "Data_Admissao"
            ]
            if len(df_funcionarios.columns) != len(colunas_esperadas):
                st.error(f"Erro: O arquivo de funcionários possui {len(df_funcionarios.columns)} colunas, mas o sistema espera {len(colunas_esperadas)}.\n\nColunas encontradas: {list(df_funcionarios.columns)}\nColunas esperadas: {colunas_esperadas}")
                return
            df_funcionarios.columns = colunas_esperadas

            df_ausencias = pd.read_excel(uploaded_ausencias)
            df_ausencias = processar_ausencias(df_ausencias)
            
            # Verificar e exibir afastamentos desconhecidos
            if not df_ausencias['Afastamentos_Desconhecidos'].str.strip().eq('').all():
                st.warning("Foram encontrados afastamentos desconhecidos na tabela de ausências:")
                st.dataframe(df_ausencias[['Matricula', 'Afastamentos_Desconhecidos']])
                st.info("Atualize os tipos de afastamento para corrigir essas inconsistências.")
            
            df_resultado = calcular_cesta_basica(df_funcionarios, df_ausencias, data_limite)
            
            st.subheader("Resultado do Cálculo de Prêmios")
            
            df_mostrar = df_resultado
            
            # Editar resultados
            df_mostrar = editar_valores_status(df_mostrar)
            
            # Mostrar métricas
            st.metric("Total de Funcionários com Direito", len(df_mostrar[df_mostrar['Status'] == "Tem direito"]))
            st.metric("Total de Funcionários sem Direito", len(df_mostrar[df_mostrar['Status'] == "Não tem direito"]))
            st.metric("Valor Total dos Prêmios", f"R$ {df_mostrar['Valor_Premio'].sum():,.2f}")
            
            # Filtros
            status_filter = st.selectbox("Filtrar por Status", options=["Todos", "Tem direito", "Não tem direito", "Aguardando decisão"])
            if status_filter != "Todos":
                df_mostrar = df_mostrar[df_mostrar['Status'] == status_filter]
            
            nome_filter = st.text_input("Filtrar por Nome")
            if nome_filter:
                df_mostrar = df_mostrar[df_mostrar['Nome'].str.contains(nome_filter, case=False)]
            
            # Mostrar tabela de resultados na interface
            st.dataframe(df_mostrar)
            
            # Exportar resultados
            if st.button("Exportar Resultados para Excel"):
                df_exportar = df_mostrar[df_mostrar['Status'] == "Tem direito"].copy()
                df_exportar['CPF'] = ""  # Adicione lógica para preencher CPF
                df_exportar['CNPJ'] = "65035552000180"  # Adicione lógica para preencher CNPJ
                df_exportar = df_exportar.rename(columns={'Valor_Premio': 'SomaDeVALOR'})
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_exportar.to_excel(writer, index=False, sheet_name='Funcionarios com Direito')
                st.download_button("Baixar Excel", output.getvalue(), "funcionarios_com_direito.xlsx")
        
        except Exception as e:
            st.error(f"Erro ao processar dados: {str(e)}")

if __name__ == "__main__":
    main()
