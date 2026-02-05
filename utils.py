import streamlit as st
import pandas as pd
import io
from datetime import datetime

def salvar_alteracoes(idx, novo_status, novo_valor, nova_obs, nome):
    """Função auxiliar para salvar alterações"""
    st.session_state.modified_df.at[idx, 'Status'] = novo_status
    st.session_state.modified_df.at[idx, 'Valor_Premio'] = novo_valor
    st.session_state.modified_df.at[idx, 'Observacoes'] = nova_obs
    st.session_state.expanded_item = idx
    st.session_state.last_saved = nome
    st.session_state.show_success = True

def editar_valores_status(df):
    if 'modified_df' not in st.session_state:
        st.session_state.modified_df = df.copy()
    
    if 'expanded_item' not in st.session_state:
        st.session_state.expanded_item = None
        
    if 'show_success' not in st.session_state:
        st.session_state.show_success = False
        
    if 'last_saved' not in st.session_state:
        st.session_state.last_saved = None
    
    st.subheader("Filtro Principal")
    
    status_options = ["Todos", "Tem direito", "Não tem direito", "Aguardando decisão"]
    
    status_principal = st.selectbox(
        "Selecione o status para visualizar:",
        options=status_options,
        index=0,
        key="status_principal_filter_unique"
    )
    
    df_filtrado = st.session_state.modified_df.copy()
    if status_principal != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Status'] == status_principal]
    
    st.subheader("Buscar Funcionários")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        matricula_busca = st.text_input("Buscar por Matrícula", key="matricula_search_unique")
    with col2:
        nome_busca = st.text_input("Buscar por Nome", key="nome_search_unique")
    with col3:
        ordem = st.selectbox(
            "Ordenar por:",
            options=["Nome (A-Z)", "Nome (Z-A)", "Matrícula (Crescente)", "Matrícula (Decrescente)"],
            key="ordem_select_unique"
        )
    
    if matricula_busca:
        df_filtrado = df_filtrado[df_filtrado['Matricula'].astype(str).str.contains(matricula_busca)]
    if nome_busca:
        df_filtrado = df_filtrado[df_filtrado['Nome'].str.contains(nome_busca, case=False)]
    
    # Ordenação
    if ordem == "Nome (A-Z)":
        df_filtrado = df_filtrado.sort_values('Nome')
    elif ordem == "Nome (Z-A)":
        df_filtrado = df_filtrado.sort_values('Nome', ascending=False)
    elif ordem == "Matrícula (Crescente)":
        df_filtrado = df_filtrado.sort_values('Matricula')
    elif ordem == "Matrícula (Decrescente)":
        df_filtrado = df_filtrado.sort_values('Matricula', ascending=False)
    
    # Métricas
    st.subheader("Métricas do Filtro Atual")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Funcionários exibidos", len(df_filtrado))
    with col2:
        st.metric("Total com direito", len(df_filtrado[df_filtrado['Status'] == 'Tem direito']))
    with col3:
        st.metric("Valor total dos prêmios", f"R$ {df_filtrado['Valor_Premio'].sum():,.2f}")
    
    # Mostrar mensagem de sucesso se houver
    if st.session_state.show_success:
        st.success(f"✅ Alterações salvas com sucesso para {st.session_state.last_saved}!")
        st.session_state.show_success = False
    
    # Editor de dados por linhas individuais
    st.subheader("Editor de Dados")
    
    for idx, row in df_filtrado.iterrows():
        with st.expander(
            f"🧑‍💼 {row['Nome']} - Matrícula: {row['Matricula']}", 
            expanded=st.session_state.expanded_item == idx
        ):
            col1, col2 = st.columns(2)
            
            with col1:
                novo_status = st.selectbox(
                    "Status",
                    options=status_options[1:],
                    index=status_options[1:].index(row['Status']) if row['Status'] in status_options[1:] else 0,
                    key=f"status_{idx}_{row['Matricula']}"
                )
                
                novo_valor = st.number_input(
                    "Valor do Prêmio",
                    min_value=0.0,
                    max_value=1000.0,
                    value=float(row['Valor_Premio']),
                    step=50.0,
                    format="%.2f",
                    key=f"valor_{idx}_{row['Matricula']}"
                )
            
            with col2:
                nova_obs = st.text_area(
                    "Observações",
                    value=row.get('Observacoes', ''),
                    key=f"obs_{idx}_{row['Matricula']}"
                )
            
            if st.button("Salvar Alterações", key=f"save_{idx}_{row['Matricula']}"):
                salvar_alteracoes(idx, novo_status, novo_valor, nova_obs, row['Nome'])
    
    # Botões de ação geral
    st.subheader("Ações Gerais")
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Reverter Todas as Alterações", key="revert_all_unique"):
            st.session_state.modified_df = df.copy()
            st.session_state.expanded_item = None
            st.session_state.show_success = False
            st.warning("⚠️ Todas as alterações foram revertidas!")
    
    with col2:
        if st.button("Exportar Arquivo Final", key="export_unique"):
            output = exportar_novo_excel(st.session_state.modified_df)
            if output:
                st.download_button(
                    label="📥 Baixar Arquivo Excel",
                    data=output,
                    file_name="funcionarios_premios.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_unique"
                )
            else:
                st.error("Erro ao gerar o arquivo Excel.")
    
    return st.session_state.modified_df

def exportar_novo_excel(df):
    try:
        output = io.BytesIO()
        
        # Garantir que cada funcionário tenha apenas uma linha no dataframe final
        # Agrupando por Matricula e conservando as informações relevantes
        if 'Matricula' in df.columns and len(df) > 0:
            # Garantir que o DataFrame já esteja agrupado por Matrícula (deve estar, após calcular_premio)
            if df['Matricula'].duplicated().any():
                st.warning("Foram encontradas múltiplas linhas por funcionário. Agrupando automaticamente...")
                
                # Funções para agregação
                def agregar_detalhes(x):
                    # Juntar todos os detalhes de afastamentos únicos
                    detalhes = []
                    for detalhe in x:
                        if isinstance(detalhe, str) and detalhe:
                            for d in detalhe.split(';'):
                                d = d.strip()
                                if d and d not in detalhes:
                                    detalhes.append(d)
                    return "; ".join(detalhes) if detalhes else ""
                
                def priorizar_status(x):
                    # Prioridade: Não tem direito > Aguardando decisão > Tem direito
                    if "Não tem direito" in x.values:
                        return "Não tem direito"
                    elif "Aguardando decisão" in x.values:
                        for status in x.values:
                            if isinstance(status, str) and "Aguardando decisão" in status:
                                return status  # Retorna com os detalhes de atraso
                        return "Aguardando decisão"
                    else:
                        return "Tem direito"
                
                def maior_valor(x):
                    return x.max()
                
                def primeiro_valor(x):
                    return x.iloc[0] if not x.empty else ""
                
                # Definir agregações por coluna
                agregacoes = {
                    'Nome': 'first',
                    'Cargo': 'first',
                    'Local': 'first',
                    'Horas_Mensais': 'first',
                    'Data_Admissao': 'first',
                    'Status': priorizar_status,
                    'Valor_Premio': maior_valor,
                    'Detalhes_Afastamentos': agregar_detalhes,
                    'Observações': 'first' if 'Observações' in df.columns else None,
                    'Observacoes': 'first' if 'Observacoes' in df.columns else None
                }
                
                # Remover colunas que não existem no DataFrame
                agregacoes = {k: v for k, v in agregacoes.items() if k in df.columns}
                
                # Agrupar o DataFrame
                df = df.groupby('Matricula').agg(agregacoes).reset_index()

        # Categorizar os funcionários por status
        df_tem_direito = df[df['Status'].str.contains('Tem direito', na=False)].copy()
        df_nao_tem_direito = df[df['Status'].str.contains('Não tem direito', na=False)].copy()
        df_aguardando_decisao = df[df['Status'].str.contains('Aguardando decisão', na=False)].copy()

        # Criar o arquivo Excel
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Aba com os funcionários com direito
            if not df_tem_direito.empty:
                df_tem_direito.to_excel(writer, index=False, sheet_name='Tem Direito')
            else:
                st.warning("Nenhum funcionário com direito foi encontrado.")

            # Aba com os funcionários sem direito
            if not df_nao_tem_direito.empty:
                df_nao_tem_direito.to_excel(writer, index=False, sheet_name='Não Tem Direito')
            else:
                st.warning("Nenhum funcionário sem direito foi encontrado.")

            # Aba com os funcionários aguardando decisão
            if not df_aguardando_decisao.empty:
                df_aguardando_decisao.to_excel(writer, index=False, sheet_name='Aguardando Decisão')
            else:
                st.warning("Nenhum funcionário aguardando decisão foi encontrado.")

            # Aba com o resumo
            resumo_data = [
                ['RESUMO DO PROCESSAMENTO'],
                [f'Data de Geração: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}'],
                [''],
                ['Métricas Gerais'],
                [f'Total de Funcionários Processados: {len(df)}'],
                [f'Total de Funcionários com Direito: {len(df_tem_direito)}'],
                [f'Total de Funcionários sem Direito: {len(df_nao_tem_direito)}'],
                [f'Total de Funcionários Aguardando Decisão: {len(df_aguardando_decisao)}'],
            ]
            
            pd.DataFrame(resumo_data).to_excel(
                writer,
                index=False,
                header=False,
                sheet_name='Resumo'
            )

        output.seek(0)
        return output.getvalue()

    except Exception as e:
        st.error(f"Erro ao exportar relatório: {e}")
        return None
