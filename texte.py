import streamlit as st
import pandas as pd
import plotly.express as px

# --- Configurações ---
NOME_ARQUIVO = 'custos.xlsx'
COLUNA_GRUPO_PLANEJAMENTO = 'Grp.planej.manutenç.'
COLUNA_VALOR = 'Valor'
COLUNA_TIPO_ORDEM = 'Tipo de Ordem'
COLUNA_L_INSTALACAO = 'Local de Instalação'
COLUNA_NUM_ORDEM = 'Ordem'
COLUNA_CAB_ORDEM = 'Cabeçalho da ordem'
# Observação: detectaremos variações do nome da coluna de cabeçalho mais abaixo

# Layout da página
st.set_page_config(layout="wide", page_title="Visualização de Custos com Filtro")
st.title("📊 Visualização de Custos PIAN Novembro")

try:
    # Carregar dados
    df = pd.read_excel(NOME_ARQUIVO)

    # Detectar colunas com possíveis variações de nomes
    def detectar_coluna(possiveis_nomes):
        for nome in possiveis_nomes:
            if nome in df.columns:
                return nome
        return None

    COLUNA_CAB_ORDEM = detectar_coluna([
        'Cabeçalho da ordem',
        
    ])
    COLUNA_EQUIP = detectar_coluna(['Equipamento', 'EQUIPAMENTO'])

    # Garantir que a coluna de valor é numérica e remover linhas sem valor
    df[COLUNA_VALOR] = pd.to_numeric(df[COLUNA_VALOR], errors='coerce')
    df.dropna(subset=[COLUNA_VALOR], inplace=True)

    st.success(f"Base de dados '{NOME_ARQUIVO}' carregada com sucesso!")

    # 🎨 Filtros na barra lateral com opção "Todos"
    st.sidebar.markdown("## 🔍 Filtros")

    # Multiselect para Grupo de Planejamento
    st.sidebar.markdown("**Grupo de Planejamento:**")
    grupos_disponiveis = sorted(df[COLUNA_GRUPO_PLANEJAMENTO].dropna().unique().tolist()) if COLUNA_GRUPO_PLANEJAMENTO in df.columns else []
    opcoes_grupo = ['Todos'] + grupos_disponiveis
    grupos_selecionados = st.sidebar.multiselect("Selecione um ou mais grupos", opcoes_grupo, default=['Todos'])

    # Multiselect para Tipo de Ordem
    st.sidebar.markdown("**Tipo de Ordem:**")
    tipos_disponiveis = sorted(df[COLUNA_TIPO_ORDEM].dropna().unique().tolist()) if COLUNA_TIPO_ORDEM in df.columns else []
    opcoes_tipo = ['Todos'] + tipos_disponiveis
    tipos_selecionados = st.sidebar.multiselect("Selecione um ou mais tipos", opcoes_tipo, default=['Todos'])

    # Assinatura no final da sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown(
        """
        <div style='font-size: 13px; line-height: 1.6; color: #555;'>
            <strong>Desenvolvido por Rafael Brandão</strong><br>
            Gerência de Manutenção | <em>Maintenance Management</em><br>
            Auxiliar Técnico II | <em>Assistant Technician</em>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Aplicar filtros
    df_filtrado = df.copy()
    if COLUNA_GRUPO_PLANEJAMENTO in df.columns and 'Todos' not in grupos_selecionados:
        df_filtrado = df_filtrado[df_filtrado[COLUNA_GRUPO_PLANEJAMENTO].isin(grupos_selecionados)]
    if COLUNA_TIPO_ORDEM in df.columns and 'Todos' not in tipos_selecionados:
        df_filtrado = df_filtrado[df_filtrado[COLUNA_TIPO_ORDEM].isin(tipos_selecionados)]

    # Subtotal
    subtotal_valor = df_filtrado[COLUNA_VALOR].sum()
    st.metric(
        label="Subtotal de Gastos (Grupos e Tipos Selecionados)",
        value=f"R$ {subtotal_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )

    # 📊 Gráfico de colunas por Grupo de Planejamento
    st.subheader("📊 Custos por Grupo de Planejamento")
    if COLUNA_GRUPO_PLANEJAMENTO in df_filtrado.columns:
        df_gp = (
            df_filtrado.groupby(COLUNA_GRUPO_PLANEJAMENTO)[COLUNA_VALOR]
            .sum()
            .sort_values()
            .reset_index()
        )
        fig_gp = px.bar(
            df_gp,
            x=COLUNA_GRUPO_PLANEJAMENTO,
            y=COLUNA_VALOR,
            labels={COLUNA_VALOR: 'Total R$', COLUNA_GRUPO_PLANEJAMENTO: 'Grupo'},
            color_discrete_sequence=['#003f5c']
        )
        st.plotly_chart(fig_gp, use_container_width=True)
    else:
        st.warning(f"A coluna '{COLUNA_GRUPO_PLANEJAMENTO}' não foi encontrada na base de dados.")

    # 📊 Gráfico de pizza por PROCESSO e por Tipo de Ordem
    st.subheader("📊 Distribuição por Processo e Tipo de Ordem")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("💸 Distribuição de Custos por Processo")
        nome_col_processo = detectar_coluna(['PROCESSO', 'PROCESSOS'])
        if nome_col_processo:
            df_proc = df_filtrado.groupby(nome_col_processo)[COLUNA_VALOR].sum().reset_index()
            fig_proc = px.pie(
                df_proc,
                names=nome_col_processo,
                values=COLUNA_VALOR,
                hole=0.4,
                color_discrete_sequence=['#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087']
            )
            st.plotly_chart(fig_proc, use_container_width=True)
        else:
            st.warning("A coluna 'PROCESSO'/'PROCESSOS' não foi encontrada na base de dados.")

    with col2:
        st.markdown("📌 Distribuição por Tipo de Ordem")
        if COLUNA_TIPO_ORDEM in df_filtrado.columns:
            df_pizza = df_filtrado[COLUNA_TIPO_ORDEM].value_counts().reset_index()
            df_pizza.columns = ['Tipo de Ordem', 'Quantidade']
            fig_pizza = px.pie(
                df_pizza,
                names='Tipo de Ordem',
                values='Quantidade',
                hole=0.4,
                color_discrete_sequence=['#003f5c', '#2f4b7c', '#665191', '#a05195', '#d45087']
            )
            st.plotly_chart(fig_pizza, use_container_width=True)
        else:
            st.warning(f"A coluna '{COLUNA_TIPO_ORDEM}' não foi encontrada na base de dados.")

    # 💰 Top 10 maiores custos por local de instalação e por equipamento
    st.subheader("💰 Top 10 Maiores Custos")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("🏭 Por Local de Instalação")
        if COLUNA_L_INSTALACAO in df_filtrado.columns:
            df_top10_local = (
                df_filtrado.groupby(COLUNA_L_INSTALACAO)[COLUNA_VALOR]
                .sum()
                .nlargest(10)
                .reset_index()
            )
            df_top10_local.insert(0, "Posição", range(1, len(df_top10_local) + 1))
            df_top10_local[COLUNA_VALOR] = df_top10_local[COLUNA_VALOR].apply(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            st.table(df_top10_local)
        else:
            st.warning(f"A coluna '{COLUNA_L_INSTALACAO}' não foi encontrada na base de dados.")

    with col2:
        st.markdown("⚙️ Por Equipamento")
        if COLUNA_EQUIP:
            df_top10_equip = (
                df_filtrado.groupby(COLUNA_EQUIP)[COLUNA_VALOR]
                .sum()
                .nlargest(10)
                .reset_index()
            )
            df_top10_equip.insert(0, "Posição", range(1, len(df_top10_equip) + 1))
            df_top10_equip[COLUNA_VALOR] = df_top10_equip[COLUNA_VALOR].apply(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            st.table(df_top10_equip)
        else:
            st.warning("A coluna 'Equipamento'/'EQUIPAMENTO' não foi encontrada na base de dados.")
       

    # 📋 Tabela: Top 10 Ordens com Maiores Custos (exibindo número da ordem e cabeçalho)
    st.subheader("📋 Top 10 Ordens com Maiores Custos")
    if COLUNA_NUM_ORDEM in df_filtrado.columns and COLUNA_CAB_ORDEM in df_filtrado.columns:
        df_ordens = (
            df_filtrado.groupby([COLUNA_NUM_ORDEM, COLUNA_CAB_ORDEM])[COLUNA_VALOR]
            .sum()
            .nlargest(10)
            .reset_index()
        )
        df_ordens.insert(0, "Posição", range(1, len(df_ordens) + 1))
        df_ordens[COLUNA_VALOR] = df_ordens[COLUNA_VALOR].apply(
            lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

        # Renomear colunas para exibição clara
        df_exibir = df_ordens.rename(columns={
            COLUNA_NUM_ORDEM: 'Ordem',
            COLUNA_CAB_ORDEM: 'Cabeçalho da ordem',
            COLUNA_VALOR: 'Total R$'
        })
        st.table(df_exibir[['Posição', 'Ordem', 'Cabeçalho da ordem', 'Total R$']])
    else:
        # Caso só exista uma das colunas, ainda entregamos a visualização
        if COLUNA_NUM_ORDEM in df_filtrado.columns:
            df_ordens = (
                df_filtrado.groupby(COLUNA_NUM_ORDEM)[COLUNA_VALOR]
                .sum()
                .nlargest(10)
                .reset_index()
            )
            df_ordens.insert(0, "Posição", range(1, len(df_ordens) + 1))
            df_ordens[COLUNA_VALOR] = df_ordens[COLUNA_VALOR].apply(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            df_ordens = df_ordens.rename(columns={COLUNA_NUM_ORDEM: 'Ordem', COLUNA_VALOR: 'Total R$'})
            st.info("Exibindo apenas 'Ordem' porque a coluna de cabeçalho não foi encontrada.")
            st.table(df_ordens[['Posição', 'Ordem', 'Total R$']])
        elif COLUNA_CAB_ORDEM in df_filtrado.columns:
            df_ordens = (
                df_filtrado.groupby(COLUNA_CAB_ORDEM)[COLUNA_VALOR]
                .sum()
                .nlargest(10)
                .reset_index()
            )
            df_ordens.insert(0, "Posição", range(1, len(df_ordens) + 1))
            df_ordens[COLUNA_VALOR] = df_ordens[COLUNA_VALOR].apply(
                lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            )
            df_ordens = df_ordens.rename(columns={COLUNA_CAB_ORDEM: 'Cabeçalho da Ordem', COLUNA_VALOR: 'Total R$'})
            st.info("Exibindo apenas 'Cabeçalho da Ordem' porque a coluna 'Ordem' não foi encontrada.")
            st.table(df_ordens[['Posição', 'Cabeçalho da ordem', 'Total R$']])
        else:
            st.warning("Nem 'ORDEM' nem 'Cabeçalho da Ordem' foram encontrados na base de dados.")

    # Rodapé com assinatura
    st.markdown("<br><br><hr>", unsafe_allow_html=True)
    st.markdown(
        """
        <div style='text-align: center; font-size: 13px; color: #888;'>
            <strong>Desenvolvido por Rafael Brandão</strong><br>
            Gerência de Manutenção | <em>Maintenance Management</em><br>
            Auxiliar Técnico II | <em>Assistant Technician</em>
        </div>
        """,
        unsafe_allow_html=True
    )

except FileNotFoundError:
    st.error(f"ERRO: O arquivo '{NOME_ARQUIVO}' não foi encontrado.")
    st.info("Verifique se o arquivo está na mesma pasta do seu script.")
except Exception as e:
    st.error(f"Ocorreu um erro ao carregar ou processar a planilha: {e}")
