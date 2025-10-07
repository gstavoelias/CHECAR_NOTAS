import streamlit as st
import pandas as pd
import re

def extrair_nota_por_produto_valor_qtd(row):
    """
    Extrai NFE com múltiplos padrões.
    Suporta: NFE/NOTA, UN/MT, texto completo/truncado
    """
    if pd.isna(row['ns1:infCpl']) or pd.isna(row['ns1:cProd']):
        return None
    
    texto = str(row['ns1:infCpl'])
    codigo_prod = str(row['ns1:cProd']).strip()
    valor_linha = round(float(row['ns1:vProd']), 2)
    qtd_linha = float(row['ns1:qCom'])
    
    # Padrão 1: CODIGO NFE NUM ... VR R$ VALOR - QTD UN
    pattern1 = rf'{re.escape(codigo_prod)}\s+NFE\s+(\d+)[^,]*?VR\s+R\$\s+([\d.]+)\s*-\s*([\d.]+)\s*UN'
    
    # Padrão 2: CODIGO NFE NUM ... VR R$ VALOR - QTD MT
    pattern2 = rf'{re.escape(codigo_prod)}\s+NFE\s+(\d+)[^,]*?VR\s+R\$\s+([\d.]+)\s*-\s*([\d.]+)\s*MT'
    
    # Tenta padrões completos com validação
    for pattern in [pattern1, pattern2]:
        for match in re.finditer(pattern, texto):
            nfe = match.group(1)
            valor_match = round(float(match.group(2)), 2)
            qtd_match = float(match.group(3))
            
            if abs(valor_match - valor_linha) < 0.01 and abs(qtd_match - qtd_linha) < 0.01:
                return nfe.zfill(6)
    
    # Padrão 3: NOTA189DIA (formato alternativo sem espaços)
    pattern3 = rf'NOTA(\d+)DIA'
    match = re.search(pattern3, texto, re.IGNORECASE)
    if match:
        return match.group(1).zfill(6)
    
    # Padrão 4: Truncado - só CODIGO NFE NUM (sem validação)
    pattern4 = rf'{re.escape(codigo_prod)}\s+NFE\s+(\d+)'
    match = re.search(pattern4, texto)
    if match:
        return match.group(1).zfill(6)
    
    return None

# Configuração da página
st.set_page_config(page_title="Verificador de NFE", page_icon="📋", layout="wide")

# Título
st.title("📋 Verificador de Produtos sem Nota de Origem")
st.markdown("Envie uma planilha Excel para identificar produtos que não possuem nota de origem no texto.")

# Upload do arquivo
uploaded_file = st.file_uploader("Escolha o arquivo Excel", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Carregar e processar
        with st.spinner("Processando planilha..."):
            df = pd.read_excel(uploaded_file).dropna(subset="TECSCI")
            df["nota_origem"] = df.apply(extrair_nota_por_produto_valor_qtd, axis=1)
        
        # Filtrar produtos sem nota
        produtos_sem_nota = df[df['nota_origem'].isna()][['ns1:cProd', 'ns1:qCom', 'ns1:vProd']].copy()
        
        # Renomear colunas para exibição
        produtos_sem_nota.columns = ['Código Produto', 'Quantidade', 'Valor']
        
        # Exibir resultados
        st.success(f"✅ Planilha processada com sucesso!")
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total de Produtos", len(df))
        with col2:
            st.metric("Produtos SEM Nota", len(produtos_sem_nota))
        
        # Mostrar produtos sem nota
        if len(produtos_sem_nota) > 0:
            st.warning(f"⚠️ **{len(produtos_sem_nota)} produto(s) não possuem nota de origem no texto:**")
            st.dataframe(produtos_sem_nota, use_container_width=True, hide_index=True)
            
            # Botão para download
            csv = produtos_sem_nota.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Baixar produtos sem nota (CSV)",
                data=csv,
                file_name="produtos_sem_nota.csv",
                mime="text/csv"
            )
            st.markdown("---")
            st.subheader("📄 Texto Completo - Informações Complementares")
            texto_completo = df['ns1:infCpl'].iloc[0] if not df['ns1:infCpl'].isna().all() else "Nenhum texto encontrado"
            st.text_area(
                label="Texto para conferência:",
                value=texto_completo,
                height=300,
                disabled=True
            )
        else:
            st.success("🎉 Todos os produtos possuem nota de origem!")
            
    except Exception as e:
        st.error(f"❌ Erro ao processar arquivo: {str(e)}")
        st.info("Certifique-se de que o arquivo contém as colunas: TECSCI, ns1:cProd, ns1:qCom, ns1:vProd, ns1:infCpl")

else:
    st.info("👆 Faça upload de uma planilha Excel para começar")
    
    # Instruções
    with st.expander("ℹ️ Como usar"):
        st.markdown("""
        1. Clique no botão acima para fazer upload do arquivo Excel
        2. A planilha será processada automaticamente
        3. Os produtos sem nota de origem serão exibidos na tabela
        4. Você pode baixar a lista de produtos sem nota em formato CSV
        
        **Colunas necessárias no Excel:**
        - TECSCI
        - ns1:cProd (Código do Produto)
        - ns1:qCom (Quantidade)
        - ns1:vProd (Valor)
        - ns1:infCpl (Informações Complementares)
        """)