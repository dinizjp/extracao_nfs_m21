import os
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from io import BytesIO

# Função para formatar datas no formato dd/mm/yyyy
def format_date(date_str):
    try:
        return pd.to_datetime(date_str).strftime('%d/%m/%Y')
    except Exception as e:
        return ''

# Função para tentar ler o conteúdo do arquivo com diferentes codificações
def read_file_content(uploaded_file):
    for encoding in ['utf-8', 'latin-1']:
        try:
            return uploaded_file.read().decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("Não foi possível decodificar o arquivo com as codificações testadas.")

# Função para extrair dados de um arquivo XML
def extract_data_from_xml(file_content):
    data = []
    try:
        # Parse the XML content
        tree = ET.ElementTree(ET.fromstring(file_content))
        root = tree.getroot()
        
        # Define namespaces para encontrar elementos
        namespaces = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

        # Extrair as informações necessárias com tratamento individual de erros
        try:
            razao_social_dest = root.find('.//nfe:dest/nfe:xNome', namespaces).text
        except:
            razao_social_dest = ''

        try:
            data_emissao = format_date(root.find('.//nfe:ide/nfe:dhEmi', namespaces).text[:10])
        except:
            data_emissao = ''

        try:
            numero_nota = int(root.find('.//nfe:ide/nfe:nNF', namespaces).text)
        except:
            numero_nota = ''

        try:
            valor_total = float(root.find('.//nfe:total/nfe:ICMSTot/nfe:vNF', namespaces).text)
        except:
            valor_total = ''

        try:
            primeiro_vencimento = format_date(root.find('.//nfe:cobr/nfe:dup/nfe:dVenc', namespaces).text)
        except:
            primeiro_vencimento = ''

        try:
            razao_social_forn = root.find('.//nfe:emit/nfe:xNome', namespaces).text
        except:
            razao_social_forn = ''
        
        # Verificar se a nota foi cancelada
        try:
            cStat = root.find('.//nfe:protNFe/nfe:infProt/nfe:cStat', namespaces).text
            status_cancelada = 'Sim' if cStat == '101' else 'Não'
        except:
            status_cancelada = 'Não'

        # Adicionar as informações extraídas à lista de dados se algum dado foi encontrado
        if any([razao_social_dest, data_emissao, numero_nota, valor_total, primeiro_vencimento, razao_social_forn]):
            data.append([razao_social_dest, data_emissao, numero_nota, valor_total, primeiro_vencimento, razao_social_forn, status_cancelada])
    except Exception as e:
        data = None

    return data

# Função principal da aplicação
def main():
    st.title("Extração de Informações de Notas Fiscais")
    
    uploaded_files = st.file_uploader("Escolha os arquivos XML", accept_multiple_files=True, type="xml")

    if uploaded_files:
        all_data = []
        error_files = []
        
        for uploaded_file in uploaded_files:
            try:
                file_content = read_file_content(uploaded_file)
                data = extract_data_from_xml(file_content)
                if data:
                    all_data.extend(data)
                else:
                    error_files.append(uploaded_file.name)
            except UnicodeDecodeError as e:
                error_files.append(uploaded_file.name)
        
        if all_data:
            df = pd.DataFrame(all_data, columns=[
                'Razao Social Destinatario', 
                'Data de Emissao', 
                'Numero da Nota', 
                'Valor Total', 
                'Primeiro Vencimento', 
                'Razao Social Fornecedor',
                'Nota Cancelada'
            ])
            df['Valor Total'] = df['Valor Total'].apply(lambda x: f'R${x:,.2f}' if isinstance(x, (int, float)) else x)
            
            st.write("### Dados Extraídos")
            st.dataframe(df)
            
            # Botão para download do DataFrame como Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Notas Fiscais')
                writer.close()
            processed_data = output.getvalue()
            
            st.download_button(
                label="Baixar Planilha",
                data=processed_data,
                file_name='extracao_notas_fiscais.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
        
        if error_files:
            st.write("### Arquivos com Erros")
            for error_file in error_files:
                st.write(error_file)

if __name__ == "__main__":
    main()
