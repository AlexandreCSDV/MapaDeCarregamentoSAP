import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfgen.canvas import Canvas
from datetime import datetime
from collections import OrderedDict
from PyPDF2 import PdfMerger
import io

# Função para gerar PDF a partir de um DataFrame
def dataframe_to_pdf(dataframe_dict, buffer):
    pdf = SimpleDocTemplate(buffer, pagesize=landscape(A4), topMargin=6, bottomMargin=6)
    elements = []
    c = Canvas("temp.pdf")
    styles = getSampleStyleSheet()
    title_style = styles["Heading2"]
    subtitle_style = styles["Heading2"]
    centered_subtitle_style = ParagraphStyle(name="CenteredSubtitle", parent=subtitle_style, alignment=1)
    centered_title_style = ParagraphStyle(name="CenteredTitle", parent=title_style, alignment=1)
    
    zona_entrega_mais_comum = df['Zona Entrega'].mode()[0]
    total_nfs = df['Nota'].nunique()
    peso_por_status = df.groupby('Status')['KG'].sum().round(2).to_dict()
    status_peso_text = ' \\\\ '.join([f"{status}: {peso} KG" for status, (peso) in peso_por_status.items()])
    
    title_text_line1 = f"Relatório de Carregamento - Zona de Entrega: {zona_entrega_mais_comum}"
    title_text_line2 = f"Peso Total: {subtotal_total} KG - NFs Totais: {total_nfs}"
    title_text_line3 = f"{status_peso_text}"
    
    header_elements = [
        Paragraph(title_text_line1, centered_title_style),
        Spacer(1, -0.3 * inch),
        Paragraph(title_text_line2, centered_title_style),
        Spacer(1, -0.3 * inch),
        Paragraph(title_text_line3, centered_title_style),
        Spacer(1, 0.00005 * inch)
    ]
    
    header_table = Table([[header_elements]], colWidths=[landscape(A4)[0]])
    header_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), colors.white)]))
    
    subtitle_line1 = f"Notas Dentro do Prazo"
    subtitle_elements = [
        Paragraph(subtitle_line1,title_style),
        Spacer(1, -0.3 * inch)
    ]
    
    subheader_table = Table([[subtitle_elements]], colWidths=[landscape(A4)[0]])
    
    first_pdf = "primeiro"
    
    for almoxarifado, dataframe in dataframe_dict.items():
        if first_pdf == "primeiro":
            elements.append(header_table)
            first_pdf = "segundo"
        
        subtotal_peso = round(dataframe['KG'].sum(), 2)
        subtotal_notas = dataframe['Nota'].nunique()
        subtitle_text = f"{almoxarifado} - Peso: {subtotal_peso} KG - Notas: {subtotal_notas}"
        elements.append(Paragraph(subtitle_text, subtitle_style))
        elements.append(Spacer(1, -0.1 * inch))
        
        cols_para_impressao = ['Data de Entrada da Nota', 'Nota', 'Número da Etiqueta Única', 'CTRC', 'Peso', 'Quantidade de volumes', 'Prioridade',
                               'ONU', 'Fornecedor', 'Almoxarifado', 'Mercadoria', 'Dt Final Entrega', 'Dt Lim Embarque']
        #['Chegada', 'Nota', 'Etq. Unica', 'CTE', 'KG', 'Vol', 'Prior', 'ONU', 'Remetente',
        #                      'Almoxarifado', 'Mercadoria', 'Data Entrega', 'End. WMS', 'Lim. Embarque', 'Status']
    
        data = dataframe[cols_para_impressao].values.tolist()
        headers = dataframe[cols_para_impressao].columns.tolist()
        
        col_widths = [50, 50, 40, 45, 30, 10, 20, 20, 120, 120, 90, 50, 70, 50, 60]
        table = Table([headers] + data, colWidths=col_widths)
        
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.black),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 0), 7),
            ('FONTSIZE', (0, 1), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 0.7),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0.7),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.gray)
        ])
        
        table.setStyle(style)
        
        elements.append(table)
        elements.append(Spacer(1, 0.0001 * inch))
    
    pdf.build(elements)

# Função para ordenar DataFrame por almoxarifado ou região WMS
def ordenar_dataframe(df):
    if 'Região WMS' in df.columns:
        ordem_regiao_wms = ['Quimico','Cantilever', 'Praça Externa 1', 'Praça Externa 2', 'Praça - Rua 00 - Posição 01']
        
        dataframes_por_regiao_wms = dict(tuple(df.groupby('Região WMS')))
        
        dataframes_ordenados_regiao_wms = OrderedDict({
            regiao: dataframes_por_regiao_wms[regiao]
            for regiao in ordem_regiao_wms if regiao in dataframes_por_regiao_wms
        })
        
        dataframes_ordenados_regiao_wms.update({
            regiao: df for regiao, df in dataframes_por_regiao_wms.items()
            if regiao not in ordem_regiao_wms
        })
        
        return dataframes_ordenados_regiao_wms
    
    else:
        return dict(tuple(df.groupby('Almoxarifado')))

# Interface do Streamlit
st.title("Gerador de PDF de Carregamento")
uploaded_file = st.file_uploader("Escolha um arquivo Excel", type="xlsx")
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    
    # Verifica se a coluna 'Região WMS' existe no DataFrame
    if 'Região WMS' in df.columns:
        colunas_para_selecao = ['Chegada', 'Nota Numero', 'Etiqueta Unica', 'Chave Conhecimento', 'Peso Nota', 'Nota Volumes',
                                'Tp. Solicitacao Coleta', 'Nº ONU', 'Razao Remetente', 'Almox. Destino', 'Mercadoria Descricao',
                                'Limite Entregar (Definitivo)', 'Endereco WMS', 'Data Limite Embarque', 'Zona Entrega', 'Região WMS']
    else:
        colunas_para_selecao = ['Data de Entrada da Nota', 'Nota', 'Número da Etiqueta Única', 'CTRC', 'Peso', 'Quantidade de volumes', 'Prioridade',
                               'ONU', 'Fornecedor', 'Almoxarifado', 'Mercadoria', 'Dt Final Entrega', 'Dt Lim Embarque']
        
        #['Chegada', 'Nota Numero', 'Etiqueta Unica', 'Chave Conhecimento', 'Peso Nota', 'Nota Volumes',
         #                       'Tp. Solicitacao Coleta', 'Nº ONU', 'Razao Remetente', 'Almox. Destino', 'Mercadoria Descricao',
          #                      'Limite Entregar (Definitivo)', 'Endereco WMS', 'Data Limite Embarque', 'Zona Entrega']

    novos_nomes = {
        'Data de Entrada da Nota': 'Chegada',
        'Nota': 'Nota Numero',
        'Número da Etiqueta Única': 'Etiqueta Unica',
        'CTRC': 'Chave Conhecimento',
        'Peso': 'Peso Nota',
        'Quantidade de Volumes': 'Nota Volumes',
        'Prioridade': 'Tp. Solicitacao Coleta',
        'ONU': 'Nº ONU',
        'Fornecedor': 'Razao Remetente',
        'Almxoarifado': 'Almox. Destino',
        'Mercadoria': 'Mercadoria Descricao',
        'Dt Final Entrega': 'Limite Entregar (Definitivo)',
        'Dt Lim Embarque': 'Data Limite Embarque'
        }
    
        df = df.rename(columns=novos_nomes)

    # Seleciona as colunas no DataFrame
    df = df[colunas_para_selecao]
    
    # Realiza as manipulações no DataFrame
    df = df.iloc[:-2]
    df['Chave Conhecimento'] = df['Chave Conhecimento'].fillna('')
    df['Tp. Solicitacao Coleta'] = df['Tp. Solicitacao Coleta'].str.slice(0, 1)
    df['Mercadoria Descricao'] = df['Mercadoria Descricao'].str.slice(0, 15)
    df['Almox. Destino'] = df['Almox. Destino'].str.slice(0, 25)
    df['Razao Remetente'] = df['Razao Remetente'].str.slice(0, 25)
    df['Limite Entregar (Definitivo)'] = pd.to_datetime(df['Limite Entregar (Definitivo)'], errors='coerce')
    df['Limite Entregar (Definitivo)'] = df['Limite Entregar (Definitivo)'].dt.strftime('%d/%m/%y')
    df['Chegada'] = pd.to_datetime(df['Chegada'], errors='coerce')
    df['Chegada'] = df['Chegada'].dt.strftime('%d/%m/%y %H:%M')
    df['Data Limite Embarque'] = pd.to_datetime(df['Data Limite Embarque'], errors='coerce')
    df['Data Limite Embarque'] = df['Data Limite Embarque'].dt.strftime('%d/%m/%y')
    df['Nota Numero'] = df['Nota Numero'].astype(int)
    df['Nota Volumes'] = df['Nota Volumes'].astype(int)
    df['Etiqueta Unica'] = pd.to_numeric(df['Etiqueta Unica'], errors='coerce').fillna('').astype('string')
    df['Etiqueta Unica'] = df['Etiqueta Unica'].apply(lambda x: x[:-2] if x.endswith('.0') else x).replace('nan', '')
    df['Nº ONU'] = pd.to_numeric(df['Nº ONU'], errors='coerce').fillna('').astype('string')
    df['Nº ONU'] = df['Nº ONU'].apply(lambda x: x[:-2] if x.endswith('.0') else x).replace('nan', '')
    # df['Endereco WMS'] = df['Endereco WMS'].fillna('')
    df['Peso Nota'] = df['Peso Nota'].round(2)
    
novos_nomes = {
    'Chegada': 'Chegada',
    'Nota Numero': 'Nota',
    'Etiqueta Unica': 'Etq. Unica',
    'Chave Conhecimento': 'CTE',
    'Peso Nota': 'KG',
    'Nota Volumes': 'Vol',
    'Tp. Solicitacao Coleta': 'Prior',
    'Nº ONU': 'ONU',
    'Razao Remetente': 'Remetente',
    'Almox. Destino': 'Almoxarifado',
    'Mercadoria Descricao': 'Mercadoria',
    'Limite Entregar (Definitivo)': 'Data Entrega',
    'Data Limite Embarque': 'Lim. Embarque',
    'Endereco WMS': 'End. WMS'
    }
    
    df = df.rename(columns=novos_nomes)
    df['Lim. Embarque'] = pd.to_datetime(df['Lim. Embarque'], format='%d/%m/%y', errors='coerce')
    df = df.sort_values(by=['Lim. Embarque', 'Etq. Unica', 'Almoxarifado'])
    subtotal_total = round(df['KG'].sum(), 2)
    df.loc[:, 'Status'] = ''
    today = datetime.today().date()
    
    for index, row in df.iterrows():
        try:
            data_embarque = pd.to_datetime(row['Lim. Embarque'], format='%d/%m/%y').date()
        except ValueError:
            data_embarque = None
        try:
            data_entrega = pd.to_datetime(row['Data Entrega'], format='%d/%m/%y').date()
        except ValueError:
            data_entrega = None
        if data_embarque == today:
            df.at[index, 'Status'] = 'Embarque Hoje'
        elif pd.isna(row['Data Entrega']):
            df.at[index, 'Status'] = 'S/ Data Entrega'
        elif data_embarque < today:
            df.at[index, 'Status'] = 'Embarque Atrasado'
        else:
            df.at[index, 'Status'] = 'Dentro do Prazo'
    
    df['Lim. Embarque'] = df['Lim. Embarque'].dt.strftime('%d/%m/%Y')
    
    # Processamento dos DataFrames e geração dos PDFs
    df_hoje = df.loc[df['Status'] == 'Dentro do Prazo']
    df_atrasado = df.loc[df['Status'] != 'Dentro do Prazo']
    df_atrasado = ordenar_dataframe(df_atrasado)
    df_hoje = ordenar_dataframe(df_hoje)
    
    # Buffer para armazenar os PDFs
    buffer_atrasado = io.BytesIO()
    buffer_hoje = io.BytesIO()
    
    # Gerar os PDFs
    dataframe_to_pdf(df_atrasado, buffer_atrasado)
    dataframe_to_pdf(df_hoje, buffer_hoje)
    
    # Mesclar os PDFs
    buffer_atrasado.seek(0)
    buffer_hoje.seek(0)
    merger = PdfMerger()
    merger.append(buffer_atrasado)
    merger.append(buffer_hoje)
    
    # Salvar o PDF final em um buffer
    output_buffer = io.BytesIO()
    merger.write(output_buffer)
    merger.close()

    data_hoje = datetime.now().strftime('%d%b%y').lower()
    zona_entrega_mais_comum = df['Zona Entrega'].mode()[0]
    zona_almoxarifado = df['Zona Entrega'].mode()[0] if 'Região WMS' in df.columns else df['Almoxarifado'].mode()[0]
    
    # Botão para download do PDF
    st.download_button(
        label="Baixar PDF",
        data=output_buffer.getvalue(),
        file_name=f"{zona_almoxarifado}_{data_hoje}.pdf",
        mime="application/pdf"
    )
