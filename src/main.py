import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.units import cm
import os
from utils import create_template

def generate_bulletins(excel_file):
    if not os.path.exists(excel_file):
        print(f"Erro: Arquivo {excel_file} não encontrado.")
        return

    df = pd.read_excel(excel_file)
    
    # Validação básica
    required_cols = ["Turma", "Nome do Aluno", "Unidade Curricular", "Nota", "Frequência Geral (%)"]
    if not all(col in df.columns for col in required_cols):
        print(f"Erro: O Excel deve conter as colunas: {required_cols}")
        return

    output_pdf = "boletins_alunos.pdf"
    doc = SimpleDocTemplate(output_pdf, pagesize=A4, rightMargin=2*cm, leftMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm)
    elements = []
    styles = getSampleStyleSheet()
    
    # Estilos customizados
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], alignment=1, spaceAfter=20)
    info_style = ParagraphStyle('InfoStyle', parent=styles['Normal'], fontSize=12, spaceAfter=10)
    
    # Agrupar por aluno
    alunos = df.groupby("Nome do Aluno")
    
    for nome, grupo in alunos:
        turma = grupo["Turma"].iloc[0]
        frequencia = grupo["Frequência Geral (%)"].iloc[0]
        
        # Cabeçalho do Boletim
        elements.append(Paragraph(f"BOLETIM ESCOLAR", title_style))
        elements.append(Spacer(1, 0.5*cm))
        
        elements.append(Paragraph(f"<b>Aluno:</b> {nome}", info_style))
        elements.append(Paragraph(f"<b>Turma:</b> {turma}", info_style))
        elements.append(Paragraph(f"<b>Frequência Geral:</b> {frequencia}%", info_style))
        elements.append(Spacer(1, 0.5*cm))
        
        # Tabela de Notas
        data = [["Unidade Curricular", "Nota"]]
        for _, row in grupo.iterrows():
            data.append([row["Unidade Curricular"], f"{row['Nota']:.1f}"])
        
        table = Table(data, colWidths=[12*cm, 3*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 2*cm))
        
        # Assinaturas
        elements.append(Paragraph("_" * 40, info_style))
        elements.append(Paragraph("Assinatura da Coordenação", info_style))
        
        # Quebra de página para o próximo aluno
        elements.append(PageBreak())

    doc.build(elements)
    print(f"PDF gerado com sucesso: {output_pdf}")

if __name__ == "__main__":
    template_path = create_template()
    generate_bulletins(template_path)
