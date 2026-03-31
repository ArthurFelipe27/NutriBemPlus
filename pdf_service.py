import os
from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.colors import HexColor 
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.utils import simpleSplit 

def limpar_valor(val):
    if val is None: return ""
    s = str(val).strip()
    if s.lower() in ['nan', 'nat', 'none']: return ""
    return s.replace('\n', ' ').replace('\r', '')

class PdfService:
    @staticmethod
    def gerar_etiquetas(lista_pacientes, caminho_saida):
        """Desenha a matriz de etiquetas de dieta."""
        c = canvas.Canvas(caminho_saida, pagesize=A4)
        largura, altura = 95*mm, 52*mm
        gap = 3*mm
        colunas, linhas = 2, 5
        
        for i, p in enumerate(lista_pacientes):
            if i > 0 and i % (colunas * linhas) == 0: c.showPage()
            pos = i % (colunas * linhas)
            x = 10*mm + ((pos % colunas) * (largura + 5*mm))
            y = A4[1] - 10*mm - (((pos // colunas) + 1) * (altura + gap))
            PdfService._desenhar_etiqueta_individual(c, x, y, largura, altura, p)
        
        c.save()

    @staticmethod
    def _desenhar_etiqueta_individual(c, x, y, w, h, p):
        TAMANHO_FONTE = 9
        c.setStrokeColorRGB(0, 0, 0); c.setLineWidth(1); c.rect(x, y, w, h)
        cor_header = HexColor('#355a31'); c.setFillColor(cor_header); c.setStrokeColor(cor_header)
        c.roundRect(x + 1*mm, y + h - 15*mm - 1*mm, w - 2*mm, 15*mm, 3*mm, fill=1, stroke=0)
        
        try:
            if os.path.exists("logo.png"): 
                c.drawImage("logo.png", x + 3*mm, y + h - 13*mm, width=12*mm, height=12*mm, mask='auto', preserveAspectRatio=True)
        except: pass
        
        c.setFillColor(colors.white); c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(x + w/2 + 5*mm, y + h - 6*mm, "SILVA E TEIXEIRA")
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(x + w/2 + 5*mm, y + h - 9.5*mm, "IDENTIFICAÇÃO DE DIETAS")
        c.drawCentredString(x + w/2 + 5*mm, y + h - 12.5*mm, "PARA PACIENTES")
        c.setFillColor(colors.black)
        
        margem_esq = x + 3*mm
        nome = limpar_valor(p.get('NOME DO PACIENTE', ''))
        enf = limpar_valor(p.get('ENFERMARIA', ''))
        leito = limpar_valor(p.get('LEITO', ''))
        dieta = limpar_valor(p.get('DIETA', ''))
        obs = limpar_valor(p.get('OBSERVAÇÕES', ''))
        if len(obs) > 120: obs = obs[:117] + "..."
        
        cursor_y = y + h - 20*mm 
        c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq, cursor_y, "PACIENTE:")
        c.setFont("Helvetica", TAMANHO_FONTE); c.drawString(margem_esq + 19*mm, cursor_y, nome[:40]) 
        cursor_y -= 5*mm 
        c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq, cursor_y, "ENF:")
        c.setFont("Helvetica", TAMANHO_FONTE); c.drawString(margem_esq + 9*mm, cursor_y, enf[:15])
        c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq + 45*mm, cursor_y, "LEITO:")
        c.setFont("Helvetica", TAMANHO_FONTE); c.drawString(margem_esq + 57*mm, cursor_y, leito)
        cursor_y -= 5*mm 
        c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq, cursor_y, "TIPO DE DIETA:")
        cursor_y -= 4*mm 
        c.setFont("Helvetica", TAMANHO_FONTE)
        
        for linha in simpleSplit(dieta, "Helvetica", TAMANHO_FONTE, w - 6*mm): 
            c.drawString(margem_esq, cursor_y, linha)
            cursor_y -= 4*mm 
            
        cursor_y -= 1*mm 
        c.setFont("Helvetica-Bold", TAMANHO_FONTE)
        c.drawRightString(x + w - 3*mm, cursor_y, f"DATA: {datetime.now().strftime('%d/%m/%Y')}")
        c.setFont("Helvetica-Bold", TAMANHO_FONTE); c.drawString(margem_esq, cursor_y, "OBS:")
        cursor_y -= 4*mm 
        c.setFont("Helvetica", TAMANHO_FONTE)
        
        for linha in simpleSplit(obs, "Helvetica", TAMANHO_FONTE, w - 6*mm): 
            if cursor_y < y + 2*mm: break 
            c.drawString(margem_esq, cursor_y, linha)
            cursor_y -= 4*mm

    @staticmethod
    def gerar_relatorio_mesclado(df, caminho_saida, titulo):
        """Gera a tabela principal de Enfermarias com células mescladas para as alas."""
        doc = SimpleDocTemplate(caminho_saida, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()
        
        try:
            if os.path.exists("logo.png"): 
                elements.append(Image("logo.png", width=15*mm, height=15*mm, hAlign='CENTER'))
                elements.append(Spacer(1, 10))
        except: pass
        
        estilo_sub = ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=TA_CENTER, fontSize=10)
        elements.append(Paragraph(f"<b>{titulo}</b> - Emitido em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", estilo_sub))
        elements.append(Spacer(1, 15))
        
        estilo_celula = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=9, leading=11)
        data = [['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']]
        
        df = df.reset_index(drop=True)

        for index, row in df.iterrows():
            nome = limpar_valor(row['NOME DO PACIENTE'])
            enf = limpar_valor(row['ENFERMARIA'])
            leito = limpar_valor(row['LEITO'])
            dieta = limpar_valor(row['DIETA'])
            obs = limpar_valor(row['OBSERVAÇÕES'])
            data.append([Paragraph(enf, estilo_celula), leito, Paragraph(nome, estilo_celula), Paragraph(dieta, estilo_celula), Paragraph(obs, estilo_celula)])
        
        t = Table(data, colWidths=[110, 50, 250, 160, 200], repeatRows=1)
        
        estilo = [
            ('BACKGROUND',(0,0),(-1,0),colors.Color(0.2,0.6,0.3)), 
            ('TEXTCOLOR',(0,0),(-1,0),colors.white), 
            ('GRID',(0,0),(-1,-1),0.5,colors.grey), 
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.whitesmoke,colors.white])
        ]
        
        # Lógica de Mesclagem
        if len(df) > 0:
            offset = 1 
            start_row = offset
            current_group = df.iloc[0]['ENFERMARIA']
            
            for i in range(1, len(df)):
                val = df.iloc[i]['ENFERMARIA']
                if val != current_group:
                    end_row = (i - 1) + offset
                    if end_row > start_row:
                        estilo.append(('SPAN', (0, start_row), (0, end_row)))
                        estilo.append(('VALIGN', (0, start_row), (0, end_row), 'MIDDLE'))
                    
                    current_group = val
                    start_row = i + offset
            
            end_row = (len(df) - 1) + offset
            if end_row > start_row:
                estilo.append(('SPAN', (0, start_row), (0, end_row)))
                estilo.append(('VALIGN', (0, start_row), (0, end_row), 'MIDDLE'))
                
        t.setStyle(TableStyle(estilo))
        elements.append(t)
        elements.append(Spacer(1,40))
        elements.append(Paragraph("_"*60, ParagraphStyle('A', parent=styles['Normal'], alignment=TA_CENTER)))
        elements.append(Paragraph("<b>NUTRICIONISTA RESPONSÁVEL</b>", ParagraphStyle('A', parent=styles['Normal'], alignment=TA_CENTER)))
        
        doc.build(elements)

    @staticmethod
    def gerar_relatorio_simples(df, caminho_saida, titulo):
        """Gera relatórios lineares (UTI, UPA)."""
        doc = SimpleDocTemplate(caminho_saida, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elements = []
        styles = getSampleStyleSheet()
        
        try:
            if os.path.exists("logo.png"): 
                elements.append(Image("logo.png", width=35*mm, height=35*mm, hAlign='CENTER'))
                elements.append(Spacer(1, 5))
        except: pass
        
        elements.append(Paragraph(f"DATA: {datetime.now().strftime('%d/%m/%Y')}", ParagraphStyle('DT', parent=styles['Normal'], alignment=TA_CENTER, fontSize=12)))
        elements.append(Paragraph(titulo, ParagraphStyle('T', parent=styles['Title'], alignment=TA_CENTER, textColor=colors.darkblue)))
        
        data = [['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']]
        style = ParagraphStyle('C', parent=styles['Normal'], fontSize=10)
        
        for i, r in df.iterrows():
            leito = limpar_valor(r['LEITO'])
            nome = limpar_valor(r['NOME DO PACIENTE'])
            dieta = limpar_valor(r['DIETA'])
            obs = limpar_valor(r['OBSERVAÇÕES'])
            data.append([leito, Paragraph(nome, style), Paragraph(dieta, style), Paragraph(obs, style)])
            
        t = Table(data, colWidths=[60, 280, 200, 230], repeatRows=1)
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.darkblue), ('TEXTCOLOR',(0,0),(-1,0),colors.white), ('GRID',(0,0),(-1,-1),0.5,colors.grey), ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.whitesmoke,colors.white])]))
        
        elements.append(t)
        elements.append(Spacer(1, 30))
        
        rodape = [[Paragraph("<b>Nº DE FUNCIONÁRIOS DIA:</b> _______________", styles['Normal']), Paragraph("<b>Nº DE FUNCIONÁRIOS NOITE:</b> _______________", styles['Normal'])]]
        elements.append(Table(rodape, colWidths=[350,350]))
        elements.append(Spacer(1, 30))
        elements.append(Paragraph("_"*60, ParagraphStyle('A', parent=styles['Normal'], alignment=TA_CENTER)))
        elements.append(Paragraph("<b>NUTRICIONISTA RESPONSÁVEL</b>", ParagraphStyle('A', parent=styles['Normal'], alignment=TA_CENTER)))
        
        doc.build(elements)