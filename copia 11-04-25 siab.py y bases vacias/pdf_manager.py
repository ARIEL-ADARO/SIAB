from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, Image
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
print("PDF_MANAGER REAL CARGADO")

class PDFManager:

    def __init__(self, usuario_actual, logo_path=None):
        self.usuario = usuario_actual
        self.styles = getSampleStyleSheet()
        self.logo_path = os.path.join(BASE_DIR, "logo.png")

        # Ruta absoluta segura
        if logo_path:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            self.logo_path = os.path.join(base_dir, logo_path)
        else:
            self.logo_path = None

    # ==================================================
    # PDF UNIFICADO INSTITUCIONAL
    # ==================================================
    def crear_pdf_unificado(self, file_path, elementos, titulo):

        doc = SimpleDocTemplate(
            file_path,
            pagesize=A4,
            rightMargin=50,
            leftMargin=50,
            topMargin=120,
            bottomMargin=60
        )

        def header_footer(canvas, doc):

            canvas.saveState()

            width, height = A4

            left = doc.leftMargin + 20
            right = width - doc.rightMargin - 20

            # Punto base superior real
            top = height - 40

            # =========================
            # LOGO
            # =========================
            canvas.drawImage(
                self.logo_path,
                left,
                top - 60,
                width=50,
                height=50,
                preserveAspectRatio=True,
                mask='auto'
            )

            # =========================
            # TÍTULO PRINCIPAL
            # =========================
            canvas.setFont("Helvetica-Bold", 16)
            canvas.drawString(
                left + 65,
                top - 25,
                "Bomberos Voluntarios de Almafuerte"
            )

            # =========================
            # SUBTÍTULO
            # =========================
            canvas.setFont("Helvetica", 11)
            canvas.drawString(
                left + 65,
                top - 40,
                "SIAB - Sistema Informático Automatizado de Bomberos"
            )

            # =========================
            # PRIMERA LÍNEA
            # =========================
            canvas.line(
                left,
                top - 70,
                right,
                top - 70
            )

            # =========================
            # DATOS USUARIO
            # =========================
            canvas.setFont("Helvetica", 8)

            fecha = datetime.now().strftime("%d/%m/%Y")
            hora = datetime.now().strftime("%H:%M")

            usuario_txt = (
                f"Usuario: {self.usuario['legajo']} - "
                f"{self.usuario['apellido']} {self.usuario['nombre']}"
            )

            canvas.drawString(
                left,
                top - 85,
                usuario_txt
            )

            canvas.drawRightString(
                right,
                top - 85,
                f"Fecha: {fecha}  Hora: {hora}"
            )

            # =========================
            # SEGUNDA LÍNEA
            # =========================
            canvas.line(
                left,
                top - 90,
                right,
                top - 90
            )

            # =========================
            # PIE
            # =========================
            canvas.drawCentredString(
                width / 2,
                20,
                f"Página {doc.page}"
            )

            canvas.restoreState()

        from reportlab.platypus import Spacer

        story = []

        story.append(Spacer(1, 15))  # empuja el título un poco hacia abajo

        from reportlab.lib.styles import ParagraphStyle

        titulo_style = ParagraphStyle(
            'TituloInforme',
            parent=self.styles['Normal'],
            fontSize=13,
            leading=15,
            spaceBefore=0,
            spaceAfter=12
        )

        story.append(Paragraph(f"<b>{titulo}</b>", titulo_style))

        story.extend(elementos)

        doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)

    # ==================================================
    # PDF TABLA ESTÁNDAR
    # ==================================================
    def exportar_tabla(self, file_path, datos, headers, titulo="", resumen=None):

        from reportlab.platypus import Paragraph
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.platypus import Table, TableStyle
        from reportlab.lib import colors

        elementos = []

        # =========================
        # ESTILO DE CELDAS
        # =========================
        style = ParagraphStyle(
            name='tabla',
            fontSize=7,
            leading=9,
        )

        data = []

        # Header en negrita
        header_row = []
        for h in headers:
            header_row.append(Paragraph(f"<b>{h}</b>", style))
        data.append(header_row)

        # Filas con wrap automático
        for fila in datos:
            nueva_fila = []
            for celda in fila:
                nueva_fila.append(Paragraph(str(celda), style))
            data.append(nueva_fila)

        # =========================
        # ANCHOS DE COLUMNA
        # =========================
        page_width = A4[0]
        usable_width = page_width - 100  # márgenes 50 y 50

        colWidths = []

        for i, h in enumerate(headers):
            if h.lower() == "actividad":
                colWidths.append(usable_width * 0.30)  # 30% del ancho total
            else:
                colWidths.append(usable_width * 0.70 / (len(headers) - 1))

        tabla = Table(
            data,
            repeatRows=1,
            colWidths=colWidths
        )

        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#a50000")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 2),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ]))

        elementos.append(tabla)

        # =========================
        # BLOQUE RESUMEN (SEPARADO)
        # =========================
        if resumen:
            from reportlab.platypus import Spacer
            from reportlab.lib.units import cm

            elementos.append(Spacer(1, 0.7 * cm))

            estilo_resumen_titulo = ParagraphStyle(
                name='resumen_titulo',
                fontSize=10,
                leading=12,
                spaceAfter=6
            )

            estilo_resumen = ParagraphStyle(
                name='resumen',
                fontSize=9,
                leading=11
            )

            elementos.append(Paragraph("<b>RESUMEN</b>", estilo_resumen_titulo))
            elementos.append(Spacer(1, 0.3 * cm))

            for linea in resumen:
                elementos.append(Paragraph(linea, estilo_resumen))

        self.crear_pdf_unificado(file_path, elementos, titulo)