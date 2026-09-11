"""PDF de manual de uso y de instalación en otra computadora."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent.parent
DOCS_DIR = ROOT / "docs"
LOGO_PATH = ROOT / "assets" / "logo-zofingen-sgr.png"
RED = colors.HexColor("#E32636")
DARK = colors.HexColor("#333333")
GRAY = colors.HexColor("#555555")

MANUAL_PDF = DOCS_DIR / "manual-sgr-tools.pdf"
SETUP_PDF = DOCS_DIR / "como-abrir-sgr-tools.pdf"


def _styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "ManualTitle",
            parent=base["Title"],
            fontName="Helvetica-Bold",
            fontSize=18,
            textColor=DARK,
            alignment=TA_CENTER,
            spaceAfter=4 * mm,
        ),
        "subtitle": ParagraphStyle(
            "ManualSubtitle",
            parent=base["Normal"],
            fontSize=11,
            textColor=GRAY,
            alignment=TA_CENTER,
            spaceAfter=8 * mm,
        ),
        "h1": ParagraphStyle(
            "ManualH1",
            parent=base["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=13,
            textColor=RED,
            spaceBefore=5 * mm,
            spaceAfter=2.5 * mm,
        ),
        "h2": ParagraphStyle(
            "ManualH2",
            parent=base["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=11,
            textColor=DARK,
            spaceBefore=3 * mm,
            spaceAfter=1.5 * mm,
        ),
        "body": ParagraphStyle(
            "ManualBody",
            parent=base["BodyText"],
            fontSize=10,
            leading=14,
            textColor=DARK,
            alignment=TA_JUSTIFY,
            spaceAfter=2 * mm,
        ),
        "step": ParagraphStyle(
            "ManualStep",
            parent=base["BodyText"],
            fontSize=10,
            leading=14,
            textColor=DARK,
            leftIndent=2 * mm,
            spaceAfter=1.2 * mm,
        ),
        "warn": ParagraphStyle(
            "ManualWarn",
            parent=base["BodyText"],
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#8B1A1A"),
            alignment=TA_LEFT,
            spaceAfter=2 * mm,
        ),
        "ok": ParagraphStyle(
            "ManualOk",
            parent=base["BodyText"],
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#1B5E20"),
            spaceAfter=2 * mm,
        ),
        "footer": ParagraphStyle(
            "ManualFooter",
            parent=base["Normal"],
            fontSize=8,
            textColor=GRAY,
            alignment=TA_CENTER,
        ),
    }


def _draw_frame(canvas: Any, document: Any) -> None:
    width, height = A4
    canvas.saveState()
    if LOGO_PATH.is_file():
        canvas.drawImage(
            ImageReader(str(LOGO_PATH)),
            document.leftMargin,
            height - 18 * mm,
            width=36 * mm,
            height=12 * mm,
            preserveAspectRatio=True,
            mask="auto",
        )
    canvas.setFillColor(GRAY)
    canvas.setFont("Helvetica-Bold", 9)
    canvas.drawRightString(
        width - document.rightMargin,
        height - 12 * mm,
        "ZOFINGEN SGR  |  SGR tools",
    )
    canvas.setStrokeColor(RED)
    canvas.setLineWidth(1.2)
    canvas.line(
        document.leftMargin,
        height - 20 * mm,
        width - document.rightMargin,
        height - 20 * mm,
    )
    canvas.setStrokeColor(colors.HexColor("#B0B3B8"))
    canvas.setLineWidth(0.4)
    canvas.line(
        document.leftMargin,
        12 * mm,
        width - document.rightMargin,
        12 * mm,
    )
    canvas.setFillColor(GRAY)
    canvas.setFont("Helvetica", 8)
    canvas.drawCentredString(width / 2, 6 * mm, f"Página {canvas.getPageNumber()}")
    canvas.restoreState()


def _items(styles: dict[str, ParagraphStyle], lines: list[str]) -> ListFlowable:
    bullets = [
        ListItem(Paragraph(line, styles["step"]), leftIndent=8, bulletColor=RED)
        for line in lines
    ]
    return ListFlowable(
        bullets,
        bulletType="bullet",
        start="•",
        leftIndent=12,
        bulletFontName="Helvetica",
        bulletFontSize=10,
        spaceAfter=3 * mm,
    )


def _steps(styles: dict[str, ParagraphStyle], lines: list[str]) -> ListFlowable:
    numbered = [
        ListItem(Paragraph(line, styles["step"]), leftIndent=10)
        for line in lines
    ]
    return ListFlowable(
        numbered,
        bulletType="1",
        leftIndent=16,
        bulletFontName="Helvetica-Bold",
        bulletFontSize=10,
        spaceAfter=3 * mm,
    )


def _build(path: Path, title: str, story: list[Any]) -> Path:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    document = SimpleDocTemplate(
        str(path),
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=26 * mm,
        bottomMargin=18 * mm,
        title=title,
        author="Zofingen SGR",
    )
    document.build(story, onFirstPage=_draw_frame, onLaterPages=_draw_frame)
    return path


def build_user_manual() -> Path:
    s = _styles()
    story: list[Any] = [
        Paragraph("Manual de SGR tools", s["title"]),
        Paragraph("Cómo usar la web, paso a paso, sin vueltas", s["subtitle"]),
        Paragraph("Cómo entrar", s["h1"]),
        Paragraph(
            "Con <b>run.bat</b> abierto, usá Chrome o Edge en esta dirección:",
            s["body"],
        ),
        Paragraph("<b>http://127.0.0.1:5055</b>", s["ok"]),
        Paragraph("Arranque de cada día", s["h1"]),
        _steps(
            s,
            [
                "Entrá a la carpeta <b>sgr-tools</b>.",
                "Hacé doble clic en <b>run.bat</b>.",
                "Se abre una ventana negra. <b>No la cierres</b> mientras uses la web.",
                "Abrí Chrome o Edge y escribí: <b>http://127.0.0.1:5055</b>",
                "Deberías ver el menú con las tarjetas: Contingente, Calendario, Conciliación, Rentas y Rendimientos.",
            ],
        ),
        Paragraph(
            "Si la ventana negra dice que no encuentra Python, instalá Python "
            "y mirá el otro PDF: <b>Cómo abrir SGR tools en tu computadora</b>.",
            s["body"],
        ),
        Paragraph("Menú principal", s["h1"]),
        Paragraph(
            "Arriba podés filtrar <b>Todos</b>, <b>SGR</b> u <b>Operaciones</b>. "
            "Hacé clic en una tarjeta para entrar. En cada pantalla hay un "
            "<b>← Volver</b> para regresar al menú.",
            s["body"],
        ),
        Paragraph("1. Contingente de avales (SGR)", s["h1"]),
        Paragraph("<b>Para qué sirve:</b> ordena el Excel de cuotas, calcula Fecha mesa y Aforo, y te deja filtrar y bajar reportes.", s["body"]),
        Paragraph("Qué hacer", s["h2"]),
        _steps(
            s,
            [
                "Clic en <b>Contingente de avales</b>.",
                "En <b>Archivo Excel</b> elegí el Excel de cuotas (hoja Base).",
                "Si hace falta, cambiá los días de Fecha mesa y el % de Aforo. Si no sabés, dejá lo que ya está.",
                "Clic en <b>Procesar</b>.",
                "Usá los filtros de las columnas (como en Excel).",
                "Descargá Excel o PDF, o mirá el calendario de vencimientos.",
            ],
        ),
        Paragraph(
            "Las filas <b>reclamadas</b> se marcan en rojo. Eso no es un error: es el dato del banco.",
            s["body"],
        ),
        Paragraph("2. Calendario Global (SGR)", s["h1"]),
        Paragraph(
            "<b>Para qué sirve:</b> junta en un solo calendario las rentas y las fechas mesa del contingente.",
            s["body"],
        ),
        _items(
            s,
            [
                "<b>Verde:</b> rentas.",
                "<b>Azul:</b> contingente.",
                "<b>Rojo:</b> contingente reclamado.",
                "<b>Amarillo:</b> evento que cargaste a mano.",
            ],
        ),
        _steps(
            s,
            [
                "Primero procesá <b>Calendario de Rentas</b> y <b>Contingente</b> (si no, este calendario sale vacío o viejo).",
                "Entrá a <b>Calendario Global</b>.",
                "Podés arrastrar un evento a otro día. Te va a pedir confirmación.",
                "Para un evento suelto: completá tipo, denominación y fecha, y agregalo.",
            ],
        ),
        Paragraph("3. Conciliación (Operaciones)", s["h1"]),
        Paragraph("<b>Para qué sirve:</b> compara custodia vs posición y te dice qué coincide y qué no.", s["body"]),
        _steps(
            s,
            [
                "Clic en <b>Conciliación</b>.",
                "Cargá el <b>Archivo de Custodia</b> (extracto).",
                "Cargá el archivo de <b>Posición</b> (tablero).",
                "Clic en <b>Conciliar</b>.",
                "Mirá coincidencias, diferencias y especies que están de un solo lado. Si hace falta, bajá el Excel.",
            ],
        ),
        Paragraph("4. Calendario de Rentas (Operaciones)", s["h1"]),
        Paragraph(
            "<b>Para qué sirve:</b> cruza tu posición con el calendario público de Bolsar y te muestra solo los pagos de lo que tenés en cartera.",
            s["body"],
        ),
        _steps(
            s,
            [
                "Clic en <b>Calendario de Rentas</b>.",
                "Cargá el archivo de <b>Posición</b>.",
                "Clic en <b>Actualizar calendario</b>.",
                "Revisá el calendario. Cada carga nueva actualiza lo guardado: agrega lo que entró y saca lo que ya no está.",
            ],
        ),
        Paragraph("5. Rendimientos (Operaciones)", s["h1"]),
        Paragraph(
            "<b>Para qué sirve:</b> arma la evolución del Patrimonio Neto y el rendimiento mensual, en pesos y en A3500.",
            s["body"],
        ),
        Paragraph("Qué archivos cargar", s["h2"]),
        _items(
            s,
            [
                "<b>Posición:</b> uno o varios Excel de tablero. El <b>+</b> al lado sirve para cargar fecha y PN a mano.",
                "<b>A3500:</b> opcional. Si no cargás nada, consulta la API del BCRA. El <b>+</b> carga un tipo de cambio a mano.",
                "<b>Aportes y retiros:</b> un solo Excel con el formato de esa solapa del reporte (Tipo, Fecha, Inversor, Monto Pesos). El <b>+</b> carga un movimiento a mano. Queda guardado hasta cargar uno nuevo.",
            ],
        ),
        Paragraph("Qué hace el cálculo (en criollo)", s["h2"]),
        Paragraph(
            "<b>Rendimiento ARS</b> = PN del mes - PN del mes anterior - aportes + retiros + pagos. "
            "Los movimientos se toman en el mes de su fecha.<br/>"
            "<b>MoM PN %</b> = Rendimiento ARS / PN pesos del mismo mes.<br/>"
            "<b>MoM PN A3500 %</b> = Rendimiento A3500 / PN A3500 del mismo mes.<br/>"
            "<b>MoM A3500 %</b> = A3500 del mes / A3500 del mes anterior - 1.",
            s["body"],
        ),
        Paragraph(
            "El primer mes no tiene rendimiento ni MoM (no hay mes anterior). "
            "Si falta un mes, se compara contra el último que hay: por ejemplo, junio contra abril si no está mayo.",
            s["body"],
        ),
        Paragraph("Cómo usarlo", s["h2"]),
        _steps(
            s,
            [
                "Cargá los archivos (podés elegir varios y también ir sumando en selecciones distintas).",
                "Si hace falta, clic en <b>+</b> al lado del archivo para una fecha y un PN, un A3500 o un movimiento a mano.",
                "Clic en <b>Procesar</b>.",
                "Cambiá Pesos / PN A3500 arriba del gráfico si querés ver la otra unidad.",
                "Abajo está la tabla. Más abajo, los aportes y retiros guardados (con filtro).",
                "Descargá con el botón <b>Descargar</b> (Excel o PDF).",
            ],
        ),
        Paragraph(
            "El Excel trae fórmulas: PN, fecha y A3500 van como dato; rendimientos y MoM se calculan solos. "
            "En <b>Aportes y retiros</b> cada movimiento trae el A3500 de su fecha y Monto A3500 = pesos / A3500. "
            "Las columnas A3500 del mes suman esa columna. La fecha de cálculo es un mes antes. Si cambiás un monto o el tipo de cambio, Excel recalcula.",
            s["body"],
        ),
        Paragraph(
            "Solo se acumulan archivos de posición del <b>último día hábil de cada mes</b>. Si cargás un día que no es cierre, se usa para el cálculo de ahora pero no queda en el histórico mensual.",
            s["body"],
        ),
        Paragraph("Si algo sale mal", s["h1"]),
        Table(
            [
                [Paragraph("<b>Qué ves</b>", s["body"]), Paragraph("<b>Qué hacer</b>", s["body"])],
                [
                    Paragraph("La página no carga / se ve fea / los botones no hacen nada", s["body"]),
                    Paragraph("Cerrá esa ventana. Confirmá que <b>run.bat</b> sigue abierto. Entrá a http://127.0.0.1:5055", s["body"]),
                ],
                [
                    Paragraph("Dice que falta el archivo", s["body"]),
                    Paragraph("Elegí el Excel otra vez y recargá. No sirve arrastrar un acceso directo vacío.", s["body"]),
                ],
                [
                    Paragraph("Procesar da error de columnas / fechas", s["body"]),
                    Paragraph("Usá el Excel de siempre (no un PDF, no una foto, no un CSV). En contingente tiene que existir la hoja Base.", s["body"]),
                ],
                [
                    Paragraph("Rendimientos sin A3500", s["body"]),
                    Paragraph("Tiene que haber internet para la API del BCRA, o cargá la serie / un A3500 manual.", s["body"]),
                ],
                [
                    Paragraph("No aparecen rentas o el calendario global está vacío", s["body"]),
                    Paragraph("Procesá primero Calendario de Rentas y/o Contingente, y después abrí Calendario Global.", s["body"]),
                ],
                [
                    Paragraph("Los números no coinciden con tu Excel", s["body"]),
                    Paragraph("MoM no se calcula sobre el PN crudo si estás mirando Neto. Fijate qué columna es. Y si falta un mes, el MoM salta al anterior disponible.", s["body"]),
                ],
            ],
            colWidths=[70 * mm, 108 * mm],
        ),
        Spacer(1, 2 * mm),
        Paragraph("Qué no hacer", s["h1"]),
        _items(
            s,
            [
                "No cierres la ventana negra de <b>run.bat</b> mientras trabajás.",
                "No borres la carpeta <b>data</b>: ahí se guardan los últimos resultados.",
                "No esperes que lo que procesaste en tu PC aparezca solo en la PC de un compañero. Cada máquina tiene su copia, salvo que copien también la carpeta <b>data</b>.",
            ],
        ),
        Paragraph(
            "Si después de esto sigue sin andar, no improvises: pedí ayuda y mandá una captura de la ventana negra y de la página.",
            s["body"],
        ),
    ]
    for item in story:
        if isinstance(item, Table):
            item.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), RED),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFF8F8")),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D0D7DE")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            break
    return _build(MANUAL_PDF, "Manual de SGR tools", story)


def build_setup_guide() -> Path:
    s = _styles()
    story: list[Any] = [
        Paragraph("Cómo abrir SGR tools en tu computadora", s["title"]),
        Paragraph("Instructivo para alguien que recibió la carpeta y nunca la usó", s["subtitle"]),
        Paragraph("Qué te tienen que haber pasado", s["h1"]),
        Paragraph(
            "Una carpeta llamada <b>sgr-tools</b> (o un ZIP). Adentro tiene que estar, como mínimo:",
            s["body"],
        ),
        _items(
            s,
            [
                "<b>run.bat</b> (el arranque)",
                "<b>app.py</b>",
                "<b>requirements.txt</b>",
                "los HTML (<b>index.html</b>, <b>rendimientos.html</b>, etc.)",
                "la carpeta <b>backend</b>",
                "la carpeta <b>assets</b>",
            ],
        ),
        Paragraph(
            "Si te pasaron solo un HTML, o un acceso directo, <b>no sirve</b>. Pedí de nuevo la carpeta completa.",
            s["warn"],
        ),
        Paragraph("Paso 1 — Instalá Python (una sola vez)", s["h1"]),
        Paragraph(
            "SGR tools no es una página de internet. Es un programita que corre en tu PC. Para eso hace falta Python.",
            s["body"],
        ),
        _steps(
            s,
            [
                "Abrí un navegador y andá a <b>https://www.python.org/downloads/</b>",
                "Descargá Python 3 (el botón grande amarillo).",
                "Ejecutá el instalador.",
                "<b>IMPORTANTE:</b> tildá la casilla <b>Add python.exe to PATH</b> (o “Agregar Python al PATH”). Si te la saltás, después no arranca.",
                "Clic en Install Now y esperá que termine.",
                "Cerrá y volvé a abrir cualquier ventana negra que tuvieras abierta.",
            ],
        ),
        Paragraph(
            "Para comprobar: tocá la tecla Windows, escribí <b>cmd</b>, Enter, y escribí: <b>python --version</b>. Tiene que aparecer algo tipo Python 3.12. Si dice que no se reconoce el comando, reinstalá Python marcando el PATH.",
            s["body"],
        ),
        Paragraph("Paso 2 — Poné la carpeta en un lugar simple", s["h1"]),
        _steps(
            s,
            [
                "Si te mandaron un ZIP, clic derecho → <b>Extraer todo</b>.",
                "Copiá la carpeta <b>sgr-tools</b> a un lugar fácil, por ejemplo: <b>Escritorio</b> o <b>Documentos</b>.",
                "Evitá dejarla “adentro” de un mail, de WhatsApp o de una carpeta de descargas con nombre raro.",
            ],
        ),
        Paragraph(
            "OneDrive está bien si la carpeta se termina de sincronizar (el iconito verde). Si sigue con flechitas, esperá.",
            s["body"],
        ),
        Paragraph("Paso 3 — Arrancá la web", s["h1"]),
        _steps(
            s,
            [
                "Entrá a la carpeta <b>sgr-tools</b>.",
                "Hacé doble clic en <b>run.bat</b>.",
                "La primera vez tarda: instala unas librerías. Hace falta internet en ese momento.",
                "Cuando esté listo vas a ver: <b>SGR tools -&gt; http://127.0.0.1:5055</b>",
                "Dejá esa ventana negra <b>abierta</b>.",
            ],
        ),
        Paragraph("Paso 4 — Abrí el navegador", s["h1"]),
        _steps(
            s,
            [
                "Abrí <b>Google Chrome</b> o <b>Microsoft Edge</b>.",
                "Arriba, en la barrita de direcciones, escribí: <b>http://127.0.0.1:5055</b>",
                "Enter.",
                "Deberías ver el menú <b>SGR tools</b> con tarjetas.",
            ],
        ),
        Paragraph(
            "No busques SGR tools en Google: no está publicada en internet. "
            "Escribí la dirección completa, con el <b>:5055</b>.",
            s["body"],
        ),
        Paragraph("Paso 5 — Usala", s["h1"]),
        Paragraph(
            "Con la ventana negra abierta y el navegador en http://127.0.0.1:5055 ya podés entrar a cada herramienta. El detalle de cada pantalla está en el <b>Manual de SGR tools</b> (botón al lado del título en el menú).",
            s["body"],
        ),
        Paragraph("Cuando terminás", s["h1"]),
        Paragraph(
            "Cerrá la pestaña del navegador y, si querés, cerrá la ventana negra (o tocá Ctrl+C ahí). Al día siguiente: otra vez <b>run.bat</b> y otra vez la dirección.",
            s["body"],
        ),
        Paragraph("Problemas típicos", s["h1"]),
        Table(
            [
                [Paragraph("<b>Qué pasa</b>", s["body"]), Paragraph("<b>Qué hacer</b>", s["body"])],
                [
                    Paragraph("Al doble clic en run.bat se cierra al toque", s["body"]),
                    Paragraph("Leé el mensaje. Casi siempre falta Python o no está en el PATH. Reinstalá Python tildando Add to PATH.", s["body"]),
                ],
                [
                    Paragraph("Error al instalar dependencias", s["body"]),
                    Paragraph("Necesitás internet la primera vez. Si la red corporativa bloquea pip, pedí a sistemas que permita Python/pip.", s["body"]),
                ],
                [
                    Paragraph("Address already in use / puerto 5055", s["body"]),
                    Paragraph("Ya hay otra ventana de SGR tools abierta. Cerrala o reiniciá la PC y volvé a ejecutar run.bat.", s["body"]),
                ],
                [
                    Paragraph("El navegador no carga nada", s["body"]),
                    Paragraph("Confirmá que la ventana negra sigue abierta y que la URL es http://127.0.0.1:5055 (con http, no https).", s["body"]),
                ],
                [
                    Paragraph("Windows bloquea run.bat", s["body"]),
                    Paragraph("Clic en Más información → Ejecutar de todas formas. No es un virus: es el arranque de la herramienta.", s["body"]),
                ],
                [
                    Paragraph("En tu PC no están los mismos números que en la otra", s["body"]),
                    Paragraph("Es normal: cada computadora guarda su propia carpeta data. Si necesitan lo mismo, copien también la carpeta data.", s["body"]),
                ],
            ],
            colWidths=[70 * mm, 108 * mm],
        ),
        Spacer(1, 3 * mm),
        Paragraph("Resumen de 15 segundos", s["h1"]),
        Paragraph(
            "1) Instalá Python con PATH. 2) Doble clic en <b>run.bat</b> y no cierres la ventana negra. "
            "3) En Chrome o Edge: <b>http://127.0.0.1:5055</b>.",
            s["ok"],
        ),
    ]
    for item in story:
        if isinstance(item, Table):
            item.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), RED),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                        ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFF8F8")),
                        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D0D7DE")),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            break
    return _build(SETUP_PDF, "Cómo abrir SGR tools en tu computadora", story)


def build_all() -> tuple[Path, Path]:
    return build_user_manual(), build_setup_guide()


if __name__ == "__main__":
    manual, setup = build_all()
    print(manual)
    print(setup)
