# -*- coding: utf-8 -*-
import base64
import os
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import pypandoc
import pythoncom
import win32com.client
from pdf2docx import Converter as PdfConverter
from PySide6.QtCore import QThread, Qt, QUrl, Signal
from PySide6.QtGui import (
    QAction,
    QColor,
    QDesktopServices,
    QDragEnterEvent,
    QDropEvent,
    QFont,
    QIcon,
    QLinearGradient,
    QPainter,
    QPainterPath,
    QPixmap,
)
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

try:
    from updater import check_for_updates, open_download_page, get_current_version
    UPDATER_AVAILABLE = True
except ImportError:
    UPDATER_AVAILABLE = False
    def get_current_version():
        return "1.0.0"

SUPPORTED_EXTENSIONS = {".md", ".markdown", ".html", ".htm", ".docx", ".pdf", ".txt"}
FILE_DIALOG_FILTER = "支持的文件 (*.md *.markdown *.html *.htm *.docx *.pdf *.txt);;所有文件 (*.*)"
OUTPUT_FORMATS = ["pdf", "docx", "html", "md", "txt"]

FORMAT_META = {
    "pdf": {"title": "PDF", "subtitle": "便携文档", "color": "#DC2626"},
    "docx": {"title": "Word", "subtitle": "可编辑文档", "color": "#2563EB"},
    "html": {"title": "HTML", "subtitle": "网页文件", "color": "#EA580C"},
    "md": {"title": "Markdown", "subtitle": "纯文本标记", "color": "#16A34A"},
    "txt": {"title": "TXT", "subtitle": "纯文本", "color": "#64748B"},
}

CSS_CONTENT = """
body {
    font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
    max-width: 900px;
    margin: 40px auto;
    padding: 20px;
    line-height: 1.8;
    color: #333;
}
h1, h2, h3 { margin-top: 28px; margin-bottom: 16px; }
h1 { font-size: 2em; border-bottom: 2px solid #eee; padding-bottom: 10px; }
h2 { font-size: 1.6em; border-bottom: 1px solid #eee; padding-bottom: 8px; }
code {
    background: #f5f5f5;
    padding: 2px 6px;
    border-radius: 3px;
    font-family: "Consolas", "Monaco", monospace;
    font-size: 0.9em;
}
pre {
    background: #f8f8f8;
    padding: 16px;
    border-radius: 5px;
    overflow-x: auto;
    border: 1px solid #e0e0e0;
}
pre code { background: none; padding: 0; }
table { border-collapse: collapse; width: 100%; margin: 20px 0; }
th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
th { background: #f5f5f5; font-weight: bold; }
img { max-width: 100%; height: auto; }
blockquote { border-left: 4px solid #ddd; padding-left: 16px; color: #666; margin: 16px 0; }
"""


def create_app_icon() -> QIcon:
    size = 128
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    path = QPainterPath()
    path.addRoundedRect(4, 4, size - 8, size - 8, 24, 24)
    grad = QLinearGradient(0, 0, size, size)
    grad.setColorAt(0, QColor("#2563EB"))
    grad.setColorAt(1, QColor("#7C3AED"))
    painter.setBrush(grad)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.drawPath(path)

    font = QFont("Segoe UI", 36, QFont.Weight.Black)
    painter.setFont(font)
    painter.setPen(QColor("#FFFFFF"))
    painter.drawText(pixmap.rect(), Qt.AlignmentFlag.AlignCenter, "F")

    painter.end()
    return QIcon(pixmap)


APP_ICON = None


def get_app_icon() -> QIcon:
    global APP_ICON
    if APP_ICON is None:
        APP_ICON = create_app_icon()
    return APP_ICON


def get_pandoc_path():
    """获取打包后的 pandoc 路径"""
    if hasattr(os, "_MEIPASS"):
        return os.path.join(os._MEIPASS, "pandoc", "pandoc.exe")
    return None


def docx_to_pdf_word(docx_path: str, pdf_path: str):
    """使用 Microsoft Word 将 DOCX 转为 PDF（完美保留格式）"""
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
    finally:
        if doc:
            doc.Close()
        if word:
            word.Quit()
        pythoncom.CoUninitialize()


def _pdf_to_docx_pdf2docx(pdf_path: str, docx_path: str):
    """使用 pdf2docx 将 PDF 转为 DOCX。"""
    cv = PdfConverter(os.path.abspath(pdf_path))
    try:
        cv.convert(os.path.abspath(docx_path))
    finally:
        cv.close()


def _unwrap_layout_tables(docx_path: str):
    """拆掉 pdf2docx 用来做版面定位的表格，只保留真正的数据表格。

    pdf2docx 为了还原版式，会把大量正文/行内代码包进单元格表格，
    并给单元格画上 1px 边框——在 Word 里就是文字周围的细/虚线方框。
    目标文档里没有这些框，正文是流式段落、只有真正的数据表才有表格。

    本函数：1) 把嵌套在单元格里的表格全部拆出来；2) 把列数 <=2 的
    顶层表格（版面定位用）拆成普通段落；3) 给剩下的真实数据表去掉
    单元格边框，统一设置目标文档同款的浅灰表格边框 (DEE2E4)。
    """
    import docx
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    try:
        document = docx.Document(docx_path)
    except Exception:
        return

    body = document.element.body
    P, TBL, TC = qn("w:p"), qn("w:tbl"), qn("w:tc")

    def cols_of(tbl):
        rows = tbl.findall(qn("w:tr"))
        return max((len(r.findall(TC)) for r in rows), default=0)

    def all_tables(root):
        return root.findall(".//" + TBL)

    def table_indent(tbl):
        """表格自身的左缩进 (twips)，拆表时需补到内部段落上。"""
        tbl_pr = tbl.find(qn("w:tblPr"))
        if tbl_pr is None:
            return 0
        ind = tbl_pr.find(qn("w:tblInd"))
        if ind is None:
            return 0
        val = ind.get(qn("w:w"))
        try:
            return int(round(float(val)))
        except (TypeError, ValueError):
            return 0

    def add_left_indent(paragraph, extra):
        """给段落左缩进追加 extra(twips)，保持拆表后水平位置不变。"""
        if extra <= 0:
            return
        ppr = paragraph.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            paragraph.insert(0, ppr)
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            ppr.append(ind)
        cur = ind.get(qn("w:left")) or ind.get(qn("w:start")) or "0"
        try:
            base = int(round(float(cur)))
        except (TypeError, ValueError):
            base = 0
        ind.set(qn("w:left"), str(base + extra))

    def unwrap(tbl):
        parent = tbl.getparent()
        idx = list(parent).index(tbl)
        extra_indent = table_indent(tbl)
        blocks = []
        for row in tbl.findall(qn("w:tr")):
            for cell in row.findall(TC):
                for child in cell.iterchildren():
                    if child.tag in (P, TBL):
                        blocks.append(child)
        for b in blocks:
            if b.tag == P:
                add_left_indent(b, extra_indent)
            b.getparent().remove(b)
        for off, b in enumerate(blocks):
            parent.insert(idx + off, b)
        parent.remove(tbl)

    def set_clean_borders(tbl):
        tbl_pr = tbl.find(qn("w:tblPr"))
        if tbl_pr is None:
            tbl_pr = OxmlElement("w:tblPr")
            tbl.insert(0, tbl_pr)
        for old in tbl_pr.findall(qn("w:tblBorders")):
            tbl_pr.remove(old)
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement("w:" + edge)
            e.set(qn("w:val"), "single")
            e.set(qn("w:sz"), "6")
            e.set(qn("w:space"), "0")
            e.set(qn("w:color"), "DEE2E4")
            borders.append(e)
        tbl_pr.append(borders)

    # 1) 反复拆除嵌套在单元格内的表格
    while True:
        nested = [t for t in all_tables(body) if t.getparent().tag == TC]
        if not nested:
            break
        for t in nested:
            unwrap(t)

    # 2) 拆除列数<=2 的顶层版面定位表格
    while True:
        layout = [
            t for t in all_tables(body) if t.getparent() is body and cols_of(t) <= 2
        ]
        if not layout:
            break
        for t in layout:
            unwrap(t)

    # 3) 规范剩余真实数据表的边框
    for t in [t for t in all_tables(body) if t.getparent() is body]:
        for tcb in list(t.iter(qn("w:tcBorders"))):
            tcb.getparent().remove(tcb)
        set_clean_borders(t)

    document.save(docx_path)


def _normalize_code_shading(docx_path: str):
    """规范化行内代码底纹，模仿目标文档的效果。

    pdf2docx 会用表格单元格承载行内代码并给整个单元格上底纹，导致
    灰底铺满单元格宽度形成"文本块"；还会把底纹误加到中文等非代码
    文字上。本函数把所有单元格级、以及加在非等宽字体上的底纹全部
    去掉，只给等宽代码字体（Consolas / Courier New）的文字本身加上
    紧贴文字的浅灰底纹（F5F7FA），效果与目标文档一致。
    """
    import docx
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    CODE_FONTS = {"Consolas", "Consola", "Courier New", "Courier"}
    GRAY = "F5F7FA"

    try:
        document = docx.Document(docx_path)
    except Exception:
        return

    def run_font(rpr):
        if rpr is None:
            return None
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            return None
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            val = rfonts.get(qn(attr))
            if val:
                return val
        return None

    def run_text(run_el):
        return "".join(t.text or "" for t in run_el.iter(qn("w:t")))

    def set_run_shading(rpr, fill):
        for shd in rpr.findall(qn("w:shd")):
            rpr.remove(shd)
        if fill is None:
            return
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
        rpr.append(shd)

    body = document.element.body

    # 1) 去掉所有单元格级底纹（含嵌套表格、文本框内的表格）
    for tc_pr in body.iter(qn("w:tcPr")):
        for shd in tc_pr.findall(qn("w:shd")):
            tc_pr.remove(shd)

    # 2) 规范每个 run 的底纹：仅等宽代码字体保留紧贴文字的浅灰底
    code_runs = []
    for run_el in body.iter(qn("w:r")):
        rpr = run_el.find(qn("w:rPr"))
        if run_font(rpr) in CODE_FONTS and run_text(run_el).strip():
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                run_el.insert(0, rpr)
            set_run_shading(rpr, GRAY)
            code_runs.append(run_el)
        elif rpr is not None:
            set_run_shading(rpr, None)

    # 3) 在行内代码与中英文之间补一个不带底纹的空格，模仿目标文档
    #    中行内代码（独立文本框）四周的留白，避免代码与正文挤在一起。
    def make_space_run():
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = " "
        r.append(t)
        return r

    def sibling_text(run_el, after):
        sib = run_el.getnext() if after else run_el.getprevious()
        if sib is None or sib.tag != qn("w:r"):
            return None
        return run_text(sib)

    for run_el in code_runs:
        prev_txt = sibling_text(run_el, after=False)
        if prev_txt and not prev_txt.endswith((" ", "　")):
            run_el.addprevious(make_space_run())
        next_txt = sibling_text(run_el, after=True)
        if next_txt and not next_txt.startswith((" ", "　")):
            run_el.addnext(make_space_run())

    document.save(docx_path)


def pdf_to_docx_word(pdf_path: str, docx_path: str):
    """将 PDF 转为 DOCX。

    使用 pdf2docx 转换（保留流式结构，不会像 Word 原生重排那样把
    文字摆成相互重叠的浮动文本框），再拆除版面定位表格（消除文字
    周围的方框/虚线边框），最后规范化行内代码底纹，使灰底紧贴代码
    文字、消除铺满单元格的"文本块"灰底与杂散灰块。
    """
    _pdf_to_docx_pdf2docx(pdf_path, docx_path)
    _unwrap_layout_tables(docx_path)
    _normalize_code_shading(docx_path)


def convert_with_pandoc(input_path: str, output_format: str, output_path: str):
    """使用 Pandoc 进行格式转换"""
    ext = os.path.splitext(input_path)[1].lower()

    if ext == ".docx" and output_format == "pdf":
        docx_to_pdf_word(input_path, output_path)
        return

    if ext == ".pdf" and output_format == "docx":
        pdf_to_docx_word(input_path, output_path)
        return

    if ext == ".pdf" and output_format == "pdf":
        import shutil

        shutil.copy2(input_path, output_path)
        return

    format_map = {
        ".md": "markdown",
        ".markdown": "markdown",
        ".html": "html",
        ".htm": "html",
        ".docx": "docx",
        ".txt": "plain",
        ".pdf": None,
    }
    input_format = format_map.get(ext, "markdown")

    if ext == ".pdf":
        docx_tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".docx", delete=False)
        docx_tmp.close()
        pdf_to_docx_word(input_path, docx_tmp.name)
        input_path = docx_tmp.name
        input_format = "docx"

    output_format_map = {
        "pdf": "docx",
        "docx": "docx",
        "html": "html",
        "md": "markdown",
        "txt": "plain",
    }
    pandoc_output_format = output_format_map.get(output_format, "html")

    resource_path = os.path.dirname(os.path.abspath(input_path))
    extra_args = [
        "--standalone",
        "--embed-resources",
        f"--resource-path={resource_path}",
    ]

    pandoc_path = get_pandoc_path()
    if pandoc_path:
        os.environ["PYPANDOC_PANDOC"] = pandoc_path

    if output_format == "pdf":
        docx_tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".docx", delete=False)
        docx_tmp.close()
        pypandoc.convert_file(
            input_path,
            "docx",
            outputfile=docx_tmp.name,
            format=input_format,
            extra_args=extra_args,
        )
        docx_to_pdf_word(docx_tmp.name, output_path)
        try:
            os.unlink(docx_tmp.name)
        except OSError:
            pass
    else:
        if output_format == "html":
            encoded_css = base64.b64encode(CSS_CONTENT.encode()).decode()
            extra_args.append(f"--css=data:text/css;base64,{encoded_css}")
        pypandoc.convert_file(
            input_path,
            pandoc_output_format,
            outputfile=output_path,
            format=input_format,
            extra_args=extra_args,
        )

    if ext == ".pdf" and os.path.exists(input_path) and input_path.endswith(".docx") and input_path != output_path:
        try:
            os.unlink(input_path)
        except OSError:
            pass


def normalized_extension(path: str) -> str:
    ext = os.path.splitext(path)[1].lower().lstrip(".")
    if ext == "markdown":
        return "md"
    if ext == "htm":
        return "html"
    return ext


def readable_size(path: str) -> str:
    size_bytes = os.path.getsize(path)
    if size_bytes >= 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    return f"{size_bytes / 1024:.2f} KB"


def elide_path(path: str, limit: int = 78) -> str:
    if len(path) <= limit:
        return path
    return "..." + path[-(limit - 3):]


def with_shadow(widget: QWidget, blur: int = 28, y_offset: int = 8):
    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(blur)
    shadow.setOffset(0, y_offset)
    shadow.setColor(Qt.GlobalColor.transparent)
    widget.setGraphicsEffect(shadow)


class FormatButton(QToolButton):
    def __init__(self, fmt: str):
        super().__init__()
        self.fmt = fmt
        self.meta = FORMAT_META[fmt]
        self.blocked = False
        self.setCheckable(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        self.setText(f"{self.meta['title']}\n{self.meta['subtitle']}")
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.setFixedHeight(64)
        self.toggled.connect(self.refresh_style)
        self.refresh_style()

    def set_blocked(self, blocked: bool):
        self.blocked = blocked
        self.setEnabled(not blocked)
        if blocked:
            self.setChecked(False)
            self.setToolTip("源文件已是该格式")
        else:
            self.setToolTip("")
        self.refresh_style()

    def refresh_style(self):
        color = self.meta["color"]
        if self.blocked:
            self.setStyleSheet(
                """
                QToolButton {
                    background: #F8FAFC;
                    color: #94A3B8;
                    border: 1px solid #E2E8F0;
                    border-radius: 10px;
                    font-family: "Microsoft YaHei UI";
                    font-size: 13px;
                    font-weight: 600;
                    padding: 7px 8px;
                }
                """
            )
        elif self.isChecked():
            self.setStyleSheet(
                f"""
                QToolButton {{
                    background: {color};
                    color: white;
                    border: 1px solid {color};
                    border-radius: 10px;
                    font-family: "Microsoft YaHei UI";
                    font-size: 13px;
                    font-weight: 700;
                    padding: 7px 8px;
                }}
                """
            )
        else:
            self.setStyleSheet(
                f"""
                QToolButton {{
                    background: #FFFFFF;
                    color: #334155;
                    border: 1px solid #CBD5E1;
                    border-radius: 10px;
                    font-family: "Microsoft YaHei UI";
                    font-size: 13px;
                    font-weight: 600;
                    padding: 7px 8px;
                }}
                QToolButton:hover {{
                    border-color: {color};
                    background: #F8FAFC;
                }}
                """
            )


class ConversionWorker(QThread):
    progress_changed = Signal(int)
    conversion_done = Signal(list, str, list)

    def __init__(self, source_file: str, output_dir: str, tasks: list[tuple[str, str]], notices: list[str]):
        super().__init__()
        self.source_file = source_file
        self.output_dir = output_dir
        self.tasks = tasks
        self.notices = notices

    def run(self):
        output_files = []
        errors = list(self.notices)
        total = len(self.tasks)
        completed = 0

        def convert_one(task):
            fmt, output_path = task
            try:
                convert_with_pandoc(self.source_file, fmt, output_path)
                if os.path.exists(output_path):
                    return "success", os.path.basename(output_path)
                return "error", f"{fmt.upper()}: 文件未生成"
            except Exception as exc:
                return "error", f"{fmt.upper()}: {str(exc)[:120]}"

        max_workers = min(3, total) if total else 1
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_task = {executor.submit(convert_one, task): task for task in self.tasks}
            for future in as_completed(future_to_task):
                status, message = future.result()
                if status == "success":
                    output_files.append(message)
                else:
                    errors.append(message)
                completed += 1
                self.progress_changed.emit(int((completed / total) * 100) if total else 100)

        self.conversion_done.emit(output_files, self.output_dir, errors)


class StyledMessageDialog(QDialog):
    def __init__(self, title: str, message: str, icon_type: str = "info", parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setWindowIcon(get_app_icon())
        self.setModal(True)
        self.setFixedWidth(440)
        self.setStyleSheet(APP_QSS)

        # 无边框窗口，隐藏标题栏
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)

        # 添加阴影效果
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(0)

        container = QFrame()
        container.setObjectName("DialogPanel")
        shadow = QGraphicsDropShadowEffect(container)
        shadow.setBlurRadius(40)
        shadow.setOffset(0, 10)
        shadow.setColor(QColor(0, 0, 0, 100))
        container.setGraphicsEffect(shadow)

        panel_layout = QVBoxLayout(container)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.setSpacing(0)
        layout.addWidget(container)

        top_bar = QFrame()
        top_bar.setObjectName("DialogTopBar")
        top_bar.setFixedHeight(5)
        panel_layout.addWidget(top_bar)

        body = QVBoxLayout()
        body.setContentsMargins(30, 26, 30, 26)
        body.setSpacing(0)
        panel_layout.addLayout(body)

        icon_emoji = {"warning": "⚠️", "info": "ℹ️", "error": "❌"}.get(icon_type, "ℹ️")
        icon = QLabel(icon_emoji)
        icon.setObjectName("MessageIcon")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(icon)
        body.addSpacing(16)

        title_label = QLabel(title)
        title_label.setObjectName("DialogTitle")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(title_label)
        body.addSpacing(10)

        message_label = QLabel(message)
        message_label.setObjectName("DialogMessage")
        message_label.setWordWrap(True)
        message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(message_label)
        body.addSpacing(24)

        ok_btn = QPushButton("确定")
        ok_btn.setObjectName("PrimaryButton")
        ok_btn.setFixedWidth(120)
        ok_btn.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(ok_btn)
        btn_layout.addStretch()
        body.addLayout(btn_layout)

        # 关闭按钮
        self.close_btn = QPushButton("×", container)
        self.close_btn.setObjectName("DialogCloseButton")
        self.close_btn.setFixedSize(32, 32)
        self.close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.close_btn.clicked.connect(self.reject)

    def showEvent(self, event):
        super().showEvent(event)
        # 将关闭按钮定位到右上角
        self.close_btn.move(self.close_btn.parent().width() - 40, 8)
        self.close_btn.raise_()
        if self.parent():
            parent_geo = self.parent().geometry()
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
            self.move(x, y)


class OverwriteDialog(QDialog):
    def __init__(self, filename: str, parent=None):
        super().__init__(parent)
        self.choice = "cancel"
        self.setWindowTitle("文件已存在")
        self.setWindowIcon(get_app_icon())
        self.setModal(True)
        self.setFixedWidth(520)
        self.setStyleSheet(APP_QSS)

        # 无边框窗口，隐藏标题栏
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)

        # 添加阴影效果
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(0)

        container = QFrame()
        container.setObjectName("DialogPanel")
        shadow = QGraphicsDropShadowEffect(container)
        shadow.setBlurRadius(40)
        shadow.setOffset(0, 10)
        shadow.setColor(QColor(0, 0, 0, 100))
        container.setGraphicsEffect(shadow)

        panel_layout = QVBoxLayout(container)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.setSpacing(0)
        layout.addWidget(container)

        top_bar = QFrame()
        top_bar.setObjectName("DialogTopBar")
        top_bar.setFixedHeight(5)
        panel_layout.addWidget(top_bar)

        body = QVBoxLayout()
        body.setContentsMargins(30, 24, 30, 26)
        body.setSpacing(0)
        panel_layout.addLayout(body)

        icon = QLabel("⚠️")
        icon.setObjectName("WarningBadge")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(icon)
        body.addSpacing(14)

        title = QLabel("文件已存在")
        title.setObjectName("DialogTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(title)
        body.addSpacing(8)

        desc = QLabel("目标目录中已有同名文件，请选择处理方式：")
        desc.setObjectName("DialogSubtitle")
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(desc)
        body.addSpacing(20)

        file_box = QFrame()
        file_box.setObjectName("NoticePanel")
        file_inner = QHBoxLayout(file_box)
        file_inner.setContentsMargins(16, 14, 16, 14)
        file_inner.setSpacing(12)

        file_icon = QLabel("📄")
        file_icon.setObjectName("FileIconLarge")
        file_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        file_inner.addWidget(file_icon)

        file_info = QVBoxLayout()
        file_info.setSpacing(4)
        file_label = QLabel(filename)
        file_label.setObjectName("DialogFileName")
        file_label.setWordWrap(True)
        file_detail = QLabel("将被新生成的文件影响")
        file_detail.setObjectName("DialogFileDetail")
        file_info.addWidget(file_label)
        file_info.addWidget(file_detail)
        file_inner.addLayout(file_info, stretch=1)
        body.addWidget(file_box)
        body.addSpacing(24)

        separator = QFrame()
        separator.setObjectName("DialogSeparator")
        separator.setFixedHeight(1)
        body.addWidget(separator)
        body.addSpacing(20)

        actions = QHBoxLayout()
        actions.setSpacing(10)

        cancel_btn = QPushButton("取消")
        cancel_btn.setObjectName("GhostButton")
        cancel_btn.setFixedHeight(40)
        cancel_btn.clicked.connect(lambda: self.finish("cancel"))
        actions.addWidget(cancel_btn)

        actions.addStretch()

        skip_btn = QPushButton("跳过")
        skip_btn.setObjectName("SecondaryButton")
        skip_btn.setFixedHeight(40)
        skip_btn.clicked.connect(lambda: self.finish("skip"))
        actions.addWidget(skip_btn)

        overwrite_btn = QPushButton("覆盖替换")
        overwrite_btn.setObjectName("DangerButton")
        overwrite_btn.setFixedHeight(40)
        overwrite_btn.clicked.connect(lambda: self.finish("overwrite"))
        actions.addWidget(overwrite_btn)

        rename_btn = QPushButton("智能重命名")
        rename_btn.setObjectName("PrimaryButton")
        rename_btn.setFixedHeight(40)
        rename_btn.clicked.connect(lambda: self.finish("rename"))
        actions.addWidget(rename_btn)

        body.addLayout(actions)

        # 关闭按钮
        self.close_btn = QPushButton("×", container)
        self.close_btn.setObjectName("DialogCloseButton")
        self.close_btn.setFixedSize(32, 32)
        self.close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.close_btn.clicked.connect(lambda: self.finish("cancel"))

    def showEvent(self, event):
        super().showEvent(event)
        # 将关闭按钮定位到右上角
        self.close_btn.move(self.close_btn.parent().width() - 40, 8)
        self.close_btn.raise_()
        if self.parent():
            parent_geo = self.parent().geometry()
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
            self.move(x, y)

    def finish(self, choice: str):
        self.choice = choice
        self.accept()


class ResultDialog(QDialog):
    def __init__(self, output_files: list[str], output_dir: str, notices: list[str], parent=None):
        super().__init__(parent)
        self.output_dir = output_dir
        self.setWindowTitle("转换完成")
        self.setWindowIcon(get_app_icon())
        self.setModal(True)
        self.setFixedWidth(540)
        self.setStyleSheet(APP_QSS)

        # 无边框窗口，隐藏标题栏
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)

        # 添加阴影效果
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(0)

        container = QFrame()
        container.setObjectName("DialogPanel")
        shadow = QGraphicsDropShadowEffect(container)
        shadow.setBlurRadius(40)
        shadow.setOffset(0, 10)
        shadow.setColor(QColor(0, 0, 0, 100))
        container.setGraphicsEffect(shadow)

        panel_layout = QVBoxLayout(container)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.setSpacing(0)
        layout.addWidget(container)

        top_bar = QFrame()
        top_bar.setObjectName("DialogTopBar")
        top_bar.setFixedHeight(5)
        panel_layout.addWidget(top_bar)

        body = QVBoxLayout()
        body.setContentsMargins(30, 26, 30, 26)
        body.setSpacing(0)
        panel_layout.addLayout(body)

        icon = QLabel("✅")
        icon.setObjectName("SuccessIcon")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(icon)
        body.addSpacing(14)

        title = QLabel("转换完成")
        title.setObjectName("DialogTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(title)
        body.addSpacing(18)

        if output_files:
            file_title = QLabel(f"成功生成 {len(output_files)} 个文件")
            file_title.setObjectName("SectionTitle")
            body.addWidget(file_title)
            body.addSpacing(10)

            file_list = QFrame()
            file_list.setObjectName("ListPanel")
            file_layout = QVBoxLayout(file_list)
            file_layout.setContentsMargins(14, 12, 14, 12)
            file_layout.setSpacing(8)
            for file_name in output_files:
                row = QLabel(file_name)
                row.setObjectName("SuccessRow")
                row.setMinimumHeight(34)
                file_layout.addWidget(row)
            body.addWidget(file_list)
            body.addSpacing(14)

        dir_label = QLabel(f"📁 {elide_path(output_dir, 70)}")
        dir_label.setObjectName("PathText")
        dir_label.setWordWrap(True)
        body.addWidget(dir_label)

        if notices:
            body.addSpacing(18)
            notice_title = QLabel("未完成项目")
            notice_title.setObjectName("SectionTitle")
            body.addWidget(notice_title)
            body.addSpacing(8)
            for item in notices[:5]:
                warning = QLabel(item)
                warning.setObjectName("WarningText")
                warning.setWordWrap(True)
                body.addWidget(warning)

        body.addSpacing(22)

        actions = QHBoxLayout()
        actions.setSpacing(10)
        actions.addStretch()

        open_btn = QPushButton("打开文件夹")
        open_btn.setObjectName("PrimaryButton")
        open_btn.setFixedHeight(40)
        open_btn.clicked.connect(self.open_folder)
        actions.addWidget(open_btn)

        ok_btn = QPushButton("确定")
        ok_btn.setObjectName("SecondaryButton")
        ok_btn.setFixedHeight(40)
        ok_btn.clicked.connect(self.accept)
        actions.addWidget(ok_btn)
        body.addLayout(actions)

        # 关闭按钮
        self.close_btn = QPushButton("×", container)
        self.close_btn.setObjectName("DialogCloseButton")
        self.close_btn.setFixedSize(32, 32)
        self.close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.close_btn.clicked.connect(self.accept)

    def showEvent(self, event):
        super().showEvent(event)
        # 将关闭按钮定位到右上角
        self.close_btn.move(self.close_btn.parent().width() - 40, 8)
        self.close_btn.raise_()
        if self.parent():
            parent_geo = self.parent().geometry()
            self.move(
                parent_geo.x() + (parent_geo.width() - self.width()) // 2,
                parent_geo.y() + (parent_geo.height() - self.height()) // 2
            )

    def open_folder(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.output_dir))
        self.accept()


class ConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_path = None
        self.custom_output_dir = None
        self.worker = None

        self.setWindowTitle(f"文件格式转换器 v{get_current_version()}")
        self.setWindowIcon(get_app_icon())
        self.setMinimumSize(1080, 720)
        self.resize(1120, 760)
        self.setAcceptDrops(True)
        self.setStyleSheet(APP_QSS)

        self.format_buttons = {}
        self.setup_ui()
        self.update_output_label()

        # 启动时检查更新（在后台线程）
        if UPDATER_AVAILABLE:
            self.check_updates_on_startup()

    def setup_ui(self):
        root = QWidget()
        root.setObjectName("AppRoot")
        self.setCentralWidget(root)

        page = QVBoxLayout(root)
        page.setContentsMargins(34, 28, 34, 28)
        page.setSpacing(22)

        header = QHBoxLayout()
        header.setSpacing(16)

        title_box = QVBoxLayout()
        title_box.setSpacing(5)
        title = QLabel("文件格式转换器")
        title.setObjectName("AppTitle")
        subtitle = QLabel("Markdown / HTML / Word / PDF / TXT")
        subtitle.setObjectName("AppSubtitle")
        title_box.addWidget(title)
        title_box.addWidget(subtitle)
        header.addLayout(title_box)
        header.addStretch()

        # 帮助按钮
        help_btn = QPushButton("帮助")
        help_btn.setObjectName("HelpButton")
        help_btn.setFixedSize(68, 34)
        help_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        help_menu = QMenu(help_btn)
        help_menu.setStyleSheet("""
            QMenu {
                background: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                padding: 8px;
            }
            QMenu::item {
                padding: 8px 20px;
                border-radius: 6px;
                color: #334155;
            }
            QMenu::item:selected {
                background: #F1F5F9;
            }
        """)

        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

        if UPDATER_AVAILABLE:
            check_update_action = QAction("检查更新", self)
            check_update_action.triggered.connect(self.manual_check_updates)
            help_menu.addAction(check_update_action)

        help_btn.setMenu(help_menu)
        header.addWidget(help_btn)

        header.addSpacing(8)

        self.status_chip = QLabel("就绪")
        self.status_chip.setObjectName("StatusChip")
        self.status_chip.setFixedSize(88, 34)
        self.status_chip.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header.addWidget(self.status_chip)
        page.addLayout(header)

        content = QHBoxLayout()
        content.setSpacing(18)
        page.addLayout(content, stretch=1)

        main_card = self.make_card()
        main_layout = QVBoxLayout(main_card)
        main_layout.setContentsMargins(24, 22, 24, 22)
        main_layout.setSpacing(18)
        content.addWidget(main_card, stretch=5)

        file_section = self.make_section("源文件")
        file_layout = QVBoxLayout(self.section_body(file_section))
        file_layout.setContentsMargins(16, 14, 16, 14)
        file_layout.setSpacing(12)

        choose_row = QHBoxLayout()
        choose_row.setSpacing(12)
        self.file_name = QLabel("未选择文件")
        self.file_name.setObjectName("FileName")
        self.file_name.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        choose_row.addWidget(self.file_name, stretch=1)

        self.choose_btn = QPushButton("选择文件")
        self.choose_btn.setObjectName("PrimaryButton")
        self.choose_btn.clicked.connect(self.select_file)
        choose_row.addWidget(self.choose_btn)
        file_layout.addLayout(choose_row)

        self.file_meta = QLabel("类型：-    大小：-")
        self.file_meta.setObjectName("MetaText")
        file_layout.addWidget(self.file_meta)

        self.file_path_label = QLabel("路径：-")
        self.file_path_label.setObjectName("PathText")
        self.file_path_label.setWordWrap(True)
        self.file_path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        file_layout.addWidget(self.file_path_label)
        main_layout.addWidget(file_section)

        format_section = self.make_section("输出格式")
        format_body = self.section_body(format_section)
        format_body.setFixedHeight(96)
        format_layout = QGridLayout(format_body)
        format_layout.setContentsMargins(14, 14, 14, 14)
        format_layout.setHorizontalSpacing(12)
        format_layout.setVerticalSpacing(0)

        for idx, fmt in enumerate(OUTPUT_FORMATS):
            button = FormatButton(fmt)
            self.format_buttons[fmt] = button
            format_layout.addWidget(button, 0, idx)
        main_layout.addWidget(format_section)

        output_section = self.make_section("保存位置")
        output_layout = QHBoxLayout(self.section_body(output_section))
        output_layout.setContentsMargins(16, 14, 16, 14)
        output_layout.setSpacing(12)

        self.output_label = QLabel("")
        self.output_label.setObjectName("PathText")
        self.output_label.setMinimumHeight(36)
        self.output_label.setWordWrap(True)
        self.output_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        output_layout.addWidget(self.output_label, stretch=1)

        output_btn = QPushButton("更改")
        output_btn.setObjectName("SecondaryButton")
        output_btn.clicked.connect(self.select_output_dir)
        output_layout.addWidget(output_btn)
        main_layout.addWidget(output_section)

        self.progress = QProgressBar()
        self.progress.setObjectName("MainProgress")
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        main_layout.addWidget(self.progress)

        bottom = QHBoxLayout()
        bottom.setSpacing(12)
        self.status_text = QLabel("就绪")
        self.status_text.setObjectName("MetaText")
        bottom.addWidget(self.status_text, stretch=1)

        self.convert_btn = QPushButton("开始转换")
        self.convert_btn.setObjectName("ConvertButton")
        self.convert_btn.clicked.connect(self.convert)
        bottom.addWidget(self.convert_btn)
        main_layout.addLayout(bottom)

        side_panel = self.make_card()
        side_panel.setObjectName("SidePanel")
        side_layout = QVBoxLayout(side_panel)
        side_layout.setContentsMargins(22, 22, 22, 22)
        side_layout.setSpacing(14)
        content.addWidget(side_panel, stretch=2)

        summary_title = QLabel("当前任务")
        summary_title.setObjectName("PanelTitle")
        side_layout.addWidget(summary_title)

        self.summary_file = QLabel("文件：-")
        self.summary_file.setObjectName("PanelText")
        self.summary_file.setWordWrap(True)
        side_layout.addWidget(self.summary_file)

        self.summary_formats = QLabel("格式：-")
        self.summary_formats.setObjectName("PanelText")
        self.summary_formats.setWordWrap(True)
        side_layout.addWidget(self.summary_formats)

        self.summary_output = QLabel("位置：-")
        self.summary_output.setObjectName("PanelText")
        self.summary_output.setWordWrap(True)
        side_layout.addWidget(self.summary_output)
        side_layout.addStretch()

        hint = QLabel("可将文件拖入窗口")
        hint.setObjectName("DropHint")
        hint.setFixedHeight(116)
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        side_layout.addWidget(hint)

        for button in self.format_buttons.values():
            button.toggled.connect(self.update_summary)

    # 移除 setup_menu，改用按钮形式

    def make_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("Card")
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        return card

    def make_section(self, title: str) -> QFrame:
        section = QFrame()
        section.setObjectName("Section")
        layout = QVBoxLayout(section)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        heading = QLabel(title)
        heading.setObjectName("SectionTitle")
        heading.setContentsMargins(0, 0, 0, 8)
        layout.addWidget(heading)

        body = QFrame()
        body.setObjectName("SectionBody")
        layout.addWidget(body)
        return section

    def section_body(self, section: QFrame) -> QFrame:
        return section.findChild(QFrame, "SectionBody")

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择要转换的文件", "", FILE_DIALOG_FILTER)
        if file_path:
            self.set_file(file_path)

    def set_file(self, file_path: str):
        self.file_path = file_path
        file_name = os.path.basename(file_path)
        file_ext = os.path.splitext(file_path)[1].upper() or "-"

        self.file_name.setText(file_name)
        self.file_meta.setText(f"类型：{file_ext}    大小：{readable_size(file_path)}")
        self.file_path_label.setText(f"路径：{file_path}")
        self.progress.setValue(0)
        self.set_status("就绪", "ready")

        if not self.custom_output_dir:
            self.update_output_label(os.path.dirname(file_path))
        else:
            self.update_output_label(self.custom_output_dir)

        source_ext = normalized_extension(file_path)
        for fmt, button in self.format_buttons.items():
            button.set_blocked(fmt == source_ext)

        self.update_summary()

    def select_output_dir(self):
        directory = QFileDialog.getExistingDirectory(self, "选择保存位置", self.current_output_dir())
        if directory:
            self.custom_output_dir = directory
            self.update_output_label(directory)
            self.update_summary()

    def update_output_label(self, directory: str | None = None):
        if directory is None:
            directory = self.current_output_dir()
        self.output_label.setText(directory if directory else "选择文件后使用源文件所在位置")

    def current_output_dir(self) -> str:
        if self.custom_output_dir:
            return self.custom_output_dir
        if self.file_path:
            return os.path.dirname(self.file_path)
        return ""

    def selected_formats(self) -> list[str]:
        return [fmt for fmt, button in self.format_buttons.items() if button.isChecked() and button.isEnabled()]

    def update_summary(self):
        file_text = os.path.basename(self.file_path) if self.file_path else "-"
        formats = self.selected_formats()
        output_dir = self.current_output_dir() or "-"

        self.summary_file.setText(f"文件：{file_text}")
        self.summary_formats.setText(f"格式：{', '.join(fmt.upper() for fmt in formats) if formats else '-'}")
        self.summary_output.setText(f"位置：{elide_path(output_dir, 64) if output_dir != '-' else '-'}")

    def convert(self):
        if self.worker and self.worker.isRunning():
            return

        if not self.file_path:
            StyledMessageDialog("提示", "请先选择要转换的文件。", "warning", self).exec()
            return

        selected = self.selected_formats()
        if not selected:
            StyledMessageDialog("提示", "请至少选择一种输出格式。", "warning", self).exec()
            return

        source_ext = normalized_extension(self.file_path)
        formats_to_convert = [fmt for fmt in selected if fmt != source_ext]
        if not formats_to_convert:
            StyledMessageDialog("提示", "输出格式与源文件格式相同，没有生成新文件。", "info", self).exec()
            return

        output_dir = self.current_output_dir()
        if not output_dir:
            StyledMessageDialog("提示", "请选择保存位置。", "warning", self).exec()
            return

        base_name = os.path.splitext(os.path.basename(self.file_path))[0]
        tasks = []
        notices = []

        for fmt in formats_to_convert:
            output_path = os.path.join(output_dir, f"{base_name}.{fmt}")
            if os.path.exists(output_path):
                choice = self.ask_overwrite(os.path.basename(output_path))
                if choice == "cancel":
                    self.set_status("已取消", "idle")
                    return
                if choice == "skip":
                    notices.append(f"{fmt.upper()}: 已跳过同名文件")
                    continue
                if choice == "rename":
                    output_path = self.available_output_path(output_dir, base_name, fmt)
            tasks.append((fmt, output_path))

        if not tasks:
            StyledMessageDialog("提示", "没有需要转换的输出格式。", "info", self).exec()
            self.set_status("就绪", "ready")
            return

        self.set_busy(True)
        self.worker = ConversionWorker(self.file_path, output_dir, tasks, notices)
        self.worker.progress_changed.connect(self.progress.setValue)
        self.worker.conversion_done.connect(self.on_conversion_done)
        self.worker.start()

    def ask_overwrite(self, filename: str) -> str:
        dialog = OverwriteDialog(filename, self)
        dialog.exec()
        return dialog.choice

    def available_output_path(self, output_dir: str, base_name: str, fmt: str) -> str:
        counter = 1
        while True:
            candidate = os.path.join(output_dir, f"{base_name}_{counter}.{fmt}")
            if not os.path.exists(candidate):
                return candidate
            counter += 1

    def set_busy(self, busy: bool):
        self.choose_btn.setEnabled(not busy)
        self.convert_btn.setEnabled(not busy)
        for button in self.format_buttons.values():
            button.setEnabled((not busy) and (not button.blocked))
            button.refresh_style()

        if busy:
            self.convert_btn.setText("转换中...")
            self.set_status("正在转换", "working")
            self.progress.setValue(0)
        else:
            self.convert_btn.setText("开始转换")

    def set_status(self, text: str, state: str):
        self.status_text.setText(text)
        self.status_chip.setText(text)
        self.status_chip.setProperty("state", state)
        self.status_chip.style().unpolish(self.status_chip)
        self.status_chip.style().polish(self.status_chip)

    def on_conversion_done(self, output_files: list[str], output_dir: str, notices: list[str]):
        self.set_busy(False)

        if output_files:
            self.progress.setValue(100)
            self.set_status("转换成功", "success")
            ResultDialog(output_files, output_dir, notices, self).exec()
        elif notices:
            self.progress.setValue(0)
            self.set_status("未生成文件", "warning")
            StyledMessageDialog("转换未完成", "\n".join(notices), "warning", self).exec()
        else:
            self.progress.setValue(0)
            self.set_status("转换失败", "error")
            StyledMessageDialog("转换失败", "所有格式转换失败。", "error", self).exec()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if os.path.splitext(path)[1].lower() in SUPPORTED_EXTENSIONS:
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if os.path.isfile(path) and os.path.splitext(path)[1].lower() in SUPPORTED_EXTENSIONS:
                self.set_file(path)
                event.acceptProposedAction()
                return
        event.ignore()

    def check_updates_on_startup(self):
        """启动时在后台检查更新"""
        from threading import Thread

        def check():
            has_update, latest, download_url, notes = check_for_updates()
            if has_update:
                # 在主线程显示更新对话框
                QApplication.instance().postEvent(
                    self,
                    UpdateAvailableEvent(latest, download_url, notes)
                )

        Thread(target=check, daemon=True).start()

    def manual_check_updates(self):
        """手动检查更新"""
        if not UPDATER_AVAILABLE:
            return

        has_update, latest, download_url, notes = check_for_updates(timeout=10)

        if has_update:
            self.show_update_dialog(latest, download_url, notes)
        else:
            StyledMessageDialog(
                "已是最新版本",
                f"当前版本 {get_current_version()} 已是最新版本。",
                "info",
                self
            ).exec()

    def show_update_dialog(self, version, download_url, notes):
        """显示更新对话框"""
        dialog = UpdateDialog(version, download_url, notes, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            open_download_page(download_url)

    def show_about(self):
        """显示关于对话框"""
        about_text = (
            f"文件格式转换器 v{get_current_version()}\n\n"
            "支持 Markdown、HTML、Word、PDF、TXT 等格式互转\n\n"
            "本软件仅供学习交流使用\n"
            "使用本软件产生的任何后果由使用者自行承担"
        )
        StyledMessageDialog("关于", about_text, "info", self).exec()

    def customEvent(self, event):
        """处理自定义事件（更新通知）"""
        if isinstance(event, UpdateAvailableEvent):
            self.show_update_dialog(event.version, event.download_url, event.notes)


class UpdateAvailableEvent:
    """自定义事件：有可用更新"""
    def __init__(self, version, download_url, notes):
        self.version = version
        self.download_url = download_url
        self.notes = notes
        self.type_id = QThread.eventType()


class UpdateDialog(QDialog):
    """更新提示对话框"""
    def __init__(self, version, download_url, notes, parent=None):
        super().__init__(parent)
        self.version = version
        self.download_url = download_url
        self.notes = notes

        self.setWindowTitle("发现新版本")
        self.setWindowIcon(get_app_icon())
        self.setModal(True)
        self.setFixedWidth(520)
        self.setStyleSheet(APP_QSS)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(0)

        container = QFrame()
        container.setObjectName("DialogPanel")
        shadow = QGraphicsDropShadowEffect(container)
        shadow.setBlurRadius(40)
        shadow.setOffset(0, 10)
        shadow.setColor(QColor(0, 0, 0, 100))
        container.setGraphicsEffect(shadow)

        panel_layout = QVBoxLayout(container)
        panel_layout.setContentsMargins(0, 0, 0, 0)
        panel_layout.setSpacing(0)
        layout.addWidget(container)

        top_bar = QFrame()
        top_bar.setObjectName("DialogTopBar")
        top_bar.setFixedHeight(5)
        panel_layout.addWidget(top_bar)

        body = QVBoxLayout()
        body.setContentsMargins(30, 24, 30, 26)
        body.setSpacing(0)
        panel_layout.addLayout(body)

        icon = QLabel("🎉")
        icon.setObjectName("MessageIcon")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(icon)
        body.addSpacing(14)

        title = QLabel("发现新版本")
        title.setObjectName("DialogTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(title)
        body.addSpacing(8)

        version_label = QLabel(f"v{version} 已发布")
        version_label.setObjectName("DialogSubtitle")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        body.addWidget(version_label)
        body.addSpacing(20)

        if notes and notes.strip():
            notes_box = QFrame()
            notes_box.setObjectName("ListPanel")
            notes_layout = QVBoxLayout(notes_box)
            notes_layout.setContentsMargins(16, 14, 16, 14)

            notes_title = QLabel("更新内容:")
            notes_title.setObjectName("SectionTitle")
            notes_layout.addWidget(notes_title)
            notes_layout.addSpacing(8)

            # 使用滚动区域防止内容过长
            from PySide6.QtWidgets import QScrollArea, QTextEdit
            notes_text = QTextEdit()
            notes_text.setObjectName("DialogMessage")
            notes_text.setReadOnly(True)
            notes_text.setPlainText(notes[:800])  # 限制长度
            notes_text.setMaximumHeight(150)
            notes_text.setMinimumHeight(100)
            notes_text.setFrameStyle(0)
            notes_text.setStyleSheet("""
                QTextEdit {
                    background: transparent;
                    border: none;
                    color: #475569;
                    font-size: 13px;
                    line-height: 1.6;
                }
                QScrollBar:vertical {
                    background: #F1F5F9;
                    width: 8px;
                    border-radius: 4px;
                    margin: 0px;
                }
                QScrollBar::handle:vertical {
                    background: #CBD5E1;
                    border-radius: 4px;
                    min-height: 20px;
                }
                QScrollBar::handle:vertical:hover {
                    background: #94A3B8;
                }
                QScrollBar::add-line:vertical,
                QScrollBar::sub-line:vertical {
                    height: 0px;
                }
                QScrollBar::add-page:vertical,
                QScrollBar::sub-page:vertical {
                    background: transparent;
                }
            """)
            notes_layout.addWidget(notes_text)

            body.addWidget(notes_box)
            body.addSpacing(20)

        actions = QHBoxLayout()
        actions.setSpacing(10)

        later_btn = QPushButton("稍后提醒")
        later_btn.setObjectName("SecondaryButton")
        later_btn.setFixedHeight(40)
        later_btn.clicked.connect(self.reject)
        actions.addWidget(later_btn)

        download_btn = QPushButton("立即下载")
        download_btn.setObjectName("PrimaryButton")
        download_btn.setFixedHeight(40)
        download_btn.clicked.connect(self.accept)
        actions.addWidget(download_btn)

        body.addLayout(actions)

        self.close_btn = QPushButton("×", container)
        self.close_btn.setObjectName("DialogCloseButton")
        self.close_btn.setFixedSize(32, 32)
        self.close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.close_btn.clicked.connect(self.reject)

    def showEvent(self, event):
        super().showEvent(event)
        self.close_btn.move(self.close_btn.parent().width() - 40, 8)
        self.close_btn.raise_()
        if self.parent():
            parent_geo = self.parent().geometry()
            x = parent_geo.x() + (parent_geo.width() - self.width()) // 2
            y = parent_geo.y() + (parent_geo.height() - self.height()) // 2
            self.move(x, y)


APP_QSS = """
QWidget#AppRoot {
    background: #F4F7FB;
    color: #111827;
    font-family: "Microsoft YaHei UI";
    font-size: 13px;
}
QLabel#AppTitle {
    color: #0F172A;
    font-size: 26px;
    font-weight: 800;
}
QLabel#AppSubtitle {
    color: #64748B;
    font-size: 13px;
}
QLabel#StatusChip {
    min-width: 86px;
    min-height: 30px;
    border-radius: 15px;
    background: #E2E8F0;
    color: #475569;
    font-weight: 700;
}
QLabel#StatusChip[state="working"] {
    background: #FEF3C7;
    color: #92400E;
}
QLabel#StatusChip[state="success"] {
    background: #DCFCE7;
    color: #166534;
}
QLabel#StatusChip[state="error"] {
    background: #FEE2E2;
    color: #991B1B;
}
QLabel#StatusChip[state="warning"] {
    background: #FFEDD5;
    color: #9A3412;
}
QPushButton#HelpButton {
    background: #2563EB;
    color: white;
    border: 1px solid #2563EB;
    border-radius: 8px;
    font-weight: 700;
    padding: 0 16px;
    text-align: center;
}
QPushButton#HelpButton:hover {
    background: #1D4ED8;
    border-color: #1D4ED8;
}
QPushButton#HelpButton::menu-indicator {
    image: none;
    width: 0px;
}
QFrame#Card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
}
QFrame#SidePanel {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
}
QFrame#Section {
    background: transparent;
    border: none;
}
QFrame#SectionBody {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
}
QLabel#SectionTitle {
    color: #0F172A;
    font-size: 14px;
    font-weight: 800;
}
QLabel#FileName {
    color: #0F172A;
    font-size: 16px;
    font-weight: 800;
}
QLabel#MetaText,
QLabel#PathText,
QLabel#PanelText {
    color: #64748B;
    font-size: 12px;
}
QLabel#PanelTitle {
    color: #0F172A;
    font-size: 17px;
    font-weight: 800;
}
QLabel#DropHint {
    min-height: 92px;
    color: #475569;
    background: #F8FAFC;
    border: 1px dashed #94A3B8;
    border-radius: 10px;
    font-weight: 700;
}
QPushButton {
    min-height: 34px;
    border-radius: 8px;
    padding: 0 16px;
    font-weight: 700;
}
QPushButton#PrimaryButton,
QPushButton#ConvertButton {
    background: #2563EB;
    color: white;
    border: 1px solid #2563EB;
}
QPushButton#PrimaryButton:hover,
QPushButton#ConvertButton:hover {
    background: #1D4ED8;
    border-color: #1D4ED8;
}
QPushButton#PrimaryButton:disabled,
QPushButton#ConvertButton:disabled {
    background: #93C5FD;
    border-color: #93C5FD;
}
QPushButton#ConvertButton {
    min-width: 180px;
    min-height: 42px;
    font-size: 14px;
}
QPushButton#SecondaryButton {
    background: #FFFFFF;
    color: #334155;
    border: 1px solid #CBD5E1;
}
QPushButton#SecondaryButton:hover {
    background: #F8FAFC;
    border-color: #94A3B8;
}
QProgressBar#MainProgress {
    height: 10px;
    background: #E2E8F0;
    border: none;
    border-radius: 5px;
}
QProgressBar#MainProgress::chunk {
    background: #2563EB;
    border-radius: 5px;
}
QFrame#DialogPanel {
    background: #FAFBFC;
    border: none;
    border-radius: 12px;
}
QFrame#DialogTopBar {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
        stop:0 #2563EB, stop:0.5 #7C3AED, stop:1 #EC4899);
    border: none;
    border-top-left-radius: 12px;
    border-top-right-radius: 12px;
}
QLabel#DialogTitle {
    color: #0F172A;
    font-size: 20px;
    font-weight: 800;
}
QLabel#DialogSubtitle {
    color: #64748B;
    font-size: 13px;
    line-height: 1.5;
}
QLabel#DialogMessage {
    color: #475569;
    font-size: 13px;
    line-height: 1.6;
}
QLabel#MessageIcon {
    font-size: 48px;
    min-height: 56px;
}
QLabel#WarningBadge {
    font-size: 48px;
    min-height: 56px;
}
QLabel#SuccessIcon {
    font-size: 48px;
    min-height: 56px;
}
QFrame#IconContainer {
    background: transparent;
    border: none;
}
QFrame#NoticePanel {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #FFFBEB, stop:1 #FEF3C7);
    border: 2px solid #FCD34D;
    border-radius: 12px;
}
QLabel#FileIconLarge {
    font-size: 28px;
    min-width: 44px;
    max-width: 44px;
    min-height: 44px;
    max-height: 44px;
}
QLabel#FileIcon {
    font-size: 22px;
    background: #EFF6FF;
    border-radius: 20px;
    border: 1px solid #BFDBFE;
}
QLabel#DialogFileName {
    color: #0F172A;
    font-size: 14px;
    font-weight: 700;
}
QLabel#DialogFileDetail {
    color: #92400E;
    font-size: 12px;
}
QFrame#DialogSeparator {
    background: #D1D5DB;
    border: none;
}
QPushButton#DangerButton {
    background: #FEE2E2;
    color: #991B1B;
    border: 2px solid #F87171;
    border-radius: 8px;
    padding: 0 20px;
    font-weight: 700;
}
QPushButton#DangerButton:hover {
    background: #FEF2F2;
    border-color: #DC2626;
    color: #7F1D1D;
}
QPushButton#GhostButton {
    background: transparent;
    color: #64748B;
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 0 20px;
    font-weight: 700;
}
QPushButton#GhostButton:hover {
    background: #F1F5F9;
    color: #334155;
}
QPushButton#DialogCloseButton {
    background: transparent;
    color: #94A3B8;
    border: none;
    border-radius: 16px;
    font-size: 24px;
    font-weight: 700;
    padding: 0;
}
QPushButton#DialogCloseButton:hover {
    background: #FEE2E2;
    color: #DC2626;
}
QFrame#ListPanel {
    background: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 10px;
}
QLabel#SuccessRow {
    color: #166534;
    background: #DCFCE7;
    border-radius: 7px;
    padding-left: 12px;
    font-weight: 700;
}
QLabel#WarningText {
    color: #B45309;
    font-size: 12px;
}
QMessageBox {
    background: #FFFFFF;
    font-family: "Microsoft YaHei UI";
}
QMessageBox QLabel {
    color: #111827;
    font-family: "Microsoft YaHei UI";
}
QDialog {
    background: #FFFFFF;
    font-family: "Microsoft YaHei UI";
}
"""


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setWindowIcon(get_app_icon())
    window = ConverterApp()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
