"""Build reproducible report assets with only the bundled pandas/numpy/Pillow/docx.

This helper mirrors the configuration in kmeans_humanitarian.py. It exists so
the report can be generated in the current environment even when scikit-learn
is not installed there. The deliverable code remains kmeans_humanitarian.py.
"""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
DATA = ROOT / "Country-data.csv"
RESULTS = ROOT / "results"
FIGURES = ROOT / "figures"
REPORT = ROOT / "Tieu_luan_KMeans_vien_tro_nhan_dao.docx"
FEATURES = [
    "child_mort", "exports", "health", "imports", "income",
    "inflation", "life_expec", "total_fer", "gdpp",
]
LABELS = {
    "child_mort": "Tử vong trẻ em",
    "exports": "Xuất khẩu",
    "health": "Chi y tế",
    "imports": "Nhập khẩu",
    "income": "Thu nhập",
    "inflation": "Lạm phát",
    "life_expec": "Tuổi thọ",
    "total_fer": "Mức sinh",
    "gdpp": "GDP/người",
}
BLUE = (31, 78, 121)
RED = (192, 80, 77)
GREEN = (46, 139, 87)
GRAY = (102, 102, 102)


def font(size: int, bold: bool = False):
    candidates = [
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("C:/Windows/Fonts/segoeui.ttf"),
    ]
    if bold:
        candidates = [Path("C:/Windows/Fonts/arialbd.ttf"), Path("C:/Windows/Fonts/segoeuib.ttf")] + candidates
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def kmeans_pp(X: np.ndarray, k: int, seed: int = 42, n_init: int = 50):
    rng = np.random.default_rng(seed)
    best = None
    for _ in range(n_init):
        centers = np.empty((k, X.shape[1]))
        idx = rng.integers(len(X))
        centers[0] = X[idx]
        d = ((X - centers[0]) ** 2).sum(axis=1)
        for j in range(1, k):
            p = d / d.sum()
            idx = rng.choice(len(X), p=p)
            centers[j] = X[idx]
            d = np.minimum(d, ((X - centers[j]) ** 2).sum(axis=1))
        for iteration in range(300):
            dist = ((X[:, None, :] - centers[None, :, :]) ** 2).sum(axis=2)
            labels = dist.argmin(axis=1)
            new_centers = np.array([
                X[labels == j].mean(axis=0) if np.any(labels == j) else centers[j]
                for j in range(k)
            ])
            if np.allclose(new_centers, centers, atol=1e-10):
                break
            centers = new_centers
        inertia = ((X - centers[labels]) ** 2).sum()
        if best is None or inertia < best[0]:
            best = inertia, labels.copy(), centers.copy()
    return best


def silhouette(X: np.ndarray, labels: np.ndarray) -> float:
    D = np.sqrt(((X[:, None, :] - X[None, :, :]) ** 2).sum(axis=2))
    values = []
    for i in range(len(X)):
        same = labels == labels[i]
        same[i] = False
        a = D[i, same].mean() if same.any() else 0.0
        b = min(D[i, labels == c].mean() for c in np.unique(labels) if c != labels[i])
        values.append((b - a) / max(a, b) if max(a, b) > 0 else 0.0)
    return float(np.mean(values))


def calinski(X: np.ndarray, labels: np.ndarray) -> float:
    n, _ = X.shape
    overall = X.mean(axis=0)
    between = 0.0
    within = 0.0
    clusters = np.unique(labels)
    for c in clusters:
        Xi = X[labels == c]
        between += len(Xi) * ((Xi.mean(axis=0) - overall) ** 2).sum()
        within += ((Xi - Xi.mean(axis=0)) ** 2).sum()
    return float((between / (len(clusters) - 1)) / (within / (n - len(clusters))))


def davies_bouldin(X: np.ndarray, labels: np.ndarray) -> float:
    centers, scatters = [], []
    for c in np.unique(labels):
        Xi = X[labels == c]
        center = Xi.mean(axis=0)
        centers.append(center)
        scatters.append(np.sqrt(((Xi - center) ** 2).sum(axis=1)).mean())
    centers = np.array(centers)
    scatters = np.array(scatters)
    distances = np.sqrt(((centers[:, None, :] - centers[None, :, :]) ** 2).sum(axis=2))
    values = []
    for i in range(len(centers)):
        ratios = (scatters[i] + scatters) / (distances[i] + 1e-12)
        ratios[i] = -np.inf
        values.append(ratios.max())
    return float(np.mean(values))


def write_png(path: Path, width: int, height: int, title: str, draw_fn):
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    draw_fn(draw, width, height)
    draw.text((35, 22), title, fill=BLUE, font=font(28, True))
    img.save(path)


def line_chart(path: Path, title: str, xs, ys, ylabel: str, color, vline=4):
    def draw(draw, w, h):
        left, top, right, bottom = 90, 82, w - 45, h - 75
        ylo, yhi = min(ys) * 0.95, max(ys) * 1.05
        for tick in range(6):
            y = bottom - tick * (bottom - top) / 5
            value = ylo + tick * (yhi - ylo) / 5
            draw.line((left, y, right, y), fill=(225, 225, 225), width=1)
            draw.text((5, y - 10), f"{value:.2f}", fill=GRAY, font=font(15))
        def xy(x, y):
            return (left + (x - min(xs)) * (right - left) / (max(xs) - min(xs)),
                    bottom - (y - ylo) * (bottom - top) / (yhi - ylo))
        if min(xs) <= vline <= max(xs):
            vx = xy(vline, ylo)[0]
            draw.line((vx, top, vx, bottom), fill=RED, width=2)
            draw.text((vx + 6, top + 5), "k = 4", fill=RED, font=font(16, True))
        points = [xy(x, y) for x, y in zip(xs, ys)]
        draw.line(points, fill=color, width=4)
        for (x, y), xv in zip(points, xs):
            draw.ellipse((x - 6, y - 6, x + 6, y + 6), fill=color, outline="white", width=2)
            draw.text((x - 7, bottom + 12), str(xv), fill=GRAY, font=font(16))
        draw.line((left, bottom, right, bottom), fill=GRAY, width=2)
        draw.line((left, top, left, bottom), fill=GRAY, width=2)
        draw.text((left, h - 42), "Số cụm k", fill=GRAY, font=font(18))
        draw.text((10, top - 30), ylabel, fill=GRAY, font=font(17))
    write_png(path, 900, 520, title, draw)


def create_figures(metrics, z, labels):
    FIGURES.mkdir(exist_ok=True)
    line_chart(FIGURES / "01_elbow.png", "Phương pháp Elbow", metrics.k.tolist(), metrics.inertia.tolist(), "SSE", BLUE)
    line_chart(FIGURES / "02_silhouette.png", "Điểm Silhouette", metrics.k.tolist(), metrics.silhouette.tolist(), "Silhouette", GREEN)

    # PCA with a two-dimensional eigen-decomposition.
    cov = np.cov(z, rowvar=False)
    values, vectors = np.linalg.eigh(cov)
    order = np.argsort(values)[::-1]
    vectors = vectors[:, order]
    coords = z @ vectors[:, :2]
    def draw_pca(draw, w, h):
        left, top, right, bottom = 90, 85, w - 45, h - 80
        colors = [(78, 121, 167), (242, 142, 43), (225, 87, 89), (89, 161, 79)]
        xlo, xhi = coords[:, 0].min() - .4, coords[:, 0].max() + .4
        ylo, yhi = coords[:, 1].min() - .4, coords[:, 1].max() + .4
        def xy(x, y):
            return (left + (x - xlo) * (right - left) / (xhi - xlo),
                    bottom - (y - ylo) * (bottom - top) / (yhi - ylo))
        draw.line((left, top, left, bottom), fill=GRAY, width=2)
        draw.line((left, bottom, right, bottom), fill=GRAY, width=2)
        for c in sorted(np.unique(labels)):
            color = colors[c % len(colors)]
            for point in coords[labels == c]:
                x, y = xy(*point)
                draw.ellipse((x - 5, y - 5, x + 5, y + 5), fill=color, outline="white", width=1)
            lx = right - 140
            ly = top + c * 30
            draw.ellipse((lx, ly, lx + 12, ly + 12), fill=color)
            draw.text((lx + 20, ly - 4), f"Cụm {c}", fill=GRAY, font=font(17))
        draw.text((left, h - 45), "PC1", fill=GRAY, font=font(18))
        draw.text((10, top - 27), "PC2", fill=GRAY, font=font(18))
    write_png(FIGURES / "03_pca_clusters.png", 900, 600, "K-Means trên mặt phẳng PCA", draw_pca)

    profile = pd.DataFrame(z, columns=FEATURES).assign(cluster=labels).groupby("cluster").mean()
    def draw_heat(draw, w, h):
        left, top = 120, 92
        cell_w, cell_h = 86, 55
        minv, maxv = -2.0, 5.0
        for j, f in enumerate(FEATURES):
            draw.text((left + j * cell_w + 5, top - 30), LABELS[f][:11], fill=GRAY, font=font(14, True))
        for i, c in enumerate(profile.index):
            draw.text((left - 64, top + i * cell_h + 16), f"Cụm {c}", fill=GRAY, font=font(17))
            for j, f in enumerate(FEATURES):
                v = float(profile.loc[c, f])
                t = max(0, min(1, (v - minv) / (maxv - minv)))
                # blue for low values, red for high values
                r = int(240 * t + 40 * (1 - t)); b = int(240 * (1 - t) + 40 * t)
                x, y = left + j * cell_w, top + i * cell_h
                draw.rectangle((x, y, x + cell_w - 3, y + cell_h - 3), fill=(r, 80, b), outline="white")
                draw.text((x + 22, y + 17), f"{v:.1f}", fill="white", font=font(16, True))
    write_png(FIGURES / "04_profile_heatmap.png", 980, 370, "Hồ sơ trung bình chuẩn hóa theo cụm", draw_heat)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, color="D9D9D9"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:color"), color)


def set_cell_text(cell, text, bold=False, color="000000", size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_borders(cell)


def add_table(doc, headers, rows, widths=None, font_size=8.5):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        set_cell_text(table.rows[0].cells[i], h, True, "FFFFFF", font_size)
        set_cell_shading(table.rows[0].cells[i], "1F4E79")
    for row_idx, row in enumerate(rows):
        cells = table.add_row().cells
        for i, value in enumerate(row):
            set_cell_text(cells[i], value, False, "000000", font_size)
            if row_idx % 2 == 1:
                set_cell_shading(cells[i], "F2F6FA")
    if widths:
        for row in table.rows:
            for i, width in enumerate(widths):
                row.cells[i].width = Inches(width)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)
    return table


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    p.add_run(text)
    return p


def add_body(doc, text, bold_lead=None):
    p = doc.add_paragraph(style="Body Text")
    if bold_lead and text.startswith(bold_lead):
        p.add_run(bold_lead).bold = True
        p.add_run(text[len(bold_lead):])
    else:
        p.add_run(text)
    return p


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(90, 90, 90)


def add_code(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.right_indent = Cm(0.5)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(5)
    for line_no, line in enumerate(text.splitlines()):
        r = p.add_run(line)
        r.font.name = "Courier New"
        r._element.rPr.rFonts.set(qn("w:ascii"), "Courier New")
        r._element.rPr.rFonts.set(qn("w:hAnsi"), "Courier New")
        r.font.size = Pt(8.2)
        if line_no < len(text.splitlines()) - 1:
            r.add_break()
    return p


def setup_document() -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.0)
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    normal.font.size = Pt(11)
    body = doc.styles["Body Text"]
    body.font.name = "Arial"
    body._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    body._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    body.font.size = Pt(11)
    body.paragraph_format.line_spacing = 1.15
    body.paragraph_format.space_after = Pt(7)
    for name, size in [("Title", 20), ("Heading 1", 15), ("Heading 2", 12), ("Heading 3", 11)]:
        style = doc.styles[name]
        style.font.name = "Arial"
        style._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
        style._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
    header = sec.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.add_run("Tiểu luận Học máy và Khai phá dữ liệu").font.size = Pt(8)
    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("K-Means và phân cụm quốc gia cần viện trợ nhân đạo").font.size = Pt(8)
    return doc


def add_title_page(doc):
    for _ in range(5):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("TIỂU LUẬN MÔN HỌC MÁY VÀ KHAI PHÁ DỮ LIỆU")
    r.bold = True; r.font.size = Pt(15); r.font.name = "Arial"
    p = doc.add_paragraph(style="Title")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Thuật toán K-Means và Ứng dụng trong phân cụm quốc gia cần viện trợ nhân đạo")
    for _ in range(2): doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Trình độ: Thạc sĩ Khoa học dữ liệu").font.size = Pt(12)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Năm học 2026").font.size = Pt(12)
    doc.add_page_break()


def build_report(df, metrics, labels, profile, priority_df, pca_ratio, centers):
    doc = setup_document()
    add_title_page(doc)
    add_heading(doc, "Mục lục", 1)
    toc = [
        "Tóm tắt",
        "1. Đặt vấn đề",
        "2. Cơ sở lý thuyết về K-Means",
        "3. Ví dụ tính tay",
        "4. Dữ liệu và bài toán ứng dụng",
        "5. Thiết kế thực nghiệm và cài đặt",
        "6. Kết quả phân cụm",
        "7. Thảo luận ứng dụng trong viện trợ nhân đạo",
        "8. Kết luận và hướng phát triển",
        "Tài liệu tham khảo",
    ]
    for i, item in enumerate(toc):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.5)
        p.add_run(item)
    doc.add_page_break()

    add_heading(doc, "Tóm tắt", 1)
    add_body(doc, "Tiểu luận trình bày nguyên lý, quy trình tính toán và cách đánh giá thuật toán K-Means trong bài toán phân cụm quốc gia theo các chỉ báo kinh tế xã hội. Dữ liệu gồm 167 quốc gia và 9 biến: tử vong trẻ em, xuất khẩu, chi y tế, nhập khẩu, thu nhập, lạm phát, tuổi thọ, mức sinh và GDP bình quân đầu người. Sau khi loại tên quốc gia khỏi ma trận đặc trưng và chuẩn hóa các biến bằng Z-score, nghiên cứu thử số cụm k từ 2 đến 10. Elbow, Silhouette, Calinski-Harabasz, Davies-Bouldin và khả năng diễn giải chính sách đều được xem xét.")
    add_body(doc, "Kết quả chọn k = 4. Bốn cụm tạo ra một nhóm có mức dễ tổn thương kinh tế xã hội cao gồm 47 quốc gia, một nhóm trung gian gồm 87 quốc gia, một nhóm có mức phát triển cao gồm 30 quốc gia và một nhóm ngoại lệ gồm Luxembourg, Malta và Singapore. Cụm dễ tổn thương có trung bình tử vong trẻ em 92,96 trên 1.000 trẻ sinh sống, thu nhập 3.942,40 USD/người, tuổi thọ 59,19 năm, mức sinh 5,01 con/phụ nữ và GDP/người 1.922,38 USD. Các kết quả này phù hợp với mục tiêu tạo một danh sách sàng lọc ban đầu cho phân bổ nghiên cứu nhu cầu viện trợ.")
    add_body(doc, "Kết luận quan trọng là K-Means chỉ phát hiện cấu trúc tương đồng trong bộ biến đã quan sát. Cụm được gọi là ưu tiên cao không đồng nghĩa chắc chắn đang có khủng hoảng nhân đạo. Quyết định viện trợ cần kết hợp thêm dữ liệu về xung đột, thiên tai, di dời, an ninh lương thực, dịch bệnh, tiếp cận dịch vụ và đánh giá thực địa.")

    add_heading(doc, "1. Đặt vấn đề", 1)
    add_body(doc, "Viện trợ nhân đạo thường phải bắt đầu trong điều kiện thông tin không đầy đủ và nguồn lực có hạn. Một tổ chức có thể cần trả lời nhanh câu hỏi quốc gia nào có đặc điểm kinh tế xã hội cho thấy sức chống chịu thấp, từ đó ưu tiên thu thập thông tin và phối hợp đánh giá. Nếu chỉ nhìn từng chỉ tiêu riêng lẻ, người phân tích dễ bỏ qua mối liên hệ giữa thu nhập, tuổi thọ, tử vong trẻ em và mức sinh. Phân cụm không giám sát cung cấp một cách tiếp cận khám phá để nhận diện các nhóm quốc gia có hồ sơ tương đồng mà không cần nhãn “cần viện trợ” có sẵn.")
    add_body(doc, "Đề tài tập trung vào K-Means vì đây là thuật toán nền tảng, dễ giải thích, phù hợp với dữ liệu số và cho phép trình bày đầy đủ chu trình từ khởi tạo tâm cụm, gán điểm, cập nhật tâm đến hội tụ. Bài toán ứng dụng được phát biểu như sau: cho mỗi quốc gia một vector gồm 9 chỉ báo, hãy chia các quốc gia thành một số nhóm đồng nhất tương đối; sau đó mô tả hồ sơ từng nhóm và xác định nhóm có biểu hiện dễ tổn thương kinh tế xã hội cao hơn để sàng lọc nghiên cứu viện trợ.")
    add_heading(doc, "1.1. Mục tiêu nghiên cứu", 2)
    add_body(doc, "Mục tiêu lý thuyết là giải thích hàm mục tiêu, quy trình Lloyd, vai trò của chuẩn hóa, khởi tạo và các chỉ số đánh giá. Mục tiêu thực nghiệm là xây dựng mã tái lập, kiểm tra chất lượng dữ liệu, chọn số cụm có cơ sở và diễn giải kết quả bằng ngôn ngữ chính sách. Mục tiêu ứng dụng không phải tạo một nhãn pháp lý hay thay thế đánh giá nhu cầu nhân đạo, mà là tạo danh sách sàng lọc có thể hỗ trợ bước phân tích tiếp theo.")

    add_heading(doc, "2. Cơ sở lý thuyết về K-Means", 1)
    add_heading(doc, "2.1. Bài toán phân cụm", 2)
    add_body(doc, "Phân cụm là nhóm các quan sát sao cho các quan sát trong cùng nhóm gần nhau theo một độ đo, còn các nhóm khác nhau có xu hướng cách xa nhau. Trong học không giám sát, dữ liệu không đi kèm biến đích. Vì vậy, thuật toán không học cách dự đoán nhãn đã biết mà tìm một cấu trúc tiềm ẩn dựa trên đặc trưng và tiêu chí tối ưu được lựa chọn.")
    add_body(doc, "Với n quan sát x_i thuộc R^p và k tâm cụm mu_j, K-Means tối thiểu hóa tổng bình phương khoảng cách trong cụm: J = tổng theo i của ||x_i - mu_{c_i}||^2. Trong đó c_i là cụm được gán cho quan sát i. Khi đã cố định tâm cụm, bước gán chọn tâm gần nhất. Khi đã cố định cách gán, tâm tối ưu của một cụm là trung bình cộng các điểm thuộc cụm đó. Hai bước luân phiên cho tới khi nhãn hoặc hàm mục tiêu không còn thay đổi đáng kể.")
    add_heading(doc, "2.2. Quy trình Lloyd", 2)
    add_body(doc, "Quy trình phổ biến gồm bốn bước. Thứ nhất, chọn k tâm ban đầu. Thứ hai, tính khoảng cách Euclid từ mỗi điểm đến từng tâm và gán điểm vào tâm gần nhất. Thứ ba, tính lại mỗi tâm bằng trung bình các điểm trong cụm. Thứ tư, lặp lại bước hai và ba cho tới khi hội tụ hoặc đạt số vòng lặp tối đa. Mỗi vòng lặp không làm tăng hàm mục tiêu, nhưng thuật toán chỉ bảo đảm hội tụ về một nghiệm cực tiểu cục bộ, không bảo đảm nghiệm toàn cục.")
    add_body(doc, "Trong thực hành, mã sử dụng khởi tạo K-Means++ và chạy nhiều lần với n_init = 50. K-Means++ ưu tiên chọn các tâm ban đầu cách xa nhau hơn, còn nhiều lần khởi tạo giúp giảm nguy cơ chọn phải nghiệm kém do seed cụ thể. Tất cả kết quả trong báo cáo cố định random_state = 42 để có thể tái lập.")
    add_heading(doc, "2.3. Chuẩn hóa và lựa chọn khoảng cách", 2)
    add_body(doc, "Dữ liệu có đơn vị và thang đo khác nhau. GDP/người và thu nhập có thể lên tới hàng chục nghìn, trong khi tỷ lệ chi y tế chỉ ở mức vài phần trăm. Nếu áp dụng khoảng cách Euclid trên dữ liệu thô, biến có độ lớn số học lớn sẽ chi phối hàm mục tiêu. Nghiên cứu biến đổi từng biến theo công thức z = (x - trung bình) / độ lệch chuẩn. Sau chuẩn hóa, mỗi biến có trung bình xấp xỉ 0 và độ lệch chuẩn xấp xỉ 1, nên đóng góp của các biến cân bằng hơn.")
    add_heading(doc, "2.4. Đánh giá chất lượng cụm", 2)
    add_body(doc, "Elbow theo dõi SSE trong cụm. SSE luôn giảm khi tăng k nên không thể chọn chỉ dựa trên giá trị nhỏ nhất; điểm gãy thể hiện mức lợi ích biên bắt đầu giảm. Silhouette kết hợp độ gắn kết trong cụm và độ tách biệt với cụm gần nhất. Giá trị càng cao càng thuận lợi, nhưng cần diễn giải theo ngữ cảnh. Calinski-Harabasz lớn hơn thường cho thấy tỷ lệ phân tán giữa cụm so với trong cụm tốt hơn. Davies-Bouldin nhỏ hơn thường tốt hơn vì cụm có độ phân tán nhỏ và xa nhau.")
    add_body(doc, "Không có một chỉ số đơn lẻ nào định nghĩa số cụm đúng. Đặc biệt, k lớn có thể làm điểm đánh giá tăng nhờ tách các ngoại lệ thành cụm rất nhỏ. Do đó, báo cáo kết hợp chỉ số nội tại, kích thước cụm, ổn định do nhiều khởi tạo và khả năng giải thích đối với mục tiêu sàng lọc viện trợ.")

    add_heading(doc, "3. Ví dụ tính tay", 1)
    add_body(doc, "Xét sáu quốc gia giả lập được mô tả bởi hai biến đã chuẩn hóa: mức nghèo P và chỉ báo phát triển D. Mục tiêu là k = 2. Ta dùng các điểm A(1,1), B(1,2), C(2,1), D(8,8), E(9,8), F(8,9), và chọn tâm ban đầu mu_1 = A(1,1), mu_2 = D(8,8).")
    add_body(doc, "Ở bước gán đầu tiên, A, B và C gần mu_1 hơn; D, E và F gần mu_2 hơn. Ví dụ khoảng cách của B đến mu_1 là sqrt((1-1)^2 + (2-1)^2) = 1, còn đến mu_2 là sqrt((1-8)^2 + (2-8)^2) = sqrt(85), nên B thuộc cụm 1. Tương tự, ba điểm cuối thuộc cụm 2.")
    add_body(doc, "Cập nhật tâm: mu_1 mới = ((1+1+2)/3, (1+2+1)/3) = (1, 1,33); mu_2 mới = ((8+9+8)/3, (8+8+9)/3) = (8,33, 8,33). Sau cập nhật, các điểm vẫn gần tâm tương ứng. Thuật toán hội tụ ở lần lặp tiếp theo. SSE cuối bằng tổng bình phương khoảng cách của sáu điểm đến tâm cụm tương ứng, xấp xỉ 2,67.")
    add_body(doc, "Ví dụ cho thấy K-Means tối ưu một tiêu chí hình học. Nhãn cụm 1 và cụm 2 không có ý nghĩa tự nhiên trước khi ta xem tâm và hồ sơ của chúng. Trong bài toán thực tế, bước diễn giải sau phân cụm quan trọng không kém bước tính toán.")

    add_heading(doc, "4. Dữ liệu và bài toán ứng dụng", 1)
    add_heading(doc, "4.1. Nguồn và cấu trúc dữ liệu", 2)
    add_body(doc, "Dữ liệu được tải từ bộ Unsupervised Learning on Country Data trên Kaggle. Tệp Country-data.csv có 167 dòng quốc gia và 10 cột, trong đó country là định danh và 9 cột còn lại là biến số. Tệp data-dictionary.csv mô tả ý nghĩa đơn vị của các cột. Bộ dữ liệu phù hợp để minh họa phân cụm kinh tế xã hội, nhưng không phải bộ dữ liệu nhu cầu nhân đạo chuyên biệt.")
    data_rows = [
        ["child_mort", "Tử vong trẻ dưới 5 tuổi trên 1.000 trẻ sinh sống", "Càng cao càng bất lợi"],
        ["exports", "Xuất khẩu hàng hóa và dịch vụ, % GDP/người", "Cấu trúc thương mại"],
        ["health", "Chi y tế tổng cộng, % GDP/người", "Nguồn lực y tế tương đối"],
        ["imports", "Nhập khẩu hàng hóa và dịch vụ, % GDP/người", "Cấu trúc thương mại"],
        ["income", "Thu nhập ròng trên mỗi người", "Mức sống kinh tế"],
        ["inflation", "Tốc độ tăng GDP hằng năm", "Bất ổn vĩ mô"],
        ["life_expec", "Tuổi thọ kỳ vọng của trẻ sơ sinh", "Càng thấp càng bất lợi"],
        ["total_fer", "Số con trên mỗi phụ nữ", "Áp lực dân số"],
        ["gdpp", "GDP bình quân đầu người", "Mức phát triển kinh tế"],
    ]
    add_table(doc, ["Biến", "Ý nghĩa", "Vai trò diễn giải"], data_rows, [1.1, 4.0, 1.5], 8.4)
    add_caption(doc, "Bảng 1. Các biến sử dụng trong ma trận đặc trưng")
    add_heading(doc, "4.2. Phát biểu input và output", 2)
    add_body(doc, "Input là ma trận X gồm 167 quốc gia và 9 biến số. Cột country chỉ dùng để truy xuất kết quả, không đưa vào khoảng cách. Tiền xử lý gồm kiểm tra kiểu dữ liệu, giá trị thiếu, dòng trùng, sau đó chuẩn hóa. Output chính là nhãn cluster của từng quốc gia, tâm cụm trên thang đo gốc, hồ sơ trung bình theo cụm, bảng đánh giá k và các biểu đồ. Output ứng dụng là danh sách nhóm có hồ sơ dễ tổn thương hơn để sàng lọc nghiên cứu nhu cầu.")
    add_heading(doc, "4.3. Kiểm tra chất lượng dữ liệu", 2)
    add_body(doc, "Tệp có 167 quan sát, 9 biến đầu vào số, không có giá trị thiếu và không có dòng trùng hoàn toàn. Dù vậy, các biến có tương quan và có ngoại lệ lớn. Xuất khẩu và nhập khẩu theo tỷ lệ GDP đặc biệt cao ở một số nền kinh tế có vai trò trung chuyển hoặc cấu trúc thương mại đặc thù. Vì vậy, kết quả phân cụm cần đọc cùng với tâm cụm và không nên suy diễn một biến thành nguyên nhân nhân quả.")

    add_heading(doc, "5. Thiết kế thực nghiệm và cài đặt", 1)
    add_heading(doc, "5.1. Quy trình thực nghiệm", 2)
    add_body(doc, "Quy trình gồm: đọc Country-data.csv; chọn 9 biến; kiểm tra thiếu và trùng; chuẩn hóa bằng StandardScaler; thử k từ 2 đến 10; chạy K-Means với K-Means++ và n_init = 50; ghi lại inertia, Silhouette, Calinski-Harabasz, Davies-Bouldin và kích thước cụm; chọn k = 4; khôi phục tâm cụm về thang đo gốc; chiếu PCA để trực quan hóa; xây dựng một chỉ báo sàng lọc minh bạch chỉ cho mục đích diễn giải.")
    add_heading(doc, "5.2. Mã cốt lõi", 2)
    add_code(doc, "scaler = StandardScaler()\nZ = scaler.fit_transform(df[FEATURES])\n\nmetrics = []\nfor k in range(2, 11):\n    model = KMeans(n_clusters=k, n_init=50, random_state=42)\n    labels = model.fit_predict(Z)\n    metrics.append({\n        'k': k,\n        'inertia': model.inertia_,\n        'silhouette': silhouette_score(Z, labels),\n    })\n\nmodel = KMeans(n_clusters=4, n_init=50, random_state=42)\nlabels = model.fit_predict(Z)\ncenters_original = scaler.inverse_transform(model.cluster_centers_)")
    add_body(doc, "Mã đầy đủ nằm trong tệp kmeans_humanitarian.py. Mã tách rõ dữ liệu đầu vào, đánh giá số cụm, mô hình cuối cùng, hồ sơ cụm, bảng sàng lọc và biểu đồ. Tệp requirements.txt ghi các thư viện cần cài để chạy lại trong môi trường Python thông thường.")

    add_heading(doc, "6. Kết quả phân cụm", 1)
    add_heading(doc, "6.1. Chọn số cụm", 2)
    rows = []
    for _, r in metrics.iterrows():
        rows.append([int(r.k), f"{r.inertia:.3f}", f"{r.silhouette:.4f}", f"{r.calinski_harabasz:.2f}", f"{r.davies_bouldin:.4f}", int(r.smallest_cluster)])
    add_table(doc, ["k", "SSE", "Silhouette", "Calinski", "Davies Bouldin", "Cụm nhỏ nhất"], rows, [0.5, 1.0, 1.1, 1.2, 1.3, 1.0], 8.2)
    add_caption(doc, "Bảng 2. Các chỉ số nội tại khi thay đổi số cụm")
    add_body(doc, "SSE giảm từ 1.050,215 ở k = 2 xuống 700,323 ở k = 4, sau đó mức giảm biên nhỏ dần. Silhouette đạt 0,3014 ở k = 4 và 0,3052 ở k = 5. Tuy k = 5 cao hơn rất nhỏ, nghiệm k = 5 tách một quốc gia thành cụm riêng và giữ một cụm chỉ gồm ba quốc gia có hồ sơ thương mại rất đặc biệt. Với mục tiêu sàng lọc và giải thích, k = 4 cân bằng hơn giữa chất lượng hình học, kích thước cụm và ý nghĩa chính sách. Đây là một lựa chọn có chủ đích, không phải khẳng định k = 4 là nghiệm duy nhất đúng.")
    doc.add_picture(str(FIGURES / "01_elbow.png"), width=Inches(5.9))
    add_caption(doc, "Hình 1. Đường Elbow với đường tham chiếu k = 4")
    doc.add_picture(str(FIGURES / "02_silhouette.png"), width=Inches(5.9))
    add_caption(doc, "Hình 2. Điểm Silhouette theo số cụm")

    add_heading(doc, "6.2. Hồ sơ các cụm", 2)
    profile_rows = []
    for raw in profile.index:
        p = profile.loc[raw]
        profile_rows.append([
            f"{int(p.cluster_priority_rank)} (mã {raw})", int(p.n_countries),
            f"{p.child_mort:.2f}", f"{p.income:,.2f}", f"{p.life_expec:.2f}", f"{p.total_fer:.2f}", f"{p.gdpp:,.2f}",
        ])
    add_table(doc, ["Ưu tiên", "Số nước", "Tử vong trẻ em", "Thu nhập", "Tuổi thọ", "Mức sinh", "GDP/người"], profile_rows, [0.7, 0.7, 1.1, 1.1, 0.9, 0.8, 1.0], 8.0)
    add_caption(doc, "Bảng 3. Trung bình cụm trên thang đo gốc; ưu tiên 1 là hồ sơ dễ tổn thương hơn")
    add_body(doc, "Cụm ưu tiên 1 có 47 quốc gia và là nhóm nổi bật nhất về tính dễ tổn thương kinh tế xã hội: tử vong trẻ em và mức sinh cao, trong khi thu nhập, tuổi thọ và GDP/người thấp. Cụm ưu tiên 2 có 87 quốc gia với mức trung gian. Cụm ưu tiên 3 có 30 quốc gia với tuổi thọ và thu nhập cao hơn, còn cụm ưu tiên 4 chỉ có Luxembourg, Malta và Singapore. Cụm nhỏ này có tỷ trọng xuất khẩu và nhập khẩu rất cao, cho thấy K-Means đang nhận diện một cấu trúc kinh tế đặc biệt chứ không phải một nhóm “nhu cầu thấp” theo mọi chiều.")
    doc.add_picture(str(FIGURES / "03_pca_clusters.png"), width=Inches(6.1))
    add_caption(doc, "Hình 3. Biểu diễn các quốc gia trên hai thành phần PCA")
    doc.add_picture(str(FIGURES / "04_profile_heatmap.png"), width=Inches(6.3))
    add_caption(doc, "Hình 4. Hồ sơ trung bình chuẩn hóa; màu đỏ biểu thị giá trị cao hơn trong dữ liệu")
    add_heading(doc, "6.3. Danh sách sàng lọc ưu tiên cao", 2)
    high = priority_df[priority_df.cluster_priority_rank == 1].copy()
    # Keep table readable by showing all names in a compact two-column table.
    first = ", ".join(high.country.iloc[:24].tolist())
    second = ", ".join(high.country.iloc[24:].tolist())
    add_table(doc, ["Nhóm", "Các quốc gia"], [["Ưu tiên 1, phần 1", first], ["Ưu tiên 1, phần 2", second]], [1.4, 5.0], 8.0)
    add_caption(doc, "Bảng 4. 47 quốc gia trong cụm có hồ sơ dễ tổn thương kinh tế xã hội cao nhất")
    add_body(doc, "Điểm sàng lọc trong results/aid_screening.csv được xây dựng từ các chiều quan sát theo hướng: tăng khi tử vong trẻ em, mức sinh và lạm phát cao; giảm khi thu nhập, tuổi thọ, GDP/người và tỷ lệ chi y tế cao. Điểm này chỉ dùng để sắp xếp tương đối các nước trong bộ dữ liệu. Nó không phải xác suất cần viện trợ, không có ngưỡng chính sách và không đại diện cho mức độ khẩn cấp theo thời gian thực.")

    add_heading(doc, "7. Thảo luận ứng dụng trong viện trợ nhân đạo", 1)
    add_heading(doc, "7.1. Cách sử dụng kết quả", 2)
    add_body(doc, "Kết quả có thể được dùng như một lớp sàng lọc trong hệ thống phân tích nhiều tầng. Ở tầng đầu, nhóm ưu tiên 1 giúp cơ quan điều phối thu hẹp phạm vi rà soát và tìm thêm dữ liệu. Ở tầng hai, nhà phân tích ghép từng quốc gia với dữ liệu tình hình hiện tại như xung đột, thiên tai, di dời, lương thực, dịch bệnh và khả năng tiếp cận dịch vụ. Ở tầng ba, các chuyên gia khu vực xác thực bối cảnh và xác định loại hỗ trợ, quy mô, thời điểm cùng ràng buộc logistics. K-Means chỉ đóng vai trò ở tầng khám phá và sắp xếp thông tin.")
    add_body(doc, "Ưu điểm của cách làm là tạo ra một hồ sơ đa biến dễ đọc. Thay vì nói riêng rằng một nước có GDP thấp, ta quan sát đồng thời tử vong trẻ em, tuổi thọ, mức sinh và thu nhập. Tâm cụm trên thang đo gốc cho phép chuyển kết quả toán học thành mô tả mà nhà hoạch định chính sách có thể kiểm tra. Các tệp CSV đi kèm cũng giúp tái lập và truy vết từ một quốc gia đến cụm, tâm cụm và điểm sàng lọc.")
    add_heading(doc, "7.2. Rủi ro diễn giải và đạo đức dữ liệu", 2)
    add_body(doc, "Một cụm là kết quả phụ thuộc vào biến, thời điểm, chuẩn hóa, khoảng cách và k. Nếu thêm hoặc bỏ một biến, ranh giới cụm có thể đổi. Dữ liệu tổng hợp theo quốc gia che khuất chênh lệch trong nội bộ quốc gia, nên không thể dùng để suy luận nhu cầu của từng tỉnh, cộng đồng hoặc hộ gia đình. Xếp hạng dựa trên chỉ báo kinh tế cũng có thể bỏ sót khủng hoảng đột ngột tại một quốc gia có trung bình kinh tế tốt.")
    add_body(doc, "Cần tránh dùng nhãn cụm như một quyết định tự động để phân bổ hoặc từ chối viện trợ. Khi triển khai, cần công bố biến đầu vào, ngày cập nhật, quy tắc chuẩn hóa, phiên bản mã, giới hạn dữ liệu và quy trình khiếu nại hoặc rà soát chuyên gia. Việc ưu tiên viện trợ phải có giám sát con người và cơ chế cập nhật khi thông tin hiện trường mâu thuẫn với hồ sơ thống kê.")
    add_heading(doc, "7.3. Hướng cải thiện mô hình", 2)
    add_body(doc, "Nghiên cứu tiếp theo có thể bổ sung chỉ báo xung đột, số người di dời, chỉ số an ninh lương thực, số ca dịch bệnh, thiệt hại thiên tai, tiếp cận nước sạch, năng lực y tế và khoảng cách logistics. Khi dữ liệu có thời gian, nên dùng phân cụm theo chuỗi thời gian hoặc cập nhật theo quý để phát hiện quốc gia có rủi ro tăng nhanh. Có thể so sánh K-Means với K-Medoids, Gaussian Mixture, phân cụm phân cấp và DBSCAN; đồng thời đánh giá độ ổn định bằng bootstrap và phân tích độ nhạy theo biến.")
    add_body(doc, "Một hướng khác là xác định trọng số với sự tham gia của chuyên gia. K-Means chuẩn hóa hiện tại coi các biến có vai trò tương đương trong khoảng cách, trong khi mục tiêu nhân đạo có thể muốn ưu tiên tử vong trẻ em hoặc tiếp cận y tế. Trọng số cần được công khai, thử nghiệm độ nhạy và kiểm tra tác động để tránh biến một giả định giá trị thành một quy tắc ẩn trong thuật toán.")

    add_heading(doc, "8. Kết luận và hướng phát triển", 1)
    add_body(doc, "Tiểu luận đã trình bày K-Means từ hàm mục tiêu đến quy trình Lloyd, minh họa bằng ví dụ tính tay, sau đó áp dụng trên 167 quốc gia với 9 biến kinh tế xã hội. Chuẩn hóa là bước cần thiết vì dữ liệu có đơn vị khác nhau. Thực nghiệm thử k từ 2 đến 10 và kết hợp nhiều chỉ số cho thấy k = 4 là lựa chọn phù hợp với mục tiêu diễn giải, dù k = 5 có Silhouette nhỉnh hơn rất nhỏ.")
    add_body(doc, "Kết quả cuối tạo ra một cụm 47 quốc gia có hồ sơ dễ tổn thương nhất theo các chỉ báo quan sát. Cụm này có trung bình tử vong trẻ em cao, mức sinh cao, tuổi thọ thấp, thu nhập thấp và GDP/người thấp. Phát hiện có thể hỗ trợ sàng lọc và lập kế hoạch thu thập dữ liệu, nhưng không thể thay thế đánh giá nhân đạo thực tế vì bộ dữ liệu thiếu các chiều khẩn cấp và bối cảnh.")
    add_body(doc, "Hướng phát triển ưu tiên là bổ sung dữ liệu động và dữ liệu tình hình, kiểm tra độ ổn định của cụm, đánh giá công bằng giữa các vùng và xây dựng quy trình human-in-the-loop. Khi đó, K-Means có thể trở thành một thành phần minh bạch trong hệ thống hỗ trợ quyết định, thay vì là một nhãn tự động đứng độc lập.")

    add_heading(doc, "Tài liệu tham khảo", 1)
    refs = [
        "Kaggle. Unsupervised Learning on Country Data. https://www.kaggle.com/datasets/rohan0301/unsupervised-learning-on-country-data",
        "Lloyd, S. (1982). Least squares quantization in PCM. IEEE Transactions on Information Theory, 28(2), 129–137. https://doi.org/10.1109/TIT.1982.1056489",
        "MacQueen, J. B. (1967). Some methods for classification and analysis of multivariate observations. Proceedings of the Fifth Berkeley Symposium on Mathematical Statistics and Probability, 1, 281–297.",
        "Rousseeuw, P. J. (1987). Silhouettes: A graphical aid to the interpretation and validation of cluster analysis. Journal of Computational and Applied Mathematics, 20, 53–65. https://doi.org/10.1016/0377-0427(87)90125-7",
        "Scikit-learn developers. Clustering documentation and silhouette analysis. https://scikit-learn.org/stable/modules/clustering.html",
        "Scikit-learn developers. Selecting the number of clusters with silhouette analysis on KMeans clustering. https://scikit-learn.org/stable/auto_examples/cluster/plot_kmeans_silhouette_analysis.html",
    ]
    for ref in refs:
        p = doc.add_paragraph(style="Body Text")
        p.paragraph_format.left_indent = Cm(0.5)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        p.add_run(ref)
    doc.save(REPORT)


def main():
    df = pd.read_csv(DATA)
    X = df[FEATURES].to_numpy(float)
    mean, std = X.mean(axis=0), X.std(axis=0, ddof=0)
    Z = (X - mean) / std
    rows = []
    saved = {}
    for k in range(2, 11):
        inertia, labels, centers = kmeans_pp(Z, k)
        rows.append({
            "k": k, "inertia": inertia, "silhouette": silhouette(Z, labels),
            "calinski_harabasz": calinski(Z, labels), "davies_bouldin": davies_bouldin(Z, labels),
            "smallest_cluster": int(np.bincount(labels).min()),
        })
        saved[k] = (inertia, labels, centers)
    metrics = pd.DataFrame(rows)
    inertia, labels, centers = saved[4]
    RESULTS.mkdir(exist_ok=True)
    FIGURES.mkdir(exist_ok=True)
    metrics.to_csv(RESULTS / "k_evaluation.csv", index=False)
    out = df[["country"] + FEATURES].copy(); out["cluster"] = labels
    out.to_csv(RESULTS / "country_clusters.csv", index=False)
    original_centers = pd.DataFrame(centers * std + mean, columns=FEATURES)
    original_centers.index.name = "cluster"
    original_centers.to_csv(RESULTS / "cluster_centers_original_scale.csv")
    profile = df.assign(cluster=labels).groupby("cluster")[FEATURES].mean()
    profile["n_countries"] = df.assign(cluster=labels).groupby("cluster").size()
    profile_z = (profile[FEATURES] - mean) / std
    profile["cluster_screening_score"] = profile_z.child_mort + profile_z.total_fer + .5 * profile_z.inflation - profile_z.income - profile_z.life_expec - profile_z.gdpp - .5 * profile_z.health
    order = profile.cluster_screening_score.sort_values(ascending=False).index
    rank = {c: i + 1 for i, c in enumerate(order)}
    profile["cluster_priority_rank"] = profile.index.map(rank)
    profile.sort_values("cluster_priority_rank").to_csv(RESULTS / "cluster_profile.csv")
    priority = df[["country"]].copy(); priority["cluster"] = labels
    zdf = pd.DataFrame(Z, columns=FEATURES)
    priority["screening_score_raw"] = zdf.child_mort + zdf.total_fer + .5 * zdf.inflation - zdf.income - zdf.life_expec - zdf.gdpp - .5 * zdf.health
    raw = priority.screening_score_raw
    priority["screening_score_0_100"] = 100 * (raw - raw.min()) / (raw.max() - raw.min())
    priority["cluster_priority_rank"] = priority.cluster.map(rank)
    priority.sort_values(["cluster_priority_rank", "screening_score_0_100"], ascending=[True, False]).to_csv(RESULTS / "aid_screening.csv", index=False)
    cov = np.cov(Z, rowvar=False); vals, vecs = np.linalg.eigh(cov); vals = vals[::-1]; vecs = vecs[:, ::-1]
    pca_coords = Z @ vecs[:, :2]
    pd.DataFrame({"country": df.country, "PC1": pca_coords[:, 0], "PC2": pca_coords[:, 1], "cluster": labels}).to_csv(RESULTS / "pca_coordinates.csv", index=False)
    create_figures(metrics, Z, labels)
    summary = {"n_countries": len(df), "n_features": len(FEATURES), "selected_k": 4, "seed": 42, "n_init": 50, "inertia": float(inertia), "silhouette": float(silhouette(Z, labels)), "calinski_harabasz": float(calinski(Z, labels)), "davies_bouldin": float(davies_bouldin(Z, labels)), "pca_variance_ratio": (vals / vals.sum())[:2].tolist(), "cluster_sizes": {str(k): int(v) for k, v in zip(*np.unique(labels, return_counts=True))}}
    (RESULTS / "run_summary.json").write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    build_report(df, metrics, labels, profile.sort_values("cluster_priority_rank"), priority.sort_values(["cluster_priority_rank", "screening_score_0_100"], ascending=[True, False]), (vals / vals.sum())[:2], original_centers)
    print(json.dumps(summary, indent=2, ensure_ascii=False))
    print(f"Created {REPORT}")


if __name__ == "__main__":
    main()
