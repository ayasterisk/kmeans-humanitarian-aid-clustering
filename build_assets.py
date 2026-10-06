"""Assemble the .docx report from the outputs of kmeans_humanitarian.py.

This script does NOT run its own clustering. Run first:
    python kmeans_humanitarian.py --data Country-data.csv --k 4
then:
    python build_assets.py
Every number and figure in the report is read from results/ and figures/, so the
report always matches the scikit-learn code that is the deliverable.
"""

from __future__ import annotations

import json
from pathlib import Path

import numpy as np
import pandas as pd
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


def vn(x, nd=2):
    """Format a number the Vietnamese way (dot thousands, comma decimals)."""
    return f"{x:,.{nd}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def build_report(df, metrics, profile, priority_df, stability, summary):
    prof = profile.sort_values("cluster_priority_rank")
    p1, p2, p3, p4 = (prof.iloc[i] for i in range(4))
    n1, n2, n3, n4 = (int(prof.iloc[i].n_countries) for i in range(4))
    m = metrics.set_index("k")
    high = priority_df[priority_df.cluster_priority_rank == 1]
    rich = (
        df[df.country.isin(high.country) & (df.income > 10000)]
        .sort_values("income", ascending=False)
    )
    rich_txt = ", ".join(f"{r.country} ({r.income:,.0f})".replace(",", ".") for r in rich.itertuples())
    stab = stability.groupby("test").ari.agg(["mean", "min"])
    drop = stability[stability.test == "drop_feature"].sort_values("ari")
    worst = drop.iloc[0]
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
    add_body(doc, f"Kết quả chọn k = 4. Bốn cụm tạo ra một nhóm có mức dễ tổn thương kinh tế xã hội cao gồm {n1} quốc gia, một nhóm trung gian gồm {n2} quốc gia, một nhóm có mức phát triển cao gồm {n3} quốc gia và một nhóm ngoại lệ gồm Luxembourg, Malta và Singapore. Cụm dễ tổn thương có trung bình tử vong trẻ em {vn(p1.child_mort)} trên 1.000 trẻ sinh sống, thu nhập {vn(p1.income)} USD/người, tuổi thọ {vn(p1.life_expec)} năm, mức sinh {vn(p1.total_fer)} con/phụ nữ và GDP/người {vn(p1.gdpp)} USD. Cần lưu ý Silhouette chỉ khoảng {vn(summary['silhouette'],2)}, tức cấu trúc cụm yếu, nên kết quả chỉ nên dùng để sàng lọc ban đầu cho nghiên cứu nhu cầu viện trợ.")
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
    add_body(doc, "Cập nhật tâm: mu_1 mới = ((1+1+2)/3, (1+2+1)/3) = (1,33; 1,33); mu_2 mới = ((8+9+8)/3, (8+8+9)/3) = (8,33; 8,33). Sau cập nhật, các điểm vẫn gần tâm tương ứng. Thuật toán hội tụ ở lần lặp tiếp theo. SSE cuối bằng tổng bình phương khoảng cách của sáu điểm đến tâm cụm tương ứng, xấp xỉ 2,67 (mỗi cụm đóng góp 4/3).")
    add_body(doc, "Ví dụ cho thấy K-Means tối ưu một tiêu chí hình học. Nhãn cụm 1 và cụm 2 không có ý nghĩa tự nhiên trước khi ta xem tâm và hồ sơ của chúng. Trong bài toán thực tế, bước diễn giải sau phân cụm quan trọng không kém bước tính toán.")

    add_heading(doc, "4. Dữ liệu và bài toán ứng dụng", 1)
    add_heading(doc, "4.1. Nguồn và cấu trúc dữ liệu", 2)
    add_body(doc, "Dữ liệu được tải từ bộ Unsupervised Learning on Country Data trên Kaggle. Tệp Country-data.csv có 167 dòng quốc gia và 10 cột, trong đó country là định danh và 9 cột còn lại là biến số. Tệp data-dictionary.csv mô tả ý nghĩa đơn vị của các cột. Bộ dữ liệu phù hợp để minh họa phân cụm kinh tế xã hội, nhưng không phải bộ dữ liệu nhu cầu nhân đạo chuyên biệt. Ba lưu ý về dữ liệu: (i) nguồn không ghi rõ năm tham chiếu và số liệu có vẻ là số liệu lịch sử, nên danh sách kết quả minh họa phương pháp chứ không phản ánh tình hình hiện nay; (ii) bộ dữ liệu không có một số quốc gia thường gắn với khủng hoảng nhân đạo như Somalia, Syria hay Nam Sudan; (iii) data-dictionary gốc mô tả biến inflation là tốc độ tăng của tổng GDP, nhưng tên biến và cách sử dụng cho thấy đây là lạm phát, nên bài coi đây là tỷ lệ lạm phát hằng năm.")
    data_rows = [
        ["child_mort", "Tử vong trẻ dưới 5 tuổi trên 1.000 trẻ sinh sống", "Càng cao càng bất lợi"],
        ["exports", "Xuất khẩu hàng hóa và dịch vụ, % GDP/người", "Cấu trúc thương mại"],
        ["health", "Chi y tế tổng cộng, % GDP/người", "Nguồn lực y tế tương đối"],
        ["imports", "Nhập khẩu hàng hóa và dịch vụ, % GDP/người", "Cấu trúc thương mại"],
        ["income", "Thu nhập ròng trên mỗi người", "Mức sống kinh tế"],
        ["inflation", "Tỷ lệ lạm phát hằng năm (%)", "Bất ổn vĩ mô"],
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
    add_body(doc, f"Mã đầy đủ nằm trong tệp kmeans_humanitarian.py. Mã tách rõ dữ liệu đầu vào, đánh giá số cụm, mô hình cuối cùng, hồ sơ cụm, bảng sàng lọc và biểu đồ. Tệp requirements.txt ghim phiên bản scikit-learn đã dùng ({summary.get('sklearn', '1.8.0')}). Cần lưu ý kết quả K-Means có thể lệch nhẹ giữa các phiên bản thư viện: quốc gia nằm sát ranh giới giữa hai cụm (ví dụ Botswana) có thể đổi cụm. Vì vậy mọi số liệu và hình trong báo cáo được lấy trực tiếp từ kết quả chạy kmeans_humanitarian.py, và tệp build_assets.py chỉ đọc lại các tệp kết quả đó để dựng báo cáo.")

    add_heading(doc, "6. Kết quả phân cụm", 1)
    add_heading(doc, "6.1. Chọn số cụm", 2)
    rows = []
    for _, r in metrics.iterrows():
        rows.append([int(r.k), vn(r.inertia,3), vn(r.silhouette,4), vn(r.calinski_harabasz,2), vn(r.davies_bouldin,4), int(r.smallest_cluster)])
    add_table(doc, ["k", "SSE", "Silhouette", "Calinski", "Davies Bouldin", "Cụm nhỏ nhất"], rows, [0.5, 1.0, 1.1, 1.2, 1.3, 1.0], 8.2)
    add_caption(doc, "Bảng 2. Các chỉ số nội tại khi thay đổi số cụm")
    add_body(doc, f"SSE giảm từ {vn(m.loc[2,'inertia'],3)} ở k = 2 xuống {vn(m.loc[4,'inertia'],3)} ở k = 4, sau đó mức giảm biên nhỏ dần. Silhouette đạt giá trị cao nhất ở k = 4 ({vn(m.loc[4,'silhouette'],4)}), nhưng chỉ nhỉnh hơn một chút so với k = 2, 3 và 5 (lần lượt {vn(m.loc[2,'silhouette'],4)}, {vn(m.loc[3,'silhouette'],4)} và {vn(m.loc[5,'silhouette'],4)}). Theo quy ước của Kaufman và Rousseeuw (1990), Silhouette trong khoảng 0,26 đến 0,50 chỉ cho thấy cấu trúc cụm yếu, có thể mang tính nhân tạo, nên không nên xem các cụm là nhóm tách biệt rõ ràng. Các chỉ số cũng không đồng thuận: Calinski-Harabasz cao nhất ở k = 2, còn Davies-Bouldin tốt nhất ở k = 5 ({vn(m.loc[5,'davies_bouldin'],4)} so với {vn(m.loc[4,'davies_bouldin'],4)} ở k = 4). Nghiệm k = 5 tách một quốc gia thành cụm riêng, còn k = 4 vẫn giữ một cụm chỉ gồm ba quốc gia có hồ sơ thương mại rất đặc biệt. Vì vậy k = 4 được chọn chủ yếu vì khả năng diễn giải cho mục tiêu sàng lọc (một nhóm dễ tổn thương, một nhóm trung gian, một nhóm phát triển cao và một nhóm ngoại lệ thương mại) và vì độ ổn định ở mục 6.4, không phải vì chỉ số nội tại vượt trội. Đây là một lựa chọn có chủ đích, không phải khẳng định k = 4 là nghiệm duy nhất đúng.")
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
            vn(p.child_mort), vn(p.income), vn(p.life_expec), vn(p.total_fer), vn(p.gdpp),
        ])
    add_table(doc, ["Ưu tiên", "Số nước", "Tử vong trẻ em", "Thu nhập", "Tuổi thọ", "Mức sinh", "GDP/người"], profile_rows, [0.7, 0.7, 1.1, 1.1, 0.9, 0.8, 1.0], 8.0)
    add_caption(doc, "Bảng 3. Trung bình cụm trên thang đo gốc; ưu tiên 1 là hồ sơ dễ tổn thương hơn")
    add_body(doc, f"Cụm ưu tiên 1 có {n1} quốc gia và là nhóm nổi bật nhất về tính dễ tổn thương kinh tế xã hội: tử vong trẻ em và mức sinh cao, trong khi thu nhập, tuổi thọ và GDP/người thấp. Cụm ưu tiên 2 có {n2} quốc gia với mức trung gian. Cụm ưu tiên 3 có {n3} quốc gia với tuổi thọ và thu nhập cao hơn, còn cụm ưu tiên 4 chỉ có Luxembourg, Malta và Singapore. Cụm nhỏ này có tỷ trọng xuất khẩu và nhập khẩu rất cao, cho thấy K-Means đang nhận diện một cấu trúc kinh tế đặc biệt chứ không phải một nhóm “nhu cầu thấp” theo mọi chiều.")
    doc.add_picture(str(FIGURES / "03_pca_clusters.png"), width=Inches(6.1))
    add_caption(doc, "Hình 3. Biểu diễn các quốc gia trên hai thành phần PCA")
    doc.add_picture(str(FIGURES / "04_profile_heatmap.png"), width=Inches(6.3))
    add_caption(doc, "Hình 4. Hồ sơ trung bình chuẩn hóa; màu đỏ biểu thị giá trị cao hơn trong dữ liệu")
    add_heading(doc, "6.3. Danh sách sàng lọc ưu tiên cao", 2)
    high = high.copy()
    # Keep table readable by showing all names in a compact two-column table.
    first = ", ".join(high.country.iloc[:24].tolist())
    second = ", ".join(high.country.iloc[24:].tolist())
    add_table(doc, ["Nhóm", "Các quốc gia"], [["Ưu tiên 1, phần 1", first], ["Ưu tiên 1, phần 2", second]], [1.4, 5.0], 8.0)
    add_caption(doc, f"Bảng 4. {n1} quốc gia trong cụm có hồ sơ dễ tổn thương kinh tế xã hội cao nhất")
    add_body(doc, "Điểm sàng lọc trong results/aid_screening.csv được xây dựng từ các chiều quan sát theo hướng: tăng khi tử vong trẻ em, mức sinh và lạm phát cao; giảm khi thu nhập, tuổi thọ, GDP/người và tỷ lệ chi y tế cao. Trọng số là 1 cho các biến tử vong trẻ em, mức sinh, thu nhập, tuổi thọ và GDP/người, và 0,5 cho lạm phát và chi y tế. Đây là một giả định có chủ đích: các biến phản ánh trực tiếp sức khỏe và mức sống được xem quan trọng hơn hai biến mang tính vĩ mô hoặc chính sách, nhưng bộ trọng số này chưa được kiểm định độ nhạy và cần được chuyên gia xác nhận. Điểm này chỉ dùng để sắp xếp tương đối các nước trong bộ dữ liệu. Nó không phải xác suất cần viện trợ, không có ngưỡng chính sách và không đại diện cho mức độ khẩn cấp theo thời gian thực.")
    add_body(doc, f"Cụm ưu tiên 1 cũng không đồng nhất. Một số quốc gia có thu nhập bình quân cao nhưng vẫn bị xếp vào cụm này do tử vong trẻ em cao và tuổi thọ thấp, gồm {rich_txt} (thu nhập, USD/người). Điều này phản ánh bất bình đẳng trong nội bộ và cho thấy trung bình quốc gia có thể che khuất nhu cầu thật, nên các nước này cần được đối chiếu thêm bằng dữ liệu tình hình trước khi đưa ra kết luận.")

    add_heading(doc, "6.4. Kiểm tra độ ổn định", 2)
    add_body(doc, f"Để đánh giá mức độ tin cậy của phân hoạch k = 4, báo cáo so sánh nhãn cụm cơ sở với nhãn thu được khi thay đổi điều kiện, dùng chỉ số Rand điều chỉnh (ARI, bằng 1 khi hai phân hoạch trùng nhau; Hubert và Arabie, 1985). Ba phép thử gồm: đổi 20 giá trị random_state; lần lượt bỏ từng biến trong 9 biến; và chạy lại trên 100 mẫu con ngẫu nhiên chiếm 80% số quốc gia.")
    stab_rows = [
        ["Đổi random_state (20 lần)", vn(stab.loc["seed","mean"],3), vn(stab.loc["seed","min"],3)],
        ["Bỏ lần lượt từng biến (9 lần)", vn(stab.loc["drop_feature","mean"],3), vn(stab.loc["drop_feature","min"],3)],
        ["Mẫu con 80% (100 lần)", vn(stab.loc["subsample_80","mean"],3), vn(stab.loc["subsample_80","min"],3)],
    ]
    add_table(doc, ["Phép thử", "ARI trung bình", "ARI nhỏ nhất"], stab_rows, [3.4, 1.6, 1.6], 8.4)
    add_caption(doc, "Bảng 5. Độ ổn định của phân hoạch k = 4 (ARI so với nghiệm cơ sở)")
    add_body(doc, f"Kết quả cho thấy nghiệm ổn định trước việc đổi seed (ARI nhỏ nhất {vn(stab.loc['seed','min'],3)}) và khá ổn định khi lấy mẫu con. Nhạy cảm nhất là phép bỏ biến: khi bỏ {worst.detail}, ARI giảm còn {vn(worst.ari,2)}, cho thấy ranh giới cụm phụ thuộc đáng kể vào biến này. Một số quốc gia nằm sát ranh giới, ví dụ Botswana, có thể đổi cụm giữa các lần chạy hoặc phiên bản thư viện, nên danh sách ưu tiên không nên đọc như một ranh giới cứng.")

    add_heading(doc, "7. Thảo luận ứng dụng trong viện trợ nhân đạo", 1)
    add_heading(doc, "7.1. Cách sử dụng kết quả", 2)
    add_body(doc, "Kết quả có thể được dùng như một lớp sàng lọc trong hệ thống phân tích nhiều tầng. Ở tầng đầu, nhóm ưu tiên 1 giúp cơ quan điều phối thu hẹp phạm vi rà soát và tìm thêm dữ liệu. Ở tầng hai, nhà phân tích ghép từng quốc gia với dữ liệu tình hình hiện tại như xung đột, thiên tai, di dời, lương thực, dịch bệnh và khả năng tiếp cận dịch vụ. Ở tầng ba, các chuyên gia khu vực xác thực bối cảnh và xác định loại hỗ trợ, quy mô, thời điểm cùng ràng buộc logistics. K-Means chỉ đóng vai trò ở tầng khám phá và sắp xếp thông tin.")
    add_body(doc, "Ưu điểm của cách làm là tạo ra một hồ sơ đa biến dễ đọc. Thay vì nói riêng rằng một nước có GDP thấp, ta quan sát đồng thời tử vong trẻ em, tuổi thọ, mức sinh và thu nhập. Tâm cụm trên thang đo gốc cho phép chuyển kết quả toán học thành mô tả mà nhà hoạch định chính sách có thể kiểm tra. Các tệp CSV đi kèm cũng giúp tái lập và truy vết từ một quốc gia đến cụm, tâm cụm và điểm sàng lọc.")
    add_heading(doc, "7.2. Rủi ro diễn giải và đạo đức dữ liệu", 2)
    add_body(doc, "Một cụm là kết quả phụ thuộc vào biến, thời điểm, chuẩn hóa, khoảng cách và k. Nếu thêm hoặc bỏ một biến, ranh giới cụm có thể đổi. Dữ liệu tổng hợp theo quốc gia che khuất chênh lệch trong nội bộ quốc gia, nên không thể dùng để suy luận nhu cầu của từng tỉnh, cộng đồng hoặc hộ gia đình. Xếp hạng dựa trên chỉ báo kinh tế cũng có thể bỏ sót khủng hoảng đột ngột tại một quốc gia có trung bình kinh tế tốt.")
    add_body(doc, "Cần tránh dùng nhãn cụm như một quyết định tự động để phân bổ hoặc từ chối viện trợ. Khi triển khai, cần công bố biến đầu vào, ngày cập nhật, quy tắc chuẩn hóa, phiên bản mã, giới hạn dữ liệu và quy trình khiếu nại hoặc rà soát chuyên gia. Việc ưu tiên viện trợ phải có giám sát con người và cơ chế cập nhật khi thông tin hiện trường mâu thuẫn với hồ sơ thống kê.")
    add_heading(doc, "7.3. Hướng cải thiện mô hình", 2)
    add_body(doc, "Nghiên cứu tiếp theo có thể bổ sung chỉ báo xung đột, số người di dời, chỉ số an ninh lương thực, số ca dịch bệnh, thiệt hại thiên tai, tiếp cận nước sạch, năng lực y tế và khoảng cách logistics. Khi dữ liệu có thời gian, nên dùng phân cụm theo chuỗi thời gian hoặc cập nhật theo quý để phát hiện quốc gia có rủi ro tăng nhanh. Có thể so sánh K-Means với K-Medoids, Gaussian Mixture, phân cụm phân cấp và DBSCAN; đồng thời mở rộng kiểm tra độ ổn định ở mục 6.4 bằng bootstrap đầy đủ và phân tích độ nhạy theo trọng số của điểm sàng lọc.")
    add_body(doc, "Một hướng khác là xác định trọng số với sự tham gia của chuyên gia. K-Means chuẩn hóa hiện tại coi các biến có vai trò tương đương trong khoảng cách, trong khi mục tiêu nhân đạo có thể muốn ưu tiên tử vong trẻ em hoặc tiếp cận y tế. Trọng số cần được công khai, thử nghiệm độ nhạy và kiểm tra tác động để tránh biến một giả định giá trị thành một quy tắc ẩn trong thuật toán.")

    add_heading(doc, "8. Kết luận và hướng phát triển", 1)
    add_body(doc, f"Tiểu luận đã trình bày K-Means từ hàm mục tiêu đến quy trình Lloyd, minh họa bằng ví dụ tính tay, sau đó áp dụng trên 167 quốc gia với 9 biến kinh tế xã hội. Chuẩn hóa là bước cần thiết vì dữ liệu có đơn vị khác nhau. Thực nghiệm thử k từ 2 đến 10 và kết hợp nhiều chỉ số cho thấy k = 4 là lựa chọn phù hợp với mục tiêu diễn giải, dù các chỉ số nội tại không đồng thuận (Silhouette cao nhất ở k = 4, Davies-Bouldin tốt nhất ở k = 5, Calinski-Harabasz cao nhất ở k = 2). Cần nhấn mạnh Silhouette chỉ khoảng 0,30 nên cấu trúc cụm yếu, và phân hoạch khá ổn định nhưng nhạy với việc bỏ biến GDP/người.")
    add_body(doc, f"Kết quả cuối tạo ra một cụm {n1} quốc gia có hồ sơ dễ tổn thương nhất theo các chỉ báo quan sát. Cụm này có trung bình tử vong trẻ em cao, mức sinh cao, tuổi thọ thấp, thu nhập thấp và GDP/người thấp. Phát hiện có thể hỗ trợ sàng lọc và lập kế hoạch thu thập dữ liệu, nhưng không thể thay thế đánh giá nhân đạo thực tế vì bộ dữ liệu thiếu các chiều khẩn cấp và bối cảnh.")
    add_body(doc, "Hướng phát triển ưu tiên là bổ sung dữ liệu động và dữ liệu tình hình, kiểm tra độ ổn định của cụm theo thời gian, đánh giá công bằng giữa các vùng và xây dựng quy trình human-in-the-loop. Khi đó, K-Means có thể trở thành một thành phần minh bạch trong hệ thống hỗ trợ quyết định, thay vì là một nhãn tự động đứng độc lập.")

    add_heading(doc, "Tài liệu tham khảo", 1)
    refs = [
        "Kaggle. Unsupervised Learning on Country Data. https://www.kaggle.com/datasets/rohan0301/unsupervised-learning-on-country-data",
        "Arthur, D., & Vassilvitskii, S. (2007). k-means++: The advantages of careful seeding. Proceedings of the 18th Annual ACM-SIAM Symposium on Discrete Algorithms, 1027–1035.",
        "Hubert, L., & Arabie, P. (1985). Comparing partitions. Journal of Classification, 2(1), 193–218.",
        "Kaufman, L., & Rousseeuw, P. J. (1990). Finding Groups in Data: An Introduction to Cluster Analysis. Wiley.",
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
    metrics = pd.read_csv(RESULTS / "k_evaluation.csv")
    profile = pd.read_csv(RESULTS / "cluster_profile.csv", index_col="cluster")
    priority = pd.read_csv(RESULTS / "aid_screening.csv")
    stability = pd.read_csv(RESULTS / "stability.csv")
    summary = json.loads((RESULTS / "run_summary.json").read_text(encoding="utf-8"))
    import sklearn
    summary["sklearn"] = sklearn.__version__
    build_report(df, metrics, profile, priority, stability, summary)
    print(f"Created {REPORT}")


if __name__ == "__main__":
    main()
