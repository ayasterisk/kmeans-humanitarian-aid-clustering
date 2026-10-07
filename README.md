# K-Means phân cụm quốc gia cần viện trợ nhân đạo

## Chạy thực nghiệm

```bash
pip install -r requirements.txt
python kmeans_humanitarian.py --data Country-data.csv --k 4   # sinh results/ và figures/                                    # dựng lại file .docx từ results/ và figures/
```

`kmeans_humanitarian.py` là nguồn duy nhất của mọi số liệu và hình trong báo cáo. 

Mã dùng 9 biến số trong dữ liệu, loại cột `country` khỏi ma trận đặc trưng, chuẩn hóa Z-score, thử `k=2..10`, huấn luyện K-Means với `n_init=50` và `random_state=42`, rồi chạy kiểm tra độ ổn định (đổi seed, bỏ từng biến, mẫu con 80%) ghi vào `results/stability.csv`.

Các hình: `01_elbow.png`, `02_silhouette.png`, `03_pca_clusters.png`, `04_profile_heatmap.png`.

## Tái lập

Kết quả đã commit được tạo với scikit-learn 1.8.0 (đã ghim trong `requirements.txt`). Quốc gia nằm sát ranh giới (ví dụ Botswana) có thể đổi cụm giữa các phiên bản thư viện khác nhau.

## Diễn giải

K-Means là phương pháp không giám sát. Silhouette chỉ khoảng 0,30, tức cấu trúc cụm yếu. Cụm có mức tử vong trẻ em cao, thu nhập và GDP bình quân đầu người thấp, tuổi thọ thấp và mức sinh cao được dùng như một nhóm ưu tiên sàng lọc về mặt kinh tế xã hội. Đây không phải quyết định phân bổ viện trợ cuối cùng: dữ liệu không có xung đột, thiên tai, di dời, an ninh lương thực, dịch bệnh hay khả năng tiếp cận dịch vụ, không ghi rõ năm tham chiếu và thiếu một số quốc gia khủng hoảng (như Somalia, Syria, Nam Sudan).

File `results/aid_screening.csv` có thêm điểm sàng lọc với trọng số heuristic (1 cho biến sức khỏe/mức sống, 0,5 cho lạm phát và chi y tế). Điểm này là chỉ báo hỗ trợ diễn giải, không phải nhãn thực tế và không thay thế đánh giá nhu cầu tại hiện trường.
