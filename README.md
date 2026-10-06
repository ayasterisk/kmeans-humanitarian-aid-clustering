# K-Means phân cụm quốc gia cần viện trợ nhân đạo

## Chạy thực nghiệm

```bash
pip install -r requirements.txt
python kmeans_humanitarian.py --data Country-data.csv --k 4
```

Kết quả được ghi vào `results/`, biểu đồ vào `figures/`. Mã dùng 9 biến số trong dữ liệu, loại cột `country` khỏi ma trận đặc trưng, chuẩn hóa Z-score, thử `k=2..10`, rồi huấn luyện K-Means với `n_init=50` và `random_state=42`.

## Diễn giải cẩn trọng

K-Means là phương pháp không giám sát. Cụm có mức tử vong trẻ em cao, thu nhập và GDP bình quân đầu người thấp, tuổi thọ thấp và mức sinh cao được dùng như một nhóm ưu tiên sàng lọc về mặt kinh tế xã hội. Đây không phải là quyết định phân bổ viện trợ cuối cùng: dữ liệu không có xung đột, thiên tai, di dời, an ninh lương thực, dịch bệnh hay khả năng tiếp cận dịch vụ.

File `results/aid_screening.csv` có thêm điểm sàng lọc minh bạch để xếp thứ tự trong dữ liệu. Điểm này là chỉ báo hỗ trợ diễn giải, không phải nhãn thực tế và không thay thế đánh giá nhu cầu tại hiện trường.
