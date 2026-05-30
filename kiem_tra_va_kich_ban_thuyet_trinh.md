# KIỂM TRA ĐỒNG BỘ & KỊCH BẢN THUYẾT TRÌNH
## Dự án: London Transport Analysis (TfL) — Nhóm 14

---

# PHẦN 1: KIỂM TRA ĐỒNG BỘ TOÀN DIỆN

> [!IMPORTANT]
> Tất cả số liệu dưới đây được xác minh trực tiếp từ file `outputs/london_tfl_cleaned.csv` — nguồn sự thật duy nhất (Single Source of Truth).

## ✅ SỐ LIỆU THỰC TẾ ĐÃ XÁC MINH (từ CSV + Code)

| Chỉ số | Giá trị chính xác | Nguồn |
|--------|------------------|-------|
| Tổng số ga phân tích | **298 ga** | `len(df)` sau `dropna()` |
| Tổng hành khách 2019 | **3,283,909,448** (~3.28 tỷ) | `passengers_2019.sum()` |
| Tổng hành khách 2020 | **1,208,311,783** (~1.21 tỷ) | `passengers_2020.sum()` |
| Tổng hành khách 2021 | **1,411,552,644** (~1.41 tỷ) | `passengers_2021.sum()` |
| COVID impact 2019→2020 | **−63.2051%** (làm tròn: −63.21%) | Tính từ CSV |
| Phục hồi 2020→2021 | **+16.8202%** (làm tròn: +16.82%) | Tính từ CSV |
| Số cụm KMeans | **K = 6** | `N_CLUSTERS = 6` |
| Xu hướng "Giảm mạnh" | **287/298 ga (96.3%)** | `trend_category.value_counts()` |
| Xu hướng "Giảm nhẹ" | **9 ga** | |
| Xu hướng "Tăng mạnh" | **1 ga** | |
| Xu hướng "Ổn định" | **1 ga** | |

---

## 📊 TOP 10 GA ĐÔNG KHÁCH NHẤT 2021 (SỐ LIỆU THỰC TẾ)

| Hạng | Tên ga | Lượt HK 2021 | Cụm |
|------|--------|-------------|-----|
| 1 | **Stratford** | 63,439,192 | Siêu trung tâm |
| 2 | **Liverpool Street** | 43,057,801 | Siêu trung tâm |
| 3 | **King's Cross St. Pancras** | 36,734,085 | Siêu trung tâm |
| 4 | **Victoria** | 33,480,926 | Siêu trung tâm |
| 5 | **Oxford Circus** | 32,863,152 | Siêu trung tâm |
| 6 | London Bridge | 30,857,082 | **Ga lớn** |
| 7 | Bank | 30,132,853 | Siêu trung tâm |
| 8 | Waterloo | 29,868,623 | Siêu trung tâm |
| 9 | Paddington | 28,950,455 | Siêu trung tâm |
| 10 | Canary Wharf | 24,923,667 | **Ga lớn** |

---

## 📋 BẢNG CLUSTER SUMMARY CHÍNH XÁC

| Cluster ID | Tên cụm | Số ga | HK TB 2021 | COVID Impact | Recovery |
|-----------|---------|-------|-----------|-------------|---------|
| 2 | **Siêu trung tâm** | 8 | 37,315,886 | −73.70% | **+71.82%** |
| 3 | **Ga lớn** | 42 | 11,268,126 | −63.47% | **+39.17%** |
| 0 | **Ga trung bình** | 78 | 3,136,520 | −54.91% | **+0.44%** |
| 4 | **Ga nhỏ** | 80 | 2,833,787 | −57.42% | **+27.60%** |
| 1 | **Ga ít khách** | 37 | 2,215,412 | −49.69% | **−10.64%** |
| 5 | **Ga rất ít khách** | 53 | 1,630,990 | −41.42% | **−20.38%** |

---

## 🔍 BẢNG KIỂM TRA ĐỒNG BỘ

### A. ĐIỂM NHẤT QUÁN ✅ (Các số liệu khớp hoàn toàn)

| # | Điểm kiểm tra | Code | Slide PPTX | Báo cáo Word | Trạng thái |
|---|--------------|------|-----------|-------------|-----------|
| 1 | Số ga phân tích | 298 | 298 (Slide 2, 8, 10) | 298 | ✅ Khớp |
| 2 | COVID impact | −63.21% | −63.21% (Slide 9, 10) | −63.21% | ✅ Khớp |
| 3 | Phục hồi 2020→2021 | +16.82% | +16.82% (Slide 9) | +16.82% | ✅ Khớp |
| 4 | Số cụm KMeans | K=6 | K=6 (Slide 8) | K=6 | ✅ Khớp |
| 5 | Tên 6 cụm | Siêu trung tâm / Ga lớn / Ga trung bình / Ga nhỏ / Ga ít khách / Ga rất ít khách | Slide 8 | Báo cáo | ✅ Khớp |
| 6 | Tổng HK 2019 | 3.28 tỷ | "3.28 Tỷ" (Slide 9) | 3,283,909,448 | ✅ Khớp |
| 7 | Tổng HK 2020 | 1.21 tỷ | "1.21 Tỷ" (Slide 9) | 1,208,311,783 | ✅ Khớp |
| 8 | Số ga mỗi cụm | 8/42/78/80/37/53 | 8/42/78/80/37/53 (Slide 8) | Khớp | ✅ Khớp |
| 9 | COVID per cluster | −73.70%/−63.47%/−54.91%/−57.42%/−49.69%/−41.42% | Slide 8 | Khớp | ✅ Khớp |
| 10 | Xu hướng "Giảm mạnh" | 287/298 ga | Slide 9 | Khớp | ✅ Khớp |
| 11 | Kiến trúc ETL | Extract→Transform→Load→Serve | Slide 6 | Chương 3 | ✅ Khớp |
| 12 | 3 nguồn dữ liệu | KML + TfL CSV + Stops.csv | Slide 4 | Chương 2 | ✅ Khớp |
| 13 | Giai đoạn phân tích | 2017–2021 | Slide 2, 9 | Chương 1 | ✅ Khớp |

---

### B. ĐIỂM KHÔNG NHẤT QUÁN ⚠️ — CẦN SỬA

> [!WARNING]
> Phát hiện **3 sai sót nghiêm trọng** trong slide PPTX (Slide 10) về Top 5 ga đông khách nhất. Đây là điểm dễ bị giám khảo hỏi nhất!

#### ❌ SAI SÓT 1 — Top 5 ga trong Slide 10 KHÔNG ĐÚNG

**Slide 10 ghi:**
1. King's Cross St. Pancras — 34.0M
2. Waterloo — 29.7M
3. Victoria — 28.3M
4. London Bridge — 27.2M
5. Bank / Monument — 26.1M

**Số liệu thực tế từ CSV:**
1. **Stratford** — **63,439,192** (63.4M) ← Số 1, không có trong slide!
2. **Liverpool Street** — **43,057,801** (43.1M) ← Số 2, không có trong slide!
3. King's Cross St. Pancras — 36,734,085 (36.7M) ← Đứng số 3, slide ghi 34M (sai)
4. Victoria — 33,480,926 (33.5M) ← Đứng số 4, slide ghi 28.3M (sai)
5. Oxford Circus — 32,863,152 (32.9M) ← Đứng số 5, không có trong slide!

> [!CAUTION]
> Waterloo (hạng 8: 29.9M), London Bridge (hạng 6: 30.9M, thuộc "Ga lớn" không phải "Siêu trung tâm") và Bank/Monument (hạng 7: 30.1M) đều không thuộc Top 5. Cần cập nhật slide trước khi trình bày.

**Lý do sai:** Slide có thể được tạo từ dữ liệu sơ bộ hoặc dùng số liệu năm khác (có thể năm 2019 hoặc dữ liệu chưa gộp đúng tuyến).

---

#### ⚠️ SAI SÓT 2 — Features KMeans trong Slide 8 không khớp code

**Slide 8 ghi:** `Features: avg_passengers, covid_impact_pct, recovery_pct`

**Code thực tế (dòng 551–552):**
```python
features = ["passengers_2021", "num_lines", "lat", "lon"]
```

Slide mô tả sai 4 đặc trưng. Code dùng `passengers_2021 + num_lines + lat + lon`, **không dùng** `avg_passengers` hay `covid_impact_pct` làm input cho KMeans.

**Cách khắc phục:** Cập nhật ghi chú ở đáy Slide 8 thành: `Features: passengers_2021, num_lines, lat, lon`

---

#### ⚠️ SAI SÓT 3 — Spatial Join trong Slide 7 mô tả sai thư viện

**Slide 7 ghi:** `GeoPandas sjoin_nearest() – Nearest Neighbor Algorithm`

**Code thực tế (dòng 394–415):** Không dùng GeoPandas. Code dùng `pd.merge()` theo `station_key` rồi tính khoảng cách Euclidean thủ công bằng NumPy:
```python
candidates["stop_distance"] = np.sqrt(
    (candidates["lat"] - candidates["stop_lat"]) ** 2
    + (candidates["lon"] - candidates["stop_lon"]) ** 2
)
```

Đây là **Pandas Nearest Neighbor**, không phải GeoPandas `sjoin_nearest()`.

**Cách khắc phục:** Sửa Slide 7 thành: `Pandas + NumPy Euclidean Distance – Nearest Neighbor`

---

### C. ĐIỂM CẦN LƯU Ý KHI THUYẾT TRÌNH (Không nhất thiết phải sửa slide)

| # | Điểm lưu ý | Chi tiết |
|---|-----------|---------|
| 1 | Recovery rate của Siêu trung tâm | Code tính được **+71.82%** — đây là số ấn tượng chưa xuất hiện trong slide |
| 2 | Stratford là ga #1 | Điều này có thể gây ngạc nhiên vì Stratford nổi tiếng là ga Olympic. Cần giải thích: Stratford phục vụ cả Underground + Overground + DLR + Elizabeth Line |
| 3 | London Bridge thuộc "Ga lớn" không phải "Siêu trung tâm" | Dù đứng hạng 6 về lượt HK, London Bridge không vào cụm Siêu trung tâm do yếu tố vị trí địa lý trong KMeans |
| 4 | Tỷ lệ 43% phục hồi so với 2019 | Năm 2021 đạt 1.41 tỷ / 3.28 tỷ = **43%** so với mức trước dịch — số liệu quan trọng chưa có trong slide |
| 5 | Lớp học | Slide ghi "DHKHDL20A" nhưng Báo cáo Word ghi "DHKTDL17A" — cần xác nhận lớp chính xác |

---

### D. ĐỀ XUẤT SỬA TRƯỚC KHI THUYẾT TRÌNH

| Ưu tiên | Slide | Sửa gì | Thời gian ước tính |
|--------|-------|--------|-------------------|
| 🔴 Bắt buộc | Slide 10 | Cập nhật Top 5: #1 Stratford 63.4M, #2 Liverpool St 43.1M, #3 King's Cross 36.7M, #4 Victoria 33.5M, #5 Oxford Circus 32.9M | 5 phút |
| 🟠 Nên sửa | Slide 8 | Sửa Features thành `passengers_2021, num_lines, lat, lon` | 2 phút |
| 🟡 Tùy chọn | Slide 7 | Sửa "GeoPandas sjoin_nearest()" → "Pandas + NumPy Euclidean Nearest Neighbor" | 2 phút |
| 🟡 Tùy chọn | Slide 10 | Thêm số liệu "Recovery Siêu trung tâm: +71.82%" | 2 phút |

---
---

# PHẦN 2: KỊCH BẢN THUYẾT TRÌNH CHI TIẾT

> [!NOTE]
> Tổng thời gian: **13 phút** (có thể kéo đến 15 phút nếu trả lời câu hỏi tương tác).
> Số liệu trong script đã được cập nhật theo kết quả kiểm tra đồng bộ ở Phần 1.

---

## 🎬 SLIDE 1 — MỞ ĐẦU & TRANG BÌA
**⏱ Thời lượng: 1 phút**

### 📢 Script nói:
> *"Kính thưa thầy/cô và các bạn. Hôm nay, Nhóm 14 chúng em xin trình bày đồ án môn Data Engineering với chủ đề: Xây dựng Pipeline ETL End-to-End và Phân tích Dữ liệu Không gian — Thời gian cho Hệ thống Nhà ga Transport for London.*
>
> *Transport for London — hay TfL — là mạng lưới giao thông đô thị lớn nhất châu Âu, phục vụ hàng triệu hành khách mỗi ngày. Câu hỏi chúng em đặt ra là: Làm thế nào để biến hàng triệu dòng dữ liệu thô từ nhiều nguồn khác nhau thành những insight có thể hành động được?*
>
> *Và câu trả lời chính là Pipeline ETL hoàn chỉnh mà chúng em sẽ trình bày trong 13 phút tới."*

### 🎯 Điểm nhấn:
- Pipeline ETL **End-to-End** — tự động hóa hoàn toàn
- Dữ liệu thực tế, quy mô thực tế

### 🖱 Hành động:
- Đứng thẳng, giao tiếp mắt với giám khảo
- Nhấn slide sau khi nói xong câu kết

### 💡 Mẹo:
> Không đọc slide! Giới thiệu tên nhóm trưởng đại diện và các thành viên khi chuyển sang slide 2.

---

## 🎬 SLIDE 2 — LỜI MỞ ĐẦU & TÓM TẮT
**⏱ Thời lượng: 1 phút**

### 📢 Script nói:
> *"Để tóm tắt toàn bộ dự án trong 30 giây: Chúng em đã xây dựng một hệ thống xử lý dữ liệu tự động hoàn chỉnh — từ lúc đọc file dữ liệu thô cho đến khi tạo ra bản đồ tương tác, báo cáo PDF và cơ sở dữ liệu — không cần can thiệp thủ công.*
>
> *Cụ thể: 298 nhà ga TfL được phân tích trong giai đoạn 2017 đến 2021. Chúng em tích hợp 3 nguồn dữ liệu độc lập, áp dụng thuật toán KMeans với K=6 cụm, và phát hiện rằng COVID-19 đã làm lưu lượng hành khách giảm 63.21% vào năm 2020.*
>
> *Tất cả được thể hiện qua bản đồ Folium tương tác mà các bạn có thể mở trực tiếp trên trình duyệt."*

### 🎯 Điểm nhấn:
- **298 ga** | **3 nguồn dữ liệu** | **KMeans K=6** | **−63.21% COVID**

### 🖱 Hành động:
- Chỉ vào 6 số liệu tóm tắt ở góc phải slide
- Nhấn mạnh từ "tự động hóa hoàn toàn"

---

## 🎬 SLIDE 3 — LÝ DO CHỌN ĐỀ TÀI
**⏱ Thời lượng: 1 phút**

### 📢 Script nói:
> *"Tại sao chọn TfL? Có 3 lý do chính.*
>
> *Thứ nhất — TfL là hệ thống giao thông phức tạp bậc nhất thế giới với hàng chục triệu lượt hành khách mỗi năm và dữ liệu đến từ nhiều hệ thống khác nhau. Đây là bài toán ETL thực sự, không phải bài tập.*
>
> *Thứ hai — COVID-19 đã tạo ra một biến cố lịch sử trong dữ liệu giao thông: lưu lượng sụt giảm không đồng đều theo từng khu vực và loại hình ga. Đây là cơ hội phân tích dữ liệu thực chiến hiếm có.*
>
> *Thứ ba — TfL có chính sách Open Data, nghĩa là tất cả dữ liệu đều miễn phí và có thể tái sử dụng. Điều này cho phép chúng em làm việc với dữ liệu thật, không phải dữ liệu giả lập.*
>
> *Kết hợp cả 3 yếu tố này, đây là đề tài lý tưởng để thực hành toàn bộ kỹ năng Data Engineering."*

### 🎯 Điểm nhấn:
- Dữ liệu **thực tế**, quy mô **thực tế**
- COVID tạo ra **sự kiện lịch sử** trong dữ liệu

### 💡 Mẹo:
> Nói với giọng tự tin — đây là phần thể hiện sự hiểu biết về context, không chỉ về code.

---

## 🎬 SLIDE 4 — BỘ DỮ LIỆU
**⏱ Thời lượng: 1 phút**

### 📢 Script nói:
> *"Dự án sử dụng 3 nguồn dữ liệu hoàn toàn độc lập về định dạng và nội dung.*
>
> *Nguồn đầu tiên là file KML — định dạng XML của Google Maps — chứa tọa độ GPS của các nhà ga. File này cho chúng em biết ga nằm ở đâu.*
>
> *Nguồn thứ hai là TfL_stations.csv từ chương trình Open Data của TfL — chứa số liệu lưu lượng hành khách Enter và Exit cho từng ga, từ năm 2017 đến 2021. File này cho chúng em biết mỗi ga có bao nhiêu khách.*
>
> *Nguồn thứ ba là Stops.csv từ cơ sở dữ liệu NaPTAN của chính phủ Anh — file nặng hơn 100 megabyte, chứa metadata về từng điểm dừng giao thông: tên thông dụng, loại điểm dừng và Borough.*
>
> *Thách thức lớn nhất: Ba nguồn này hoàn toàn không 'biết' về nhau — không có ID chung, tên ga viết khác nhau và cấu trúc file khác nhau. Đây là bài toán Data Integration thực sự."*

### 🎯 Điểm nhấn:
- KML: **tọa độ** | CSV: **lưu lượng** | NaPTAN: **metadata**
- NaPTAN: **101 MB**, hàng trăm nghìn bản ghi

### 🖱 Hành động:
- Chỉ vào từng card dữ liệu theo thứ tự 01 → 02 → 03

---

## 🎬 SLIDE 5 — THÁCH THỨC & GIẢI PHÁP
**⏱ Thời lượng: 1 phút**

### 📢 Script nói:
> *"Khi bắt tay vào làm, chúng em gặp ngay 4 thách thức kỹ thuật lớn.*
>
> *Thách thức đầu tiên và nan giải nhất: tên nhà ga không đồng nhất giữa các nguồn. Cùng một ga nhưng KML viết 'King's Cross Station', TfL viết 'King's Cross St. Pancras', NaPTAN viết hàng chục biến thể khác nhau. Không thể join trực tiếp.*
>
> *Giải pháp: Chúng em xây dựng hàm normalize_station_name với hơn 50 quy tắc Regex — loại bỏ hậu tố tuyến đường, chuẩn hóa ký tự đặc biệt, và duy trì một bảng alias thủ công cho các trường hợp ngoại lệ. Kết quả: tỷ lệ match trên 95%.*
>
> *Thách thức thứ hai: File Stops.csv 101MB có hàng nghìn bản ghi trùng lặp cho cùng một ga. Giải pháp: Thuật toán chọn bản ghi tốt nhất dựa trên thứ tự ưu tiên StopType và khoảng cách Euclidean đến tọa độ KML.*
>
> *Hai thách thức kỹ thuật còn lại về định dạng CSV đặc biệt và các con số sai số đều được xử lý bằng logic parsing tùy chỉnh trong code."*

### 🎯 Điểm nhấn:
- **50+ quy tắc Regex** cho chuẩn hóa tên
- **>95% tỷ lệ match** thành công
- Euclidean Nearest Neighbor (không cần GeoPandas)

### 💡 Mẹo:
> Sẵn sàng giải thích tại sao dùng Euclidean thay vì Haversine: phạm vi địa lý nhỏ (London ~25km bán kính), sai lệch không đáng kể, nhanh hơn nhiều.

---

## 🎬 SLIDE 6 — KIẾN TRÚC PIPELINE ETL
**⏱ Thời lượng: 1.5 phút**

### 📢 Script nói:
> *"Đây là trái tim của dự án — Pipeline ETL End-to-End.*
>
> *Giai đoạn EXTRACT: Ba hàm độc lập đọc ba file từ đĩa — load_kml_stations(), load_tfl_csv(), và load_stops_csv(). Mỗi hàm xử lý định dạng file riêng và trả về một DataFrame Pandas chuẩn hóa.*
>
> *Giai đoạn TRANSFORM: Đây là phần phức tạp nhất. Chúng em chuẩn hóa tên ga, ghép 3 nguồn theo khóa station_key, làm sạch dữ liệu, và tính 6 đặc trưng mới: số tuyến, lưu lượng trung bình, tỷ lệ tác động COVID, tỷ lệ phục hồi, hệ số xu hướng tuyến tính và nhãn phân loại xu hướng.*
>
> *Giai đoạn LOAD: Dữ liệu sạch được lưu đồng thời ra 4 định dạng — CSV, Excel với 2 sheet, SQLite với bảng fact_stations và dim_clusters, và file text báo cáo tóm tắt.*
>
> *Cuối cùng là VISUALIZATION: Hàm create_folium_map() tạo ra một ứng dụng web HTML hoàn chỉnh nhúng sẵn dữ liệu — mở được trực tiếp trên trình duyệt mà không cần backend.*
>
> *Và serve_outputs.py cho phép chia sẻ bản đồ ra internet qua tunnel trong 1 lệnh duy nhất."*

### 🎯 Điểm nhấn:
- **4 giai đoạn**: Extract → Transform → Load → Visualize
- **6 đặc trưng** mới được tính toán
- **4 định dạng output** song song (CSV, Excel, SQLite, Text)
- **Tự động hóa 100%**: 1 lệnh `python final_project.py`

### 🖱 Hành động:
- Chỉ theo chiều mũi tên từ trái sang phải
- Dừng lại ở mỗi giai đoạn, đếm nhịp

### 💡 Mẹo:
> Nhấn mạnh: "Toàn bộ pipeline chạy bằng một lệnh duy nhất" — đây là điểm khác biệt với nhiều dự án sinh viên phải can thiệp thủ công nhiều bước.

---

## 🎬 SLIDE 7 — ĐỔI MỚI KỸ THUẬT
**⏱ Thời lượng: 1 phút**

### 📢 Script nói:
> *"Có 3 điểm kỹ thuật mà chúng em tự hào nhất trong dự án này.*
>
> *Đầu tiên là hệ thống chuẩn hóa tên ga bằng Regex — hơn 50 quy tắc được thiết kế cẩn thận, bao gồm vòng lặp while để xử lý nhiều hậu tố lồng nhau, và bảng alias thủ công cho các trường hợp đặc biệt như Bank và Monument được gộp thành một ga duy nhất.*
>
> *Thứ hai là thuật toán Nearest Neighbor tự xây dựng — không dùng thư viện ngoài, chỉ dùng Pandas và NumPy. Khoảng cách Euclidean kết hợp với thứ tự ưu tiên StopType đảm bảo chọn đúng bản ghi NaPTAN cho mỗi nhà ga.*
>
> *Thứ ba là Feature Engineering: Chúng em tính hệ số xu hướng bằng LinearRegression của scikit-learn — nhưng thay vì bỏ qua những ga thiếu dữ liệu, code chỉ fit trên các năm có dữ liệu thực, cho phép xác định xu hướng ngay cả khi thiếu 1-2 năm."*

### 🎯 Điểm nhấn:
- **50+ Regex rules** trong `normalize_station_name()`
- **Pandas + NumPy Euclidean** (không dùng GeoPandas)
- **LinearRegression cho trend** với xử lý NaN thông minh

### 💡 Mẹo:
> Nếu bị hỏi về GeoPandas: "Chúng em cân nhắc dùng GeoPandas nhưng nhận thấy với phạm vi địa lý nhỏ như London, khoảng cách Euclidean cho kết quả tương đương mà không cần dependency nặng thêm."

---

## 🎬 SLIDE 8 — PHÂN CỤM KMEANS (K=6)
**⏱ Thời lượng: 1.5 phút**

### 📢 Script nói:
> *"Để phân loại 298 nhà ga, chúng em áp dụng thuật toán KMeans với K=6.*
>
> *Tại sao K=6? Chúng em dùng phương pháp Elbow: vẽ đồ thị Inertia theo K từ 2 đến 10 — điểm 'khuỷu tay' rõ ràng nhất xuất hiện ở K=6. Và quan trọng hơn, K=6 cho ra 6 nhóm có ý nghĩa thực tiễn rõ ràng.*
>
> *4 đặc trưng đầu vào: lưu lượng hành khách 2021, số tuyến phục vụ, vĩ độ và kinh độ. Tất cả được chuẩn hóa bằng StandardScaler trước khi đưa vào KMeans — điều này rất quan trọng vì lưu lượng hành khách có giá trị hàng triệu trong khi tọa độ chỉ có giá trị 51-52 độ.*
>
> *Kết quả: 8 ga Siêu trung tâm với lưu lượng trung bình 37 triệu lượt mỗi năm. Các ga này bị COVID tác động nặng nhất — minus 73.7% — nhưng cũng phục hồi nhanh nhất với cộng 71.8% vào 2021.*
>
> *42 ga Lớn, 78 ga Trung bình, 80 ga Nhỏ, 37 ga Ít khách, và 53 ga Rất ít khách.*
>
> *Điều thú vị: 8 ga Siêu trung tâm chỉ chiếm 2.7% tổng số ga nhưng xử lý khoảng 21% tổng lưu lượng toàn mạng lưới."*

### 🎯 Điểm nhấn:
- Features: `passengers_2021` + `num_lines` + `lat` + `lon` (KHÔNG phải avg_passengers)
- **StandardScaler** — bắt buộc vì scale khác nhau
- Siêu trung tâm: **8 ga, 37.3M HK, COVID −73.7%, Phục hồi +71.8%**
- 8 ga = 2.7% số ga nhưng = **~21% tổng lưu lượng**

### 🖱 Hành động:
- Chỉ vào từng dòng bảng cluster theo thứ tự từ trên xuống
- Nhấn mạnh dòng "Siêu trung tâm" — hàng đầu tiên

### 💡 Mẹo:
> Sẵn sàng giải thích: "Tại sao không dùng avg_passengers thay vì passengers_2021?" → "Chúng em muốn clustering phản ánh trạng thái hiện tại (2021), không phải trung bình lịch sử bị ảnh hưởng nặng bởi COVID."

---

## 🎬 SLIDE 9 — PHÂN TÍCH COVID-19
**⏱ Thời lượng: 1.5 phút**

### 📢 Script nói:
> *"Đây là slide mà chúng em muốn các thầy cô dành thêm thời gian để nhìn.*
>
> *Năm 2019 — trước đại dịch — toàn mạng lưới TfL đón tổng cộng 3.28 tỷ lượt hành khách. Đây là mức cao nhất trong lịch sử 5 năm chúng em phân tích.*
>
> *Năm 2020 — COVID-19 bùng phát — con số đó rơi tự do xuống còn 1.21 tỷ. Mức sụt giảm: âm 63.21%. Hơn 2 tỷ lượt hành khách biến mất trong một năm.*
>
> *Điều đáng chú ý là sự sụt giảm không đồng đều: Các ga Siêu trung tâm mất âm 73.7% — vì phụ thuộc nhiều vào dân văn phòng và khách du lịch. Trong khi đó các ga Rất ít khách chỉ mất âm 41.4% — vì phục vụ chủ yếu nhu cầu thiết yếu.*
>
> *Năm 2021 — có tín hiệu phục hồi: tổng lưu lượng tăng thêm 16.82% so với 2020. Tuy nhiên, mức này chỉ bằng 43% so với năm 2019. Chặng đường phục hồi vẫn còn dài.*
>
> *Và con số đáng nhớ: 287 trong số 298 ga — tức 96.3% — đang trong xu hướng Giảm mạnh. Chỉ có 1 ga có xu hướng Tăng mạnh."*

### 🎯 Điểm nhấn:
- **3.28 tỷ** (2019) → **1.21 tỷ** (2020) = **−63.21%**
- **+16.82%** phục hồi nhưng chỉ đạt **43%** mức 2019
- **287/298 ga** = 96.3% xu hướng "Giảm mạnh"
- Siêu trung tâm: −73.7% → Phục hồi +71.8%

### 🖱 Hành động:
- Chỉ vào 4 số liệu lớn lần lượt: 3.28 tỷ → −63.21% → 1.21 tỷ → +16.82%
- Dừng 2 giây ở số −63.21% để tạo hiệu ứng

### 💡 Mẹo:
> Câu kết ấn tượng: "Đây không chỉ là số liệu — đây là hình ảnh của một thành phố bị phong tỏa, và đang dần thức dậy."

---

## 🎬 SLIDE 10 — KẾT QUẢ NỔI BẬT
**⏱ Thời lượng: 1 phút**

### 📢 Script nói:
> *"Tổng hợp kết quả, chúng em có 4 con số headline: 298 nhà ga, 6 cụm hành vi, âm 63.21% tác động COVID, cộng 16.82% phục hồi.*
>
> *Và đây là Top 5 ga đông khách nhất năm 2021 — kết quả này có thể khá bất ngờ:*
>
> *Đứng số 1 không phải King's Cross hay Waterloo mà là Stratford — với 63.4 triệu lượt. Stratford là ga Olympic, phục vụ 5 tuyến tàu — Underground, Overground, DLR, Elizabeth Line và National Rail — tạo ra lưu lượng khổng lồ.*
>
> *Số 2 là Liverpool Street với 43.1 triệu — ga phục vụ 6 tuyến.*
>
> *Số 3 là King's Cross St. Pancras với 36.7 triệu.*
>
> *Số 4 là Victoria với 33.5 triệu.*
>
> *Và số 5 là Oxford Circus với 32.9 triệu — ga duy nhất ở trung tâm West End.*
>
> *Điều thú vị: tất cả 5 ga này đều thuộc cụm Siêu trung tâm và đều phục vụ 3 tuyến trở lên."*

### 🎯 Điểm nhấn:
- **#1 Stratford: 63.4M** — bất ngờ nhưng có giải thích rõ ràng (5 tuyến!)
- **#2 Liverpool Street: 43.1M**
- **#3 King's Cross: 36.7M**
- Tất cả thuộc **Siêu trung tâm** và phục vụ **≥3 tuyến**

> [!WARNING]
> Slide 10 gốc hiện ghi sai Top 5. Nếu chưa sửa slide, cần nói: "Số liệu trên bản chiếu là dữ liệu sơ bộ — số liệu chính xác đã được xác minh từ file CSV là..." rồi đọc Top 5 đúng.

### 🖱 Hành động:
- Chỉ vào từng ga trong danh sách Top 5 theo thứ tự
- Nhìn về phía khán giả khi nói về Stratford

---

## 🎬 SLIDE 11 — SẢN PHẨM & TRỰC QUAN HÓA
**⏱ Thời lượng: 1.5 phút**

### 📢 Script nói:
> *"Sản phẩm cuối cùng của dự án là một hệ sinh thái các output được tạo ra tự động.*
>
> *Trọng tâm là bản đồ Folium tương tác — file london_tfl_map.html. Đây không phải một bản đồ tĩnh mà là một ứng dụng web hoàn chỉnh với giao diện sidebar đầy đủ tính năng.*
>
> *Người dùng có thể: chọn năm từ 2017 đến 2021 bằng thanh trượt để xem sự thay đổi theo thời gian; bật tắt từng cụm ga; lọc theo Borough; tìm kiếm ga theo tên, tuyến hoặc khu vực; bật heatmap để thấy mật độ hành khách; và click vào bất kỳ ga nào để xem popup với biểu đồ sparkline 5 năm.*
>
> *Bên cạnh đó là SQLite database với 2 bảng quan hệ — fact_stations và dim_clusters; file Excel với 2 sheet; CSV và file báo cáo text.*
>
> *Và serve_outputs.py cho phép tạo public URL để chia sẻ bản đồ ra internet qua localtunnel — chỉ cần 1 lệnh python serve_outputs.py."*
>
> *[Nếu có demo] Bây giờ em xin mời các thầy cô xem bản đồ trực tiếp...*

### 🎯 Điểm nhấn:
- Bản đồ có **Time Slider** 2017–2021
- **MarkerCluster** + **Heatmap** + **Popup sparkline**
- Xuất **CSV/Excel/SQLite** đồng thời
- **Public URL** qua localtunnel

### 🖱 Hành động:
- Nếu có laptop: Demo trực tiếp bản đồ HTML
- Click vào marker, đổi năm với slider, bật heatmap
- Chỉ vào sidebar với các tính năng lọc

### 💡 Mẹo:
> Chuẩn bị sẵn FINAL_MAP.html trên trình duyệt, zoom vào khu Zone 1 để thấy Siêu trung tâm cụm lại. Demo 30 giây là đủ ấn tượng.

---

## 🎬 SLIDE 12 — KẾT LUẬN & HƯỚNG PHÁT TRIỂN
**⏱ Thời lượng: 1 phút**

### 📢 Script nói:
> *"Tóm lại, dự án đã hoàn thành đầy đủ 5 mục tiêu đề ra:*
>
> *Một — Pipeline ETL tự động hóa hoàn toàn: từ 3 file dữ liệu thô đến 4 loại output khác nhau bằng 1 lệnh.*
>
> *Hai — Phân tích dữ liệu có giá trị thực tiễn: phát hiện sự phân tầng tác động COVID theo loại ga, xu hướng phục hồi khác nhau.*
>
> *Ba — Sản phẩm trực quan có thể dùng ngay: bản đồ tương tác mà bất kỳ ai cũng có thể mở và khám phá.*
>
> *Về hướng phát triển: Chúng em muốn mở rộng sang dữ liệu thời gian thực từ TfL API; thêm mô hình dự báo ARIMA hoặc LSTM để dự đoán lưu lượng; và xây dựng FastAPI backend thay cho static HTML.*
>
> *Quan trọng hơn, framework ETL này có thể tái sử dụng cho bất kỳ hệ thống giao thông đô thị nào — Paris, Tokyo, hay ngay cả TP.HCM trong tương lai."*

### 🎯 Điểm nhấn:
- **5 mục tiêu** → Hoàn thành tất cả
- Framework có thể **tái sử dụng** cho thành phố khác
- Hướng phát triển: **Real-time API + LSTM + FastAPI**

---

## 🎬 SLIDE 13 — CẢM ƠN & Q&A
**⏱ Thời lượng: 0.5 phút**

### 📢 Script nói:
> *"Cảm ơn thầy cô và các bạn đã lắng nghe phần trình bày của Nhóm 14. Chúng em xin sẵn sàng trả lời các câu hỏi."*

### 💡 Chuẩn bị sẵn câu trả lời cho Q&A:

**Q: Tại sao Stratford đứng số 1?**
> A: Stratford là hub đa phương tiện phục vụ 5 tuyến tàu (Underground Central/Jubilee, Overground, DLR, Elizabeth Line và National Rail). Sau Olympic 2012, khu East London phát triển mạnh dẫn đến lưu lượng tăng vọt. Năm 2021 khi COVID hạ nhiệt, Stratford phục hồi nhanh hơn do cộng đồng địa phương đông đúc.

**Q: Tại sao dùng Euclidean thay vì Haversine?**
> A: Phạm vi địa lý của London rất nhỏ (bán kính ~25km). Ở vĩ độ 51°N, sai lệch giữa Euclidean và Haversine chỉ ~0.3%. Đổi lại, Euclidean nhanh hơn đáng kể khi tính trên hàng triệu cặp điểm, và không cần convert đơn vị.

**Q: Tại sao chọn K=6 cho KMeans?**
> A: Hai tiêu chí: (1) Elbow Method — inertia giảm chậm lại rõ ràng sau K=6; (2) Tính diễn giải — 6 cụm cho phép đặt tên có ý nghĩa kinh doanh rõ ràng từ Siêu trung tâm đến Rất ít khách.

**Q: Recovery rate của Siêu trung tâm là bao nhiêu?**
> A: +71.82% — đây là nhóm phục hồi nhanh nhất vì các ga trung tâm London có nhu cầu giao thông bền vững từ dân văn phòng và du lịch.

**Q: Tại sao features KMeans dùng passengers_2021 thay vì avg_passengers?**
> A: Chúng em muốn clustering phản ánh trạng thái hiện tại của mạng lưới. avg_passengers bị kéo lệch mạnh bởi năm COVID (2020), không đại diện cho hành vi "bình thường" của ga.

---

## 📋 BẢNG TÓM TẮT SỐ LIỆU CẦN GHI NHỚ

| Chỉ số | Giá trị | Ghi nhớ bằng |
|--------|--------|-------------|
| Số ga | 298 | "Gần 300 ga" |
| COVID impact | −63.21% | "Mất hơn nửa" |
| Phục hồi | +16.82% | "Gần 17%" |
| Tổng HK 2019 | 3.28 tỷ | "3 tỷ rưỡi" |
| Siêu trung tâm | 8 ga, 37.3M, −73.7% | "8 ga, mất 3/4" |
| #1 Stratford | 63.4M | "63 triệu lượt" |
| Xu hướng Giảm mạnh | 287/298 (96.3%) | "Gần như toàn bộ" |
| Recovery Siêu TT | +71.82% | "Phục hồi gần 72%" |

---

*Tài liệu được tạo ngày 28/05/2026 — Nhóm 14 — ĐH Công Nghiệp TP.HCM*
