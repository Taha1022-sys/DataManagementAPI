# Excel Data Management API

## Overview
Excel Data Management API is a robust .NET 9 Web API for uploading, reading, editing, comparing, and exporting Excel files. It is designed for enterprise-level data management, enabling users to interact with Excel data programmatically and securely.

## Features
- **Upload Excel Files:** Upload .xlsx/.xls files and store them securely.
- **Read Data:** Read data from specific sheets or all sheets in an Excel file.
- **Database Storage:** Persist Excel data in a SQL database for fast access and versioning.
- **Edit & Update:** Update individual rows or perform bulk updates on Excel data.
- **Soft Delete:** Soft delete for both files and rows, preserving data history.
- **Export:** Export filtered or full data back to Excel format.
- **Audit System:** All changes are tracked in an audit table for traceability.
- **Swagger UI:** Interactive API documentation and testing.
- **CORS Support:** Secure cross-origin requests for frontend integration.

## API Endpoints
- `POST /api/excel/upload` — Upload an Excel file
- `GET /api/excel/files` — List uploaded files
- `POST /api/excel/read/{fileName}` — Read data from a file (all sheets or specific sheet)
- `GET /api/excel/data/{fileName}` — Get paginated data
- `GET /api/excel/data/{fileName}/all` — Get all data for a file
- `PUT /api/excel/data` — Update a row
- `PUT /api/excel/data/bulk` — Bulk update
- `POST /api/excel/data` — Add a new row
- `DELETE /api/excel/data/{id}` — Delete a row
- `DELETE /api/excel/files/{fileName}` — Delete a file
- `POST /api/excel/export` — Export data to Excel
- `GET /api/excel/sheets/{fileName}` — List sheets in a file
- `GET /api/excel/statistics/{fileName}` — Get statistics

## Technologies Used
- .NET 9 Web API
- Entity Framework Core (SQL Server)
- EPPlus (Excel operations)
- Swagger (OpenAPI)

## Getting Started
1. **Clone the repository**
2. **Configure your connection string** in `appsettings.json`
3. **Run database migrations:**
   ```
   dotnet ef database update
   ```
4. **Run the API:**
   ```
   dotnet run --project ExcelDataManagementAPI
   ```
5. **Access Swagger UI:**
   - [http://localhost:5002/swagger](http://localhost:5002/swagger)

## Example Use Cases
- Enterprise data import/export
- Data cleaning and transformation
- Versioned data management
- Backend for Excel-based business workflows

## License
This project uses the [EPPlus NonCommercial License](https://epplussoftware.com/developers/licenseexception/).

---

# Excel Data Management API

## Genel Bakış
Excel Data Management API, Excel dosyalarını yüklemek, okumak, düzenlemek, karşılaştırmak ve dışa aktarmak için geliştirilmiş kurumsal seviyede bir .NET 9 Web API projesidir. Kullanıcıların Excel verileriyle programatik ve güvenli şekilde etkileşim kurmasını sağlar.

## Özellikler
- **Excel Dosyası Yükleme:** .xlsx/.xls dosyalarını yükleyip güvenli şekilde saklayın.
- **Veri Okuma:** Belirli bir sheet veya tüm sheet'lerden veri okuyun.
- **Veritabanı Saklama:** Excel verilerini SQL veritabanında saklayarak hızlı erişim ve versiyonlama imkanı.
- **Düzenleme & Güncelleme:** Satır bazında veya toplu veri güncelleme işlemleri.
- **Soft Delete:** Dosya ve satır bazında silme işlemleri, veri geçmişi korunur.
- **Dışa Aktarma:** Filtrelenmiş veya tüm verileri tekrar Excel formatında dışa aktarın.
- **Audit Sistemi:** Tüm değişiklikler audit tablosunda izlenir.
- **Swagger UI:** Etkileşimli API dokümantasyonu ve test ortamı.
- **CORS Desteği:** Frontend entegrasyonu için güvenli cross-origin istekler.

## API Uç Noktaları
- `POST /api/excel/upload` — Excel dosyası yükle
- `GET /api/excel/files` — Yüklenen dosyaları listele
- `POST /api/excel/read/{fileName}` — Dosyadan veri oku (tüm sheet'ler veya belirli bir sheet)
- `GET /api/excel/data/{fileName}` — Sayfalı veri çek
- `GET /api/excel/data/{fileName}/all` — Tüm verileri çek
- `PUT /api/excel/data` — Satır güncelle
- `PUT /api/excel/data/bulk` — Toplu güncelleme
- `POST /api/excel/data` — Yeni satır ekle
- `DELETE /api/excel/data/{id}` — Satır sil
- `DELETE /api/excel/files/{fileName}` — Dosya sil
- `POST /api/excel/export` — Verileri Excel'e aktar
- `GET /api/excel/sheets/{fileName}` — Dosyadaki sheet'leri listele
- `GET /api/excel/statistics/{fileName}` — İstatistikleri getir

## Kullanılan Teknolojiler
- .NET 9 Web API
- Entity Framework Core (SQL Server)
- EPPlus (Excel işlemleri)
- Swagger (OpenAPI)

## Başlarken
1. **Projeyi klonlayın**
2. `appsettings.json` dosyasında bağlantı cümlenizi ayarlayın
3. **Veritabanı migrasyonlarını çalıştırın:**
   ```
   dotnet ef database update
   ```
4. **API'yi başlatın:**
   ```
   dotnet run --project ExcelDataManagementAPI
   ```
5. **Swagger UI'ya erişin:**
   - [http://localhost:5002/swagger](http://localhost:5002/swagger)

## Kullanım Senaryoları
- Kurumsal veri içe/dışa aktarma
- Veri temizleme ve dönüştürme
- Versiyonlu veri yönetimi
- Excel tabanlı iş süreçleri için backend

## Lisans
Bu projede [EPPlus NonCommercial License](https://epplussoftware.com/developers/licenseexception/) kullanılmaktadır.

---

