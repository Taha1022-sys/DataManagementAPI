# Excel Data Management API

## Genel Bakýþ
Excel Data Management API, Excel dosyalarýný yüklemek, okumak, düzenlemek, karþýlaþtýrmak ve dýþa aktarmak için geliþtirilmiþ kurumsal seviyede bir .NET 9 Web API projesidir. Kullanýcýlarýn Excel verileriyle programatik ve güvenli þekilde etkileþim kurmasýný saðlar.

## Özellikler
- **Excel Dosyasý Yükleme:** .xlsx/.xls dosyalarýný yükleyip güvenli þekilde saklayýn.
- **Veri Okuma:** Belirli bir sheet veya tüm sheet'lerden veri okuyun.
- **Veritabaný Saklama:** Excel verilerini SQL veritabanýnda saklayarak hýzlý eriþim ve versiyonlama imkaný.
- **Düzenleme & Güncelleme:** Satýr bazýnda veya toplu veri güncelleme iþlemleri.
- **Soft Delete:** Dosya ve satýr bazýnda silme iþlemleri, veri geçmiþi korunur.
- **Dýþa Aktarma:** Filtrelenmiþ veya tüm verileri tekrar Excel formatýnda dýþa aktarýn.
- **Audit Sistemi:** Tüm deðiþiklikler audit tablosunda izlenir.
- **Swagger UI:** Etkileþimli API dokümantasyonu ve test ortamý.
- **CORS Desteði:** Frontend entegrasyonu için güvenli cross-origin istekler.

## API Uç Noktalarý
- `POST /api/excel/upload` — Excel dosyasý yükle
- `GET /api/excel/files` — Yüklenen dosyalarý listele
- `POST /api/excel/read/{fileName}` — Dosyadan veri oku (tüm sheet'ler veya belirli bir sheet)
- `GET /api/excel/data/{fileName}` — Sayfalý veri çek
- `GET /api/excel/data/{fileName}/all` — Tüm verileri çek
- `PUT /api/excel/data` — Satýr güncelle
- `PUT /api/excel/data/bulk` — Toplu güncelleme
- `POST /api/excel/data` — Yeni satýr ekle
- `DELETE /api/excel/data/{id}` — Satýr sil
- `DELETE /api/excel/files/{fileName}` — Dosya sil
- `POST /api/excel/export` — Verileri Excel'e aktar
- `GET /api/excel/sheets/{fileName}` — Dosyadaki sheet'leri listele
- `GET /api/excel/statistics/{fileName}` — Ýstatistikleri getir

## Kullanýlan Teknolojiler
- .NET 9 Web API
- Entity Framework Core (SQL Server)
- EPPlus (Excel iþlemleri)
- Swagger (OpenAPI)

## Baþlarken
1. **Projeyi klonlayýn**
2. `appsettings.json` dosyasýnda baðlantý cümlenizi ayarlayýn
3. **Veritabaný migrasyonlarýný çalýþtýrýn:**
   ```
   dotnet ef database update
   ```
4. **API'yi baþlatýn:**
   ```
   dotnet run --project ExcelDataManagementAPI
   ```
5. **Swagger UI'ya eriþin:**
   - [http://localhost:5002/swagger](http://localhost:5002/swagger)

## Kullaným Senaryolarý
- Kurumsal veri içe/dýþa aktarma
- Veri temizleme ve dönüþtürme
- Versiyonlu veri yönetimi
- Excel tabanlý iþ süreçleri için backend

## Lisans
Bu projede [EPPlus NonCommercial License](https://epplussoftware.com/developers/licenseexception/) kullanýlmaktadýr.

---

# For English, see the README.md file.
