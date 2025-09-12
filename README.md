# ?? Excel Data Management API

**Excel dosyalarýný yönetmek, veri karþýlaþtýrmasý yapmak ve gerçek zamanlý veri iþleme için geliþtirilmiþ .NET 9 tabanlý RESTful API**

## ?? Özellikler

### ?? Excel Dosya Yönetimi
- ? Excel dosyasý yükleme (.xlsx, .xls)
- ? Çoklu sheet desteði 
- ? Otomatik dosya doðrulama
- ? Dosya listesi ve durumu sorgulama

### ?? Veri Ýþleme
- ? Excel verilerini veritabanýna aktarma
- ? Gerçek zamanlý veri düzenleme
- ? Toplu veri güncelleme
- ? Sayfalý veri getirme (pagination)
- ? Veri filtreleme ve arama

### ?? Karþýlaþtýrma Sistemi
- ? Ýki Excel dosyasýný karþýlaþtýrma
- ? Deðiþiklik takibi
- ? Fark raporlarý oluþturma
- ? Versiyon geçmiþi

### ?? Audit ve Ýzleme
- ? Tüm deðiþikliklerin otomatik kaydý
- ? Kullanýcý bazlý iþlem takibi
- ? IP ve zaman damgasý kayýtlarý
- ? Detaylý aktivite raporlarý

### ?? Güvenlik
- ? CORS yapýlandýrmasý
- ? Dosya türü doðrulama
- ? SQL injection korumasý
- ? Entity Framework güvenliði

## ??? Teknoloji Stack

- **Framework:** .NET 9.0
- **ORM:** Entity Framework Core 9.0
- **Veritabaný:** SQL Server LocalDB
- **Excel Ýþleme:** EPPlus 7.5.0
- **API Dokümantasyonu:** Swagger/OpenAPI
- **Mimari:** Clean Architecture, Repository Pattern

## ? Hýzlý Baþlangýç

### 1. Gereksinimler
```bash
- .NET 9.0 SDK
- SQL Server LocalDB (Visual Studio ile birlikte gelir)
- Visual Studio 2022 (önerilen) veya VS Code
```

### 2. Kurulum
```bash
# Repository'yi klonlayýn
git clone https://github.com/Taha1022-sys/IsnaDataManagementAPI.git
cd IsnaDataManagementAPI

# Baðýmlýlýklarý yükleyin
dotnet restore

# Veritabanýný oluþturun
dotnet ef database update

# Uygulamayý baþlatýn
dotnet run --project ExcelDataManagementAPI
```

### 3. Uygulama Eriþimi
- **API Base URL:** `http://localhost:5002/api`
- **Swagger UI:** `http://localhost:5002/swagger`
- **HTTPS URL:** `https://localhost:7002`

## ?? API Kullaným Örnekleri

### ?? API Test Etme
```http
GET /api/excel/test
```

### ?? Excel Dosyasý Yükleme
```http
POST /api/excel/upload
Content-Type: multipart/form-data

file: [Excel dosyanýz]
uploadedBy: "kullanici_adi"
```

### ?? Yüklü Dosyalarý Listeleme
```http
GET /api/excel/files
```

### ?? Excel Dosyasýný Okuma (Tüm Sayfalar)
```http
POST /api/excel/read/{fileName}
```

### ?? Belirli Sayfayý Okuma
```http
POST /api/excel/read/{fileName}?sheetName=Sayfa1
```

### ?? Veri Getirme (Sayfalama ile)
```http
GET /api/excel/data/{fileName}?page=1&pageSize=50&sheetName=Sayfa1
```

### ?? Tüm Verileri Getirme
```http
GET /api/excel/data/{fileName}/all?sheetName=Sayfa1
```

### ?? Veri Güncelleme
```http
PUT /api/excel/data
Content-Type: application/json

{
  "id": 1,
  "data": {
    "kolon1": "yeni_deger1",
    "kolon2": "yeni_deger2"
  },
  "modifiedBy": "kullanici_adi",
  "changeReason": "Güncelleme nedeni"
}
```

### ?? Yeni Satýr Ekleme
```http
POST /api/excel/data
Content-Type: application/json

{
  "fileName": "dosya_adi",
  "sheetName": "sayfa_adi",
  "rowData": {
    "kolon1": "deger1",
    "kolon2": "deger2"
  },
  "addedBy": "kullanici_adi"
}
```

### ?? Veri Silme
```http
DELETE /api/excel/data/{id}?deletedBy=kullanici_adi
```

### ?? Son Deðiþiklikleri Görme
```http
GET /api/excel/recent-changes?hours=24&limit=100
```

### ?? Dashboard Verileri
```http
GET /api/excel/dashboard
```

### ?? Ýki Dosyayý Karþýlaþtýrma
```http
POST /api/comparison/compare-from-files
Content-Type: multipart/form-data

file1: [Ýlk Excel dosyasý]
file2: [Ýkinci Excel dosyasý]
comparedBy: "kullanici_adi"
```

## ?? Yapýlandýrma

### appsettings.json
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=(localdb)\\mssqllocaldb;Database=ISNADATAMANAGEMENT;Trusted_Connection=true;MultipleActiveResultSets=true;TrustServerCertificate=true;"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*"
}
```

### CORS Ayarlarý
API aþaðýdaki portlarý destekler:
- `http://localhost:5174` (Frontend)
- `http://localhost:3000` (React Dev Server)
- `http://localhost:5173` (Vite Dev Server)

## ?? Veritabaný Yapýsý

### Ana Tablolar
1. **ExcelFiles** - Yüklenen dosya bilgileri
2. **ExcelDataRows** - Excel verilerinin satýr bazýnda saklanmasý
3. **GerceklesenRaporlar** - Tüm deðiþikliklerin audit kaydý

### Migration Komutlarý
```bash
# Yeni migration oluþturma
dotnet ef migrations add MigrationAdi

# Migration uygulama
dotnet ef database update

# Migration geri alma
dotnet ef database update PreviousMigrationName
```

## ?? Sorun Giderme

### Yaygýn Sorunlar

#### 1. Veritabaný Baðlantý Hatasý
```bash
# LocalDB'nin çalýþtýðýndan emin olun
sqllocaldb info
sqllocaldb start mssqllocaldb

# Migration'larý tekrar uygulayýn
dotnet ef database update
```

#### 2. Excel Dosyasý Yüklenmiyor
- Dosya boyutunun 100MB'dan küçük olduðundan emin olun
- Sadece `.xlsx` ve `.xls` dosyalarý desteklenir
- Dosya bozuk olmadýðýndan emin olun

#### 3. CORS Hatasý
- `appsettings.json`'da doðru URL'lerin tanýmlý olduðundan emin olun
- Development ortamýnda `DevelopmentPolicy` kullanýlýr

#### 4. Port Çakýþmasý
```bash
# Kullanýlan portlarý kontrol edin
netstat -an | findstr :5002

# launchSettings.json'da port deðiþtirin
```

## ?? Log Ýzleme

### Konsol Çýktýsý
Uygulama baþlatýldýðýnda görülebilecek mesajlar:
- ? Migration durumu
- ? Veritabaný baðlantýsý
- ? Audit tablosu kontrolü
- ?? API endpoint'leri

### Log Seviyeleri
- **Information:** Genel iþlem bilgileri
- **Warning:** Uyarýlar ve dikkat edilmesi gerekenler
- **Error:** Hatalar ve exception'lar

## ?? Güncellemeler

### v1.0.0 Özellikleri
- Excel dosya yükleme ve okuma
- Çoklu sheet desteði
- Veri düzenleme (CRUD)
- Karþýlaþtýrma sistemi
- Audit sistemi
- Dashboard ve raporlama

## ?? Katkýda Bulunma

1. Repository'yi fork edin
2. Feature branch oluþturun (`git checkout -b feature/YeniOzellik`)
3. Deðiþikliklerinizi commit edin (`git commit -am 'Yeni özellik eklendi'`)
4. Branch'i push edin (`git push origin feature/YeniOzellik`)
5. Pull Request oluþturun

## ?? Ýletiþim

- **GitHub:** [Taha1022-sys](https://github.com/Taha1022-sys)
- **Repository:** [IsnaDataManagementAPI](https://github.com/Taha1022-sys/IsnaDataManagementAPI)

## ?? Lisans

Bu proje MIT lisansý altýnda lisanslanmýþtýr.

---

## ?? Gelecek Özellikler

- [ ] Authentication ve Authorization
- [ ] Redis Cache entegrasyonu
- [ ] Email bildirimleri
- [ ] Dosya sürüm yönetimi
- [ ] Bulk operations API
- [ ] Export to different formats (PDF, CSV)
- [ ] Real-time notifications (SignalR)
- [ ] Advanced filtering and search
- [ ] Data validation rules
- [ ] Scheduled data imports

---

**? Bu projeyi beðendiyseniz yýldýz vermeyi unutmayýn!**