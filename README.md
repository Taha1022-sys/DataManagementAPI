# ?? Excel Data Management API

**Excel dosyalar�n� y�netmek, veri kar��la�t�rmas� yapmak ve ger�ek zamanl� veri i�leme i�in geli�tirilmi� .NET 9 tabanl� RESTful API**

## ?? �zellikler

### ?? Excel Dosya Y�netimi
- ? Excel dosyas� y�kleme (.xlsx, .xls)
- ? �oklu sheet deste�i 
- ? Otomatik dosya do�rulama
- ? Dosya listesi ve durumu sorgulama

### ?? Veri ��leme
- ? Excel verilerini veritaban�na aktarma
- ? Ger�ek zamanl� veri d�zenleme
- ? Toplu veri g�ncelleme
- ? Sayfal� veri getirme (pagination)
- ? Veri filtreleme ve arama

### ?? Kar��la�t�rma Sistemi
- ? �ki Excel dosyas�n� kar��la�t�rma
- ? De�i�iklik takibi
- ? Fark raporlar� olu�turma
- ? Versiyon ge�mi�i

### ?? Audit ve �zleme
- ? T�m de�i�ikliklerin otomatik kayd�
- ? Kullan�c� bazl� i�lem takibi
- ? IP ve zaman damgas� kay�tlar�
- ? Detayl� aktivite raporlar�

### ?? G�venlik
- ? CORS yap�land�rmas�
- ? Dosya t�r� do�rulama
- ? SQL injection korumas�
- ? Entity Framework g�venli�i

## ??? Teknoloji Stack

- **Framework:** .NET 9.0
- **ORM:** Entity Framework Core 9.0
- **Veritaban�:** SQL Server LocalDB
- **Excel ��leme:** EPPlus 7.5.0
- **API Dok�mantasyonu:** Swagger/OpenAPI
- **Mimari:** Clean Architecture, Repository Pattern

## ? H�zl� Ba�lang��

### 1. Gereksinimler
```bash
- .NET 9.0 SDK
- SQL Server LocalDB (Visual Studio ile birlikte gelir)
- Visual Studio 2022 (�nerilen) veya VS Code
```

### 2. Kurulum
```bash
# Repository'yi klonlay�n
git clone https://github.com/Taha1022-sys/IsnaDataManagementAPI.git
cd IsnaDataManagementAPI

# Ba��ml�l�klar� y�kleyin
dotnet restore

# Veritaban�n� olu�turun
dotnet ef database update

# Uygulamay� ba�lat�n
dotnet run --project ExcelDataManagementAPI
```

### 3. Uygulama Eri�imi
- **API Base URL:** `http://localhost:5002/api`
- **Swagger UI:** `http://localhost:5002/swagger`
- **HTTPS URL:** `https://localhost:7002`

## ?? API Kullan�m �rnekleri

### ?? API Test Etme
```http
GET /api/excel/test
```

### ?? Excel Dosyas� Y�kleme
```http
POST /api/excel/upload
Content-Type: multipart/form-data

file: [Excel dosyan�z]
uploadedBy: "kullanici_adi"
```

### ?? Y�kl� Dosyalar� Listeleme
```http
GET /api/excel/files
```

### ?? Excel Dosyas�n� Okuma (T�m Sayfalar)
```http
POST /api/excel/read/{fileName}
```

### ?? Belirli Sayfay� Okuma
```http
POST /api/excel/read/{fileName}?sheetName=Sayfa1
```

### ?? Veri Getirme (Sayfalama ile)
```http
GET /api/excel/data/{fileName}?page=1&pageSize=50&sheetName=Sayfa1
```

### ?? T�m Verileri Getirme
```http
GET /api/excel/data/{fileName}/all?sheetName=Sayfa1
```

### ?? Veri G�ncelleme
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
  "changeReason": "G�ncelleme nedeni"
}
```

### ?? Yeni Sat�r Ekleme
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

### ?? Son De�i�iklikleri G�rme
```http
GET /api/excel/recent-changes?hours=24&limit=100
```

### ?? Dashboard Verileri
```http
GET /api/excel/dashboard
```

### ?? �ki Dosyay� Kar��la�t�rma
```http
POST /api/comparison/compare-from-files
Content-Type: multipart/form-data

file1: [�lk Excel dosyas�]
file2: [�kinci Excel dosyas�]
comparedBy: "kullanici_adi"
```

## ?? Yap�land�rma

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

### CORS Ayarlar�
API a�a��daki portlar� destekler:
- `http://localhost:5174` (Frontend)
- `http://localhost:3000` (React Dev Server)
- `http://localhost:5173` (Vite Dev Server)

## ?? Veritaban� Yap�s�

### Ana Tablolar
1. **ExcelFiles** - Y�klenen dosya bilgileri
2. **ExcelDataRows** - Excel verilerinin sat�r baz�nda saklanmas�
3. **GerceklesenRaporlar** - T�m de�i�ikliklerin audit kayd�

### Migration Komutlar�
```bash
# Yeni migration olu�turma
dotnet ef migrations add MigrationAdi

# Migration uygulama
dotnet ef database update

# Migration geri alma
dotnet ef database update PreviousMigrationName
```

## ?? Sorun Giderme

### Yayg�n Sorunlar

#### 1. Veritaban� Ba�lant� Hatas�
```bash
# LocalDB'nin �al��t���ndan emin olun
sqllocaldb info
sqllocaldb start mssqllocaldb

# Migration'lar� tekrar uygulay�n
dotnet ef database update
```

#### 2. Excel Dosyas� Y�klenmiyor
- Dosya boyutunun 100MB'dan k���k oldu�undan emin olun
- Sadece `.xlsx` ve `.xls` dosyalar� desteklenir
- Dosya bozuk olmad���ndan emin olun

#### 3. CORS Hatas�
- `appsettings.json`'da do�ru URL'lerin tan�ml� oldu�undan emin olun
- Development ortam�nda `DevelopmentPolicy` kullan�l�r

#### 4. Port �ak��mas�
```bash
# Kullan�lan portlar� kontrol edin
netstat -an | findstr :5002

# launchSettings.json'da port de�i�tirin
```

## ?? Log �zleme

### Konsol ��kt�s�
Uygulama ba�lat�ld���nda g�r�lebilecek mesajlar:
- ? Migration durumu
- ? Veritaban� ba�lant�s�
- ? Audit tablosu kontrol�
- ?? API endpoint'leri

### Log Seviyeleri
- **Information:** Genel i�lem bilgileri
- **Warning:** Uyar�lar ve dikkat edilmesi gerekenler
- **Error:** Hatalar ve exception'lar

## ?? G�ncellemeler

### v1.0.0 �zellikleri
- Excel dosya y�kleme ve okuma
- �oklu sheet deste�i
- Veri d�zenleme (CRUD)
- Kar��la�t�rma sistemi
- Audit sistemi
- Dashboard ve raporlama

## ?? Katk�da Bulunma

1. Repository'yi fork edin
2. Feature branch olu�turun (`git checkout -b feature/YeniOzellik`)
3. De�i�ikliklerinizi commit edin (`git commit -am 'Yeni �zellik eklendi'`)
4. Branch'i push edin (`git push origin feature/YeniOzellik`)
5. Pull Request olu�turun

## ?? �leti�im

- **GitHub:** [Taha1022-sys](https://github.com/Taha1022-sys)
- **Repository:** [IsnaDataManagementAPI](https://github.com/Taha1022-sys/IsnaDataManagementAPI)

## ?? Lisans

Bu proje MIT lisans� alt�nda lisanslanm��t�r.

---

## ?? Gelecek �zellikler

- [ ] Authentication ve Authorization
- [ ] Redis Cache entegrasyonu
- [ ] Email bildirimleri
- [ ] Dosya s�r�m y�netimi
- [ ] Bulk operations API
- [ ] Export to different formats (PDF, CSV)
- [ ] Real-time notifications (SignalR)
- [ ] Advanced filtering and search
- [ ] Data validation rules
- [ ] Scheduled data imports

---

**? Bu projeyi be�endiyseniz y�ld�z vermeyi unutmay�n!**