# 📚 Excel Data Management API - Kod Dokümantasyonu

Bu dokümantasyon, Excel Data Management API projesindeki her kod parçasının ne işe yaradığını, nerede kullanıldığını ve nasıl çalıştığını sıfırdan açıklar.

---

## 📁 Proje Yapısı

### Ana Dizin Yapısı
```
ExcelDataManagementAPI/
├── Controllers/          # API endpoint'leri (HTTP isteklerini karşılar)
├── Services/            # İş mantığı katmanı
├── Models/              # Veri modelleri ve DTO'lar
├── Data/                # Veritabanı context'i
├── Migrations/          # Veritabanı migration dosyaları
├── uploads/             # Yüklenen Excel dosyaları
├── Program.cs           # Uygulama başlangıç noktası
└── appsettings.json     # Yapılandırma dosyası
```

---

## 🚀 Program.cs - Uygulama Başlangıç Noktası

### Ne İşe Yarar?
Program.cs, .NET 9 uygulamasının başlangıç noktasıdır. Tüm servisleri yapılandırır, middleware'leri ayarlar ve uygulamayı başlatır.

### Kod Açıklaması

#### 1. Temel Konfigürasyon
```csharp
var builder = WebApplication.CreateBuilder(args);
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
```
- `WebApplication.CreateBuilder()`: .NET 9'un yeni hosting modelini başlatır
- `ExcelPackage.LicenseContext`: EPPlus kütüphanesini ticari olmayan kullanım için ayarlar

#### 2. Servis Kayıtları
```csharp
builder.Services.AddControllers();                    // MVC Controller'ları aktif eder
builder.Services.AddEndpointsApiExplorer();          // API keşif özelliklerini açar
builder.Services.AddSwaggerGen();                    // Swagger dokümantasyonunu ekler
```

#### 3. Veritabanı Konfigürasyonu
```csharp
builder.Services.AddDbContext<ExcelDataContext>(options =>
{
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"));
});
```
- Entity Framework Core'u SQL Server ile kullanım için yapılandırır
- Connection string'i appsettings.json'dan okur

#### 4. Dependency Injection
```csharp
builder.Services.AddScoped<IExcelService, ExcelService>();
builder.Services.AddScoped<IDataComparisonService, DataComparisonService>();
builder.Services.AddScoped<IAuditService, AuditService>();
```
- Servisleri DI container'a kaydeder
- `AddScoped`: Her HTTP isteği için bir instance oluşturur

#### 5. CORS Konfigürasyonu
```csharp
builder.Services.AddCors(options =>
{
    options.AddPolicy("ApiPolicy", policy =>
    {
        policy.WithOrigins("http://localhost:5174", "http://localhost:3000")
              .AllowAnyHeader()
              .AllowAnyMethod()
              .AllowCredentials();
    });
});
```
- Cross-Origin Resource Sharing ayarları
- Frontend uygulamalarının API'ye erişimini sağlar

#### 6. Middleware Pipeline
```csharp
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();           // Swagger JSON endpoint'i açar
    app.UseSwaggerUI();         // Swagger UI'ını aktif eder
}
app.UseCors();                  // CORS middleware'ini aktif eder
app.UseAuthorization();         // Yetkilendirme middleware'i
app.MapControllers();           // Controller routing'i aktif eder
```

#### 7. Veritabanı Başlatma
```csharp
using var scope = app.Services.CreateScope();
var context = scope.ServiceProvider.GetRequiredService<ExcelDataContext>();
await context.Database.MigrateAsync();
```
- Uygulama başlarken bekleyen migration'ları otomatik uygular
- Veritabanı bağlantısını test eder

---

## 🎯 Controllers/ - API Endpoint'leri

### ExcelController.cs

#### Ne İşe Yarar?
Excel dosyalarıyla ilgili tüm HTTP isteklerini karşılar. RESTful API prensiplerine uygun endpoint'ler sağlar.

#### Ana Metodlar

##### 1. Test Endpoint
```csharp
[HttpGet("test")]
public IActionResult Test()
```
**Ne yapar:** API'nin çalışıp çalışmadığını test eder
**Nerede kullanılır:** Sistem sağlığı kontrolü, bağlantı testleri
**Döner:** Mevcut işlemler listesi ve zaman damgası

##### 2. Dosya Yükleme
```csharp
[HttpPost("upload")]
[Consumes("multipart/form-data")]
public async Task<IActionResult> Upload(IFormFile file, [FromForm] string? uploadedBy = null)
```
**Ne yapar:** Bilgisayardan Excel dosyası yükler
**Validasyonlar:**
- Dosya boş olmamalı
- Sadece .xlsx ve .xls uzantıları kabul edilir
- Dosya boyutu kontrolleri

##### 3. Dosya Okuma
```csharp
[HttpPost("read/{fileName}")]
public async Task<IActionResult> ReadExcelData(string fileName, [FromQuery] string? sheetName = null)
```
**Ne yapar:** Yüklenen Excel dosyasını okuyup veritabanına aktarır
**Özellikler:**
- Tekrar okuma önleme (cache kontrolü)
- Çoklu sheet desteği
- Zorla okuma seçeneği

##### 4. Veri Getirme
```csharp
[HttpGet("data/{fileName}")]
public async Task<IActionResult> GetExcelData(string fileName, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
```
**Ne yapar:** Veritabanından Excel verilerini sayfalı olarak getirir
**Parametreler:**
- `fileName`: Dosya adı
- `page`: Sayfa numarası (varsayılan: 1)
- `pageSize`: Sayfa başına kayıt (varsayılan: 50)

##### 5. Veri Güncelleme
```csharp
[HttpPut("data")]
public async Task<IActionResult> UpdateExcelData([FromBody] ExcelDataUpdateDto updateDto)
```
**Ne yapar:** Mevcut veriyi günceller
**Özellikler:**
- Versiyon kontrolü
- Audit log'a kayıt
- Eşzamanlılık kontrolü

### ComparisonController.cs

#### Ne İşe Yarar?
İki Excel dosyasını karşılaştırma işlemlerini yönetir.

##### Ana Metod
```csharp
[HttpPost("compare-from-files")]
public async Task<IActionResult> CompareFromFiles([FromForm] CompareExcelFilesDto request)
```
**Ne yapar:** İki dosyayı karşılaştırıp farkları gösterir
**Çıktı:** Değişen, eklenen ve silinen kayıtların listesi

---

## ⚙️ Services/ - İş Mantığı Katmanı

### IExcelService.cs - Interface Tanımı

#### Ne İşe Yarar?
Excel işlemleri için contract (sözleşme) tanımlar. Dependency Injection için gerekli.

```csharp
public interface IExcelService
{
    Task<ExcelFile> UploadExcelFileAsync(IFormFile file, string? uploadedBy = null);
    Task<List<ExcelDataResponseDto>> ReadExcelDataAsync(string fileName, string? sheetName = null);
    Task<List<ExcelDataResponseDto>> GetExcelDataAsync(string fileName, string? sheetName = null, int page = 1, int pageSize = 50);
    // ... diğer metodlar
}
```

### ExcelService.cs - İş Mantığı Implementasyonu

#### Ne İşe Yarar?
Excel dosyalarıyla ilgili tüm iş mantığını içerir. EPPlus kütüphanesini kullanarak Excel işlemlerini gerçekleştirir.

#### Ana Metodlar

##### 1. Dosya Yükleme
```csharp
public async Task<ExcelFile> UploadExcelFileAsync(IFormFile file, string? uploadedBy = null)
```
**Yapar:**
1. Dosyayı fiziksel olarak uploads/ klasörüne kaydeder
2. Benzersiz dosya adı oluşturur (timestamp ile)
3. Veritabanına dosya bilgilerini kaydeder

##### 2. Excel Okuma
```csharp
public async Task<List<ExcelDataResponseDto>> ReadExcelDataAsync(string fileName, string? sheetName = null)
```
**Yapar:**
1. EPPlus ile Excel dosyasını açar
2. Tüm satırları okur
3. JSON formatında veritabanına kaydeder
4. DTO listesi döner

##### 3. Veri Güncelleme
```csharp
public async Task<ExcelDataResponseDto> UpdateExcelDataAsync(ExcelDataUpdateDto updateDto, HttpContext? httpContext = null)
```
**Yapar:**
1. Mevcut kaydı bulur
2. Değişiklikleri uygular
3. Audit log'a kaydeder
4. Versiyon numarasını arttırır

### DataComparisonService.cs

#### Ne İşe Yarar?
İki Excel dosyasını karşılaştırma algoritmasını içerir.

##### Ana Algoritma
```csharp
public async Task<ComparisonResultDto> CompareExcelFilesAsync(...)
```
**Yapar:**
1. Her iki dosyayı okur
2. Satır bazında karşılaştırır
3. Farkları kategorize eder (eklenen, silinen, değişen)
4. Detaylı rapor oluşturur

### AuditService.cs

#### Ne İşe Yarar?
Tüm veri değişikliklerini audit tablosuna kaydeder.

```csharp
public async Task LogChangeAsync(...)
```
**Kaydeder:**
- Hangi kullanıcı
- Ne zaman
- Hangi veriyi
- Nasıl değiştirdi
- IP adresi ve User Agent bilgileri

---

## 📊 Models/ - Veri Modelleri

### ExcelDataModels.cs

#### 1. ExcelFile Entity
```csharp
public class ExcelFile
{
    public int Id { get; set; }                    // Primary Key
    public string FileName { get; set; }           // Benzersiz dosya adı
    public string OriginalFileName { get; set; }   // Orijinal dosya adı
    public string FilePath { get; set; }           // Fiziksel dosya yolu
    public long FileSize { get; set; }             // Dosya boyutu (byte)
    public DateTime UploadDate { get; set; }       // Yüklenme tarihi
    public string? UploadedBy { get; set; }        // Yükleyen kullanıcı
    public bool IsActive { get; set; }             // Aktif/pasif durumu
}
```

#### 2. ExcelDataRow Entity
```csharp
public class ExcelDataRow
{
    public int Id { get; set; }                    // Primary Key
    public string FileName { get; set; }           // Hangi dosyaya ait
    public string SheetName { get; set; }          // Hangi sayfa
    public int RowIndex { get; set; }              // Satır numarası
    public string Data { get; set; }               // JSON formatında veri
    public DateTime CreatedDate { get; set; }      // Oluşturulma tarihi
    public DateTime? ModifiedDate { get; set; }    // Son değişiklik tarihi
    public string? ModifiedBy { get; set; }        // Değiştiren kullanıcı
    public bool IsDeleted { get; set; }            // Silinmiş mi?
    public int Version { get; set; }               // Versiyon numarası
}
```

#### 3. GerceklesenRaporlar (Audit Table)
```csharp
public class GerceklesenRaporlar
{
    public int Id { get; set; }                    // Primary Key
    public string? FileName { get; set; }          // Dosya adı
    public string? SheetName { get; set; }         // Sayfa adı
    public int? RowIndex { get; set; }             // Satır numarası
    public string? OperationType { get; set; }     // İşlem tipi (Create/Update/Delete)
    public DateTime ChangeDate { get; set; }       // Değişiklik tarihi
    public string? ModifiedBy { get; set; }        // Değiştiren kullanıcı
    public string? UserIP { get; set; }            // IP adresi
    public string? UserAgent { get; set; }         // Browser bilgisi
    public string? ChangeReason { get; set; }      // Değişiklik sebebi
    public string? OldValue { get; set; }          // Eski değer
    public string? NewValue { get; set; }          // Yeni değer
    public string? ChangedColumns { get; set; }    // Değişen kolonlar
    public bool IsSuccess { get; set; }            // İşlem başarılı mı?
    public string? ErrorMessage { get; set; }      // Hata mesajı (varsa)
}
```

### DTOs/ - Data Transfer Objects

#### ExcelDataDTOs.cs

##### 1. ExcelDataResponseDto
```csharp
public class ExcelDataResponseDto
{
    public int Id { get; set; }                           // Kayıt ID'si
    public string FileName { get; set; }                  // Dosya adı
    public string SheetName { get; set; }                 // Sayfa adı
    public int RowIndex { get; set; }                     // Satır numarası
    public Dictionary<string, string> Data { get; set; }  // Key-Value veri
    public DateTime CreatedDate { get; set; }             // Oluşturulma tarihi
    public DateTime? ModifiedDate { get; set; }           // Değişiklik tarihi
    public string? ModifiedBy { get; set; }               // Değiştiren
    public int Version { get; set; }                      // Versiyon
}
```

##### 2. ExcelDataUpdateDto
```csharp
public class ExcelDataUpdateDto
{
    public int Id { get; set; }                           // Güncellenecek kayıt ID'si
    public Dictionary<string, string> Data { get; set; }  // Yeni veriler
    public string? ModifiedBy { get; set; }               // Değiştiren kullanıcı
    public string? ChangeReason { get; set; }             // Değişiklik sebebi
}
```

---

## 🗄️ Data/ - Veritabanı Katmanı

### ExcelDataContext.cs

#### Ne İşe Yarar?
Entity Framework Core context'i. Veritabanı ile uygulama arasındaki köprü görevi görür.

```csharp
public class ExcelDataContext : DbContext
{
    public DbSet<ExcelFile> ExcelFiles { get; set; }                    // Dosya tablosu
    public DbSet<ExcelDataRow> ExcelDataRows { get; set; }              // Veri tablosu
    public DbSet<GerceklesenRaporlar> GerceklesenRaporlar { get; set; }  // Audit tablosu

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // Tablo konfigürasyonları
        modelBuilder.Entity<ExcelFile>(entity =>
        {
            entity.HasKey(e => e.Id);                          // Primary Key
            entity.Property(e => e.FileName).IsRequired();     // Zorunlu alan
            entity.HasIndex(e => e.FileName).IsUnique();       // Benzersiz index
        });

        modelBuilder.Entity<ExcelDataRow>(entity =>
        {
            entity.HasKey(e => e.Id);
            entity.Property(e => e.Data).HasColumnType("nvarchar(max)");  // JSON için uzun text
            entity.HasIndex(e => new { e.FileName, e.SheetName, e.RowIndex }); // Compound index
        });
    }
}
```

---

## 🔄 Migrations/ - Veritabanı Şeması

### Migration Dosyaları

#### 20250807122033_InitialCreate.cs
**Ne yapar:** İlk veritabanı şemasını oluşturur
- ExcelFiles tablosu
- ExcelDataRows tablosu  
- GerceklesenRaporlar tablosu
- Index'ler ve kısıtlamalar

#### Migration Komutları
```bash
# Yeni migration oluşturma
dotnet ef migrations add MigrationName

# Migration uygulama
dotnet ef database update

# Migration kaldırma
dotnet ef migrations remove
```

---

## ⚙️ Yapılandırma Dosyaları

### appsettings.json
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Server=(localdb)\\mssqllocaldb;Database=ISNADATAMANAGEMENT;Trusted_Connection=true;MultipleActiveResultSets=true;TrustServerCertificate=true;"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",                              // Genel log seviyesi
      "Microsoft.AspNetCore": "Warning",                     // ASP.NET Core logları
      "Microsoft.EntityFrameworkCore.Database.Command": "Information"  // SQL sorgu logları
    }
  },
  "AllowedHosts": "*"                                       // İzin verilen host'lar
}
```

### appsettings.Development.json
Development ortamına özel ayarlar (genelde daha detaylı loglama)

---

## 📦 NuGet Paketleri ve Kullanım Amaçları

### ExcelDataManagementAPI.csproj
```xml
<PackageReference Include="EPPlus" Version="7.5.0" />
<!-- Excel dosyalarını okuma/yazma için -->

<PackageReference Include="Microsoft.EntityFrameworkCore.SqlServer" Version="9.0.7" />
<!-- SQL Server veritabanı provider'ı -->

<PackageReference Include="Microsoft.EntityFrameworkCore.Tools" Version="9.0.7" />
<!-- Migration komutları için (Package Manager Console) -->

<PackageReference Include="Microsoft.EntityFrameworkCore.Design" Version="9.0.7" />
<!-- Design-time EF Core işlemleri için -->

<PackageReference Include="Swashbuckle.AspNetCore" Version="9.0.3" />
<!-- Swagger/OpenAPI dokümantasyonu için -->
```

---

## 🔄 Veri Akışı ve İşlem Sırası

### 1. Excel Dosyası Yükleme İşlemi
```
1. Frontend'den dosya seçilir
2. ExcelController.Upload() çağrılır
3. Dosya validasyonu yapılır
4. ExcelService.UploadExcelFileAsync() çalışır
5. Dosya fiziksel olarak kaydedilir
6. Veritabanına dosya bilgileri eklenir
7. Başarı response'u döner
```

### 2. Excel Okuma İşlemi
```
1. ExcelController.ReadExcelData() çağrılır
2. Cache kontrolü yapılır (daha önce okunmuş mu?)
3. ExcelService.ReadExcelDataAsync() çalışır
4. EPPlus ile Excel dosyası açılır
5. Tüm satırlar okunur ve JSON'a çevrilir
6. Veritabanına satır satır kaydedilir
7. DTO listesi response olarak döner
```

### 3. Veri Güncelleme İşlemi
```
1. ExcelController.UpdateExcelData() çağrılır
2. Validasyonlar yapılır
3. ExcelService.UpdateExcelDataAsync() çalışır
4. Mevcut kayıt bulunur
5. Versiyon kontrolü yapılır
6. Değişiklikler uygulanır
7. AuditService.LogChangeAsync() çağrılır
8. Audit tablosuna log kaydedilir
9. Güncel veri response olarak döner
```

---

## 🔍 Hata Yönetimi ve Logging

### Global Exception Handling
```csharp
try
{
    // İş mantığı
}
catch (Exception ex)
{
    _logger.LogError(ex, "Hata mesajı: {FileName}", fileName);
    return StatusCode(500, new { success = false, message = ex.Message });
}
```

### Log Seviyeleri
- **LogTrace**: En detaylı bilgiler
- **LogDebug**: Debug bilgileri
- **LogInformation**: Genel bilgiler
- **LogWarning**: Uyarılar
- **LogError**: Hatalar
- **LogCritical**: Kritik hatalar

---

## 🚀 Performance Optimizasyonları

### 1. Database Indexing
```csharp
// ExcelDataContext.cs içinde
entity.HasIndex(e => new { e.FileName, e.SheetName, e.RowIndex });
```

### 2. Pagination
```csharp
// ExcelController.cs içinde
var data = await query
    .Skip((page - 1) * pageSize)
    .Take(pageSize)
    .ToListAsync();
```

### 3. Async Programming
Tüm veritabanı işlemleri async/await pattern'i ile yazılmıştır.

---

## 🔐 Güvenlik Önlemleri

### 1. SQL Injection Koruması
Entity Framework Core parameterized queries kullanır.

### 2. Dosya Validasyonu
```csharp
var allowedExtensions = new[] { ".xlsx", ".xls" };
var fileExtension = Path.GetExtension(file.FileName).ToLowerInvariant();
if (!allowedExtensions.Contains(fileExtension))
{
    return BadRequest("Sadece Excel dosyaları desteklenir");
}
```

### 3. CORS Yapılandırması
Sadece belirtilen origin'lere izin verir.

---

## 📈 Monitoring ve Audit

### Audit Log Sistemi
Her veri değişikliği otomatik olarak `GerceklesenRaporlar` tablosuna kaydedilir:
- Kim değiştirdi
- Ne zaman değiştirdi
- Neyi değiştirdi
- Eski ve yeni değerler
- IP adresi ve browser bilgileri

### Dashboard Endpoint'i
```csharp
[HttpGet("dashboard")]
public async Task<IActionResult> GetDashboardData()
```
Sistem istatistiklerini ve güncel durumu gösterir.

---

Bu dokümantasyon, projedeki her kod parçasının amacını ve işleyişini detaylı olarak açıklar. Herhangi bir geliştirici bu bilgilerle projeyi anlayabilir ve üzerinde çalışabilir.