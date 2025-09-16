# Excel Data Management API

## Genel Bak��
Excel Data Management API, Excel dosyalar�n� y�klemek, okumak, d�zenlemek, kar��la�t�rmak ve d��a aktarmak i�in geli�tirilmi� kurumsal seviyede bir .NET 9 Web API projesidir. Kullan�c�lar�n Excel verileriyle programatik ve g�venli �ekilde etkile�im kurmas�n� sa�lar.

## �zellikler
- **Excel Dosyas� Y�kleme:** .xlsx/.xls dosyalar�n� y�kleyip g�venli �ekilde saklay�n.
- **Veri Okuma:** Belirli bir sheet veya t�m sheet'lerden veri okuyun.
- **Veritaban� Saklama:** Excel verilerini SQL veritaban�nda saklayarak h�zl� eri�im ve versiyonlama imkan�.
- **D�zenleme & G�ncelleme:** Sat�r baz�nda veya toplu veri g�ncelleme i�lemleri.
- **Soft Delete:** Dosya ve sat�r baz�nda silme i�lemleri, veri ge�mi�i korunur.
- **D��a Aktarma:** Filtrelenmi� veya t�m verileri tekrar Excel format�nda d��a aktar�n.
- **Audit Sistemi:** T�m de�i�iklikler audit tablosunda izlenir.
- **Swagger UI:** Etkile�imli API dok�mantasyonu ve test ortam�.
- **CORS Deste�i:** Frontend entegrasyonu i�in g�venli cross-origin istekler.

## API U� Noktalar�
- `POST /api/excel/upload` � Excel dosyas� y�kle
- `GET /api/excel/files` � Y�klenen dosyalar� listele
- `POST /api/excel/read/{fileName}` � Dosyadan veri oku (t�m sheet'ler veya belirli bir sheet)
- `GET /api/excel/data/{fileName}` � Sayfal� veri �ek
- `GET /api/excel/data/{fileName}/all` � T�m verileri �ek
- `PUT /api/excel/data` � Sat�r g�ncelle
- `PUT /api/excel/data/bulk` � Toplu g�ncelleme
- `POST /api/excel/data` � Yeni sat�r ekle
- `DELETE /api/excel/data/{id}` � Sat�r sil
- `DELETE /api/excel/files/{fileName}` � Dosya sil
- `POST /api/excel/export` � Verileri Excel'e aktar
- `GET /api/excel/sheets/{fileName}` � Dosyadaki sheet'leri listele
- `GET /api/excel/statistics/{fileName}` � �statistikleri getir

## Kullan�lan Teknolojiler
- .NET 9 Web API
- Entity Framework Core (SQL Server)
- EPPlus (Excel i�lemleri)
- Swagger (OpenAPI)

## Ba�larken
1. **Projeyi klonlay�n**
2. `appsettings.json` dosyas�nda ba�lant� c�mlenizi ayarlay�n
3. **Veritaban� migrasyonlar�n� �al��t�r�n:**
   ```
   dotnet ef database update
   ```
4. **API'yi ba�lat�n:**
   ```
   dotnet run --project ExcelDataManagementAPI
   ```
5. **Swagger UI'ya eri�in:**
   - [http://localhost:5002/swagger](http://localhost:5002/swagger)

## Kullan�m Senaryolar�
- Kurumsal veri i�e/d��a aktarma
- Veri temizleme ve d�n��t�rme
- Versiyonlu veri y�netimi
- Excel tabanl� i� s�re�leri i�in backend

## Lisans
Bu projede [EPPlus NonCommercial License](https://epplussoftware.com/developers/licenseexception/) kullan�lmaktad�r.

---

# For English, see the README.md file.
