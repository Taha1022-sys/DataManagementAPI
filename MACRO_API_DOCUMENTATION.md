# MacroController API Dokümantasyonu

MacroController, Excel dosyalarýnda **dosya numarasý bazlý** iþlemler yapmak için tasarlanmýþtýr. Bu API ile belirli bir dosya numarasýna ait tüm kayýtlarý bulabilir, düzenleyebilir ve yönetebilirsiniz.

## ?? Baþlangýç

### 1. Excel Dosyalarýnýzý Hazýrlayýn
Öncelikle Excel dosyalarýnýzý sistemde hazýr bulundurmanýz gerekiyor:

**Option A: Mevcut Excel API ile Upload:**
```http
POST /api/excel/upload
Content-Type: multipart/form-data

file: [Excel dosyanýz]
uploadedBy: "kullanici_adi"
```

**Option B: Dosyayý okutma:**
```http
POST /api/excel/read/{fileName}
```

### 2. Mevcut Dosyalarý Kontrol Edin
```http
GET /api/macro/available-files
```

## ?? API Endpoints

### ?? Dosya Numarasý Arama

#### Temel Arama
```http
GET /api/macro/search-by-document/{documentNumber}
```

**Örnek:**
```http
GET /api/macro/search-by-document/1010007261000100
```

#### Filtrelenmiþ Arama
```http
GET /api/macro/search-by-document/{documentNumber}?fileName=dosya.xlsx&sheetName=Sheet1
```

**Response:**
```json
{
  "success": true,
  "documentNumber": "1010007261000100",
  "totalRows": 15,
  "data": [
    {
      "id": 1,
      "fileName": "dosya1.xlsx",
      "sheetName": "Sheet1",
      "rowIndex": 1,
      "data": {
        "dosya_no": "1010007261000100",
        "ad": "Ahmet",
        "soyad": "Yýlmaz",
        "tutar": "1500"
      },
      "createdDate": "2024-01-01T10:00:00",
      "modifiedDate": null,
      "version": 1,
      "modifiedBy": null
    }
  ],
  "groupedData": {
    "dosya1.xlsx - Sheet1": [
      // ... veriler ...
    ]
  },
  "searchInfo": {
    "foundInFiles": ["dosya1.xlsx"],
    "foundInSheets": ["dosya1.xlsx - Sheet1"]
  }
}
```

### ?? Veri Güncelleme

#### Tekil Güncelleme
```http
PUT /api/macro/update-document-data
Content-Type: application/json

{
  "documentNumber": "1010007261000100",
  "rowId": 1,
  "updateData": {
    "ad": "Mehmet",
    "soyad": "Kaya",
    "tutar": "2000"
  },
  "updatedBy": "kullanici_adi"
}
```

#### Toplu Güncelleme
```http
PUT /api/macro/bulk-update-document
Content-Type: application/json

{
  "documentNumber": "1010007261000100",
  "updates": [
    {
      "rowId": 1,
      "updateData": {
        "ad": "Ali"
      }
    },
    {
      "rowId": 2,
      "updateData": {
        "soyad": "Demir"
      }
    }
  ],
  "updatedBy": "kullanici_adi"
}
```

### ?? Ýstatistikler

```http
GET /api/macro/document-statistics/{documentNumber}
```

**Response:**
```json
{
  "success": true,
  "documentNumber": "1010007261000100",
  "statistics": {
    "totalRows": 15,
    "filesCount": 2,
    "fileBreakdown": [
      { "fileName": "dosya1.xlsx", "count": 10 },
      { "fileName": "dosya2.xlsx", "count": 5 }
    ],
    "sheetBreakdown": [
      { "fileName": "dosya1.xlsx", "sheetName": "Sheet1", "count": 10 },
      { "fileName": "dosya2.xlsx", "sheetName": "Data", "count": 5 }
    ],
    "lastModified": {
      "modifiedDate": "2024-01-15T14:30:00",
      "modifiedBy": "kullanici_adi"
    }
  }
}
```

## ??? Kullaným Senaryolarý

### Senaryo 1: Dosya Numarasý Bazlý Düzenleme
1. **Arama:** `GET /api/macro/search-by-document/1010007261000100`
2. **Ýnceleme:** Response'taki `data` arrayinde tüm kayýtlarý görün
3. **Güncelleme:** Deðiþtirmek istediðiniz kayýt için `rowId` kullanarak `PUT /api/macro/update-document-data`

### Senaryo 2: Toplu Ýþlem
1. **Arama:** Dosya numarasýna ait tüm kayýtlarý getirin
2. **Planla:** Hangi `rowId`'lerde hangi deðiþiklikleri yapacaðýnýzý planlayýn
3. **Toplu Güncelle:** `PUT /api/macro/bulk-update-document` ile tek seferde tüm deðiþiklikleri yapýn

### Senaryo 3: Dosya Bazlý Filtreleme
Excel dosyalarýnýz farklý departmanlardan geliyorsa:
```http
GET /api/macro/search-by-document/1010007261000100?fileName=muhasebe.xlsx
GET /api/macro/search-by-document/1010007261000100?fileName=satis.xlsx
```

## ?? Önemli Notlar

1. **Dosya Numarasý Formatý:** Sistem string tabanlý arama yapar, tam eþleþme gerekli deðil
2. **Row ID:** Her güncelleme iþlemi için mutlaka `rowId` gereklidir
3. **Versiyon Kontrolü:** Her güncelleme `version` numarasýný artýrýr
4. **Eþzamanlýlýk:** Ayný anda birden fazla kullanýcý ayný satýrý güncelleyemez
5. **Soft Delete:** Silinen veriler `isDeleted=true` olarak iþaretlenir, fiziksel olarak silinmez

## ?? Hata Yönetimi

### Yaygýn Hatalar:
- **404:** Dosya numarasý bulunamadý
- **400:** Geçersiz `rowId` veya boþ `updateData`
- **409:** Eþzamanlýlýk çakýþmasý (baþka kullanýcý ayný anda güncelledi)
- **500:** Beklenmeyen sunucu hatasý

### Hata Response Örneði:
```json
{
  "success": false,
  "message": "Güncellenecek satýr bulunamadý",
  "rowId": 999
}
```

## ?? Test Adýmlarý

1. **API Test:** `GET /api/macro/test`
2. **Dosya Kontrolü:** `GET /api/macro/available-files`
3. **Veri Yükleme:** Excel dosyasý upload edin ve okutun
4. **Arama Testi:** `GET /api/macro/search-by-document/{test_document_number}`
5. **Güncelleme Testi:** Bulunan bir `rowId` ile güncelleme yapýn

---

Bu API ile dosya numarasý bazlý Excel veri yönetiminizi kolayca gerçekleþtirebilirsiniz! ??