# MacroController API Dok�mantasyonu

MacroController, Excel dosyalar�nda **dosya numaras� bazl�** i�lemler yapmak i�in tasarlanm��t�r. Bu API ile belirli bir dosya numaras�na ait t�m kay�tlar� bulabilir, d�zenleyebilir ve y�netebilirsiniz.

## ?? Ba�lang��

### 1. Excel Dosyalar�n�z� Haz�rlay�n
�ncelikle Excel dosyalar�n�z� sistemde haz�r bulundurman�z gerekiyor:

**Option A: Mevcut Excel API ile Upload:**
```http
POST /api/excel/upload
Content-Type: multipart/form-data

file: [Excel dosyan�z]
uploadedBy: "kullanici_adi"
```

**Option B: Dosyay� okutma:**
```http
POST /api/excel/read/{fileName}
```

### 2. Mevcut Dosyalar� Kontrol Edin
```http
GET /api/macro/available-files
```

## ?? API Endpoints

### ?? Dosya Numaras� Arama

#### Temel Arama
```http
GET /api/macro/search-by-document/{documentNumber}
```

**�rnek:**
```http
GET /api/macro/search-by-document/1010007261000100
```

#### Filtrelenmi� Arama
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
        "soyad": "Y�lmaz",
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

### ?? Veri G�ncelleme

#### Tekil G�ncelleme
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

#### Toplu G�ncelleme
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

### ?? �statistikler

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

## ??? Kullan�m Senaryolar�

### Senaryo 1: Dosya Numaras� Bazl� D�zenleme
1. **Arama:** `GET /api/macro/search-by-document/1010007261000100`
2. **�nceleme:** Response'taki `data` arrayinde t�m kay�tlar� g�r�n
3. **G�ncelleme:** De�i�tirmek istedi�iniz kay�t i�in `rowId` kullanarak `PUT /api/macro/update-document-data`

### Senaryo 2: Toplu ��lem
1. **Arama:** Dosya numaras�na ait t�m kay�tlar� getirin
2. **Planla:** Hangi `rowId`'lerde hangi de�i�iklikleri yapaca��n�z� planlay�n
3. **Toplu G�ncelle:** `PUT /api/macro/bulk-update-document` ile tek seferde t�m de�i�iklikleri yap�n

### Senaryo 3: Dosya Bazl� Filtreleme
Excel dosyalar�n�z farkl� departmanlardan geliyorsa:
```http
GET /api/macro/search-by-document/1010007261000100?fileName=muhasebe.xlsx
GET /api/macro/search-by-document/1010007261000100?fileName=satis.xlsx
```

## ?? �nemli Notlar

1. **Dosya Numaras� Format�:** Sistem string tabanl� arama yapar, tam e�le�me gerekli de�il
2. **Row ID:** Her g�ncelleme i�lemi i�in mutlaka `rowId` gereklidir
3. **Versiyon Kontrol�:** Her g�ncelleme `version` numaras�n� art�r�r
4. **E�zamanl�l�k:** Ayn� anda birden fazla kullan�c� ayn� sat�r� g�ncelleyemez
5. **Soft Delete:** Silinen veriler `isDeleted=true` olarak i�aretlenir, fiziksel olarak silinmez

## ?? Hata Y�netimi

### Yayg�n Hatalar:
- **404:** Dosya numaras� bulunamad�
- **400:** Ge�ersiz `rowId` veya bo� `updateData`
- **409:** E�zamanl�l�k �ak��mas� (ba�ka kullan�c� ayn� anda g�ncelledi)
- **500:** Beklenmeyen sunucu hatas�

### Hata Response �rne�i:
```json
{
  "success": false,
  "message": "G�ncellenecek sat�r bulunamad�",
  "rowId": 999
}
```

## ?? Test Ad�mlar�

1. **API Test:** `GET /api/macro/test`
2. **Dosya Kontrol�:** `GET /api/macro/available-files`
3. **Veri Y�kleme:** Excel dosyas� upload edin ve okutun
4. **Arama Testi:** `GET /api/macro/search-by-document/{test_document_number}`
5. **G�ncelleme Testi:** Bulunan bir `rowId` ile g�ncelleme yap�n

---

Bu API ile dosya numaras� bazl� Excel veri y�netiminizi kolayca ger�ekle�tirebilirsiniz! ??