using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Services;
using ExcelDataManagementAPI.Models.DTOs;
using System.Text.Json;

namespace ExcelDataManagementAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class MacroController : ControllerBase
    {
        private readonly IExcelService _excelService;
        private readonly ExcelDataContext _context;
        private readonly ILogger<MacroController> _logger;

        public MacroController(IExcelService excelService, ExcelDataContext context, ILogger<MacroController> logger)
        {
            _excelService = excelService;
            _context = context;
            _logger = logger;
        }

        [HttpGet("test")]
        public IActionResult Test()
        {
            return Ok(new {
                message = "Macro API �al���yor!",
                timestamp = DateTime.Now,
                availableOperations = new[]
                {
                    "H�zl� arama (sadece yeni dosyalarda): GET /api/macro/quick-search/{documentNumber}",
                    "Dosya numaras� arama (t�m dosyalarda): GET /api/macro/search-by-document/{documentNumber}",
                    "Dosya numaras� bazl� veri g�ncelleme: PUT /api/macro/update-document-data",
                    "Mevcut dosyalar� listeleme: GET /api/macro/available-files",
                    "Dosya numaras� istatistikleri: GET /api/macro/document-statistics/{documentNumber}",
                    "Dosya filtre durumu: GET /api/macro/file-filter-status"
                },
                description = "Excel dosyalar�n�z� upload ettikten sonra dosya numaras� bazl� i�lemler yapabilirsiniz",
                newFiles = new[]
                {
                    "gerceklesenhesap_20250905104743.xlsx",
                    "gerceklesenmakrodata_20250915153256.xlsx"
                },
                importantNote = "Sadece 'gerceklesenmakro' ve 'gerceklesenhesap' i�eren dosyalarda arama yap�l�r. GER�EKLE�EN dosyalar� hari� tutulur.",
                quickSearchNote = "H�zl� sonu� i�in /quick-search endpoint'ini kullan�n - sadece yeni dosyalarda arar"
            });
        }

        /// <summary>
        /// Mevcut Excel dosyalar�n� listeler - sadece gerceklesenmakro ve gerceklesenhesap dosyalar�n� g�sterir
        /// Yeni dosyalar �ncelikli g�sterilir
        /// </summary>
        [HttpGet("available-files")]
        public async Task<IActionResult> GetAvailableFiles()
        {
            try
            {
                var files = await _excelService.GetExcelFilesAsync();
                var fileStats = new List<object>();

                // �ncelikli dosya isimleri - yeni y�klenen dosyalar
                var priorityFiles = new[]
                {
                    "gerceklesenhesap_20250905104743.xlsx",
                    "gerceklesenmakrodata_20250915153256.xlsx"
                };

                // Sadece gerceklesenmakro ve gerceklesenhesap i�eren dosyalar� filtrele
                var filteredFiles = files.Where(f => f.IsActive &&
                    (f.FileName.ToLower().Contains("gerceklesenmakro") ||
                     f.FileName.ToLower().Contains("gerceklesenhesap")) &&
                    !f.FileName.ToUpper().Contains("GER�EKLE�EN"))
                    .OrderBy(f => priorityFiles.Contains(f.FileName) ? 0 : 1) // Yeni dosyalar �nce
                    .ThenBy(f => f.FileName);

                foreach (var file in filteredFiles)
                {
                    var dataCount = await _context.ExcelDataRows
                        .Where(r => r.FileName == file.FileName && !r.IsDeleted)
                        .CountAsync();

                    var sheets = await _context.ExcelDataRows
                        .Where(r => r.FileName == file.FileName && !r.IsDeleted)
                        .Select(r => r.SheetName)
                        .Distinct()
                        .ToListAsync();

                    var isNewFile = priorityFiles.Contains(file.FileName);

                    fileStats.Add(new
                    {
                        file.FileName,
                        file.OriginalFileName,
                        file.UploadDate,
                        DataRowCount = dataCount,
                        AvailableSheets = sheets,
                        ReadyForSearch = dataCount > 0,
                        IsNewPriorityFile = isNewFile,
                        Status = isNewFile ? "YEN� - �NCEL�KL�" : "Normal"
                    });
                }

                var newFiles = fileStats.Where(f => (bool)f.GetType().GetProperty("IsNewPriorityFile")?.GetValue(f, null)!).ToList();
                var otherFiles = fileStats.Where(f => !(bool)f.GetType().GetProperty("IsNewPriorityFile")?.GetValue(f, null)!).ToList();

                return Ok(new
                {
                    success = true,
                    data = fileStats,
                    totalFiles = fileStats.Count,
                    breakdown = new
                    {
                        newPriorityFiles = newFiles.Count,
                        otherFiles = otherFiles.Count
                    },
                    priorityFiles = priorityFiles,
                    message = "Macro ama�l� kullan�labilir Excel dosyalar� (yeni dosyalar �ncelikli)",
                    filteredOut = "GER�EKLE�EN dosyalar� hari� tutuldu",
                    quickSearchNote = "H�zl� arama i�in /quick-search endpoint'i sadece �ncelikli dosyalarda arar"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Macro dosyalar� listelenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Dosya numaras�na g�re verileri arar - SADECE gerceklesenmakro ve gerceklesenhesap dosyalar�nda
        /// </summary>
        [HttpGet("search-by-document/{documentNumber}")]
        public async Task<IActionResult> SearchByDocumentNumber(string documentNumber, [FromQuery] string? fileName = null, [FromQuery] string? sheetName = null)
        {
            try
            {
                documentNumber = Uri.UnescapeDataString(documentNumber);

                _logger.LogInformation("START: Dosya numaras� aran�yor: {DocumentNumber}, FileName: {FileName}, SheetName: {SheetName}",
                    documentNumber, fileName, sheetName);

                if (string.IsNullOrWhiteSpace(documentNumber))
                {
                    return BadRequest(new { success = false, message = "Dosya numaras� bo� olamaz" });
                }

                // Temel sorgu
                var query = _context.ExcelDataRows.Where(r => !r.IsDeleted);

                // KES�N F�LTRE: Sadece belirli dosya adlar�n� i�eren sat�rlar� al.
                // Bu, �nceki t�m karma��k filtrelerin yerini al�r.
                var targetFiles = new[] { "gerceklesenmakro", "gerceklesenhesap" };
                query = query.Where(r => targetFiles.Any(f => r.FileName.ToLower().Contains(f)));

                _logger.LogInformation("ADIM 1: Macro dosyalar� filtrelendi. Mevcut sat�r say�s�: {Count}", await query.CountAsync());

                // Dosya filtresi (iste�e ba�l�)
                if (!string.IsNullOrEmpty(fileName))
                {
                    query = query.Where(r => r.FileName == fileName);
                    _logger.LogInformation("ADIM 2: Belirtilen dosya ad� '{FileName}' filtrelendi. Mevcut sat�r say�s�: {Count}", fileName, await query.CountAsync());
                }

                // Sheet filtresi (iste�e ba�l�)
                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                    _logger.LogInformation("ADIM 3: Belirtilen sheet ad� '{SheetName}' filtrelendi. Mevcut sat�r say�s�: {Count}", sheetName, await query.CountAsync());
                }

                // Dosya numaras� aramas� - JSON verisi i�inde arama
                query = query.Where(r => r.RowData.Contains(documentNumber));
                _logger.LogInformation("ADIM 4: Dosya numaras� '{DocumentNumber}' arand�. Sonu� say�s�: {Count}", documentNumber, await query.CountAsync());


                var foundRows = await query
                    .OrderBy(r => r.FileName)
                    .ThenBy(r => r.SheetName)
                    .ThenBy(r => r.RowIndex)
                    .ToListAsync();

                if (!foundRows.Any())
                {
                    _logger.LogWarning("SONU�: '{DocumentNumber}' dosya numaras� i�in kay�t bulunamad�.", documentNumber);
                    return NotFound(new
                    {
                        success = false,
                        message = $"'{documentNumber}' dosya numaras� belirtilen macro dosyalar�nda bulunamad�.",
                        suggestion = "Dosya numaras�n� kontrol edin veya dosyalar�n veritaban�na do�ru y�klendi�inden emin olun. API sadece 'gerceklesenmakro' ve 'gerceklesenhesap' i�eren dosyalarda arama yapar.",
                        searchedIn = new
                        {
                            fileName = fileName ?? "T�m 'gerceklesenmakro' ve 'gerceklesenhesap' dosyalar�",
                            sheetName = sheetName ?? "T�m sheet'ler",
                            documentNumber = documentNumber
                        }
                    });
                }

                _logger.LogInformation("SONU�: '{DocumentNumber}' i�in {Count} kay�t bulundu.", documentNumber, foundRows.Count);

                // JSON verileri parse et ve response DTO'lar�na d�n��t�r
                var responseData = foundRows.Select(row =>
                {
                    var rowData = JsonSerializer.Deserialize<Dictionary<string, string>>(row.RowData) ?? new Dictionary<string, string>();

                    return new ExcelDataResponseDto
                    {
                        Id = row.Id,
                        FileName = row.FileName,
                        SheetName = row.SheetName,
                        RowIndex = row.RowIndex,
                        Data = rowData,
                        CreatedDate = row.CreatedDate,
                        ModifiedDate = row.ModifiedDate,
                        Version = row.Version,
                        ModifiedBy = row.ModifiedBy
                    };
                }).ToList();

                return Ok(new
                {
                    success = true,
                    documentNumber = documentNumber,
                    totalRows = responseData.Count,
                    data = responseData,
                    message = $"'{documentNumber}' dosya numaras� i�in {responseData.Count} kay�t bulundu."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya numaras� aran�rken kritik hata: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = $"Sunucu hatas�: {ex.Message}" });
            }
        }

        /// <summary>
        /// Dosya numaras�na g�re veri g�nceller - sadece macro dosyalar�nda
        /// </summary>
        [HttpPut("update-document-data")]
        public async Task<IActionResult> UpdateDocumentData([FromBody] MacroUpdateRequestDto request)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.DocumentNumber))
                {
                    return BadRequest(new { success = false, message = "Dosya numaras� gerekli" });
                }

                if (request.RowId <= 0)
                {
                    return BadRequest(new { success = false, message = "Ge�erli bir Row ID gerekli" });
                }

                if (request.UpdateData == null || !request.UpdateData.Any())
                {
                    return BadRequest(new { success = false, message = "G�ncellenecek veri gerekli" });
                }

                _logger.LogInformation("Macro g�ncelleme: DocumentNumber={DocumentNumber}, RowId={RowId}, UpdatedBy={UpdatedBy}",
                    request.DocumentNumber, request.RowId, request.UpdatedBy);

                // �nce ilgili sat�r� bul ve do�rula - sadece macro dosyalar�nda
                var targetRow = await _context.ExcelDataRows
                    .FirstOrDefaultAsync(r => r.Id == request.RowId && !r.IsDeleted &&
                        (r.FileName.ToLower().Contains("gerceklesenmakro") ||
                         r.FileName.ToLower().Contains("gerceklesenhesap")) &&
                        !r.FileName.ToUpper().Contains("GER�EKLE�EN"));

                if (targetRow == null)
                {
                    return NotFound(new
                    {
                        success = false,
                        message = "G�ncellenecek sat�r macro dosyalar�nda bulunamad�",
                        rowId = request.RowId,
                        searchScope = "Sadece gerceklesenmakro ve gerceklesenhesap dosyalar�",
                        excludedScope = "GER�EKLE�EN dosyalar� hari�"
                    });
                }

                // Dosya numaras�n�n bu sat�rda olup olmad���n� kontrol et
                if (!targetRow.RowData.Contains(request.DocumentNumber))
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "Belirtilen sat�r bu dosya numaras�na ait de�il",
                        documentNumber = request.DocumentNumber,
                        rowId = request.RowId,
                        rowFileName = targetRow.FileName,
                        rowSheetName = targetRow.SheetName
                    });
                }

                // G�ncelleme DTO'su olu�tur
                var updateDto = new ExcelDataUpdateDto
                {
                    Id = request.RowId,
                    Data = request.UpdateData,
                    ModifiedBy = request.UpdatedBy
                };

                // Mevcut ExcelService kullanarak g�ncelle
                var result = await _excelService.UpdateExcelDataAsync(updateDto, HttpContext);

                return Ok(new
                {
                    success = true,
                    data = result,
                    documentNumber = request.DocumentNumber,
                    updatedFields = request.UpdateData.Keys.ToArray(),
                    message = "Macro dosyas�ndaki veri ba�ar�yla g�ncellendi",
                    version = result.Version,
                    modifiedDate = result.ModifiedDate,
                    updatedInFile = result.FileName
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Macro g�ncelleme hatas�: DocumentNumber={DocumentNumber}, RowId={RowId}",
                    request.DocumentNumber, request.RowId);

                if (ex.Message.Contains("bulunamad�"))
                {
                    return NotFound(new { success = false, message = ex.Message });
                }
                else if (ex.Message.Contains("e�zamanl�l�k"))
                {
                    return Conflict(new { success = false, message = ex.Message });
                }

                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet("document-statistics/{documentNumber}")]
        public async Task<IActionResult> GetDocumentStatistics(string documentNumber, [FromQuery] string? fileName = null)
        {
            try
            {
                documentNumber = Uri.UnescapeDataString(documentNumber);

                var query = _context.ExcelDataRows.Where(r => !r.IsDeleted && r.RowData.Contains(documentNumber));

                // Sadece gerceklesenmakro ve gerceklesenhesap dosyalar�nda arama yap
                // GER�EKLE�EN dosyalar�n� hari� tut
                query = query.Where(r =>
                    (r.FileName.ToLower().Contains("gerceklesenmakro") ||
                     r.FileName.ToLower().Contains("gerceklesenhesap")) &&
                    !r.FileName.ToUpper().Contains("GER�EKLE�EN"));

                if (!string.IsNullOrEmpty(fileName))
                {
                    query = query.Where(r => r.FileName == fileName);
                }

                var totalRows = await query.CountAsync();

                if (totalRows == 0)
                {
                    return NotFound(new
                    {
                        success = false,
                        message = $"'{documentNumber}' dosya numaras� macro dosyalar�nda bulunamad�",
                        documentNumber = documentNumber,
                        searchScope = "Sadece gerceklesenmakro ve gerceklesenhesap dosyalar�",
                        excludedScope = "GER�EKLE�EN dosyalar� hari�"
                    });
                }

                var fileGroups = await query
                    .GroupBy(r => r.FileName)
                    .Select(g => new { FileName = g.Key, Count = g.Count() })
                    .ToListAsync();

                var sheetGroups = await query
                    .GroupBy(r => new { r.FileName, r.SheetName })
                    .Select(g => new {
                        FileName = g.Key.FileName,
                        SheetName = g.Key.SheetName,
                        Count = g.Count()
                    })
                    .ToListAsync();

                var lastModified = await query
                    .Where(r => r.ModifiedDate.HasValue)
                    .OrderByDescending(r => r.ModifiedDate)
                    .Select(r => new { r.ModifiedDate, r.ModifiedBy })
                    .FirstOrDefaultAsync();

                return Ok(new
                {
                    success = true,
                    documentNumber = documentNumber,
                    statistics = new
                    {
                        totalRows = totalRows,
                        filesCount = fileGroups.Count,
                        fileBreakdown = fileGroups,
                        sheetBreakdown = sheetGroups,
                        lastModified = lastModified
                    },
                    searchScope = "Sadece gerceklesenmakro ve gerceklesenhesap dosyalar�",
                    excludedScope = "GER�EKLE�EN dosyalar� hari�",
                    message = $"'{documentNumber}' dosya numaras� i�in macro dosyalar� istatistikleri"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya numaras� istatistikleri al�n�rken hata: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPut("bulk-update-document")]
        public async Task<IActionResult> BulkUpdateDocument([FromBody] MacroBulkUpdateRequestDto request)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.DocumentNumber))
                {
                    return BadRequest(new { success = false, message = "Dosya numaras� gerekli" });
                }

                if (request.Updates == null || !request.Updates.Any())
                {
                    return BadRequest(new { success = false, message = "G�ncellenecek veri listesi gerekli" });
                }

                _logger.LogInformation("Macro toplu g�ncelleme: DocumentNumber={DocumentNumber}, UpdateCount={UpdateCount}",
                    request.DocumentNumber, request.Updates.Count);

                var results = new List<ExcelDataResponseDto>();
                var errors = new List<string>();

                foreach (var update in request.Updates)
                {
                    try
                    {
                        // Her g�ncelleme i�in ayr� ayr� kontrol - sadece macro dosyalar�nda
                        var targetRow = await _context.ExcelDataRows
                            .FirstOrDefaultAsync(r => r.Id == update.RowId && !r.IsDeleted &&
                                (r.FileName.ToLower().Contains("gerceklesenmakro") ||
                                 r.FileName.ToLower().Contains("gerceklesenhesap")) &&
                                !r.FileName.ToUpper().Contains("GER�EKLE�EN"));

                        if (targetRow == null)
                        {
                            errors.Add($"Row ID {update.RowId}: Sat�r macro dosyalar�nda bulunamad�");
                            continue;
                        }

                        if (!targetRow.RowData.Contains(request.DocumentNumber))
                        {
                            errors.Add($"Row ID {update.RowId}: Bu dosya numaras�na ait de�il");
                            continue;
                        }

                        var updateDto = new ExcelDataUpdateDto
                        {
                            Id = update.RowId,
                            Data = update.UpdateData,
                            ModifiedBy = request.UpdatedBy
                        };

                        var result = await _excelService.UpdateExcelDataAsync(updateDto, HttpContext);
                        results.Add(result);
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Row ID {update.RowId}: {ex.Message}");
                    }
                }

                return Ok(new
                {
                    success = true,
                    documentNumber = request.DocumentNumber,
                    successfulUpdates = results.Count,
                    totalRequested = request.Updates.Count,
                    data = results,
                    errors = errors,
                    searchScope = "Sadece gerceklesenmakro ve gerceklesenhesap dosyalar�",
                    excludedScope = "GER�EKLE�EN dosyalar� hari�",
                    message = $"Macro dosyalar�nda {results.Count}/{request.Updates.Count} g�ncelleme ba�ar�l�"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Macro toplu g�ncelleme hatas�: DocumentNumber={DocumentNumber}", request.DocumentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Macro API'sinin dosya filtre durumunu g�sterir
        /// </summary>
        [HttpGet("file-filter-status")]
        public async Task<IActionResult> GetFileFilterStatus()
        {
            try
            {
                var allFiles = await _excelService.GetExcelFilesAsync();

                var macroFiles = allFiles.Where(f => f.IsActive &&
                    (f.FileName.ToLower().Contains("gerceklesenmakro") ||
                     f.FileName.ToLower().Contains("gerceklesenhesap")) &&
                    !f.FileName.ToUpper().Contains("GER�EKLE�EN")).ToList();

                var excludedFiles = allFiles.Where(f => f.IsActive &&
                    (f.FileName.ToUpper().Contains("GER�EKLE�EN") ||
                     (!f.FileName.ToLower().Contains("gerceklesenmakro") &&
                      !f.FileName.ToLower().Contains("gerceklesenhesap")))).ToList();

                return Ok(new
                {
                    success = true,
                    filterCriteria = new
                    {
                        included = "Dosya ad�nda 'gerceklesenmakro' veya 'gerceklesenhesap' ge�en dosyalar",
                        excluded = "Dosya ad�nda 'GER�EKLE�EN' ge�en dosyalar veya macro dosyas� olmayan dosyalar"
                    },
                    macroFiles = macroFiles.Select(f => new
                    {
                        f.FileName,
                        f.OriginalFileName,
                        f.UploadDate,
                        status = "Macro i�lemlerde KULLANILIR"
                    }),
                    excludedFiles = excludedFiles.Select(f => new
                    {
                        f.FileName,
                        f.OriginalFileName,
                        f.UploadDate,
                        status = "Macro i�lemlerde KULLANILMAZ"
                    }),
                    counts = new
                    {
                        totalFiles = allFiles.Count(f => f.IsActive),
                        macroFilesCount = macroFiles.Count,
                        excludedFilesCount = excludedFiles.Count
                    },
                    message = "Macro API dosya filtre durumu"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya filtre durumu al�n�rken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// H�zl� arama - SADECE belirtilen yeni dosyalarda arar
        /// </summary>
        [HttpGet("quick-search/{documentNumber}")]
        public async Task<IActionResult> QuickSearchInNewFiles(string documentNumber, [FromQuery] string? sheetName = null)
        {
            try
            {
                documentNumber = Uri.UnescapeDataString(documentNumber);

                // Sadece yeni ve belirli dosyalar
                var newFiles = new[]
                {
                    "gerceklesenmakrodata_20250915153256.xlsx",
                    "gerceklesenhesap_20250905104743.xlsx"
                };

                var query = _context.ExcelDataRows.Where(r => !r.IsDeleted &&
                    newFiles.Contains(r.FileName) &&
                    r.RowData.Contains(documentNumber));

                // Sheet filtresi (e�er belirtilmi�se)
                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                }

                var foundRows = await query
                    .OrderBy(r => r.FileName)
                    .ThenBy(r => r.SheetName)
                    .ThenBy(r => r.RowIndex)
                    .ToListAsync();

                if (!foundRows.Any())
                {
                    return NotFound(new
                    {
                        success = false,
                        message = $"'{documentNumber}' dosya numaras� belirtilen yeni dosyalarda bulunamad�.",
                        searchedFiles = newFiles,
                        suggestion = "T�m dosyalarda arama yapmak i�in /api/macro/search-by-document/{documentNumber} kullan�n."
                    });
                }

                // JSON verileri parse et ve response DTO'lar�na d�n��t�r
                var responseData = foundRows.Select(row =>
                {
                    var rowData = JsonSerializer.Deserialize<Dictionary<string, string>>(row.RowData) ?? new Dictionary<string, string>();

                    return new ExcelDataResponseDto
                    {
                        Id = row.Id,
                        FileName = row.FileName,
                        SheetName = row.SheetName,
                        RowIndex = row.RowIndex,
                        Data = rowData,
                        CreatedDate = row.CreatedDate,
                        ModifiedDate = row.ModifiedDate,
                        Version = row.Version,
                        ModifiedBy = row.ModifiedBy
                    };
                }).ToList();

                return Ok(new
                {
                    success = true,
                    documentNumber = documentNumber,
                    totalRows = responseData.Count,
                    data = responseData,
                    searchedFiles = newFiles,
                    message = $"'{documentNumber}' i�in yeni dosyalarda {responseData.Count} kay�t bulundu."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "H�zl� arama hatas�: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// H�zl� arama - SADECE belirtilen makro dosyas�nda arar
        /// </summary>
        [HttpGet("quick-search-makro/{documentNumber}")]
        public async Task<IActionResult> QuickSearchMakroOnly(string documentNumber, [FromQuery] string? sheetName = null)
        {
            try
            {
                documentNumber = Uri.UnescapeDataString(documentNumber);

                // Sadece makro dosyas�nda arama yap
                var newFiles = new[]
                {
                    "gerceklesenmakrodata_20250915153256.xlsx"
                };

                var query = _context.ExcelDataRows.Where(r => !r.IsDeleted &&
                    newFiles.Contains(r.FileName) &&
                    r.RowData.Contains(documentNumber));

                // Sheet filtresi (e�er belirtilmi�se)
                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                }

                var foundRows = await query
                    .OrderBy(r => r.FileName)
                    .ThenBy(r => r.SheetName)
                    .ThenBy(r => r.RowIndex)
                    .ToListAsync();

                if (!foundRows.Any())
                {
                    return NotFound(new
                    {
                        success = false,
                        message = $"'{documentNumber}' dosya numaras� belirtilen makro dosyas�nda bulunamad�.",
                        searchedFiles = newFiles,
                        suggestion = "T�m dosyalarda arama yapmak i�in /api/macro/search-by-document/{documentNumber} kullan�n."
                    });
                }

                // JSON verileri parse et ve response DTO'lar�na d�n��t�r
                var responseData = foundRows.Select(row =>
                {
                    var rowData = JsonSerializer.Deserialize<Dictionary<string, string>>(row.RowData) ?? new Dictionary<string, string>();

                    return new ExcelDataResponseDto
                    {
                        Id = row.Id,
                        FileName = row.FileName,
                        SheetName = row.SheetName,
                        RowIndex = row.RowIndex,
                        Data = rowData,
                        CreatedDate = row.CreatedDate,
                        ModifiedDate = row.ModifiedDate,
                        Version = row.Version,
                        ModifiedBy = row.ModifiedBy
                    };
                }).ToList();

                return Ok(new
                {
                    success = true,
                    documentNumber = documentNumber,
                    totalRows = responseData.Count,
                    data = responseData,
                    searchedFiles = newFiles,
                    message = $"'{documentNumber}' i�in makro dosyas�nda {responseData.Count} kay�t bulundu."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "H�zl� arama hatas�: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Dosya numaras�na g�re verileri arar - sadece 'gerceklesenhesap_20250905104743.xlsx' dosyas�nda arama yapar
        /// </summary>
        [HttpGet("search-in-hesap/{documentNumber}")]
        public async Task<IActionResult> SearchInHesapFile(string documentNumber, [FromQuery] string? sheetName = null)
        {
            try
            {
                documentNumber = Uri.UnescapeDataString(documentNumber);
                var targetFile = "gerceklesenhesap_20250905104743.xlsx";

                var query = _context.ExcelDataRows.Where(r => !r.IsDeleted &&
                    r.FileName == targetFile &&
                    r.RowData.Contains(documentNumber));

                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                }

                var foundRows = await query
                    .OrderBy(r => r.SheetName)
                    .ThenBy(r => r.RowIndex)
                    .ToListAsync();

                if (!foundRows.Any())
                {
                    return NotFound(new
                    {
                        success = false,
                        message = $"'{documentNumber}' dosya numaras� {targetFile} dosyas�nda bulunamad�.",
                        searchedFile = targetFile
                    });
                }

                var responseData = foundRows.Select(row =>
                {
                    var rowData = JsonSerializer.Deserialize<Dictionary<string, string>>(row.RowData) ?? new Dictionary<string, string>();
                    return new ExcelDataResponseDto
                    {
                        Id = row.Id,
                        FileName = row.FileName,
                        SheetName = row.SheetName,
                        RowIndex = row.RowIndex,
                        Data = rowData,
                        CreatedDate = row.CreatedDate,
                        ModifiedDate = row.ModifiedDate,
                        Version = row.Version,
                        ModifiedBy = row.ModifiedBy
                    };
                }).ToList();

                return Ok(new
                {
                    success = true,
                    documentNumber = documentNumber,
                    totalRows = responseData.Count,
                    data = responseData,
                    searchedFile = targetFile
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Hesap dosyas�nda arama hatas�: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}