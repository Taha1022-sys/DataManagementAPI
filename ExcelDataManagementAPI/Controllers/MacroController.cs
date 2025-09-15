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
                message = "Macro API çalýþýyor!",
                timestamp = DateTime.Now,
                availableOperations = new[]
                {
                    "Hýzlý arama (sadece yeni dosyalarda): GET /api/macro/quick-search/{documentNumber}",
                    "Dosya numarasý arama (tüm dosyalarda): GET /api/macro/search-by-document/{documentNumber}",
                    "Dosya numarasý bazlý veri güncelleme: PUT /api/macro/update-document-data",
                    "Mevcut dosyalarý listeleme: GET /api/macro/available-files",
                    "Dosya numarasý istatistikleri: GET /api/macro/document-statistics/{documentNumber}",
                    "Dosya filtre durumu: GET /api/macro/file-filter-status"
                },
                description = "Excel dosyalarýnýzý upload ettikten sonra dosya numarasý bazlý iþlemler yapabilirsiniz",
                newFiles = new[]
                {
                    "gerceklesenhesap_20250905104743.xlsx",
                    "gerceklesenmakrodata_20250915153256.xlsx"
                },
                importantNote = "Sadece 'gerceklesenmakro' ve 'gerceklesenhesap' içeren dosyalarda arama yapýlýr. GERÇEKLEÞEN dosyalarý hariç tutulur.",
                quickSearchNote = "Hýzlý sonuç için /quick-search endpoint'ini kullanýn - sadece yeni dosyalarda arar"
            });
        }

        /// <summary>
        /// Mevcut Excel dosyalarýný listeler - sadece gerceklesenmakro ve gerceklesenhesap dosyalarýný gösterir
        /// Yeni dosyalar öncelikli gösterilir
        /// </summary>
        [HttpGet("available-files")]
        public async Task<IActionResult> GetAvailableFiles()
        {
            try
            {
                var files = await _excelService.GetExcelFilesAsync();
                var fileStats = new List<object>();

                // Öncelikli dosya isimleri - yeni yüklenen dosyalar
                var priorityFiles = new[]
                {
                    "gerceklesenhesap_20250905104743.xlsx",
                    "gerceklesenmakrodata_20250915153256.xlsx"
                };

                // Sadece gerceklesenmakro ve gerceklesenhesap içeren dosyalarý filtrele
                var filteredFiles = files.Where(f => f.IsActive &&
                    (f.FileName.ToLower().Contains("gerceklesenmakro") ||
                     f.FileName.ToLower().Contains("gerceklesenhesap")) &&
                    !f.FileName.ToUpper().Contains("GERÇEKLEÞEN"))
                    .OrderBy(f => priorityFiles.Contains(f.FileName) ? 0 : 1) // Yeni dosyalar önce
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
                        Status = isNewFile ? "YENÝ - ÖNCELÝKLÝ" : "Normal"
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
                    message = "Macro amaçlý kullanýlabilir Excel dosyalarý (yeni dosyalar öncelikli)",
                    filteredOut = "GERÇEKLEÞEN dosyalarý hariç tutuldu",
                    quickSearchNote = "Hýzlý arama için /quick-search endpoint'i sadece öncelikli dosyalarda arar"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Macro dosyalarý listelenirken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Dosya numarasýna göre verileri arar - SADECE gerceklesenmakro ve gerceklesenhesap dosyalarýnda
        /// </summary>
        [HttpGet("search-by-document/{documentNumber}")]
        public async Task<IActionResult> SearchByDocumentNumber(string documentNumber, [FromQuery] string? fileName = null, [FromQuery] string? sheetName = null)
        {
            try
            {
                documentNumber = Uri.UnescapeDataString(documentNumber);

                _logger.LogInformation("START: Dosya numarasý aranýyor: {DocumentNumber}, FileName: {FileName}, SheetName: {SheetName}",
                    documentNumber, fileName, sheetName);

                if (string.IsNullOrWhiteSpace(documentNumber))
                {
                    return BadRequest(new { success = false, message = "Dosya numarasý boþ olamaz" });
                }

                // Temel sorgu
                var query = _context.ExcelDataRows.Where(r => !r.IsDeleted);

                // KESÝN FÝLTRE: Sadece belirli dosya adlarýný içeren satýrlarý al.
                // Bu, önceki tüm karmaþýk filtrelerin yerini alýr.
                var targetFiles = new[] { "gerceklesenmakro", "gerceklesenhesap" };
                query = query.Where(r => targetFiles.Any(f => r.FileName.ToLower().Contains(f)));

                _logger.LogInformation("ADIM 1: Macro dosyalarý filtrelendi. Mevcut satýr sayýsý: {Count}", await query.CountAsync());

                // Dosya filtresi (isteðe baðlý)
                if (!string.IsNullOrEmpty(fileName))
                {
                    query = query.Where(r => r.FileName == fileName);
                    _logger.LogInformation("ADIM 2: Belirtilen dosya adý '{FileName}' filtrelendi. Mevcut satýr sayýsý: {Count}", fileName, await query.CountAsync());
                }

                // Sheet filtresi (isteðe baðlý)
                if (!string.IsNullOrEmpty(sheetName))
                {
                    query = query.Where(r => r.SheetName == sheetName);
                    _logger.LogInformation("ADIM 3: Belirtilen sheet adý '{SheetName}' filtrelendi. Mevcut satýr sayýsý: {Count}", sheetName, await query.CountAsync());
                }

                // Dosya numarasý aramasý - JSON verisi içinde arama
                query = query.Where(r => r.RowData.Contains(documentNumber));
                _logger.LogInformation("ADIM 4: Dosya numarasý '{DocumentNumber}' arandý. Sonuç sayýsý: {Count}", documentNumber, await query.CountAsync());


                var foundRows = await query
                    .OrderBy(r => r.FileName)
                    .ThenBy(r => r.SheetName)
                    .ThenBy(r => r.RowIndex)
                    .ToListAsync();

                if (!foundRows.Any())
                {
                    _logger.LogWarning("SONUÇ: '{DocumentNumber}' dosya numarasý için kayýt bulunamadý.", documentNumber);
                    return NotFound(new
                    {
                        success = false,
                        message = $"'{documentNumber}' dosya numarasý belirtilen macro dosyalarýnda bulunamadý.",
                        suggestion = "Dosya numarasýný kontrol edin veya dosyalarýn veritabanýna doðru yüklendiðinden emin olun. API sadece 'gerceklesenmakro' ve 'gerceklesenhesap' içeren dosyalarda arama yapar.",
                        searchedIn = new
                        {
                            fileName = fileName ?? "Tüm 'gerceklesenmakro' ve 'gerceklesenhesap' dosyalarý",
                            sheetName = sheetName ?? "Tüm sheet'ler",
                            documentNumber = documentNumber
                        }
                    });
                }

                _logger.LogInformation("SONUÇ: '{DocumentNumber}' için {Count} kayýt bulundu.", documentNumber, foundRows.Count);

                // JSON verileri parse et ve response DTO'larýna dönüþtür
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
                    message = $"'{documentNumber}' dosya numarasý için {responseData.Count} kayýt bulundu."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya numarasý aranýrken kritik hata: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = $"Sunucu hatasý: {ex.Message}" });
            }
        }

        /// <summary>
        /// Dosya numarasýna göre veri günceller - sadece macro dosyalarýnda
        /// </summary>
        [HttpPut("update-document-data")]
        public async Task<IActionResult> UpdateDocumentData([FromBody] MacroUpdateRequestDto request)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.DocumentNumber))
                {
                    return BadRequest(new { success = false, message = "Dosya numarasý gerekli" });
                }

                if (request.RowId <= 0)
                {
                    return BadRequest(new { success = false, message = "Geçerli bir Row ID gerekli" });
                }

                if (request.UpdateData == null || !request.UpdateData.Any())
                {
                    return BadRequest(new { success = false, message = "Güncellenecek veri gerekli" });
                }

                _logger.LogInformation("Macro güncelleme: DocumentNumber={DocumentNumber}, RowId={RowId}, UpdatedBy={UpdatedBy}",
                    request.DocumentNumber, request.RowId, request.UpdatedBy);

                // Önce ilgili satýrý bul ve doðrula - sadece macro dosyalarýnda
                var targetRow = await _context.ExcelDataRows
                    .FirstOrDefaultAsync(r => r.Id == request.RowId && !r.IsDeleted &&
                        (r.FileName.ToLower().Contains("gerceklesenmakro") ||
                         r.FileName.ToLower().Contains("gerceklesenhesap")) &&
                        !r.FileName.ToUpper().Contains("GERÇEKLEÞEN"));

                if (targetRow == null)
                {
                    return NotFound(new
                    {
                        success = false,
                        message = "Güncellenecek satýr macro dosyalarýnda bulunamadý",
                        rowId = request.RowId,
                        searchScope = "Sadece gerceklesenmakro ve gerceklesenhesap dosyalarý",
                        excludedScope = "GERÇEKLEÞEN dosyalarý hariç"
                    });
                }

                // Dosya numarasýnýn bu satýrda olup olmadýðýný kontrol et
                if (!targetRow.RowData.Contains(request.DocumentNumber))
                {
                    return BadRequest(new
                    {
                        success = false,
                        message = "Belirtilen satýr bu dosya numarasýna ait deðil",
                        documentNumber = request.DocumentNumber,
                        rowId = request.RowId,
                        rowFileName = targetRow.FileName,
                        rowSheetName = targetRow.SheetName
                    });
                }

                // Güncelleme DTO'su oluþtur
                var updateDto = new ExcelDataUpdateDto
                {
                    Id = request.RowId,
                    Data = request.UpdateData,
                    ModifiedBy = request.UpdatedBy
                };

                // Mevcut ExcelService kullanarak güncelle
                var result = await _excelService.UpdateExcelDataAsync(updateDto, HttpContext);

                return Ok(new
                {
                    success = true,
                    data = result,
                    documentNumber = request.DocumentNumber,
                    updatedFields = request.UpdateData.Keys.ToArray(),
                    message = "Macro dosyasýndaki veri baþarýyla güncellendi",
                    version = result.Version,
                    modifiedDate = result.ModifiedDate,
                    updatedInFile = result.FileName
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Macro güncelleme hatasý: DocumentNumber={DocumentNumber}, RowId={RowId}",
                    request.DocumentNumber, request.RowId);

                if (ex.Message.Contains("bulunamadý"))
                {
                    return NotFound(new { success = false, message = ex.Message });
                }
                else if (ex.Message.Contains("eþzamanlýlýk"))
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

                // Sadece gerceklesenmakro ve gerceklesenhesap dosyalarýnda arama yap
                // GERÇEKLEÞEN dosyalarýný hariç tut
                query = query.Where(r =>
                    (r.FileName.ToLower().Contains("gerceklesenmakro") ||
                     r.FileName.ToLower().Contains("gerceklesenhesap")) &&
                    !r.FileName.ToUpper().Contains("GERÇEKLEÞEN"));

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
                        message = $"'{documentNumber}' dosya numarasý macro dosyalarýnda bulunamadý",
                        documentNumber = documentNumber,
                        searchScope = "Sadece gerceklesenmakro ve gerceklesenhesap dosyalarý",
                        excludedScope = "GERÇEKLEÞEN dosyalarý hariç"
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
                    searchScope = "Sadece gerceklesenmakro ve gerceklesenhesap dosyalarý",
                    excludedScope = "GERÇEKLEÞEN dosyalarý hariç",
                    message = $"'{documentNumber}' dosya numarasý için macro dosyalarý istatistikleri"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Dosya numarasý istatistikleri alýnýrken hata: {DocumentNumber}", documentNumber);
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
                    return BadRequest(new { success = false, message = "Dosya numarasý gerekli" });
                }

                if (request.Updates == null || !request.Updates.Any())
                {
                    return BadRequest(new { success = false, message = "Güncellenecek veri listesi gerekli" });
                }

                _logger.LogInformation("Macro toplu güncelleme: DocumentNumber={DocumentNumber}, UpdateCount={UpdateCount}",
                    request.DocumentNumber, request.Updates.Count);

                var results = new List<ExcelDataResponseDto>();
                var errors = new List<string>();

                foreach (var update in request.Updates)
                {
                    try
                    {
                        // Her güncelleme için ayrý ayrý kontrol - sadece macro dosyalarýnda
                        var targetRow = await _context.ExcelDataRows
                            .FirstOrDefaultAsync(r => r.Id == update.RowId && !r.IsDeleted &&
                                (r.FileName.ToLower().Contains("gerceklesenmakro") ||
                                 r.FileName.ToLower().Contains("gerceklesenhesap")) &&
                                !r.FileName.ToUpper().Contains("GERÇEKLEÞEN"));

                        if (targetRow == null)
                        {
                            errors.Add($"Row ID {update.RowId}: Satýr macro dosyalarýnda bulunamadý");
                            continue;
                        }

                        if (!targetRow.RowData.Contains(request.DocumentNumber))
                        {
                            errors.Add($"Row ID {update.RowId}: Bu dosya numarasýna ait deðil");
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
                    searchScope = "Sadece gerceklesenmakro ve gerceklesenhesap dosyalarý",
                    excludedScope = "GERÇEKLEÞEN dosyalarý hariç",
                    message = $"Macro dosyalarýnda {results.Count}/{request.Updates.Count} güncelleme baþarýlý"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Macro toplu güncelleme hatasý: DocumentNumber={DocumentNumber}", request.DocumentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Macro API'sinin dosya filtre durumunu gösterir
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
                    !f.FileName.ToUpper().Contains("GERÇEKLEÞEN")).ToList();

                var excludedFiles = allFiles.Where(f => f.IsActive &&
                    (f.FileName.ToUpper().Contains("GERÇEKLEÞEN") ||
                     (!f.FileName.ToLower().Contains("gerceklesenmakro") &&
                      !f.FileName.ToLower().Contains("gerceklesenhesap")))).ToList();

                return Ok(new
                {
                    success = true,
                    filterCriteria = new
                    {
                        included = "Dosya adýnda 'gerceklesenmakro' veya 'gerceklesenhesap' geçen dosyalar",
                        excluded = "Dosya adýnda 'GERÇEKLEÞEN' geçen dosyalar veya macro dosyasý olmayan dosyalar"
                    },
                    macroFiles = macroFiles.Select(f => new
                    {
                        f.FileName,
                        f.OriginalFileName,
                        f.UploadDate,
                        status = "Macro iþlemlerde KULLANILIR"
                    }),
                    excludedFiles = excludedFiles.Select(f => new
                    {
                        f.FileName,
                        f.OriginalFileName,
                        f.UploadDate,
                        status = "Macro iþlemlerde KULLANILMAZ"
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
                _logger.LogError(ex, "Dosya filtre durumu alýnýrken hata");
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Hýzlý arama - SADECE belirtilen yeni dosyalarda arar
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

                // Sheet filtresi (eðer belirtilmiþse)
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
                        message = $"'{documentNumber}' dosya numarasý belirtilen yeni dosyalarda bulunamadý.",
                        searchedFiles = newFiles,
                        suggestion = "Tüm dosyalarda arama yapmak için /api/macro/search-by-document/{documentNumber} kullanýn."
                    });
                }

                // JSON verileri parse et ve response DTO'larýna dönüþtür
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
                    message = $"'{documentNumber}' için yeni dosyalarda {responseData.Count} kayýt bulundu."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Hýzlý arama hatasý: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Hýzlý arama - SADECE belirtilen makro dosyasýnda arar
        /// </summary>
        [HttpGet("quick-search-makro/{documentNumber}")]
        public async Task<IActionResult> QuickSearchMakroOnly(string documentNumber, [FromQuery] string? sheetName = null)
        {
            try
            {
                documentNumber = Uri.UnescapeDataString(documentNumber);

                // Sadece makro dosyasýnda arama yap
                var newFiles = new[]
                {
                    "gerceklesenmakrodata_20250915153256.xlsx"
                };

                var query = _context.ExcelDataRows.Where(r => !r.IsDeleted &&
                    newFiles.Contains(r.FileName) &&
                    r.RowData.Contains(documentNumber));

                // Sheet filtresi (eðer belirtilmiþse)
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
                        message = $"'{documentNumber}' dosya numarasý belirtilen makro dosyasýnda bulunamadý.",
                        searchedFiles = newFiles,
                        suggestion = "Tüm dosyalarda arama yapmak için /api/macro/search-by-document/{documentNumber} kullanýn."
                    });
                }

                // JSON verileri parse et ve response DTO'larýna dönüþtür
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
                    message = $"'{documentNumber}' için makro dosyasýnda {responseData.Count} kayýt bulundu."
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Hýzlý arama hatasý: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        /// <summary>
        /// Dosya numarasýna göre verileri arar - sadece 'gerceklesenhesap_20250905104743.xlsx' dosyasýnda arama yapar
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
                        message = $"'{documentNumber}' dosya numarasý {targetFile} dosyasýnda bulunamadý.",
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
                _logger.LogError(ex, "Hesap dosyasýnda arama hatasý: {DocumentNumber}", documentNumber);
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}