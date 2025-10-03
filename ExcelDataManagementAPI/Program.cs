using Microsoft.EntityFrameworkCore;
using Microsoft.OpenApi.Models;
using ExcelDataManagementAPI.Data;
using ExcelDataManagementAPI.Services;
using OfficeOpenXml;

namespace ExcelDataManagementAPI
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            builder.Services.AddControllers();
            builder.Services.AddEndpointsApiExplorer();
            
            builder.Services.AddSwaggerGen(c =>
            {
                c.SwaggerDoc("v1", new OpenApiInfo
                {
                    Title = "Excel Data Management API",
                    Version = "v1",
                    Description = "Excel dosyalarını yönetmek ve karşılaştırmak için API"
                });
            });

            builder.Services.AddDbContext<ExcelDataContext>(options =>
            {
                options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"));
            });

            builder.Services.AddScoped<IExcelService, ExcelService>();
            builder.Services.AddScoped<IDataComparisonService, DataComparisonService>();
            builder.Services.AddScoped<IAuditService, AuditService>();    

            builder.Services.AddHttpContextAccessor();  

            builder.Services.Configure<IISServerOptions>(options =>
            {
                options.MaxRequestBodySize = 100 * 1024 * 1024; 
            });

            builder.Services.AddCors(options =>
            {
                options.AddPolicy("ApiPolicy", policy =>
                {
                    policy.WithOrigins(
                            "http://localhost:5174",   
                            "http://localhost:3000",   
                            "http://localhost:5173",   
                            "https://localhost:7002",  
                            "http://localhost:5002",   
                            "http://172.20.10.7:5002"   
                        )
                        .AllowAnyHeader()
                        .AllowAnyMethod()
                        .AllowCredentials(); 
                });

                options.AddPolicy("DevelopmentPolicy", policy =>
                {
                    policy.AllowAnyOrigin()
                          .AllowAnyHeader()
                          .AllowAnyMethod();
                });
            });

            var app = builder.Build();
            
            app.UseSwagger();
            app.UseSwaggerUI(c =>
            {
                c.SwaggerEndpoint("/swagger/v1/swagger.json", "Excel Data Management API v1");
                c.RoutePrefix = "swagger";
            });

            if (app.Environment.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
                app.UseCors("DevelopmentPolicy");
            }
            else
            {
                app.UseHttpsRedirection();
                app.UseCors("ApiPolicy");
            }

            app.UseStaticFiles();
            app.UseAuthorization();
            app.MapControllers();

            app.MapGet("/", context =>
            {
                context.Response.Redirect("/swagger");
                return Task.CompletedTask;
            });

            try
            {
                using var scope = app.Services.CreateScope();
                var context = scope.ServiceProvider.GetRequiredService<ExcelDataContext>();
                
                var pendingMigrations = await context.Database.GetPendingMigrationsAsync();
                if (pendingMigrations.Any())
                {
                    Console.WriteLine("🔄 Bekleyen migration'lar uygulanıyor...");
                    foreach (var migration in pendingMigrations)
                    {
                        Console.WriteLine($"   - {migration}");
                    }
                    await context.Database.MigrateAsync();
                    Console.WriteLine("✅ Migration'lar başarıyla uygulandı!");
                }
                else
                {
                    Console.WriteLine("✅ Veritabanı güncel - migration gerekmiyor!");
                }
                
                var canConnect = await context.Database.CanConnectAsync();
                if (canConnect)
                {
                    Console.WriteLine("✅ Veritabanı bağlantısı başarılı!");
                    
                    var auditTableExists = await context.Database
                        .SqlQueryRaw<int>("SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'GerceklesenRaporlarKopya'")
                        .FirstOrDefaultAsync();
                    
                    if (auditTableExists > 0)
                    {
                        Console.WriteLine("✅ GerceklesenRaporlarKopya audit tablosu hazır!");
                    }
                    else
                    {
                        Console.WriteLine("⚠️  GerceklesenRaporlarKopya tablosu bulunamadı - migration gerekebilir");
                    }
                }
                else
                {
                    Console.WriteLine("❌ Veritabanı bağlantısı başarısız!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Veritabanı hatası: {ex.Message}");
                Console.WriteLine("💡 Lütfen migration komutlarını manuel olarak çalıştırın:");
                Console.WriteLine("   1. cd ExcelDataManagementAPI");
                Console.WriteLine("   2. dotnet ef migrations add AddAuditTable");
                Console.WriteLine("   3. dotnet ef database update");
            }

            Console.WriteLine("🚀 Excel Data Management API başlatıldı!");
            Console.WriteLine("📖 Swagger UI: http://localhost:5002/swagger");
            Console.WriteLine("🌐 API Base URL: http://localhost:5002/api");
            Console.WriteLine("🔒 HTTPS Swagger UI: https://localhost:7002/swagger");
            Console.WriteLine("🔒 HTTPS API Base URL: https://localhost:7002/api");
            Console.WriteLine("🌐 Frontend URL: http://localhost:5174");
            Console.WriteLine("✅ CORS yapılandırması aktif - Frontend bağlantısı hazır!");
            Console.WriteLine("📊 Audit System aktif - Tüm değişiklikler GerceklesenRaporlarKopya tablosunda!");
            Console.WriteLine("💡 LaunchSettings.json'daki portlar kullanılıyor");

            app.Run();
        }
    }
}
