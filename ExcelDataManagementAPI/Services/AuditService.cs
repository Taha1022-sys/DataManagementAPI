using ExcelDataManagementAPI.Data;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace ExcelDataManagementAPI.Services
{
    public interface IAuditService
    {

    }

    public class AuditService : IAuditService
    {
        public AuditService(ExcelDataContext context, ILogger<AuditService> logger) { }
    }
}