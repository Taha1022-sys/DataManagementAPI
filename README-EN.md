# ?? Excel Data Management API

**A comprehensive .NET 9 based RESTful API for Excel file management, data comparison, and real-time data processing**

## ?? Features

### ?? Excel File Management
- ? Excel file upload (.xlsx, .xls)
- ? Multi-sheet support
- ? Automatic file validation
- ? File listing and status queries

### ?? Data Processing
- ? Import Excel data to database
- ? Real-time data editing
- ? Bulk data updates
- ? Paginated data retrieval
- ? Data filtering and search

### ?? Comparison System
- ? Compare two Excel files
- ? Change tracking
- ? Difference reporting
- ? Version history

### ?? Audit and Monitoring
- ? Automatic logging of all changes
- ? User-based operation tracking
- ? IP and timestamp records
- ? Detailed activity reports

### ?? Security
- ? CORS configuration
- ? File type validation
- ? SQL injection protection
- ? Entity Framework security

## ??? Technology Stack

- **Framework:** .NET 9.0
- **ORM:** Entity Framework Core 9.0
- **Database:** SQL Server LocalDB
- **Excel Processing:** EPPlus 7.5.0
- **API Documentation:** Swagger/OpenAPI
- **Architecture:** Clean Architecture, Repository Pattern

## ? Quick Start

### 1. Prerequisites
```bash
- .NET 9.0 SDK
- SQL Server LocalDB (comes with Visual Studio)
- Visual Studio 2022 (recommended) or VS Code
```

### 2. Installation
```bash
# Clone the repository
git clone https://github.com/Taha1022-sys/IsnaDataManagementAPI.git
cd IsnaDataManagementAPI

# Install dependencies
dotnet restore

# Create database
dotnet ef database update

# Run the application
dotnet run --project ExcelDataManagementAPI
```

### 3. Application Access
- **API Base URL:** `http://localhost:5002/api`
- **Swagger UI:** `http://localhost:5002/swagger`
- **HTTPS URL:** `https://localhost:7002`

## ?? API Usage Examples

### ?? Test API
```http
GET /api/excel/test
```

### ?? Upload Excel File
```http
POST /api/excel/upload
Content-Type: multipart/form-data

file: [Your Excel file]
uploadedBy: "username"
```

### ?? List Uploaded Files
```http
GET /api/excel/files
```

### ?? Read Excel File (All Sheets)
```http
POST /api/excel/read/{fileName}
```

### ?? Read Specific Sheet
```http
POST /api/excel/read/{fileName}?sheetName=Sheet1
```

### ?? Get Data (with Pagination)
```http
GET /api/excel/data/{fileName}?page=1&pageSize=50&sheetName=Sheet1
```

### ?? Get All Data
```http
GET /api/excel/data/{fileName}/all?sheetName=Sheet1
```

### ?? Update Data
```http
PUT /api/excel/data
Content-Type: application/json

{
  "id": 1,
  "data": {
    "column1": "new_value1",
    "column2": "new_value2"
  },
  "modifiedBy": "username",
  "changeReason": "Update reason"
}
```

### ?? Add New Row
```http
POST /api/excel/data
Content-Type: application/json

{
  "fileName": "file_name",
  "sheetName": "sheet_name",
  "rowData": {
    "column1": "value1",
    "column2": "value2"
  },
  "addedBy": "username"
}
```

### ?? Delete Data
```http
DELETE /api/excel/data/{id}?deletedBy=username
```

### ?? Get Recent Changes
```http
GET /api/excel/recent-changes?hours=24&limit=100
```

### ?? Dashboard Data
```http
GET /api/excel/dashboard
```

### ?? Compare Two Files
```http
POST /api/comparison/compare-from-files
Content-Type: multipart/form-data

file1: [First Excel file]
file2: [Second Excel file]
comparedBy: "username"
```

## ?? Configuration

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

### CORS Settings
The API supports the following ports:
- `http://localhost:5174` (Frontend)
- `http://localhost:3000` (React Dev Server)
- `http://localhost:5173` (Vite Dev Server)

## ?? Database Structure

### Main Tables
1. **ExcelFiles** - Uploaded file information
2. **ExcelDataRows** - Row-based storage of Excel data
3. **GerceklesenRaporlar** - Audit log of all changes

### Migration Commands
```bash
# Create new migration
dotnet ef migrations add MigrationName

# Apply migration
dotnet ef database update

# Rollback migration
dotnet ef database update PreviousMigrationName
```

## ?? Troubleshooting

### Common Issues

#### 1. Database Connection Error
```bash
# Make sure LocalDB is running
sqllocaldb info
sqllocaldb start mssqllocaldb

# Reapply migrations
dotnet ef database update
```

#### 2. Excel File Not Uploading
- Ensure file size is less than 100MB
- Only `.xlsx` and `.xls` files are supported
- Make sure the file is not corrupted

#### 3. CORS Error
- Ensure correct URLs are defined in `appsettings.json`
- `DevelopmentPolicy` is used in development environment

#### 4. Port Conflict
```bash
# Check used ports
netstat -an | findstr :5002

# Change port in launchSettings.json
```

## ?? Log Monitoring

### Console Output
Messages you can see when the application starts:
- ? Migration status
- ? Database connection
- ? Audit table check
- ?? API endpoints

### Log Levels
- **Information:** General operation information
- **Warning:** Warnings and things to watch out for
- **Error:** Errors and exceptions

## ?? Updates

### v1.0.0 Features
- Excel file upload and reading
- Multi-sheet support
- Data editing (CRUD)
- Comparison system
- Audit system
- Dashboard and reporting

## ?? Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/NewFeature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/NewFeature`)
5. Create a Pull Request

## ?? Contact

- **GitHub:** [Taha1022-sys](https://github.com/Taha1022-sys)
- **Repository:** [IsnaDataManagementAPI](https://github.com/Taha1022-sys/IsnaDataManagementAPI)

## ?? License

This project is licensed under the MIT License.

---

## ?? Future Features

- [ ] Authentication and Authorization
- [ ] Redis Cache integration
- [ ] Email notifications
- [ ] File version management
- [ ] Bulk operations API
- [ ] Export to different formats (PDF, CSV)
- [ ] Real-time notifications (SignalR)
- [ ] Advanced filtering and search
- [ ] Data validation rules
- [ ] Scheduled data imports

---

**? If you like this project, don't forget to give it a star!**