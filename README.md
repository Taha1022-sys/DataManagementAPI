# Excel Data Management API

## Overview
Excel Data Management API is a robust .NET 9 Web API for uploading, reading, editing, comparing, and exporting Excel files. It is designed for enterprise-level data management, enabling users to interact with Excel data programmatically and securely.

## Features
- **Upload Excel Files:** Upload .xlsx/.xls files and store them securely.
- **Read Data:** Read data from specific sheets or all sheets in an Excel file.
- **Database Storage:** Persist Excel data in a SQL database for fast access and versioning.
- **Edit & Update:** Update individual rows or perform bulk updates on Excel data.
- **Soft Delete:** Soft delete for both files and rows, preserving data history.
- **Export:** Export filtered or full data back to Excel format.
- **Audit System:** All changes are tracked in an audit table for traceability.
- **Swagger UI:** Interactive API documentation and testing.
- **CORS Support:** Secure cross-origin requests for frontend integration.

## API Endpoints
- `POST /api/excel/upload` — Upload an Excel file
- `GET /api/excel/files` — List uploaded files
- `POST /api/excel/read/{fileName}` — Read data from a file (all sheets or specific sheet)
- `GET /api/excel/data/{fileName}` — Get paginated data
- `GET /api/excel/data/{fileName}/all` — Get all data for a file
- `PUT /api/excel/data` — Update a row
- `PUT /api/excel/data/bulk` — Bulk update
- `POST /api/excel/data` — Add a new row
- `DELETE /api/excel/data/{id}` — Delete a row
- `DELETE /api/excel/files/{fileName}` — Delete a file
- `POST /api/excel/export` — Export data to Excel
- `GET /api/excel/sheets/{fileName}` — List sheets in a file
- `GET /api/excel/statistics/{fileName}` — Get statistics

## Technologies Used
- .NET 9 Web API
- Entity Framework Core (SQL Server)
- EPPlus (Excel operations)
- Swagger (OpenAPI)

## Getting Started
1. **Clone the repository**
2. **Configure your connection string** in `appsettings.json`
3. **Run database migrations:**
   ```
   dotnet ef database update
   ```
4. **Run the API:**
   ```
   dotnet run --project ExcelDataManagementAPI
   ```
5. **Access Swagger UI:**
   - [http://localhost:5002/swagger](http://localhost:5002/swagger)

## Example Use Cases
- Enterprise data import/export
- Data cleaning and transformation
- Versioned data management
- Backend for Excel-based business workflows

## License
This project uses the [EPPlus NonCommercial License](https://epplussoftware.com/developers/licenseexception/).

---

# Türkçesi için README.tr.md dosyasýna bakýnýz.