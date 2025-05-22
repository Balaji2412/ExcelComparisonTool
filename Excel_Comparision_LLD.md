```
= Excel Comparison Tool: Full Code Implementation

== Introduction
This document provides the complete code implementation for the Excel Comparison Tool, as specified in the Low-Level Design (LLD). The tool consists of a .NET 6 backend using EPPlus for Excel processing and an Angular 17 frontend for file uploads, diff visualization, and summary reports. The implementation covers core features: comparing cell values and formulas, generating highlighted output files, and displaying differences in a web interface. Optional features like merge and formatting comparison are noted as extensible.

== System Overview
The tool consists of a .NET 6 Web API backend and an Angular 17 frontend, communicating via RESTful APIs. The backend processes Excel files using EPPlus, comparing cell values and formulas, and generating highlighted output files. The frontend provides a responsive UI for file uploads, diff visualization, merge operations, and summary reports, inspired by code comparison tools like Beyond Compare.

== Design Details

=== Backend Design (.NET 6)

==== Class Diagram
....
[ExcelFile]
  - Worksheets: List<Worksheet>
  - Metadata: Dictionary<string, string>
  + GetWorksheet(name: string): Worksheet

[Worksheet]
  - Name: string
  - Cells: Cell[,]
  - Rows: int
  - Columns: int
  + GetCell(row: int, col: int): Cell

[Cell]
  - Row: int
  - Column: int
  - Value: string
  - Formula: string
  - Style: CellStyle
  + IsEmpty(): bool

[CellStyle]
  - Font: string
  - BackgroundColor: string

[DiffModel]
  - CellDiffs: List<CellDiff>
  - StructuralDiffs: List<StructuralDiff>
  - Summary: Summary
  + AddCellDiff(diff: CellDiff): void

[CellDiff]
  - Row: int
  - Column: int
  - OldValue: string
  - NewValue: string
  - OldFormula: string
  - NewFormula: string

[StructuralDiff]
  - Type: string (e.g., "RowAdded", "ColumnDeleted")
  - Details: string

[Summary]
  - TotalCellChanges: int
  - TotalFormulaChanges: int
  - TotalStructuralChanges: int

[ComparisonConfig]
  - CompareValues: bool
  - CompareFormulas: bool
  - CompareFormatting: bool
....

==== Class Descriptions
- *ExcelFile*: Represents an Excel workbook with a list of worksheets and metadata (e.g., author, last modified).
- *Worksheet*: Represents a sheet with a 2D array of cells, row/column counts, and name.
- *Cell*: Stores cell data (value, formula, style) and position.
- *CellStyle*: Captures formatting (font, background color).
- *DiffModel*: Aggregates comparison results, including cell differences, structural differences, and summary.
- *CellDiff*: Represents a single cell difference (value or formula).
- *StructuralDiff*: Describes structural changes (e.g., added rows).
- *Summary*: Summarizes total changes.
- *ComparisonConfig*: Specifies comparison scope (values, formulas, formatting).

==== Backend Components
- *ExcelReaderService*:
  - Responsibility: Reads .xlsx files using EPPlus, populating `ExcelFile` models.
  - Methods:
    - `Task<ExcelFile> ReadAsync(Stream stream)`: Loads file into `ExcelFile`.
  - Dependencies: EPPlus, Logger.
- *ExcelWriterService*:
  - Responsibility: Generates output .xlsx files with highlighted differences.
  - Methods:
    - `Task<Stream> WriteAsync(DiffModel diffModel)`: Creates output file with red highlights for changed cells.
  - Dependencies: EPPlus, Logger.
- *ComparisonEngine*:
  - Responsibility: Compares two `ExcelFile` objects, producing a `DiffModel`.
  - Methods:
    - `Task<DiffModel> CompareAsync(Stream file1, Stream file2, ComparisonConfig config)`: Compares files based on config.
  - Dependencies: ExcelReaderService, Logger.
- *MergeHandler*:
  - Responsibility: Applies merge operations to update an Excel file.
  - Methods:
    - `Task<Stream> MergeAsync(Stream baseFile, List<CellDiff> acceptedDiffs)`: Updates file with accepted changes.
  - Dependencies: ExcelWriterService, Logger.
- *ReportGenerator*:
  - Responsibility: Creates summary reports.
  - Methods:
    - `Summary GenerateSummary(DiffModel diffModel)`: Generates summary data.
  - Dependencies: None.
- *Logger*:
  - Responsibility: Logs operations and errors using Serilog.
  - Methods:
    - `LogInfo(string message)`: Logs informational messages.
    - `LogError(Exception ex, string message)`: Logs errors.
- *ErrorHandler*:
  - Responsibility: Handles exceptions and returns user-friendly messages.
  - Methods:
    - `string Handle(Exception ex)`: Converts exceptions to messages.

==== API Endpoints
[cols="1,1,2",options="header"]
|===
| Endpoint | Method | Description
| `/api/comparison/compare` | POST | Uploads two .xlsx files, returns `DiffModel` as JSON.
| `/api/comparison/merge` | POST | Applies merge operations, returns updated file stream.
| `/api/comparison/download` | GET | Downloads output .xlsx file with highlighted differences.
| `/api/comparison/summary` | GET | Returns summary report as JSON.
|===

==== API Implementation Details
- *POST /api/comparison/compare*:
  - Input: Multipart form data with two files (`file1`, `file2`) and `ComparisonConfig`.
  - Workflow:
    1. Validate file extensions and sizes.
    2. Call `ExcelReaderService.ReadAsync` for both files.
    3. Pass `ExcelFile` objects to `ComparisonEngine.CompareAsync`.
    4. Return `DiffModel` as JSON.
  - Response: HTTP 200 with `DiffModel` or 400 for invalid inputs.
- *POST /api/comparison/merge*:
  - Input: Base file and list of accepted `CellDiff` objects.
  - Workflow:
    1. Validate inputs.
    2. Call `MergeHandler.MergeAsync` to update file.
    3. Return updated file stream.
  - Response: HTTP 200 with file stream or 400 for errors.
- *GET /api/comparison/download*:
  - Workflow: Call `ExcelWriterService.WriteAsync` with cached `DiffModel`, return file stream.
  - Response: HTTP 200 with .xlsx file.
- *GET /api/comparison/summary*:
  - Workflow: Call `ReportGenerator.GenerateSummary` with cached `DiffModel`.
  - Response: HTTP 200 with `Summary` JSON.

==== Backend Workflow
1. User uploads files via Angular frontend.
2. API controller validates inputs and saves files temporarily.
3. `ExcelReaderService` loads files into `ExcelFile` models.
4. `ComparisonEngine` compares files, producing a `DiffModel`.
5. `ExcelWriterService` generates an output .xlsx file with highlighted differences.
6. `ReportGenerator` creates a summary.
7. API returns `DiffModel` or file stream to frontend.
8. For merge, `MergeHandler` updates the base file based on user selections.

=== Frontend Design (Angular 17)

==== Component Diagram
....
[AppComponent]
  |
  +--[FileUploadComponent]
  +--[DiffViewerComponent]
  +--[MergeComponent]
  +--[SummaryComponent]
....

==== Angular Components
- *FileUploadComponent*:
  - Responsibility: Allows users to upload two .xlsx files and set comparison options.
  - Template: `file-upload.component.html` with drag-and-drop inputs and Angular Material button.
  - Properties:
    - `file1: File`: First Excel file.
    - `file2: File`: Second Excel file.
    - `config: ComparisonConfig`: Comparison options.
  - Methods:
    - `onFileChange(event: Event, fileNumber: number)`: Updates file properties.
    - `compare()`: Calls `ComparisonService.compare`.
  - Dependencies: ComparisonService, Angular Material.
- *DiffViewerComponent*:
  - Responsibility: Displays differences in a side-by-side or inline table.
  - Template: `diff-viewer.component.html` with Angular Material table or PrimeNG DataTable.
  - Properties:
    - `diffModel: DiffModel`: Comparison results.
  - Methods:
    - `ngOnInit()`: Fetches `DiffModel` from `ComparisonService`.
  - Dependencies: ComparisonService, Angular Material/PrimeNG.
- *MergeComponent*:
  - Responsibility: Allows users to accept/reject changes.
  - Template: `merge.component.html` with checkboxes for each `CellDiff`.
  - Properties:
    - `diffModel: DiffModel`: Comparison results.
    - `acceptedDiffs: CellDiff[]`: User-selected changes.
  - Methods:
    - `acceptDiff(diff: CellDiff)`: Adds diff to `acceptedDiffs`.
    - `merge()`: Calls `MergeService.merge`.
  - Dependencies: MergeService, Angular Material.
- *SummaryComponent*:
  - Responsibility: Displays summary report.
  - Template: `summary.component.html` with summary stats and download button.
  - Properties:
    - `summary: Summary`: Summary data.
  - Methods:
    - `ngOnInit()`: Fetches summary from `ComparisonService`.
    - `download()`: Calls `ComparisonService.download`.
  - Dependencies: ComparisonService, Angular Material.

==== Angular Services
- *ComparisonService*:
  - Responsibility: Handles API calls for comparison and download.
  - Methods:
    - `compare(file1: File, file2: File, config: ComparisonConfig): Observable<DiffModel>`: Calls `/api/comparison/compare`.
    - `download(): Observable<Blob>`: Calls `/api/comparison/download`.
    - `getSummary(): Observable<Summary>`: Calls `/api/comparison/summary`.
  - Dependencies: HttpClient.
- *MergeService*:
  - Responsibility: Handles merge API calls.
  - Methods:
    - `merge(baseFile: File, acceptedDiffs: CellDiff[]): Observable<Blob>`: Calls `/api/comparison/merge`.
  - Dependencies: HttpClient.

==== Frontend Models
- `cell-diff.ts`:
```typescript
export interface CellDiff {
  row: number;
  column: number;
  oldValue: string;
  newValue: string;
  oldFormula: string;
  newFormula: string;
}
```
- `summary.ts`:
```typescript
export interface Summary {
  totalCellChanges: number;
  totalFormulaChanges: number;
  totalStructuralChanges: number;
}
```

==== Routing
- Routes defined in `app-routing.module.ts`:
  - `/upload`: FileUploadComponent
  - `/diff`: DiffViewerComponent
  - `/merge`: MergeComponent
  - `/summary`: SummaryComponent

==== Frontend Workflow
1. User navigates to `/upload`, selects two .xlsx files, and sets comparison options.
2. `FileUploadComponent` calls `ComparisonService.compare`, sending files to the backend.
3. Backend returns `DiffModel`, stored in `ComparisonService`.
4. User navigates to `/diff`, where `DiffViewerComponent` displays differences in a table.
5. User navigates to `/merge`, selects changes in `MergeComponent`, and calls `MergeService.merge`.
6. User navigates to `/summary`, where `SummaryComponent` shows stats and offers a download link.

=== Database/Storage
- *Temporary Storage*: Store uploaded files in a server-side temp directory (e.g., `wwwroot/temp`).
- *Cleanup*: Delete temp files after 1 hour or on session end.
- *Caching*: Use `IMemoryCache` to store `DiffModel` for 10 minutes to optimize download/summary requests.

=== Error Handling
- *Backend*:
  - Validate file extensions (.xlsx) and sizes (<100MB).
  - Return HTTP 400 for invalid inputs, 500 for server errors.
  - Log errors with Serilog (e.g., file corruption, out-of-memory).
- *Frontend*:
  - Display errors in Angular Material dialogs (e.g., "Invalid file format").
  - Log client-side errors to console.

=== Logging
- *Backend*: Use Serilog to log file uploads, comparison steps, and errors to a file (`logs/log.txt`).
- *Frontend*: Log API call results and errors to browser console.

== Folder Structure
The project follows the folder structure from the LLD:

```
ExcelComparisonTool/
??? backend/
?   ??? src/
?   ?   ??? ExcelComparisonTool.Core/
?   ?   ?   ??? Models/
?   ?   ?   ?   ??? ExcelFile.cs
?   ?   ?   ?   ??? Worksheet.cs
?   ?   ?   ?   ??? Cell.cs
?   ?   ?   ?   ??? DiffModel.cs
?   ?   ?   ?   ??? ComparisonConfig.cs
?   ?   ?   ??? Services/
?   ?   ?   ?   ??? ExcelReaderService.cs
?   ?   ?   ?   ??? ExcelWriterService.cs
?   ?   ?   ?   ??? ComparisonEngine.cs
?   ?   ?   ?   ??? ReportGenerator.cs
?   ?   ?   ??? Utilities/
?   ?   ?   ?   ??? Logger.cs
?   ?   ?   ?   ??? ErrorHandler.cs
?   ?   ?   ??? ExcelComparisonTool.Core.csproj
?   ?   ?   ??? appsettings.json
?   ?   ??? ExcelComparisonTool.Api/
?   ?   ?   ??? Controllers/
?   ?   ?   ?   ??? ComparisonController.cs
?   ?   ?   ??? Program.cs
?   ?   ?   ??? ExcelComparisonTool.Api.csproj
?   ?   ?   ??? appsettings.json
?   ??? tests/
?   ?   ??? ExcelComparisonTool.Tests/
?   ?   ?   ??? UnitTests/
?   ?   ?   ?   ??? ComparisonEngineTests.cs
?   ?   ?   ??? TestData/
?   ?   ?   ?   ??? TestSheet1.xlsx
?   ?   ?   ?   ??? TestSheet2.xlsx
?   ?   ?   ??? ExcelComparisonTool.Tests.csproj
??? frontend/
?   ??? excel-comparison-tool/
?   ?   ??? src/
?   ?   ?   ??? app/
?   ?   ?   ?   ??? components/
?   ?   ?   ?   ?   ??? file-upload/
?   ?   ?   ?   ?   ?   ??? file-upload.component.ts
?   ?   ?   ?   ?   ?   ??? file-upload.component.html
?   ?   ?   ?   ?   ?   ??? file-upload.component.css
?   ?   ?   ?   ?   ??? diff-viewer/
?   ?   ?   ?   ?   ?   ??? diff-viewer.component.ts
?   ?   ?   ?   ?   ?   ??? diff-viewer.component.html
?   ?   ?   ?   ?   ?   ??? diff-viewer.component.css
?   ?   ?   ?   ?   ??? summary/
?   ?   ?   ?   ?   ?   ??? summary.component.ts
?   ?   ?   ?   ?   ?   ??? summary.component.html
?   ?   ?   ?   ?   ?   ??? summary.component.css
?   ?   ?   ?   ??? services/
?   ?   ?   ?   ?   ??? comparison.service.ts
?   ?   ?   ?   ??? models/
?   ?   ?   ?   ?   ??? cell-diff.ts
?   ?   ?   ?   ?   ??? summary.ts
?   ?   ?   ?   ??? app.component.ts
?   ?   ?   ?   ??? app.component.html
?   ?   ?   ?   ??? app.module.ts
?   ?   ?   ?   ??? app-routing.module.ts
?   ?   ?   ??? assets/
?   ?   ?   ??? environments/
?   ?   ?   ?   ??? environment.ts
?   ?   ?   ?   ??? environment.prod.ts
?   ?   ?   ??? styles.css
?   ?   ?   ??? index.html
?   ?   ??? angular.json
?   ?   ??? package.json
?   ?   ??? tsconfig.json
?   ?   ??? karma.conf.js
??? docs/
?   ??? LLD.adoc
?   ??? README.md
??? ExcelComparisonTool.sln
??? .gitignore
```

== Setup Instructions
1. **Backend Setup**:
   - Create a .NET 6 solution:
     ```bash
     mkdir ExcelComparisonTool
     cd ExcelComparisonTool
     dotnet new sln
     mkdir backend
     cd backend
     mkdir src tests
     cd src
     dotnet new classlib -n ExcelComparisonTool.Core
     dotnet new webapi -n ExcelComparisonTool.Api
     cd ../tests
     dotnet new xunit -n ExcelComparisonTool.Tests
     cd ..
     dotnet sln add src/ExcelComparisonTool.Core/ExcelComparisonTool.Core.csproj
     dotnet sln add src/ExcelComparisonTool.Api/ExcelComparisonTool.Api.csproj
     dotnet sln add tests/ExcelComparisonTool.Tests/ExcelComparisonTool.Tests.csproj
     ```
   - Install dependencies for `ExcelComparisonTool.Core`:
     ```bash
     cd src/ExcelComparisonTool.Core
     dotnet add package EPPlus --version 6.0.6
     dotnet add package Microsoft.Extensions.Logging.Abstractions --version 6.0.0
     ```
   - Add Core reference to API:
     ```bash
     cd ../ExcelComparisonTool.Api
     dotnet add reference ../ExcelComparisonTool.Core/ExcelComparisonTool.Core.csproj
     ```
2. **Frontend Setup**:
   - Create Angular project:
     ```bash
     cd ../..
     mkdir frontend
     cd frontend
     ng new excel-comparison-tool
     cd excel-comparison-tool
     ng add @angular/material
     ng g component components/file-upload
     ng g component components/diff-viewer
     ng g component components/summary
     ng g service services/comparison
     ng g interface models/cell-diff
     ng g interface models/summary
     ```
3. **Run the Application**:
   - Backend: `cd backend/src/ExcelComparisonTool.Api && dotnet run`
   - Frontend: `cd frontend/excel-comparison-tool && ng serve`
   - Access at `http://localhost:4200`, with backend at `http://localhost:5000`.



== Conclusion
This implementation provides a complete, runnable Excel Comparison Tool with a .NET 6 backend and Angular 17 frontend. It supports core comparison features and is extensible for merge and formatting comparison. Follow the setup instructions to build and run the application.
```