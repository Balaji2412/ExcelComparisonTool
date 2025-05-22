```
= Excel Comparison Tool: High-Level Design


== Overview
The Excel Comparison Tool is a web-based application designed to compare two Excel files (.xlsx) for differences in cell values, formulas, and optionally formatting or structure. The backend, built with .NET 6 and EPPlus, exposes RESTful APIs for file processing. The frontend, developed with Angular 17, provides a user-friendly interface for file uploads, difference visualization, and merge operations. Inspired by code comparison tools like Beyond Compare, the tool is tailored for Excel files, ensuring scalability, extensibility, and a seamless user experience.

== Objectives
- Compare cell values, formulas, and optionally formatting between two Excel files.
- Provide a responsive Angular UI for file uploads, diff visualization, and merge operations.
- Generate highlighted output files or visual reports.
- Support multi-sheet comparisons and structural differences (e.g., added/deleted rows).
- Ensure scalability for large Excel files and extensibility for additional features.
- Maintain a modular architecture for backend and frontend.

== Architecture Overview
The application follows a client-server architecture with a .NET 6 backend API and an Angular frontend, communicating via RESTful APIs.

=== Layers
- *Frontend (Angular)*: Handles user interaction, file uploads, and diff visualization using Angular components.
- *Backend (Business Logic Layer)*: Manages comparison logic, merge operations, and report generation.
- *Backend (Data Access Layer)*: Uses EPPlus to read/write Excel files and normalize data.
- *Core Utilities*: Includes logging, configuration, and error handling.

=== Components
- *Frontend Components*:
  - File Upload: For uploading two .xlsx files.
  - Diff Viewer: Displays differences in a table or side-by-side view.
  - Merge: Enables users to accept/reject changes.
  - Summary: Shows a summary of differences.
- *Backend Services*:
  - Excel Reader Service: Reads Excel files into models using EPPlus.
  - Comparison Engine: Compares cells, formulas, and metadata.
  - Excel Writer Service: Generates output Excel files with highlighted differences.
  - Merge Handler: Processes merge requests.
  - Report Generator: Creates summary reports.
- *Utilities*:
  - Logger: Logs operations and errors (e.g., using Serilog).
  - Error Handler: Provides user-friendly error messages.
  - Configuration Manager: Manages settings (e.g., comparison scope).

=== Technology Stack
- *Backend*:
  - Language: C# (.NET 6)
  - Library: EPPlus
  - API Framework: ASP.NET Core Web API
  - Logging: Serilog or Microsoft.Extensions.Logging
- *Frontend*:
  - Framework: Angular 17
  - UI Library: Angular Material or PrimeNG
  - HTTP Client: Angular HttpClient
- *Testing*:
  - Backend: xUnit or NUnit
  - Frontend: Jasmine/Karma, Cypress
- *Build Tools*:
  - Backend: .NET CLI
  - Frontend: Angular CLI

== System Design

=== Data Flow
. *Input*: User uploads two .xlsx files via the Angular frontend, along with comparison options.
. *Processing*:
  - *Frontend*: Sends files to the backend API via multipart form data.
  - *Backend*: Excel Reader Service loads files, Comparison Engine compares cells/formulas, Excel Writer Service generates output.
  - *Response*: Backend returns a JSON DiffModel to the frontend.
. *Output*: Frontend renders differences in a table; backend provides a downloadable .xlsx file with highlighted differences.

=== Data Model
- *Backend Models*:
  - `ExcelFile`: Contains Worksheets, Metadata (e.g., author).
  - `Worksheet`: Includes Rows, Columns, Cells, NamedRanges.
  - `Cell`: Has Value, Formula, Style (font, fill).
  - `DiffModel`: Stores CellDiffs, StructuralDiffs, Summary.
  - `ComparisonConfig`: Defines comparison scope (values, formulas, formatting).
- *Frontend Models*:
  - TypeScript interfaces mirroring backend models (e.g., `CellDiff`, `Summary`).

=== API Endpoints
- `POST /api/comparison/compare`: Uploads files and returns DiffModel.
- `POST /api/comparison/merge`: Applies merge operations.
- `GET /api/comparison/download`: Downloads output Excel file.
- `GET /api/comparison/summary`: Returns summary report.

=== Component Diagram
....
[Angular Frontend]
  | File Upload | Diff Viewer | Merge | Summary |
         |
         v
[ASP.NET Core API]
  | ComparisonController | MergeController |
         |
         v
[Business Logic Layer]
  | Comparison Engine | Merge Handler | Report Generator |
         |
         v
[Data Access Layer]
  | Excel Reader (EPPlus) | Excel Writer (EPPlus) |
         |
         v
[Excel Files (.xlsx)]
....

== Key Features
- *File Upload*: Upload .xlsx files via Angular drag-and-drop or file input.
- *Cell Value Comparison*: Compare cell values, handling nulls and mismatched dimensions.
- *Formula Comparison*: Compare cell formulas (e.g., `=A1+B1` vs. `=A1*B1`).
- *Formatting Comparison* (optional): Compare cell styles (e.g., font, background).
- *Diff Visualization*: Display differences in Angular tables with highlighting.
- *Merge Functionality*: Accept/reject changes via UI, updating the backend.
- *Output Generation*: Download Excel file with highlighted differences.
- *Summary Report*: Show total changes (e.g., "10 cells changed, 2 rows added").

== Scalability and Performance
- *Backend*:
  - Use EPPlus streaming for large files.
  - Implement async/await for API endpoints.
  - Cache DiffModel using IMemoryCache.
- *Frontend*:
  - Use Angular virtual scrolling for large diff tables.
  - Optimize API calls with RxJS operators (e.g., `debounceTime`).
- *File Handling*: Store uploaded files temporarily with cleanup mechanism.

== Extensibility
- Support .xls files using NPOI.
- Allow custom comparison rules via ComparisonConfig.
- Support additional output formats (CSV, JSON).

== User Interface (Angular)
- *File Upload Page*: Drag-and-drop or file input for .xlsx files.
- *Diff Viewer Page*: Side-by-side or inline table with highlighting (e.g., red for changes).
- *Merge Page*: Interactive UI for accepting/rejecting changes.
- *Summary Page*: Report with download button.
- *Tech*: Angular Material, PrimeNG, RxJS.

== Error Handling and Logging
- *Backend*:
  - Validate file formats and sizes.
  - Return HTTP 400/500 with user-friendly messages.
  - Log operations/errors using Serilog.
- *Frontend*:
  - Display errors in Angular Material dialogs.
  - Log client-side errors to console or logging service.

== Non-Functional Requirements
- *Performance*: Process 10,000 cells in under 10 seconds, render in under 2 seconds.
- *Scalability*: Handle 100MB files and multiple concurrent users.
- *Security*: Validate file uploads to prevent malicious content.
- *Usability*: Intuitive, responsive UI.

== Future Enhancements
- Support for charts, pivot tables, VBA macros.
- Cloud storage integration (e.g., Azure Blob Storage).
- Real-time collaboration using SignalR.

== Assumptions and Constraints
- *Assumptions*:
  - Input files are valid .xlsx files.
  - Users have modern browsers (e.g., Chrome, Edge).
  - EPPlus is used under a non-commercial or commercial license.
- *Constraints*:
  - EPPlus is not thread-safe; use locks for parallel processing.
  - Large file uploads may require server-side storage limits.

== Conclusion
The Excel Comparison Tool provides a robust solution for comparing Excel files, with a modular .NET 6 backend and Angular frontend. It ensures scalability, extensibility, and a user-friendly experience, suitable for data analysts and developers.
```