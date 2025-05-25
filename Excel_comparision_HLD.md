```
= Excel Comparison Tool: High-Level Design


== Overview
The Excel Comparison Tool is a web-based application designed to compare two Excel files (.xlsx) for differences in cell values, formulas, and merged cell ranges. Built with a .NET 6 backend using EPPlus for Excel processing and an Angular 17 frontend, the tool provides a user-friendly interface for file uploads, difference visualization, and output generation. Inspired by code comparison tools like Beyond Compare, it targets data analysts and developers, offering a scalable and extensible solution for Excel file comparison.

== Objectives
- Compare cell values, formulas, and merged cell ranges between two Excel files.
- Provide a responsive Angular UI for file uploads, difference visualization, and summary reports.
- Generate highlighted output Excel files showing differences.
- Support multi-sheet comparisons and structural differences (e.g., missing sheets).
- Ensure scalability for large Excel files and extensibility for additional features (e.g., formatting comparison).
- Maintain a modular architecture for maintainability and future enhancements.

== Architecture Overview
The application follows a client-server architecture with a .NET 6 backend exposing RESTful APIs and an Angular 17 frontend for user interaction.

=== Layers
- *Frontend (Angular)*: Handles user interface, file uploads, and visualization of differences using Angular components.
- *Backend (Business Logic Layer)*: Manages comparison logic, output generation, and report creation.
- *Backend (Data Access Layer)*: Uses EPPlus to read and write Excel files.
- *Core Utilities*: Includes logging, error handling, and configuration management.

=== Technology Stack
- *Backend*:
  - Language: C# (.NET 6)
  - Library: EPPlus (for Excel processing)
  - API Framework: ASP.NET Core Web API
  - Logging: Microsoft.Extensions.Logging
- *Frontend*:
  - Framework: Angular 17
  - UI Library: Angular Material
  - HTTP Client: Angular HttpClient
- *Testing*:
  - Backend: xUnit
  - Frontend: Jasmine/Karma
- *Build Tools*:
  - Backend: .NET CLI
  - Frontend: Angular CLI

== System Design

=== Data Flow
1. *Input*: User uploads two .xlsx files via the Angular frontend, along with comparison options (values, formulas, merged cells).
2. *Processing*:
   - *Frontend*: Sends files to the backend API via multipart form data.
   - *Backend*: Reads files using EPPlus, compares worksheets, cells, and merged cells, and generates a difference model.
   - *Response*: Backend returns a JSON difference model to the frontend.
3. *Output*: Frontend renders differences in a table; backend provides a downloadable .xlsx file with highlighted differences.

=== API Endpoints
- `POST /api/comparison/compare`: Uploads two .xlsx files and returns a difference model (JSON).
- `GET /api/comparison/download`: Downloads an output Excel file with highlighted differences.
- `GET /api/comparison/summary`: Returns a summary report (JSON).

=== Component Diagram
....
[Angular Frontend]
  | File Upload | Diff Viewer | Summary |
         |
         v
[ASP.NET Core API]
  | ComparisonController |
         |
         v
[Business Logic Layer]
  | Comparison Engine | Report Generator |
         |
         v
[Data Access Layer]
  | Excel Reader (EPPlus) | Excel Writer (EPPlus) |
         |
         v
[Excel Files (.xlsx)]
....

== Key Features
- *File Upload*: Upload two .xlsx files via Angular drag-and-drop or file input.
- *Cell Value Comparison*: Compare cell values across worksheets, handling nulls and mismatched dimensions.
- *Formula Comparison*: Compare cell formulas (e.g., `=A1+B1` vs. `=A1*B1`).
- *Merged Cell Comparison*: Compare merged cell ranges (e.g., `A1:B2`), including range boundaries and values/formulas.
- *Diff Visualization*: Display differences in an Angular table with columns for row, column, merged range, and old/new values/formulas.
- *Output Generation*: Generate a downloadable .xlsx file with highlighted differences (e.g., red background for changed cells).
- *Summary Report*: Provide a summary of total changes (cells, formulas, merged cells, structural).

== Scalability and Performance
- *Backend*:
  - Use EPPlus streaming mode for large files (>10,000 cells).
  - Implement async/await for API endpoints to handle concurrent requests.
  - Cache difference models to reduce redundant processing.
- *Frontend*:
  - Use Angular virtual scrolling for large difference tables.
  - Optimize API calls with RxJS operators (e.g., `debounceTime`).
- *File Handling*: Store uploaded files temporarily with automated cleanup.

== Extensibility
- Add support for formatting comparison (e.g., fonts, colors) by extending the comparison engine.
- Implement merge functionality to allow users to accept/reject changes.
- Support additional file formats (e.g., .xls) using libraries like NPOI.
- Integrate cloud storage (e.g., Azure Blob Storage) for file management.

== User Interface (Angular)
- *File Upload Page*: Drag-and-drop or file input for .xlsx files with comparison options.
- *Diff Viewer Page*: Table-based view of differences, including merged cell ranges, with highlighting.
- *Summary Page*: Report with statistics (e.g., total cell changes) and a download button.
- *Tech*: Angular Material for responsive design, RxJS for reactive programming.

== Error Handling and Logging
- *Backend*:
  - Validate file formats (.xlsx) and sizes (<100MB).
  - Return HTTP 400/500 with user-friendly error messages.
  - Log operations and errors using Microsoft.Extensions.Logging.
- *Frontend*:
  - Display errors in Angular Material dialogs (e.g., "Invalid file format").
  - Log client-side errors to the browser console.

== Non-Functional Requirements
- *Performance*: Process 10,000 cells in under 10 seconds, render UI in under 2 seconds.
- *Scalability*: Handle 100MB files and multiple concurrent users.
- *Security*: Validate file uploads to prevent malicious content; use HTTPS for API communication.
- *Usability*: Intuitive, responsive UI compatible with modern browsers (Chrome, Edge).

== Future Enhancements
- Support for comparing charts, pivot tables, or VBA macros.
- Real-time collaboration using SignalR for multi-user scenarios.
- Integration with cloud storage for persistent file access.
- Advanced merge functionality with interactive change acceptance.

== Assumptions and Constraints
- *Assumptions*:
  - Input files are valid .xlsx files.
  - Users have modern browsers (e.g., Chrome, Edge).
  - EPPlus is used under a non-commercial or commercial license.
- *Constraints*:
  - EPPlus is not thread-safe; use locks or single-threaded processing for concurrent operations.
  - Large file uploads require server-side storage limits (e.g., 100MB).
  - Merged cell comparison assumes consistent values/formulas within ranges.

== Conclusion
The Excel Comparison Tool provides a robust, user-friendly solution for comparing Excel files, with support for cell values, formulas, and merged cell ranges. Its modular .NET 6 backend and Angular 17 frontend ensure scalability, extensibility, and maintainability, making it suitable for data analysts and developers. The architecture supports future enhancements like merge functionality and formatting comparison.
```