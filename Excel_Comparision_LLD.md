```
= Excel Comparison Tool: Full Code Implementation with Merged Cell Support

== Introduction
This document provides the updated code implementation for the Excel Comparison Tool, extending the previous implementation to support merged cell use cases. The tool compares two Excel files (.xlsx) for differences in cell values, formulas, and now merged cell ranges, using a .NET 6 backend with EPPlus and an Angular 17 frontend. The implementation covers file uploads, comparison, diff visualization, output generation, summary reports, and merged cell comparison. The merge functionality remains a future enhancement.

== Folder Structure
The folder structure remains the same as in the previous implementation, with updates to specific files:

```
ExcelComparisonTool/
├── backend/
│   ├── src/
│   │   ├── ExcelComparisonTool.Core/
│   │   │   ├── Models/
│   │   │   │   ├── ExcelFile.cs
│   │   │   │   ├── Worksheet.cs
│   │   │   │   ├── Cell.cs
│   │   │   │   ├── DiffModel.cs
│   │   │   │   ├── CellDiff.cs
│   │   │   │   ├── StructuralDiff.cs
│   │   │   │   ├── Summary.cs
│   │   │   │   ├── ComparisonConfig.cs
│   │   │   │   └── MergedCellRange.cs
│   │   │   ├── Services/
│   │   │   │   ├── ExcelReaderService.cs
│   │   │   │   ├── ExcelWriterService.cs
│   │   │   │   ├── ComparisonEngine.cs
│   │   │   │   └── ReportGenerator.cs
│   │   │   ├── Utilities/
│   │   │   │   ├── Logger.cs
│   │   │   │   └── ErrorHandler.cs
│   │   │   ├── ExcelComparisonTool.Core.csproj
│   │   │   └── appsettings.json
│   │   ├── ExcelComparisonTool.Api/
│   │   │   ├── Controllers/
│   │   │   │   └── ComparisonController.cs
│   │   │   ├── Program.cs
│   │   │   ├── ExcelComparisonTool.Api.csproj
│   │   │   └── appsettings.json
│   ├── tests/
│   │   ├── ExcelComparisonTool.Tests/
│   │   │   ├── UnitTests/
│   │   │   │   ├── ComparisonEngineTests.cs
│   │   │   ├── TestData/
│   │   │   │   ├── TestSheet1.xlsx
│   │   │   │   └── TestSheet2.xlsx
│   │   │   └── ExcelComparisonTool.Tests.csproj
├── frontend/
│   ├── excel-comparison-tool/
│   │   ├── src/
│   │   │   ├── app/
│   │   │   │   ├── components/
│   │   │   │   │   ├── file-upload/
│   │   │   │   │   │   ├── file-upload.component.ts
│   │   │   │   │   │   ├── file-upload.component.html
│   │   │   │   │   │   └── file-upload.component.css
│   │   │   │   │   ├── diff-viewer/
│   │   │   │   │   │   ├── diff-viewer.component.ts
│   │   │   │   │   │   ├── diff-viewer.component.html
│   │   │   │   │   │   └── diff-viewer.component.css
│   │   │   │   │   ├── summary/
│   │   │   │   │   │   ├── summary.component.ts
│   │   │   |   │   │   ├── summary.component.html
│   │   │   │   │   │   └── summary.component.css
│   │   │   │   ├── services/
│   │   │   │   │   ├── comparison.service.ts
│   │   │   │   ├── models/
│   │   │   │   │   ├── cell-diff.ts
│   │   │   │   │   └── summary.ts
│   │   │   │   ├── app.component.ts
│   │   │   │   ├── app.component.html
│   │   │   │   ├── app.module.ts
│   │   │   │   └── app-routing.module.ts
│   │   │   ├── assets/
│   │   │   ├── environments/
│   │   │   │   ├── environment.ts
│   │   │   │   └── environment.prod.ts
│   │   │   ├── styles.css
│   │   │   ├── index.html
│   │   ├── angular.json
│   │   ├── package.json
│   │   ├── tsconfig.json
│   │   └── karma.conf.js
├── docs/
│   ├── LLD.adoc
│   └── README.md
├── ExcelComparisonTool.sln
└── .gitignore
```

== Setup Instructions
Follow the same setup as in the previous implementation, with no additional dependencies required for merged cell support (EPPlus already supports merged cells).

1. **Backend Setup**:
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
   cd src/ExcelComparisonTool.Core
   dotnet add package EPPlus --version 6.0.6
   dotnet add package Microsoft.Extensions.Logging.Abstractions --version 6.0.0
   cd ../ExcelComparisonTool.Api
   dotnet add reference ../ExcelComparisonTool.Core/ExcelComparisonTool.Core.csproj
   ```
2. **Frontend Setup**:
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
   - Backend: `cd backend/src/ExcelComparisonTool.Api && dotnet run` (runs at `http://localhost:5000`).
   - Frontend: `cd frontend/excel-comparison-tool && ng serve` (runs at `http://localhost:4200`).

== Backend Code (Updated Files Only)

=== ExcelComparisonTool.Core/Models/MergedCellRange.cs
[source,csharp]
----
namespace ExcelComparisonTool.Core.Models
{
    public class MergedCellRange
    {
        public string StartCell { get; set; } // e.g., "A1"
        public string EndCell { get; set; } // e.g., "B2"
        public string Value { get; set; } = string.Empty;
        public string Formula { get; set; } = string.Empty;

        public string Range => $"{StartCell}:{EndCell}";
    }
}
----

=== ExcelComparisonTool.Core/Models/Worksheet.cs
[source,csharp]
----
using System.Collections.Generic;

namespace ExcelComparisonTool.Core.Models
{
    public class Worksheet
    {
        public string Name { get; set; }
        public Cell[,] Cells { get; set; }
        public int Rows { get; set; }
        public int Columns { get; set; }
        public List<MergedCellRange> MergedCells { get; set; } = new List<MergedCellRange>();

        public Cell GetCell(int row, int col)
        {
            return (row <= Rows && col <= Columns && row > 0 && col > 0) ? Cells[row - 1, col - 1] : new Cell { Row = row, Column = col };
        }
    }
}
----

=== ExcelComparisonTool.Core/Models/CellDiff.cs
[source,csharp]
----
namespace ExcelComparisonTool.Core.Models
{
    public class CellDiff
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public string OldValue { get; set; } = string.Empty;
        public string NewValue { get; set; } = string.Empty;
        public string OldFormula { get; set; } = string.Empty;
        public string NewFormula { get; set; } = string.Empty;
        public string MergedRange { get; set; } = string.Empty; // e.g., "A1:B2" for merged cells
    }
}
----

=== ExcelComparisonTool.Core/Models/DiffModel.cs
[source,csharp]
----
using System.Collections.Generic;

namespace ExcelComparisonTool.Core.Models
{
    public class DiffModel
    {
        public List<CellDiff> CellDiffs { get; set; } = new List<CellDiff>();
        public List<StructuralDiff> StructuralDiffs { get; set; } = new List<StructuralDiff>();
        public Summary Summary { get; set; } = new Summary();

        public void AddCellDiff(CellDiff diff)
        {
            CellDiffs.Add(diff);
            if (!string.IsNullOrEmpty(diff.MergedRange))
                Summary.TotalMergedCellChanges++;
            else if (!string.IsNullOrEmpty(diff.OldValue) || !string.IsNullOrEmpty(diff.NewValue))
                Summary.TotalCellChanges++;
            else if (!string.IsNullOrEmpty(diff.OldFormula) || !string.IsNullOrEmpty(diff.NewFormula))
                Summary.TotalFormulaChanges++;
        }
    }
}
----

=== ExcelComparisonTool.Core/Models/Summary.cs
[source,csharp]
----
namespace ExcelComparisonTool.Core.Models
{
    public class Summary
    {
        public int TotalCellChanges { get; set; }
        public int TotalFormulaChanges { get; set; }
        public int TotalStructuralChanges { get; set; }
        public int TotalMergedCellChanges { get; set; }
    }
}
----

=== ExcelComparisonTool.Core/Services/ExcelReaderService.cs
[source,csharp]
----
using OfficeOpenXml;
using System;
using System.Threading.Tasks;
using ExcelComparisonTool.Core.Models;
using Microsoft.Extensions.Logging;

namespace ExcelComparisonTool.Core.Services
{
    public class ExcelReaderService
    {
        private readonly ILogger<ExcelReaderService> _logger;

        public ExcelReaderService(ILogger<ExcelReaderService> logger)
        {
            _logger = logger;
        }

        public async Task<ExcelFile> ReadAsync(Stream stream)
        {
            _logger.LogInformation("Reading Excel file");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(stream);
            var excelFile = new ExcelFile();

            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var ws = new Worksheet
                {
                    Name = worksheet.Name,
                    Rows = worksheet.Dimension?.Rows ?? 0,
                    Columns = worksheet.Dimension?.Columns ?? 0
                };
                ws.Cells = new Cell[ws.Rows, ws.Columns];

                // Read cells
                for (int row = 1; row <= ws.Rows; row++)
                {
                    for (int col = 1; col <= ws.Columns; col++)
                    {
                        ws.Cells[row - 1, col - 1] = new Cell
                        {
                            Row = row,
                            Column = col,
                            Value = worksheet.Cells[row, col].Text ?? string.Empty,
                            Formula = worksheet.Cells[row, col].Formula ?? string.Empty
                        };
                    }
                }

                // Read merged cells
                foreach (var mergedRange in worksheet.MergedCells)
                {
                    var range = worksheet.Cells[mergedRange];
                    ws.MergedCells.Add(new MergedCellRange
                    {
                        StartCell = range.Start.Address,
                        EndCell = range.End.Address,
                        Value = range.Text ?? string.Empty,
                        Formula = range.Formula ?? string.Empty
                    });
                }

                excelFile.Worksheets.Add(ws);
            }

            _logger.LogInformation("Excel file read successfully");
            return excelFile;
        }
    }
}
----

=== ExcelComparisonTool.Core/Services/ExcelWriterService.cs
[source,csharp]
----
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Threading.Tasks;
using ExcelComparisonTool.Core.Models;
using Microsoft.Extensions.Logging;

namespace ExcelComparisonTool.Core.Services
{
    public class ExcelWriterService
    {
        private readonly ILogger<ExcelWriterService> _logger;

        public ExcelWriterService(ILogger<ExcelWriterService> logger)
        {
            _logger = logger;
        }

        public async Task<Stream> WriteAsync(DiffModel diffModel)
        {
            _logger.LogInformation("Writing output Excel file");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Differences");

            foreach (var diff in diffModel.CellDiffs)
            {
                var cell = worksheet.Cells[diff.Row, diff.Column];
                string displayValue = string.IsNullOrEmpty(diff.MergedRange)
                    ? $"{diff.OldValue} -> {diff.NewValue}"
                    : $"Merged {diff.MergedRange}: {diff.OldValue} -> {diff.NewValue}";
                cell.Value = displayValue;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
            }

            var stream = new MemoryStream();
            await package.SaveAsync(stream);
            stream.Position = 0;
            _logger.LogInformation("Output Excel file written successfully");
            return stream;
        }
    }
}
----

=== ExcelComparisonTool.Core/Services/ComparisonEngine.cs
[source,csharp]
----
using System;
using System.Linq;
using System.Threading.Tasks;
using ExcelComparisonTool.Core.Models;
using Microsoft.Extensions.Logging;

namespace ExcelComparisonTool.Core.Services
{
    public class ComparisonEngine
    {
        private readonly ExcelReaderService _reader;
        private readonly ILogger<ComparisonEngine> _logger;

        public ComparisonEngine(ExcelReaderService reader, ILogger<ComparisonEngine> logger)
        {
            _reader = reader;
            _logger = logger;
        }

        public async Task<DiffModel> CompareAsync(Stream file1, Stream file2, ComparisonConfig config)
        {
            _logger.LogInformation("Starting comparison");
            var excel1 = await _reader.ReadAsync(file1);
            var excel2 = await _reader.ReadAsync(file2);
            var diffModel = new DiffModel();

            foreach (var sheet1 in excel1.Worksheets)
            {
                var sheet2 = excel2.GetWorksheet(sheet1.Name);
                if (sheet2 == null)
                {
                    diffModel.StructuralDiffs.Add(new StructuralDiff { Type = "SheetMissing", Details = sheet1.Name });
                    diffModel.Summary.TotalStructuralChanges++;
                    continue;
                }

                // Compare merged cells
                foreach (var merge1 in sheet1.MergedCells)
                {
                    var merge2 = sheet2.MergedCells.FirstOrDefault(m => m.Range == merge1.Range);
                    if (merge2 == null)
                    {
                        diffModel.AddCellDiff(new CellDiff
                        {
                            Row = GetRowFromAddress(merge1.StartCell),
                            Column = GetColumnFromAddress(merge1.StartCell),
                            MergedRange = merge1.Range,
                            OldValue = merge1.Value,
                            NewValue = string.Empty
                        });
                        continue;
                    }

                    if (config.CompareValues && merge1.Value != merge2.Value)
                    {
                        diffModel.AddCellDiff(new CellDiff
                        {
                            Row = GetRowFromAddress(merge1.StartCell),
                            Column = GetColumnFromAddress(merge1.StartCell),
                            MergedRange = merge1.Range,
                            OldValue = merge1.Value,
                            NewValue = merge2.Value
                        });
                    }

                    if (config.CompareFormulas && merge1.Formula != merge2.Formula)
                    {
                        diffModel.AddCellDiff(new CellDiff
                        {
                            Row = GetRowFromAddress(merge1.StartCell),
                            Column = GetColumnFromAddress(merge1.StartCell),
                            MergedRange = merge1.Range,
                            OldFormula = merge1.Formula,
                            NewFormula = merge2.Formula
                        });
                    }
                }

                // Compare non-merged cells
                int maxRows = Math.Max(sheet1.Rows, sheet2.Rows);
                int maxCols = Math.Max(sheet1.Columns, sheet2.Columns);

                for (int row = 1; row <= maxRows; row++)
                {
                    for (int col = 1; col <= maxCols; col++)
                    {
                        // Skip if cell is part of a merged range
                        if (sheet1.MergedCells.Any(m => IsCellInRange(row, col, m.Range)) ||
                            sheet2.MergedCells.Any(m => IsCellInRange(row, col, m.Range)))
                            continue;

                        var cell1 = sheet1.GetCell(row, col);
                        var cell2 = sheet2.GetCell(row, col);

                        if (config.CompareValues && cell1.Value != cell2.Value)
                        {
                            diffModel.AddCellDiff(new CellDiff
                            {
                                Row = row,
                                Column = col,
                                OldValue = cell1.Value,
                                NewValue = cell2.Value
                            });
                        }

                        if (config.CompareFormulas && cell1.Formula != cell2.Formula)
                        {
                            diffModel.AddCellDiff(new CellDiff
                            {
                                Row = row,
                                Column = col,
                                OldFormula = cell1.Formula,
                                NewFormula = cell2.Formula
                            });
                        }
                    }
                }
            }

            _logger.LogInformation($"Comparison completed: {diffModel.Summary.TotalCellChanges} cell changes, {diffModel.Summary.TotalMergedCellChanges} merged cell changes");
            return diffModel;
        }

        private int GetRowFromAddress(string address)
        {
            return int.Parse(address.Substring(1)); // e.g., "A1" -> 1
        }

        private int GetColumnFromAddress(string address)
        {
            string col = address.Substring(0, address.Length - address.Skip(1).TakeWhile(char.IsDigit).Count());
            return col.Aggregate(0, (current, c) => current * 26 + (c - 'A' + 1));
        }

        private bool IsCellInRange(int row, int col, string range)
        {
            var parts = range.Split(':');
            int startRow = int.Parse(parts[0].Substring(1));
            int endRow = int.Parse(parts[1].Substring(1));
            int startCol = GetColumnFromAddress(parts[0]);
            int endCol = GetColumnFromAddress(parts[1]);
            return row >= startRow && row <= endRow && col >= startCol && col <= endCol;
        }
    }
}
----

=== ExcelComparisonTool.Tests/UnitTests/ComparisonEngineTests.cs
[source,csharp]
----
using System.IO;
using System.Threading.Tasks;
using ExcelComparisonTool.Core.Services;
using ExcelComparisonTool.Core.Models;
using Microsoft.Extensions.Logging;
using Moq;
using Xunit;

namespace ExcelComparisonTool.Tests.UnitTests
{
    public class ComparisonEngineTests
    {
        [Fact]
        public async Task CompareAsync_DifferentValues_ReturnsCellDiffs()
        {
            var logger = new Mock<ILogger<ComparisonEngine>>().Object;
            var readerMock = new Mock<ExcelReaderService>(logger);
            readerMock.Setup(r => r.ReadAsync(It.IsAny<Stream>())).ReturnsAsync(new ExcelFile
            {
                Worksheets = { new Worksheet
                {
                    Name = "Sheet1",
                    Rows = 1,
                    Columns = 1,
                    Cells = new Cell[,] { { new Cell { Row = 1, Column = 1, Value = "A" } } }
                }}
            });

            var engine = new ComparisonEngine(readerMock.Object, logger);
            var config = new ComparisonConfig { CompareValues = true };

            using var stream1 = new MemoryStream();
            using var stream2 = new MemoryStream();
            var diffModel = await engine.CompareAsync(stream1, stream2, config);

            Assert.NotNull(diffModel);
            Assert.Equal(1, diffModel.Summary.TotalCellChanges);
        }

        [Fact]
        public async Task CompareAsync_DifferentMergedCells_ReturnsMergedCellDiffs()
        {
            var logger = new Mock<ILogger<ComparisonEngine>>().Object;
            var readerMock = new Mock<ExcelReaderService>(logger);
            readerMock.Setup(r => r.ReadAsync(It.IsAny<Stream>())).ReturnsAsync(new ExcelFile
            {
                Worksheets = { new Worksheet
                {
                    Name = "Sheet1",
                    Rows = 2,
                    Columns = 2,
                    Cells = new Cell[2, 2],
                    MergedCells = { new MergedCellRange { StartCell = "A1", EndCell = "B2", Value = "Merged" } }
                }}
            });

            var engine = new ComparisonEngine(readerMock.Object, logger);
            var config = new ComparisonConfig { CompareValues = true };

            using var stream1 = new MemoryStream();
            using var stream2 = new MemoryStream();
            var diffModel = await engine.CompareAsync(stream1, stream2, config);

            Assert.NotNull(diffModel);
            Assert.Equal(1, diffModel.Summary.TotalMergedCellChanges);
        }
    }
}
----

== Frontend Code (Updated Files Only)

=== frontend/excel-comparison-tool/src/app/models/cell-diff.ts
[source,typescript]
----
export interface CellDiff {
  row: number;
  column: number;
  oldValue: string;
  newValue: string;
  oldFormula: string;
  newFormula: string;
  mergedRange: string; // e.g., "A1:B2"
}

export interface DiffModel {
  cellDiffs: CellDiff[];
  structuralDiffs: { type: string; details: string }[];
  summary: Summary;
}
----

=== frontend/excel-comparison-tool/src/app/models/summary.ts
[source,typescript]
----
export interface Summary {
  totalCellChanges: number;
  totalFormulaChanges: number;
  totalStructuralChanges: number;
  totalMergedCellChanges: number;
}
----

=== frontend/excel-comparison-tool/src/app/components/diff-viewer/diff-viewer.component.ts
[source,typescript]
----
import { Component, OnInit } from '@angular/core';
import { ComparisonService } from '../../services/comparison.service';
import { CellDiff } from '../../models/cell-diff';
import { Router } from '@angular/router';

@Component({
  selector: 'app-diff-viewer',
  templateUrl: './diff-viewer.component.html',
  styleUrls: ['./diff-viewer.component.css']
})
export class DiffViewerComponent implements OnInit {
  displayedColumns: string[] = ['row', 'column', 'mergedRange', 'oldValue', 'newValue', 'oldFormula', 'newFormula'];
  dataSource: CellDiff[] = [];

  constructor(private comparisonService: ComparisonService, private router: Router) {}

  ngOnInit() {
    const diffModel = this.comparisonService.getDiffModel();
    if (diffModel) {
      this.dataSource = diffModel.cellDiffs;
    } else {
      this.router.navigate(['/upload']);
    }
  }

  goToSummary() {
    this.router.navigate(['/summary']);
  }

  download() {
    this.comparisonService.download().subscribe(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'differences.xlsx';
      a.click();
      window.URL.revokeObjectURL(url);
    });
  }
}
----

=== frontend/excel-comparison-tool/src/app/components/diff-viewer/diff-viewer.component.html
[source,html]
----
<div class="container">
  <h2>Differences</h2>
  <mat-table [dataSource]="dataSource">
    <ng-container matColumnDef="row">
      <mat-header-cell *matHeaderCellDef>Row</mat-header-cell>
      <mat-cell *matCellDef="let diff">{{diff.row}}</mat-cell>
    </ng-container>
    <ng-container matColumnDef="column">
      <mat-header-cell *matHeaderCellDef>Column</mat-header-cell>
      <mat-cell *matCellDef="let diff">{{diff.column}}</mat-cell>
    </ng-container>
    <ng-container matColumnDef="mergedRange">
      <mat-header-cell *matHeaderCellDef>Merged Range</mat-header-cell>
      <mat-cell *matCellDef="let diff">{{diff.mergedRange || '-'}}</mat-cell>
    </ng-container>
    <ng-container matColumnDef="oldValue">
      <mat-header-cell *matHeaderCellDef>Old Value</mat-header-cell>
      <mat-cell *matCellDef="let diff">{{diff.oldValue}}</mat-cell>
    </ng-container>
    <ng-container matColumnDef="newValue">
      <mat-header-cell *matHeaderCellDef>New Value</mat-header-cell>
      <mat-cell *matCellDef="let diff">{{diff.newValue}}</mat-cell>
    </ng-container>
    <ng-container matColumnDef="oldFormula">
      <mat-header-cell *matHeaderCellDef>Old Formula</mat-header-cell>
      <mat-cell *matCellDef="let diff">{{diff.oldFormula}}</mat-cell>
    </ng-container>
    <ng-container matColumnDef="newFormula">
      <mat-header-cell *matHeaderCellDef>New Formula</mat-header-cell>
      <mat-cell *matCellDef="let diff">{{diff.newFormula}}</mat-cell>
    </ng-container>
    <mat-header-row *matHeaderRowDef="displayedColumns"></mat-header-row>
    <mat-row *matRowDef="let row; columns: displayedColumns;"></mat-row>
  </mat-table>
  <button mat-raised-button color="primary" (click)="goToSummary()">View Summary</button>
  <button mat-raised-button color="accent" (click)="download()">Download Output</button>
</div>
----

=== frontend/excel-comparison-tool/src/app/components/summary/summary.component.html
[source,html]
----
<div class="container">
  <h2>Summary</h2>
  <div *ngIf="summary">
    <p>Total Cell Changes: {{summary.totalCellChanges}}</p>
    <p>Total Formula Changes: {{summary.totalFormulaChanges}}</p>
    <p>Total Structural Changes: {{summary.totalStructuralChanges}}</p>
    <p>Total Merged Cell Changes: {{summary.totalMergedCellChanges}}</p>
  </div>
  <button mat-raised-button color="primary" (click)="router.navigate(['/upload'])">Back to Upload</button>
</div>
----

== Unchanged Files
All other files (`ExcelFile.cs`, `Cell.cs`, `StructuralDiff.cs`, `ComparisonConfig.cs`, `ReportGenerator.cs`, `Logger.cs`, `ErrorHandler.cs`, `ComparisonController.cs`, `Program.cs`, `ExcelComparisonTool.Core.csproj`, `ExcelComparisonTool.Api.csproj`, `appsettings.json`, `ExcelComparisonTool.Tests.csproj`, `app.component.ts`, `app.component.html`, `app.module.ts`, `app-routing.module.ts`, `file-upload.component.ts`, `file-upload.component.html`, `file-upload.component.css`, `comparison.service.ts`, `environment.ts`, `environment.prod.ts`, `styles.css`, `index.html`, `angular.json`, `package.json`, `tsconfig.json`, `karma.conf.js`, `ExcelComparisonTool.sln`, `.gitignore`) remain identical to the previous implementation.

== Notes
- *Merged Cell Support*:
  - Detects merged cell ranges using EPPlus’s `MergedCells` property.
  - Compares merged cell ranges by their boundaries (`A1:B2`), values, and formulas.
  - Displays merged cell differences in the frontend with a `Merged Range` column.
  - Highlights merged cell differences in the output Excel file.
- *Features Implemented*: File upload, cell value/formula comparison, merged cell comparison, diff visualization, output generation, summary report.
- *Not Implemented*: Merge functionality, formatting comparison (extend `CellStyle`).
- *Testing*: Added a unit test for merged cell comparison. Create test files in `backend/tests/TestData/` with merged cells (e.g., merge `A1:B2` in one file with different values).
- *Performance*: Handles small to medium files. For large files, enable EPPlus streaming and Angular virtual scrolling.
- *Security*: Validates file extensions. Add file scanning for production.

== Testing Merged Cell Support
1. Create two .xlsx files:
   - `TestSheet1.xlsx`: Merge cells `A1:B2` with value "Merged1".
   - `TestSheet2.xlsx`: Merge cells `A1:B2` with value "Merged2".
2. Upload via the frontend at `http://localhost:4200/upload`.
3. Verify the diff table shows a merged cell difference for `A1:B2`.
4. Check the output file (`differences.xlsx`) for highlighted merged cell changes.
5. Run unit tests: `cd backend/tests/ExcelComparisonTool.Tests && dotnet test`.

== Conclusion
This updated implementation extends the Excel Comparison Tool to support merged cell use cases, maintaining all existing features. The .NET 6 backend handles merged cell detection and comparison, while the Angular 17 frontend displays these differences clearly. The solution is modular, scalable, and ready for further enhancements like merge functionality or formatting comparison.
```