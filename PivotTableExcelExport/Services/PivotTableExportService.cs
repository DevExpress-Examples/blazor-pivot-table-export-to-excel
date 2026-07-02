using DevExpress.Blazor.PivotTable;
using DevExpress.Spreadsheet;

namespace PivotTableExcelExport.Services;

public class PivotTableExportService {
    public byte[] ExportToExcel<T>(IEnumerable<T> data, PivotTablePersistentLayout layout, PivotTableExportConfig? config = null) {
        config ??= new PivotTableExportConfig();

        using var workbook = new Workbook();
        var dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "PivotTableData";
        var properties = typeof(T).GetProperties();

        for (int col = 0; col < properties.Length; col++) {
            dataSheet.Cells[0, col].Value = properties[col].Name;
        }

        var dataList = data.ToList();
        for (int row = 0; row < dataList.Count; row++) {
            for (int col = 0; col < properties.Length; col++) {
                var value = properties[col].GetValue(dataList[row]);
                var cell = dataSheet.Cells[row + 1, col];
                switch (value) {
                    case int intVal:
                        cell.Value = intVal;
                        break;
                    case DateTime dateVal:
                        cell.Value = dateVal;
                        cell.NumberFormat = "m/d/yyyy";
                        break;
                    case string strVal:
                        cell.Value = strVal;
                        break;
                }
            }
        }

        int lastRow = dataList.Count;
        int lastCol = properties.Length - 1;
        var sourceRange = dataSheet.Range.FromLTRB(0, 0, lastCol, lastRow);

        var pivotSheet = workbook.Worksheets.Add("PivotTable");
        workbook.Worksheets.ActiveWorksheet = pivotSheet;

        var pivotTable = pivotSheet.PivotTables.Add(sourceRange, pivotSheet["A1"]);
        pivotTable.BeginUpdate();

        var layoutFields = layout.Fields
            .Where(f => f.Area.HasValue && f.Visible != false)
            .OrderBy(f => f.AreaIndex)
            .ToList();

        var addedColumnFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var layoutField in layoutFields) {
            var fieldName = layoutField.Field;
            if (string.IsNullOrEmpty(fieldName))
                continue;

            var pivotField = FindPivotField(pivotTable, fieldName);
            if (pivotField == null)
                continue;

            switch (layoutField.Area) {
                case PivotTableArea.Row:
                    pivotTable.RowFields.Add(pivotField);
                    break;
                case PivotTableArea.Column:
                    if (!addedColumnFields.Contains(fieldName)) {
                        pivotTable.ColumnFields.Add(pivotField);
                        addedColumnFields.Add(fieldName);
                    }
                    break;
                case PivotTableArea.Data:
                    var dataField = pivotTable.DataFields.Add(pivotField);
                    if (config.DataFieldConfigs.TryGetValue(fieldName, out var dfConfig)) {
                        dataField.SummarizeValuesBy = dfConfig.SummaryFunction;
                        if (!string.IsNullOrEmpty(dfConfig.Caption))
                            dataField.Name = dfConfig.Caption;
                    }
                    break;
                case PivotTableArea.Filter:
                    pivotTable.PageFields.Add(pivotField);
                    break;
            }
        }

        pivotTable.Style = workbook.TableStyles[BuiltInPivotStyleId.PivotStyleMedium2];

        pivotTable.EndUpdate();

        foreach (var kvp in config.ColumnFieldGroupings) {
            var pivotField = FindPivotField(pivotTable, kvp.Key);
            if (pivotField != null) {
                pivotField.GroupItems(kvp.Value.GroupBy);
            }
        }

        using var stream = new MemoryStream();
        workbook.SaveDocument(stream, DocumentFormat.Xlsx);
        return stream.ToArray();
    }

    private static PivotField? FindPivotField(PivotTable pivotTable, string fieldName) {
        for (int i = 0; i < pivotTable.Fields.Count; i++) {
            if (string.Equals(pivotTable.Fields[i].Name, fieldName, StringComparison.OrdinalIgnoreCase))
                return pivotTable.Fields[i];
        }
        return null;
    }
}

public class DataFieldConfig
{
    public PivotDataConsolidationFunction SummaryFunction { get; set; } = PivotDataConsolidationFunction.Sum;
    public string? Caption { get; set; }
}

public class ColumnFieldGroupingConfig
{
    public PivotFieldGroupByType GroupBy { get; set; }
}

public class PivotTableExportConfig
{
    public Dictionary<string, DataFieldConfig> DataFieldConfigs { get; set; } = new();
    public Dictionary<string, ColumnFieldGroupingConfig> ColumnFieldGroupings { get; set; } = new();
}
