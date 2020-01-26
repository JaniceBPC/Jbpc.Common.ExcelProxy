using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Jbpc.Common.Excel
{
    public class WorksheetDataTable
    {
        private object[,] matrix;
        private readonly  DataTable dataTable = new DataTable();
        public DataTable CopyUsedRangeIntoDataTable(string fullyQualifiedWorkbookName, string worksheetName = null)
        {
            matrix = WorksheetValues.GetEntireSheet(fullyQualifiedWorkbookName, worksheetName);

            AddDataTableColumn();

            PopulateRows();

            return dataTable;
        }
        private void PopulateRows()
        {
            var names = ColumnHeaders();

            for (int i = 2; i < matrix.GetLength(0)+1; i++)
            {
                var dataRow = dataTable.Rows.Add();

                for (int j = 1; j < names.Count+1; j++)
                {
                    var cell = matrix[i, j];
                    var columnName = names[j - 1];

                    dataRow[columnName] = cell == null ? DBNull.Value : Convert.ChangeType(cell, cell.GetType());
                }
            }
        }
        private void AddDataTableColumn()
        {
            var columnHeadings = ColumnHeaders();

            foreach (var columnName in columnHeadings)
                dataTable.Columns.Add(columnName);
        }
        private List<string> ColumnHeaders()
        {
            var list = new List<string>();

            for (int i = 1; i < matrix.GetLength(1)+1; i++)
            {
                if (matrix[1, i] == null)
                {
                    break;
                }
                list.Add(matrix[1, i].ToString());
            }

            var q = Enumerable.Range(1, list.Count).Zip(list, (x, y) => new {NthCol = x, Name = y}).ToList();

            var missingHeader = q.Where(x => x.Name == "").ToList();

            var msg = string.Join(", ", missingHeader.Select(x => $"{x.NthCol}) {x.Name}"));

            if (missingHeader.Any()) throw new ApplicationException($"Missing column headers for Excel import file={msg}");
            
            return list;
        }
    }
}
