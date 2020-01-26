using System;

namespace Jbpc.Common.Excel
{
    public static class CellAddress
    {
        public static string A1_A1(int row, int col) => $"{ConvertColumnNumberToLetters(col)}:{row}";

        public static string ConvertColumnNumberToLetters(int col)
        {
            if (col>16384) throw new ApplicationException($"{col} > max number of Excel columns: 16,384");

            int dividend = col;
            string columnName = "";

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = $"{Convert.ToChar(65 + modulo)}{columnName}";
                dividend = (int) ((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
