using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Jbpc.Common.Excel;

namespace Jbpc.Common.Excel.ExtensionMethods
{
    public static class RangeCellOffsets
    {
        public static string A1_A1(this Range rngHome)
        {
            var start = rngHome.A1();
            var end = rngHome.A1(rngHome.Rows.Count - 1, rngHome.Columns.Count - 1);

            return $"{start}:{end}";
        }
        public static string A1(this Range rngHome)
        {
            return rngHome.A1(0, 0);
        }
        private static string A1(this Range rngRange, int rowOffset, int colOffset)
        {
            int row = rngRange.Row + rowOffset;
            int col = rngRange.Column + colOffset;

            if (row>1048576) throw new ApplicationException($"{row} > max number of excel rows: 1,048,576");

            return $"{CellAddress.ConvertColumnNumberToLetters(col)}{row}";
        }
    }
}
