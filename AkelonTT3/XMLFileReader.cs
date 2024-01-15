using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTT3
{
    internal class XMLFileReader
    {
        private XLWorkbook wb;

        public XMLFileReader(XLWorkbook wb)
        {
            this.wb = wb;
        }

        public IEnumerable<IXLRangeRow> GetSheetRows(string sheetName)
        {
            // Получаем таблицу
            var ws = wb.Worksheet(sheetName);
            // вытаскиваем все заполненные строки кроме шапки
            var rows = ws.RangeUsed().RowsUsed().Skip(1);
            return rows;
        }
    }
}
