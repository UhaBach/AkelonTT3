using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTT3
{
    internal class XMLFileWriter
    {
        private XLWorkbook wb;

        public XMLFileWriter(XLWorkbook wb)
        {
            this.wb = wb;
        }

        public void WtiteDataInCell(IXLCell cell, XLCellValue cellValue)
        {
            cell.SetValue(cellValue);
        }

        public void SaveFileData()
        {
            wb.Save();
        }
    }
}
