using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficePositionAttributes;

namespace ExcelView
{
    class InspectionRecordEM
    {
        [ExcelCellPosition(Row = 5, Col = 4)]
        public string Quantity { get; set; }

        [ExcelCellPosition(Row = 5, Col = 5)]
        public string Weight { get; set; }

        [ExcelCellPosition(Row = 5, Col = 6)]
        public string UnitPrice { get; set; }
    }
}
