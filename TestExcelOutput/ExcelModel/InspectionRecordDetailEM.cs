using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficePositionAttributes;

namespace ExcelView
{
    class InspectionRecordDetailEM
    {
        private readonly int StartRow = 6;


        [ExcelColPosition(Col = 4)]
        public string Quantity { get; set; }

        [ExcelColPosition(Col = 5)]
        public string Weight { get; set; }

        [ExcelColPosition(Col = 6)]
        public string UnitPrice { get; set; }
    }
}
