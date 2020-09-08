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
        public int RowIndex { get; set; }

        [ExcelColPosition(Col = 1)]
        public string GoodName { get; set; }

        [ExcelColPosition(Col = 2)]
        public string ReceptionTime { get; set; }

        [ExcelColPosition(Col = 3)]
        public string PackingRemark { get; set; }
    }
}
