using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VendorEDI
{
    class VendorItemMap : CsvClassMap<VendorItem>
    {
        public VendorItemMap()
        {
            Map(m => m.Skn).Index(0);
            Map(m => m.VendorSku).Index(1);
            Map(m => m.Description).Index(2);
            Map(m => m.ItemCost).Index(3);
            Map(m => m.PaidToVendor).Index(4);
        }
    }
}
