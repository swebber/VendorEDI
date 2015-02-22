using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VendorEDI
{
    class VendorItem
    {
        public string Skn { get; set; }
        public string VendorSku { get; set; }
        public string Description { get; set; }
        public decimal ItemCost { get; set; }
        public decimal PaidToVendor { get; set; }

        public long UnitsShipped { get; set; }
        public bool IsDuplicateSkn { get; set; }
    }
}
