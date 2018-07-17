using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Invoice_Income_Correction
{
    public class PeymentsEntryReconsilation
    {
        public string PaymentDocEntry { get; set; }
        public string PaymentNumber { get; set; }
        public string JournalEntryNumber { get; set; }
        public string BpCardCode { get; set; }
    }
}
