using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication1.BLL
{
    class WORK_ORDER
    {
        private static String strSQL;
       
        //     Get Work Order
        public static int OrdNo { get; set; }
        public static string UserOrderNo { get; set; }
        public static decimal OrdQty { get; set; }
        public static string OrdStatus { get; set; }
        public static string UserPartNo { get; set; }
        public static string PartDesc { get; set; }
        public static int PartNo { get; set; }
        public static int CustNo { get; set; }
        public static string UserCustNo { get; set; }

    }
}
