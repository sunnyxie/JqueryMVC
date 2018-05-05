using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace JqueryMVC2.Models
{
    // The property Tax;
    public class InvestingTax
    {
        public int ID { get; set; }

         [Required]
         public int Type { get; set; }

         [Required]
         public decimal Price { get; set; }

        public double CommissionFee { get; set; }

        public int Count { get; set; }

        [Required]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString ="{0:yyyy-MM-dd}",ApplyFormatInEditMode = true)]
        public DateTime SettleDate { get; set; }

        [DataType(DataType.Date)]
        public DateTime TradeDate { get; set; }

        public double Score { get; set; }

        public string PicturePath { get; set; }


        public static Dictionary<int, string> TypeList = 
            new Dictionary<int, string> { { 1, "Buy" },
                                          { 2, "Sell" }
                };

        static InvestingTax()
        {
        
        }

    }
}