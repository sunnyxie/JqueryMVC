using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace JqueryMVC2.Models
{
    public class Movie
    {
        public int ID { get; set; }

        public string Name { get; set; }

        public double Score { get; set; }

        public string PicturePath { get; set; }

        //[DataType(DataType.Date)]
        //[DisplayFormat(DataFormatString ="{0:yyyy-MM-dd}",ApplyFormatInEditMode = true)]
        public DateTime? DateOfBirth { get; set; }
    }
}