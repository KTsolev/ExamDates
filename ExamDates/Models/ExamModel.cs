using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExamDates.Models
{
    public class ExamModel
    {
        public string Name { get; set; }
        public DateTime FirstDate { get; set; }
        public DateTime SecondDate { get; set; }
        public Boolean Overlap { get; set; }
    }
}