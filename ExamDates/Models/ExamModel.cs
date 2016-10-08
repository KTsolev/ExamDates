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
        public RegistrationDates RegPerFirstDate { get; set; }
        public DateTime SecondDate { get; set; }
        public RegistrationDates RegPerSecondDate { get; set; }
        public Boolean Overlap { get; set; }
        public String Session { get; set; }
    }

    public class RegistrationDates 
    {
        public DateTime RegStartDate { get; set; }
        public DateTime RegEndDate { get; set; }
    }
}