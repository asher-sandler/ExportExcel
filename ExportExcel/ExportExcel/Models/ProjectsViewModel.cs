using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportExcel.Models
{
    public class ProjectsViewModel
    {
        public int PROJECT_ID { get; set; }
        public string EMPLOYEE_NAME { get; set; }
        public System.DateTime WORK_DATE { get; set; }
        public string CLIENT_NAME { get; set; }
        public string WORK_DESCRIPTION { get; set; }
        public Nullable<decimal> WORK_HOURS { get; set; }
        public Nullable<decimal> HOUR_COST { get; set; }
        // public int? DAYOFWEEK { get; set; }
    }
}