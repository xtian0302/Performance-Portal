using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace HCL_HRIS.Models
{
    public class userMetadata
    { 
        [Key]
        public int user_id { get; set; }
        [Display(Name = "SAP Number")]
        public int sap_id { get; set; }
        [Display(Name = "Full Name")]
        public string name { get; set; }
        [DataType(DataType.Password)]
        [Display(Name = "Password")]
        public string password { get; set; }
    }
}