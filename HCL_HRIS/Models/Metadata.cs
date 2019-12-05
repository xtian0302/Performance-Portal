using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace HCL_HRIS.Models
{
    public class userMetadata
    { 
        [Key]
        public int user_id { get; set; }
        [Required]
        [Index(IsUnique = true)]
        [Display(Name = "SAP Number")]
        public int sap_id { get; set; } 
        [Display(Name = "Full Name")]
        public string name { get; set; }
        [Required]
        [DataType(DataType.Password)]
        [Display(Name = "Password")]
        public string password { get; set; }
    }

    public class chatMetadata
    { 
        [Key]
        public int chat_id { get; set; }
        [Required]
        [Display(Name = "Chat From")]
        public Nullable<int> chat_from { get; set; }
        [Required]
        [Display(Name = "Chat For")]
        public Nullable<int> chat_to { get; set; }
        [Required]
        [Display(Name = "Chat Message")]
        public string chat_message { get; set; }
        public Nullable<System.DateTime> datetime_sent { get; set; }
        public Nullable<System.DateTime> datetime_read { get; set; }
    }
    public class announcementMetadata
    {
        [Key]
        public int announcement_id { get; set; }

        [Display(Name = "Title")]
        public string announcement_title { get; set; }

        [Display(Name = "Details")]
        public string announcement_details { get; set; }

        [Display(Name = "Date")]
        public Nullable<System.DateTime> announcement_date { get; set; }
    }
}