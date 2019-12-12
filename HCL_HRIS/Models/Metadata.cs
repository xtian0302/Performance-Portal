using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Web.Script.Serialization;

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
        [Display(Name = "Team")]
        public Nullable<int> group_id { get; set; }  
        [Display(Name = "Role")]
        public string user_role { get; set; } 
        public virtual group group { get; set; }
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

    public class groupMetadata
    {
        public int group_id { get; set; }
        [Display(Name = "Name")]
        public string group_name { get; set; }
        [Display(Name = "Leader")]
        public Nullable<int> group_leader { get; set; }
        [Display(Name = "Track")]
        public Nullable<int> track_id { get; set; } 
        public virtual user user { get; set; } 
        public virtual track track { get; set; }
    }
    public class trackMetadata
    {
        public int track_id { get; set; }
        [Display(Name = "Name")]
        public string track_name { get; set; }
        [Display(Name = "Manager")]
        public Nullable<int> track_manager { get; set; } 
        public virtual user user { get; set; }
    }
    public class auditMetadata
    {
        public int audit_id { get; set; }
        [Required]
        [Display(Name = "Case ID")]
        public string case_id { get; set; }
        [Required]
        [Display(Name = "Audit Type")]
        public string audit_type { get; set; }
        [Required]
        [Display(Name = "Type of Monitoring")]
        public string type_of_monitoring { get; set; }
        [Required]
        [Display(Name = "Auditor")]
        public Nullable<int> auditor_sap { get; set; }
        [Required]
        [Display(Name = "Agent")]
        public Nullable<int> agent_sap { get; set; }
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "Date of Audit")]
        public Nullable<System.DateTime> audit_date { get; set; }
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        [Display(Name = "Date of Transaction")]
        public Nullable<System.DateTime> transaction_date { get; set; }
        [Display(Name = "B.C. Fail?")]
        public Nullable<bool> bc { get; set; }
        [Display(Name = "E.U.C. Fail?")]
        public Nullable<bool> euc { get; set; }
        [Display(Name = "C.C. Fail?")]
        public Nullable<bool> cc { get; set; }
        [Display(Name = "Fatal Accuracy Fail?")]
        public Nullable<bool> fatal { get; set; }
    }
}
