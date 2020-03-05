using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;
using System.Web.Mvc;
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
        [ScriptIgnore]
        [Display(Name = "Manager")]
        public virtual group group { get; set; }
        [Display(Name = "Designation")]
        public string designation { get; set; }
        [Display(Name = "Division")]
        public string division { get; set; }
        [Display(Name = "Status")]
        public string status { get; set; }

        [Display(Name = "Sub Department")]
        public string sub_department { get; set; }
        [Display(Name = "Phase")]
        public string phase { get; set; }
        [Display(Name = "Band")]
        public string band { get; set; }
        [Display(Name = "Tenurity")]
        public string tenurity { get; set; }
        [Display(Name = "Hire Date")]
        public string hcl_hire_date { get; set; }
        [Display(Name = "ABAY Start Date")]
        public string abay_start_date { get; set; }
        [Display(Name = "CMS ID")]
        public string cms_id { get; set; }
        [Display(Name = "Citrix")]
        public string citrix { get; set; }
        [Display(Name = "NT Login")]
        public string nt_login { get; set; }
        [Display(Name = "Finesse Extension")]
        public string finesse_extension { get; set; }
        [Display(Name = "Finesse Names")]
        public string finesse_names { get; set; }
        [Display(Name = "Finesse Enterprise Names")]
        public string finesse_enterprise_names { get; set; }
        [Display(Name = "Badge ID")]
        public string badge_id { get; set; }
        [Display(Name = "Birth Date")]
        public string birth_date { get; set; }
        [Display(Name = "Address")]
        public string address { get; set; }
        [Display(Name = "Contact Number")]
        public string contact_number { get; set; }
        [Display(Name = "NDA")]
        public string nda { get; set; }
        [Display(Name = "NHO Policies Sign Off")]
        public string nho_policies_sign_off { get; set; }
        [Display(Name = "BGV")]
        public string bgv { get; set; }
        [Display(Name = "Versant")]
        public string versant { get; set; }
        [Display(Name = "Typing")]
        public string typing { get; set; }
        [Display(Name = "Aptitude")]
        public string aptitude { get; set; }
        [Display(Name = "Group Policy")]
        public string group_policy { get; set; }
        [HiddenInput(DisplayValue = false)]
        public Nullable<bool> agreement_read { get; set; }
        [Display(Name = "Email")]
        public string email { get; set; }
        [Display(Name = "C201")]
        public string C201 { get; set; }
        [Display(Name = "Basic Requirements")]
        public string basicreq { get; set; }
        [Display(Name = "BGV Results")]
        public string bgv_result { get; set; }
        [Display(Name = "Police / NBI Clearance")]
        public string police_nbi { get; set; }
        [Display(Name = "Versant TIN")]
        public string versant_tin { get; set; }
        [Display(Name = "BGV Status")]
        public string bgv_status { get; set; }
         
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
