//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace HCL_HRIS.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class user
    {
        public int user_id { get; set; }
        public int sap_id { get; set; }
        public string name { get; set; }
        public string password { get; set; }
        public Nullable<int> group_id { get; set; }
        public string user_role { get; set; }
        public string designation { get; set; }
        public string division { get; set; }
        public string status { get; set; }
        public string sub_department { get; set; }
        public string phase { get; set; }
        public string band { get; set; }
        public string tenurity { get; set; }
        public string hcl_hire_date { get; set; }
        public string abay_start_date { get; set; }
        public string cms_id { get; set; }
        public string citrix { get; set; }
        public string nt_login { get; set; }
        public string finesse_extension { get; set; }
        public string finesse_names { get; set; }
        public string finesse_enterprise_names { get; set; }
        public string badge_id { get; set; }
        public string birth_date { get; set; }
        public string address { get; set; }
        public string contact_number { get; set; }
        public string nda { get; set; }
        public string nho_policies_sign_off { get; set; }
        public string bgv { get; set; }
        public string versant { get; set; }
        public string typing { get; set; }
        public string aptitude { get; set; }
        public string group_policy { get; set; }
    
        public virtual group group { get; set; }
    }
}
