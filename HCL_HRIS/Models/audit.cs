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
    
    public partial class audit
    {
        public int audit_id { get; set; }
        public string case_id { get; set; }
        public string audit_type { get; set; }
        public string type_of_monitoring { get; set; }
        public Nullable<int> auditor_sap { get; set; }
        public Nullable<int> agent_sap { get; set; }
        public Nullable<System.DateTime> audit_date { get; set; }
        public Nullable<System.DateTime> transaction_date { get; set; }
        public Nullable<bool> bc { get; set; }
        public Nullable<bool> euc { get; set; }
        public Nullable<bool> cc { get; set; }
        public Nullable<bool> fatal { get; set; }
    }
}