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
        public int group_id { get; set; }
        public string user_role { get; set; }
    
        public virtual group group { get; set; }
    }
}
