﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace HCL_HRIS.Models
{
    [MetadataType(typeof(userMetadata))]
    public partial class user
    {
    }
    [MetadataType(typeof(chatMetadata))]
    public partial class chat
    {
    }
    [MetadataType(typeof(announcementMetadata))]
    public partial class announcement
    {
    }
    [MetadataType(typeof(groupMetadata))]
    public partial class group
    {
    }
    [MetadataType(typeof(trackMetadata))]
    public partial class track
    {
    }
    [MetadataType(typeof(auditMetadata))]
    public partial class audit
    {
    }
}