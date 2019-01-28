using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class PPSRScale
{
        public PPSRScale(int scale_number, string description)
        {
            Scale_number = scale_number;
            Description = description ?? throw new ArgumentNullException(nameof(description));
        }

        [DisplayName("Scale Number")]
        public int Scale_number { get; set; }

        [DisplayName("Scale Description")]
        public string Description { get; set; }
    }
}
