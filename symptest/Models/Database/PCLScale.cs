using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class PCLScale
{
        public PCLScale(int number, string description)
        {
            Number = number;
            Description = description ?? throw new ArgumentNullException(nameof(description));
        }

        [DisplayName("PCL Scale Number")]
        public int Number { get; set; }

        [DisplayName("Scale Description")]
        public string Description { get; set; }
    }
}
