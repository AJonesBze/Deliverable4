using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class Shelter
{
        public Shelter(string location)
        {
            Location = location ?? throw new ArgumentNullException(nameof(location));
        }

        [DisplayName("Shelter Location")]
        public string Location { get; set; }
    }
}
