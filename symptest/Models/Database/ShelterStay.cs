using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class ShelterStay

    {
        public ShelterStay(string location, string clientID, DateTime enrollment_date, DateTime? exit_date)
        {
            this.Location = location ?? throw new ArgumentNullException(nameof(location));
            ClientID = clientID ?? throw new ArgumentNullException(nameof(clientID));
            this.Enrollment_date = enrollment_date;
            this.Exit_date = exit_date ?? throw new ArgumentNullException(nameof(exit_date));
        }

        [DisplayName("Shelter Location")]
        public string Location { get; set; }

        [DisplayName("Client ID")]
        public string ClientID { get; set; }

        [DisplayName("Shelter Enrollment Date")]
        public DateTime Enrollment_date { get; set; }

        [DisplayName("Shelter Exit Date")]
        public DateTime? Exit_date { get; set; }





    }
}
