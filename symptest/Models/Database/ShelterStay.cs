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
            this.location = location ?? throw new ArgumentNullException(nameof(location));
            ClientID = clientID ?? throw new ArgumentNullException(nameof(clientID));
            this.enrollment_date = enrollment_date;
            this.exit_date = exit_date ?? throw new ArgumentNullException(nameof(exit_date));
        }

        [DisplayName("Shelter Location")]
        public string location { get; set; }

        [DisplayName("Client ID")]
        public string ClientID { get; set; }

        [DisplayName("Shelter Enrollment Date")]
        public DateTime enrollment_date { get; set; }

        [DisplayName("Shelter Exit Date")]
        public DateTime? exit_date { get; set; }





    }
}
