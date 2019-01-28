using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class Client
    {
        public Client(string clientID, DateTime date_of_birth, string relationship_status, string advocate_name, string notes, string gender, string ethnicity, string race, string partner_gender)
        {
            ClientID = clientID ?? throw new ArgumentNullException(nameof(clientID));
            Date_of_birth = date_of_birth;
            Relationship_status = relationship_status ?? throw new ArgumentNullException(nameof(relationship_status));
            Advocate_name = advocate_name ?? throw new ArgumentNullException(nameof(advocate_name));
            Notes = notes ?? throw new ArgumentNullException(nameof(notes));
            Gender = gender ?? throw new ArgumentNullException(nameof(gender));
            Ethnicity = ethnicity ?? throw new ArgumentNullException(nameof(ethnicity));
            Race = race ?? throw new ArgumentNullException(nameof(race));
            Partner_gender = partner_gender ?? throw new ArgumentNullException(nameof(partner_gender));
        }

        [DisplayName("Client ID")]
        public string ClientID { get; set; }

        [DisplayName("Date of Birth")]
        public DateTime Date_of_birth { get; set; }

        [DisplayName("Relationship Status")]
        public string Relationship_status { get; set; }

        [DisplayName("Advocate Name")]
        public string Advocate_name { get; set; }

        [DisplayName("Notes on Client")]
        public string Notes { get; set; }

        [DisplayName("Gender")]
        public string Gender { get; set; }

        [DisplayName("Ethnicity")]
        public string Ethnicity { get; set; }

        [DisplayName("Race")]
        public string Race { get; set; }

        [DisplayName("Partner Gender")]
        public string Partner_gender { get; set; }

    }

}
