using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class ProgramEnrollment
{
        public ProgramEnrollment(string clientID, string name, DateTime enrollment_date, DateTime? exit_date)
        {
            ClientID = clientID ?? throw new ArgumentNullException(nameof(clientID));
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Enrollment_date = enrollment_date;
            Exit_date = exit_date ?? throw new ArgumentNullException(nameof(exit_date));
        }

        [DisplayName("Client ID")]
        public string ClientID { get; set; }

        [DisplayName("Program Name")]
        public string Name { get; set; }

        [DisplayName("Program Enrollment Date")]
        public DateTime Enrollment_date { get; set; }

        [DisplayName("Program Exit Date")]
        public DateTime? Exit_date { get; set; }

    }
}
