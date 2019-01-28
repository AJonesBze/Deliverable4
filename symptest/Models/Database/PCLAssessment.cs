using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class PCLAssessment
{
        public PCLAssessment(int sequence_number, string clientID, DateTime date, string type, string event_description, DateTime? event_date, string edited_by, DateTime edit_date)
        {
            Sequence_number = sequence_number;
            ClientID = clientID ?? throw new ArgumentNullException(nameof(clientID));
            Date = date;
            Type = type ?? throw new ArgumentNullException(nameof(type));
            Event_description = event_description ?? throw new ArgumentNullException(nameof(event_description));
            Event_date = event_date ?? throw new ArgumentNullException(nameof(event_date));
            Edited_by = edited_by ?? throw new ArgumentNullException(nameof(edited_by));
            Edit_date = edit_date;
        }

        [DisplayName("PCL Sequence Number")]
        public int Sequence_number { get; set; }

        [DisplayName("Client ID")]
        public string ClientID { get; set; }

        [DisplayName("Assessment Date")]
        public DateTime Date { get; set; }

        [DisplayName("Assessment Type")]
        public string Type { get; set; }

        [DisplayName("Event Description")]
        public string Event_description { get; set; }

        [DisplayName("Event Date")]
        public DateTime? Event_date { get; set; }

        [DisplayName("Edited By")]
        public string Edited_by { get; set; }

        [DisplayName("Edit Date")]
        public DateTime Edit_date { get; set; }





    }
}
