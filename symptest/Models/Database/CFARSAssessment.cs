using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class CFARSAssessment
{
        public CFARSAssessment(int serial_number, DateTime date, string clientID, DateTime preassessment_date, string type, string program_name, string conducted_by, int number_of_group_sessions, int number_of_advocacy_meetings, int number_of_counseling_sessions, string edited_by, DateTime? edit_date)
        {
            Serial_number = serial_number;
            Date = date;
            ClientID = clientID ?? throw new ArgumentNullException(nameof(clientID));
            Preassessment_date = preassessment_date;
            Type = type ?? throw new ArgumentNullException(nameof(type));
            Program_name = program_name ?? throw new ArgumentNullException(nameof(program_name));
            Conducted_by = conducted_by ?? throw new ArgumentNullException(nameof(conducted_by));
            Number_of_group_sessions = number_of_group_sessions;
            Number_of_advocacy_meetings = number_of_advocacy_meetings;
            Number_of_counseling_sessions = number_of_counseling_sessions;
            Edited_by = edited_by ?? throw new ArgumentNullException(nameof(edited_by));
            Edit_date = edit_date ?? throw new ArgumentNullException(nameof(edit_date));
        }

        [DisplayName("CFARS Serial Number")]
        public int Serial_number { get; set; }

        [DisplayName("Assessment Date")]
        public DateTime Date { get; set; }

        [DisplayName("Client ID")]
        public string ClientID { get; set; }

        [DisplayName("Pre-Assessment Date")]
        public DateTime Preassessment_date { get; set; }

        [DisplayName("Assessment Type")]
        public string Type { get; set; }

        [DisplayName("Program Name")]
        public string Program_name { get; set; }

        [DisplayName("Conducted By")]
        public string Conducted_by { get; set; }

        [DisplayName("Number of Group Sessions")]
        public int Number_of_group_sessions { get; set; }

        [DisplayName("Number of Advocacy Meetings")]
        public int Number_of_advocacy_meetings { get; set; }

        [DisplayName("Number of Counseling Meetings")]
        public int Number_of_counseling_sessions { get; set; }

        [DisplayName("Edited By")]
        public string Edited_by { get; set; }

        [DisplayName("Edit Date")]
        public DateTime? Edit_date { get; set; }





    }
}
