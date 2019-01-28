using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class PPSRDetailRecord
{
        public PPSRDetailRecord(int number, int serial_number, string statements, int order_number, string letter_identifier, int scale_number)
        {
            Number = number;
            Serial_number = serial_number;
            Statements = statements ?? throw new ArgumentNullException(nameof(statements));
            Order_number = order_number;
            Letter_identifier = letter_identifier ?? throw new ArgumentNullException(nameof(letter_identifier));
            Scale_number = scale_number;
        }

        [DisplayName("PPSR Detail Record Number")]
        public int Number { get; set; }

        [DisplayName("PPSR Serial Number")]
        public int Serial_number { get; set; }

        [DisplayName("PPSR Statements")]
        public string Statements { get; set; }

        [DisplayName("Order Number")]
        public int Order_number { get; set; }

        [DisplayName("Letter Identifier")]
        public string Letter_identifier { get; set; }

        [DisplayName("Scale Number")]
        public int Scale_number { get; set; }


    }
}
