using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class PPSRStatements
{
        public PPSRStatements(int order_number, string statements, string letter_identifier)
        {
            Order_number = order_number;
            Statements = statements ?? throw new ArgumentNullException(nameof(statements));
            Letter_identifier = letter_identifier ?? throw new ArgumentNullException(nameof(letter_identifier));
        }

        [DisplayName("Order Number")]
        public int Order_number { get; set; }

        [DisplayName("PPSR Statements")]
        public string Statements { get; set; }

        [DisplayName("Letter Identifier")]
        public string Letter_identifier { get; set; }
    }
}
