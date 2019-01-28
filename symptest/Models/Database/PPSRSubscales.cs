using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class PPSRSubscales
{
        public PPSRSubscales(string letter_identifier, string description)
        {
            Letter_identifier = letter_identifier ?? throw new ArgumentNullException(nameof(letter_identifier));
            Description = description ?? throw new ArgumentNullException(nameof(description));
        }

        [DisplayName("Letter Identifier")]
        public string Letter_identifier { get; set; }

        [DisplayName("Subscale Description")]
        public string Description { get; set; }
    }
}
