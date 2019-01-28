using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class ClientEthnicity
{
        public ClientEthnicity(string ethnicity)
        {
            Ethnicity = ethnicity ?? throw new ArgumentNullException(nameof(ethnicity));
        }

        [DisplayName("Ethnicity")]
        public string Ethnicity { get; set; }
    }
}
