using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class ClientRace
{
        public ClientRace(string race)
        {
            Race = race ?? throw new ArgumentNullException(nameof(race));
        }

        [DisplayName("Client Race")]
    public string Race { get; set; }
}
}
