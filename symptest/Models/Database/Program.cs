using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class Program
    {
        public Program(string program_name)
        {
            Program_name = program_name ?? throw new ArgumentNullException(nameof(program_name));
        }

        [DisplayName("Program Name")]
        public string Program_name { get; set; }
    }
}
