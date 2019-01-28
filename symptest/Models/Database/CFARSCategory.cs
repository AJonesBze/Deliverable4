using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class CFARSCategory
{
        public CFARSCategory(string name)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        [DisplayName("CFARS Category Name")]
        public string Name { get; set; }
    }
}
