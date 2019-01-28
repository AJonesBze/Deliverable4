using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class CFARSKeyword
{
        public CFARSKeyword(string keyword, string definition, string subcategory_name)
        {
            Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
            Definition = definition ?? throw new ArgumentNullException(nameof(definition));
            Subcategory_name = subcategory_name ?? throw new ArgumentNullException(nameof(subcategory_name));
        }

        [DisplayName("CFARS Keyword")]
        public string Keyword { get; set; }

        [DisplayName("Definition")]
        public string Definition { get; set; }

        [DisplayName("Subcategory Name")]
        public string Subcategory_name { get; set; }












    }
}
