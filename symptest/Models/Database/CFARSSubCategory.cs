using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class CFARSSubCategory
{
        public CFARSSubCategory(string name, string category_name)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Category_name = category_name ?? throw new ArgumentNullException(nameof(category_name));
        }

        [DisplayName("Subcategory Name")]
        public string Name { get; set; }

        [DisplayName("Category Name")]
        public string Category_name { get; set; }
    }
}
