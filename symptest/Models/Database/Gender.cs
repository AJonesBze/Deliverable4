using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class Gender
{
        public Gender(string genderType)
        {
            GenderType = genderType ?? throw new ArgumentNullException(nameof(genderType));
        }

        [DisplayName("Gender")]
        public string GenderType { get; set; }
    }
}
