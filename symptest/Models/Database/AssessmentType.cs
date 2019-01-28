using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class AssessmentType
{
        public AssessmentType(string type)
        {
            Type = type ?? throw new ArgumentNullException(nameof(type));
        }

        [DisplayName("Assessment Type")]
        public string Type { get; set; }
    }
}
