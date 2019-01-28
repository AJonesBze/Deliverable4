using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class CFARSDetailRecord
{
        public CFARSDetailRecord(int number, int scale_number, int assessment_number, string keyword)
        {
            Number = number;
            Scale_number = scale_number;
            Assessment_number = assessment_number;
            Keyword = keyword ?? throw new ArgumentNullException(nameof(keyword));
        }

        [DisplayName("CFARS Detail Record Number")]
        public int Number { get; set; }

        [DisplayName("CFARS Scale Number")]
        public int Scale_number { get; set; }

        [DisplayName("CFARS Assessment Number")]
        public int Assessment_number { get; set; }

        [DisplayName("Keyword")]
        public string Keyword { get; set; }





    }
}
