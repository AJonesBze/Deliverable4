using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;

namespace HH_client_manager.Models.Database
{
    public class PCLDetailRecord
{
        public PCLDetailRecord(int number, int sequence_number, int scale_number, int bothered_by_order)
        {
            Number = number;
            Sequence_number = sequence_number;
            Scale_number = scale_number;
            Bothered_by_order = bothered_by_order;
        }

        [DisplayName("PCL Detail Record Number")]
        public int Number { get; set; }

        [DisplayName("PCL Sequence Number")]
        public int Sequence_number { get; set; }

        [DisplayName("Scale Number")]
        public int Scale_number { get; set; }

        [DisplayName("PCL Bothered By Order")]
        public int Bothered_by_order { get; set; }

    }
}
