using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ModbusTCP_Client
{
    class Brazil_Setup
    {
        public string Models { get; set; }
        public string Model { get; set; }
        public string Rated_Pressure { get; set; }
        public string Cmin { get; set; }
        public string Cmax { get; set; }
        public string Powermin { get; set; }
        public string Powermax { get; set; }
        public string Nozzel_Diameter { get; set; }
        public string Nozzel_Coefficient { get; set; }
        public string Oil_Type { get; set; }
    }

    class EmpConstants
    {
        private const string DOMAIN_NAME = "xyz.com";
    }
}
