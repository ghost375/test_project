using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject1
{
   class Contact
    {
        public int Vid { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string LifeCycleStage { get; set; }
        public Company Company { get; set; }
    }
}
