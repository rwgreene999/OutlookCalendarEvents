using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCalendarEvents
{
    class RunOptions
    {
        public int secondsBetweenCheck { get; set; } = 60000; 
        public DateTime nextCheck { get; set; } = DateTime.MinValue; 
    }
}
