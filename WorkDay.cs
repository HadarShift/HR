using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HR
{
    class WorkDay
    {
        public DateTime date { get; set; }
        public bool arriveToWork { get; set; }
        public string absenceReason { get; set; }
        public WorkDay()
        {
            absenceReason = "";
            arriveToWork = false;

        }
        public bool checkIfWeekEnd()
        {
            return (int)date.DayOfWeek >= 5;
        }
    }
}
