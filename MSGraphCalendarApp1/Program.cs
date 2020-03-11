using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSGraphCalendarApp1
{
    class Program
    {
        static async Task Main(string[] args)
        {
            GraphHelper.CreateCalendar();
            //GraphHelper.CreateEvents();
            //await GraphHelper.GetEventsAsync();
            //GraphHelper.GetUserDetailsAsync()
            Console.ReadKey();
        }
    }
}
