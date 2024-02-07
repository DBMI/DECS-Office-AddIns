using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Computes estimated time of completion.
     * https://stackoverflow.com/a/6822458/18749636
     */
    static class StopWatchUtils
    {
        /// <summary>
        /// Computes estimated time of completion.
        /// </summary>
        /// <param name="sw">@c Stopwatch object</param>
        /// <param name="counter">int Number of interations completed</param>
        /// <param name="counterGoal">int Number of interations to be performed</param>
        /// <returns>@c TimeSpan</returns>
        public static TimeSpan GetEta(this Stopwatch sw, int counter, int counterGoal)
        {
            /* 
             * (TimeTaken / linesProcessed) * linesLeft=timeLeft
             *
             * pulled from http://stackoverflow.com/questions/473355/calculate-time-remaining/473369#473369
             */

            if (counter == 0)
                return TimeSpan.Zero;

            float elapsedMin = ((float)sw.ElapsedMilliseconds / 1000) / 60;
            float minLeft = (elapsedMin / counter) * (counterGoal - counter);
            TimeSpan ret = TimeSpan.FromMinutes(minLeft);
            return ret;
        }
    }
}
