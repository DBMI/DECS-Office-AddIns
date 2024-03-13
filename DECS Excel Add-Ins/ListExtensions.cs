using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DECS_Excel_Add_Ins
{
    public static class ListExtensions
    {
       public static List<string> Except(this List<string> list, string notThisOne)
        {
            List<string> remainingList = new List<string>();

            foreach(string item in list)
            {
                if (item == notThisOne)
                {
                    continue;
                }

                remainingList.Add(item);
            }

            return remainingList;
        }

        public static List<string> Except(this List<string> list, List<string> notThese)
        {
            List<string> remainingList = new List<string>();

            foreach (string item in list)
            {
                if (notThese.Contains(item))
                {
                    continue;
                }

                remainingList.Add(item);
            }

            return remainingList;
        }
    }
}
