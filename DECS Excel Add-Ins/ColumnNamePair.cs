using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DECS_Excel_Add_Ins
{
    internal class ColumnNamePair
    {
        private string _name1;
        private string _name2;
        private string _nameInCommon;

        internal ColumnNamePair(string name1, string name2, string nameInCommon)
        {
            _name1 = name1;
            _name2 = name2;
            _nameInCommon = nameInCommon;
        }

        internal string CommonName()
        {
            return _nameInCommon;
        }

        internal bool Contains(string name)
        {
            return _name1.Contains(name) || _name2.Contains(name);
        }

        internal string Name1()
        {
            return _name1;
        }

        internal string Name2()
        {
            return _name2;
        }
    }

    internal class ColumnNamePairs 
    {
        private List<ColumnNamePair> _pairs;

        internal ColumnNamePairs()
        {
            _pairs = new List<ColumnNamePair>();
        }
        
        internal ColumnNamePairs(List<string> columnNames, List<string> ignoredWords)
        {
            _pairs = new List<ColumnNamePair>();
            List<string> rangeNames = Utilities.DistinctElements(columnNames, ignoredWords);

            foreach (string name in rangeNames)
            {
                string candidateStartName = name + " Start Date";
                string candidateEndName = name + " End Date";

                if (columnNames.Contains(candidateStartName) && columnNames.Contains(candidateEndName))
                {
                    _pairs.Add(new ColumnNamePair(candidateStartName, candidateEndName, name));
                }
            }
        }

        internal void Clear()
        {
            _pairs.Clear();
        }

        internal bool Contains(string name)
        {
            return _pairs.Any(p => p.Contains(name));
        }

        internal int Count()
        {
            return _pairs.Count;
        }

        internal List<string> GetColumnNames()
        {
            List<string> names = new List<string>();

            foreach (ColumnNamePair pair in _pairs)
            {
                names.Add(pair.Name1());
                names.Add(pair.Name2());
            }

            return names;
        }

        internal List<ColumnNamePair> GetColumnPairs()
        {
            return _pairs;
        }

        internal void Remove(string name)
        {
            var itemToRemove = _pairs.SingleOrDefault(p => p.Contains(name));

            if (itemToRemove != null)
                _pairs.Remove(itemToRemove);
        }
    }
}
