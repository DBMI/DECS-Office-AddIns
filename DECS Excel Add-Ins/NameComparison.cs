using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DECS_Excel_Add_Ins
{
    // https://stackoverflow.com/a/78483552/20241849
    internal class NameComparison
    {
        private string desiredNameSorted;
        private Fastenshtein.Levenshtein lev;

        internal NameComparison(string desiredName)
        {
            string[] desiredWords = NormalizeAndSplitName(RemoveSalutations(desiredName));
            Array.Sort(desiredWords);
            desiredNameSorted = string.Join(" ", desiredWords);
            lev = new Fastenshtein.Levenshtein(desiredNameSorted);
        }

        public string FindBestMatch(List<string> names, double maxDistanceAllowed = 2.0)
        {
            double lowestScore = 1000000;
            string bestMatch = string.Empty;

            foreach (string thisName in names)
            {
                string[] words = NormalizeAndSplitName(RemoveSalutations(thisName));
                Array.Sort(words);
                string nameSorted = string.Join(" ", words);

                double wordLength = Math.Min(desiredNameSorted.Length, nameSorted.Length);
                int levenshteinDistance = lev.DistanceFrom(nameSorted);
                double relativeDistance = levenshteinDistance / wordLength;

                if (relativeDistance < lowestScore && relativeDistance < maxDistanceAllowed)
                {
                    bestMatch = thisName;
                    lowestScore = relativeDistance;

                    // No need to check further if we've found an exact match.
                    if (relativeDistance == 0) { break; }
                }
            }

            return bestMatch;
        }

        internal string[] NormalizeAndSplitName(string name)
        {
            name = Regex.Replace(name.Trim(), @"[\s,]+", " ").ToUpper();

            string[] words = name.Split(' ');

            if (words.Length == 2)
                return words;
            else if (words.Length == 3)
            {

                if (words[1].Length == 1)
                    return new string[] { words[0], words[2], words[1] };
                else
                    return words;
            }
            else if (words.Length == 4)
            {
                if (words[1].Length == 1 && words[3].Length == 1)
                    return new string[] { words[0], words[2], words[1], words[3] };
                else
                    return words;
            }
            else if (words.Length == 5 && words[2].Length == 1)
                return new string[] { words[0], words[2], words[1], words[3], words[4] };
            else
                return new string[0];
        }

        internal string RemoveSalutations(string name)
        {
            string[] salutations = { ", DO", ", LAC", ", LMFT", ", LMT", ", MD", ", MPH", ", NP", ", PA", ", PHD", ", PSYD", ", RN" };

            foreach (string salutation in salutations)
            {
                if (name.ToUpper().EndsWith(salutation.ToUpper()))
                {
                    name = name.Remove(name.Length - salutation.Length).TrimEnd();
                    break;
                }
            }
            return name;
        }
    }
}
