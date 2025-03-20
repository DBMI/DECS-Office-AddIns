using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace DECS_Excel_Add_Ins
{
    public enum ThresholdCondition
    {
        [Description("<")]
        lt,

        [Description("≤")]
        lte,

        [Description("Unknown")]
        Unknown
    }

    public class FollowUpTimeframeThresholds
    {
        public int highUrgencyUpperThresholdValue { get; set; }
        public ThresholdCondition highUrgencyUpperThresholdCondition { get; set; }
        public int mediumUrgencyUpperThresholdValue { get; set; }
        public ThresholdCondition mediumUrgencyUpperThresholdCondition { get; set; }

        public FollowUpTimeframeThresholds() 
        {
            // Set default values.
            highUrgencyUpperThresholdValue = 2;
            highUrgencyUpperThresholdCondition = ThresholdCondition.lt;
            mediumUrgencyUpperThresholdValue = 4;
            mediumUrgencyUpperThresholdCondition = ThresholdCondition.lte;
        }

        public FollowUpTimeframeThresholds(int huutv, int muutv, ThresholdCondition huutc, ThresholdCondition muutc)
        {
            highUrgencyUpperThresholdValue = huutv;
            mediumUrgencyUpperThresholdValue = muutv;
            highUrgencyUpperThresholdCondition = huutc;
            mediumUrgencyUpperThresholdCondition = muutc;
        }

        /// <summary>
        /// Use threshold values and conditions to turn a TimeSpan into a priority level.
        /// </summary>
        /// <param name="timeSpan">Time between when note written and when action scheduled</param>
        /// <returns>@c TriagePriority</returns>

        internal TriagePriority ParsePriority(TimeSpan timeSpan)
        {
            double deltaWeeks = (timeSpan.Days / 7.0);

            if (mediumUrgencyUpperThresholdCondition == ThresholdCondition.lt)
            {
                if (deltaWeeks >= mediumUrgencyUpperThresholdValue)
                {
                    return TriagePriority.Routine;
                }
            }
            else
            {
                if (deltaWeeks > mediumUrgencyUpperThresholdValue)
                {
                    return TriagePriority.Routine;
                }
            }

            if (highUrgencyUpperThresholdCondition == ThresholdCondition.lt)
            {
                if (deltaWeeks >= highUrgencyUpperThresholdValue)
                {
                    return TriagePriority.Medium;
                }

                return TriagePriority.High;
            }
            else
            {
                if (deltaWeeks > highUrgencyUpperThresholdValue)
                {
                    return TriagePriority.Medium;
                }

                return TriagePriority.High;
            }
        }

        /// <summary>
        /// Use threshold values and conditions to turn # of weeks into a priority level.
        /// </summary>
        /// <param name="numWeeks"># weeks between when note written and when action scheduled</param>
        /// <returns>@c TriagePriority</returns>

        internal TriagePriority ParsePriority(int numWeeks)
        {
            if (mediumUrgencyUpperThresholdCondition == ThresholdCondition.lt)
            {
                if (numWeeks >= mediumUrgencyUpperThresholdValue)
                {
                    return TriagePriority.Routine;
                }
            }
            else
            {
                if (numWeeks > mediumUrgencyUpperThresholdValue)
                {
                    return TriagePriority.Routine;
                }
            }

            if (highUrgencyUpperThresholdCondition == ThresholdCondition.lt)
            {
                if (numWeeks >= highUrgencyUpperThresholdValue)
                {
                    return TriagePriority.Medium;
                }

                return TriagePriority.High;
            }
            else
            {
                if (numWeeks > highUrgencyUpperThresholdValue)
                {
                    return TriagePriority.Medium;
                }

                return TriagePriority.High;
            }
        }

        /// <summary>
        /// Reads the stored config file & parses it into a @c FollowUpTimeframeThresholds object.
        /// </summary>
        /// <param name="filePath">Full path to config file</param>
        /// <returns>@c FollowUpTimeframeThresholds</returns>

        internal static FollowUpTimeframeThresholds ReadConfigFile(string filePath)
        {
            // Declare this variable outside the 'using' block so we can access it later
            FollowUpTimeframeThresholds config = new FollowUpTimeframeThresholds();

            if (string.IsNullOrEmpty(filePath))
            {
                return config;
            }

            using (var reader = new StreamReader(filePath))
            {
                try
                {
                    config = (FollowUpTimeframeThresholds)new XmlSerializer(typeof(FollowUpTimeframeThresholds)).Deserialize(reader);
                }
                catch (InvalidOperationException)
                {
                    // Looks like the xml file is not as expected--
                    // so skip it and return the default values.
                }
            }

            return config;
        }

        /// <summary>
        /// Writes @c FollowUpTimeframeThresholds object to a config file.
        /// </summary>
        /// <param name="filePath">Full path to config file</param>
        /// <returns>void</returns>

        internal void WriteConfigFile(string filePath)
        {
            using (var writer = new System.IO.StreamWriter(filePath))
            {
                var serializer = new XmlSerializer(typeof(FollowUpTimeframeThresholds));
                serializer.Serialize(writer, this);
                writer.Flush();
            }
        }
    }


    internal class TimeSorterSettings
    {
        private string configFilepath;
        private FollowUpTimeframeThresholds thresholds;

        internal TimeSorterSettings()
        {
            thresholds = new FollowUpTimeframeThresholds();

            Microsoft.Office.Interop.Excel.Application application = Globals.ThisAddIn.Application;

            // Read the config file.
            configFilepath = Path.Combine(application.DefaultFilePath, "urgency_thresholds.xml");

            if (File.Exists(configFilepath))
            {
                thresholds = FollowUpTimeframeThresholds.ReadConfigFile(configFilepath);
            }
        }

        internal FollowUpTimeframeThresholds Set()
        {
            using (ChooseTimeThresholdsForm form = new ChooseTimeThresholdsForm(thresholds))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    thresholds = new FollowUpTimeframeThresholds(form.highUpperThresholdValue, 
                                                                 form.mediumUpperThresholdValue,
                                                                 form.highUpperThresholdCondition,
                                                                 form.mediumUpperThresholdCondition);
                    thresholds.WriteConfigFile(configFilepath);
                }
            }

            return thresholds;
        }

        internal FollowUpTimeframeThresholds Thresholds()
        {
            return thresholds;
        }
    }
}
