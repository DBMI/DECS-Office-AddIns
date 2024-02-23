using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.UI;

namespace DECS_Excel_Add_Ins
{
    /// <summary>
    /// Enum to capture whether a GUI component holds @c NewColumn, @c Pattern or @c Replace definitions.
    /// </summary>
    public enum RuleComponent
    {
        NewColumn,
        Pattern,
        Replace
    }

    /// <summary>
    /// What's the purpose of this rule?
    /// </summary>
    public enum RuleType
    {
        Cleaning,
        Extract
    }

    /**
     * @brief Describes result of attempting to form a RegEx.
     */
    public class RuleValidationResult
    {
        private bool valid = true;
        private string message = string.Empty;

        /// <summary>
        /// Constructor: If @c ArgumentException is not null, then @c valid property is set to false and @c message holds the exception message.
        /// </summary>
        /// <param name="ex">ArgumentException object</param>
        public RuleValidationResult(ArgumentException ex = null)
        {
            if (ex != null)
            {
                valid = false;
                message = ex.Message;
            }
        }

        /// <summary>
        /// What's the result?
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return message;
        }

        /// <summary>
        /// Is the rule valid?
        /// </summary>
        /// <returns></returns>
        internal bool Valid()
        {
            return valid;
        }
    }

    /**
     * @brief Describes just why & where a rule has been found invalid.
     */ 
    public class RuleValidationError
    {
        private RuleType ruleType;
        private int index;
        private string message;
        private RuleComponent ruleComponent;

        public RuleValidationError(
            RuleType _ruleType,
            int _index,
            RuleComponent _ruleComponent,
            string _message
        )
        {
            ruleType = _ruleType;
            index = _index;
            ruleComponent = _ruleComponent;
            message = _message;
        }

        public override string ToString()
        {
            string explanation =
                ruleType.ToString()
                + " rule #"
                + index.ToString()
                + " has invalid "
                + ruleComponent.ToString()
                + " component";

            if (string.IsNullOrEmpty(message))
            {
                explanation += ".";
            }
            else
            {
                explanation += " because " + message;
            }

            return explanation;
        }
    }
}
