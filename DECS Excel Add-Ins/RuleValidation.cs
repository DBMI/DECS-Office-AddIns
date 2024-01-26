using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;

namespace DECS_Excel_Add_Ins
{
    public enum RuleComponent
    {
        NewColumn,
        Pattern,
        Replace
    }

    public enum RuleType
    {
        Cleaning,
        Extract
    }

    // Describes result of attempting to form a RegEx.
    public class RuleValidationResult
    {
        private bool valid = true;
        private string message = string.Empty;

        public RuleValidationResult(ArgumentException ex = null)
        {
            if (ex != null)
            {
                valid = false;
                message = ex.Message;
            }
        }

        public override string ToString()
        {
            return message;
        }

        internal bool Valid()
        {
            return valid;
        }
    }

    // Describes just why & where a rule has been found invalid.
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
