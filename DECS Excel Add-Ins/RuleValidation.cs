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
                this.valid = false;
                this.message = ex.Message;
            }
        }

        public override string ToString()
        {
            return this.message;
        }

        internal bool Valid()
        {
            return this.valid;
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
            RuleType ruleType,
            int index,
            RuleComponent ruleComponent,
            string message
        )
        {
            this.ruleType = ruleType;
            this.index = index;
            this.ruleComponent = ruleComponent;
            this.message = message;
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

            if (string.IsNullOrEmpty(this.message))
            {
                explanation += ".";
            }
            else
            {
                explanation += " because " + this.message;
            }

            return explanation;
        }
    }
}
