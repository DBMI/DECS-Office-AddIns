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
    // Describes just why & where a rule has been found invalid.
    public class RuleValidationError
    {
        private RuleType ruleType;
        private int index;
        private RuleComponent ruleComponent;

        public RuleValidationError(RuleType ruleType, int index, RuleComponent ruleComponent)
        {
            this.ruleType = ruleType;
            this.index = index;
            this.ruleComponent = ruleComponent;
        }
        public override string ToString()
        {
            return ruleType.ToString() + " rule #" + index.ToString() + " has invalid " + ruleComponent.ToString() + " component.";
        }
    }
}