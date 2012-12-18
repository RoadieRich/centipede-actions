using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Mathcad;
using Centipede;

namespace Centipede.MathCadActions
{
    [ActionCategory("Other")]
    public class MathCadActions : Action
    {
        public ActionTemplate(Dictionary<String, Object> variables) 
            : base("MathCadActions", variables)
        { }

        [ActionArgument(displayName="Visible Argument 1", usage="Help on the argument")]
        public String VisibleAttribute1 = "";

        public override void DoAction()
        {
            throw new NotImplementedException();
        }
    }
}
