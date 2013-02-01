using System;
using System.Collections.Generic;
using Centipede;
using Action = Centipede.Action;


namespace MathCadActions
{
    [ActionCategory("Other")]
    public class MathCadActions : Action
    {
        public MathCadActions(Dictionary<String, Object> variables) 
            : base("MathCadActions", variables)
        { }

        [ActionArgument(displayName="Visible Argument 1", usage="Help on the argument")]
        public String VisibleAttribute1 = "";

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }
}
