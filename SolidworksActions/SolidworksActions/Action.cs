using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Centipede;
using SldWorks;
using SwConst;

namespace Centipede.SolidworksActions
{   
    public abstract class SolidworksAction : Action
    {
        public SolidworksAction(String name, Dictionary<String, Object> variables)
            : base(name, variables)
        { }

        protected SldWorks.ModelDoc2 SolidWorksDoc = null;

        private SldWorks.SldWorks SolidWorks = null;

        private static readonly String _SwAppVar = "_SolidWorksApp";

        protected override void InitAction()
        {
            Object obj;
            if (!Variables.TryGetValue(_SwAppVar, out obj))
            {           
                SolidWorks = new SldWorks.SldWorks();
            }
            else
            {
                SolidWorks = obj as SldWorks.SldWorks;
            }

        }

        protected override void CleanupAction()
        {
            Variables[_SwAppVar] = SolidWorks;
        }
                
        public override int Complexity
        {
            get
            {
                return 3;
            }
        }

        protected override void Dispose()
        {
            if (SolidWorks != null)
            {
                try
                {
                    SolidWorks.ExitApp();
                    SolidWorks = null;
                }
                catch
                { }
            }
            base.Dispose();
        }
    }

    [ActionCategory("SolidWorks", iconName="solidworks")]
    public class OpenSolidWorksFile : SolidworksAction
    {
        public OpenSolidWorksFile(Dictionary<String, Object> v)
            : base("Open SolidWorks File", v)
        { }

        protected override void DoAction()
        {
 	        throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", iconName = "solidworks")]
    public class ActivateSolidWorksDoc : SolidworksAction
    {
        public ActivateSolidWorksDoc(Dictionary<String, Object> v)
            : base("Activate SolidWorks Document", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }


    [ActionCategory("SolidWorks", displayName="Insert Design Table")]
    public class SWInsertDesignTable : SolidworksAction
    {
        public SWInsertDesignTable(Dictionary<String, Object> v)
            : base("Insert Design Table", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Set Active Config")]
    public class SetActiveSolidWorksConfig : SolidworksAction
    {
        public SetActiveSolidWorksConfig(Dictionary<String, Object> v)
            : base("Set Active Config", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Set Component Config")]
    public class SetSWComponentConfig : SolidworksAction
    {
        public SetSWComponentConfig(Dictionary<String, Object> v)
            : base("Set Component Config", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Set Dimension")]
    public class SetSWDimension : SolidworksAction
    {
        public SetSWDimension(Dictionary<String, Object> v)
            : base("Set Dimension", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    
    [ActionCategory("SolidWorks", displayName="Rebuild")]
    public class RebuildSolidWorks : SolidworksAction
    {
        public RebuildSolidWorks(Dictionary<String, Object> v)
            : base("Rebuild", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }
    
    [ActionCategory("SolidWorks", displayName="Set Suppression State")]
    public class SetSWSuppressionState : SolidworksAction
    {
        public SetSWSuppressionState(Dictionary<String, Object> v)
            : base("Set Suppression State", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Delete Item")]
    public class DeleteSWItem : SolidworksAction
    {
        public DeleteSWItem(Dictionary<String, Object> v)
            : base("Delete Item", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Delete Inactive Configurations")]
    public class DeleteInactiveSWConfigurations : SolidworksAction
    {
        public DeleteInactiveSWConfigurations(Dictionary<String, Object> v)
            : base("Delete Inactive Configurations", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Delete Inactive Components")]
    public class DeleteInactiveSWComponents : SolidworksAction
    {
        public DeleteInactiveSWComponents(Dictionary<String, Object> v)
            : base("Delete Inactive Components", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Save")]
    public class SaveSolidWorks : SolidworksAction
    {
        public SaveSolidWorks(Dictionary<String, Object> v)
            : base("Save", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Save As")]
    public class SolidWorksSaveAs : SolidworksAction
    {
        public SolidWorksSaveAs(Dictionary<String, Object> v)
            : base("Save As", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Close Active Document")]
    public class CloseActiveSWDoc : SolidworksAction
    {
        public CloseActiveSWDoc(Dictionary<String, Object> v)
            : base("Close Active Document", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Close All")]
    public class SWCloseAll : SolidworksAction
    {
        public SWCloseAll(Dictionary<String, Object> v)
            : base("Close All", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }

    [ActionCategory("SolidWorks", displayName="Quit")]
    public class Quit : SolidworksAction
    {
        public Quit(Dictionary<String, Object> v)
            : base("Quit", v)
        { }

        protected override void DoAction()
        {
            throw new NotImplementedException();
        }
    }
}

