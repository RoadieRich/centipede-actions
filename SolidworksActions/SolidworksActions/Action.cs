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

        
        [ActionArgument(displayName="SolidWorks Document Variable")]
        public String SolidWorksDocVar = "SldWrksDoc";

        protected SldWorks.ModelDoc2 SolidWorksDoc = null;

        protected SldWorks.SldWorks SolidWorks = null;

        protected static readonly String _SwAppVar = "_SolidWorksApp";

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

            obj = null;
            Variables.TryGetValue( SolidWorksDocVar, out obj);
            SolidWorksDoc = obj as ModelDoc2;

        }

        protected override void CleanupAction()
        {
            Variables[_SwAppVar] = SolidWorks;
            Variables[SolidWorksDocVar] = SolidWorksDoc;
        }
                
        public override int Complexity
        {
            get
            {
                return 3;
            }
        }

        public override void Dispose()
        {
            if (SolidWorks != null)
            {
                try
                {
                    SolidWorksDoc.Close();
                    SolidWorks.ExitApp();
                }
                catch
                { 
                    //Throwing exceptions from dispose is bad and wrong.
                }
                finally
                {
                    SolidWorksDoc = null;
                    SolidWorks = null;
                }
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


        [ActionArgument]
        public String Filename = "";
    
        protected override void DoAction()
        {
            Int32 errorCode = 0, warnings = 0;
            swFileLoadError_e errors;

            swDocumentTypes_e docType;
            switch (System.IO.Path.GetExtension(Filename).ToUpper())
            {
                case ".SLDPRT":
                    docType = swDocumentTypes_e.swDocPART;
                    break;
                case ".SLDASM":
                    docType = swDocumentTypes_e.swDocASSEMBLY;
                    break;
                case ".SLDDRW":
                    docType = swDocumentTypes_e.swDocDRAWING;
                    break;
                default:
                    throw new ActionException(String.Format("Unknown document type: {0} (must be *.SLDPRT, *.SLDASM or *.SLDDRW)", Filename), this);
            }

            SolidWorksDoc = SolidWorks.OpenDoc6(Filename, (int)docType, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref errorCode, ref warnings);

            errors = (swFileLoadError_e)errorCode;

            if (errors != 0)
            {
                Object _null = null;
                object messages;
                SolidWorks.GetErrorMessages(out messages, out _null, out _null);
                String errorType = ((swFileLoadError_e) errors).ToString();
                throw new ActionException(String.Format("Error messages: {0}", String.Join("\n", messages as String[])), this);
            }
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

