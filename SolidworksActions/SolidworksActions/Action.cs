using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Centipede;
using SldWorks;
using SwConst;
using SolidworksActions;
using System.IO;

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

        protected SolidWorksWrapper SolidWorks = null;

        protected override void InitAction()
        {
            
            SolidWorks = SolidWorksWrapper.Instance;

            Object obj = null;
            Variables.TryGetValue( SolidWorksDocVar, out obj);
            SolidWorksDoc = obj as ModelDoc2;

        }

        protected override void CleanupAction()
        {
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
                    (SolidWorks as IDisposable).Dispose();
                }
                catch
                { 
                    //Throwing exceptions from dispose is bad and wrong.
                }
                finally
                {
                    SolidWorksDoc = null;
                }
            }
            base.Dispose();
        }

        protected String SelectFeature(String featureName, String instanceId, String itemType)
        {
            
            String processedName = "";
            if (itemType != "COMPONENT")
            {

                if ((SolidWorksWrapper.FileType)SolidWorksDoc.GetType() != SolidWorksWrapper.FileType.Assembly)
                {
                    throw new SolidWorksActionException("Cannot set component config because current document is not an assembly.", this);
                }

                if (instanceId != null && instanceId != "")
                {
                    featureName = String.Format("{0}-{1}", featureName, instanceId);
                }
                
                String temp = (from Component2 component in (SolidWorksDoc as AssemblyDoc).GetComponents(true) as IEnumerable<Component2>
                        where component.Name.ToUpper() == featureName.ToUpper()
                        select component.Name).First();

                processedName = String.Format("{0}@{1}", temp, Path.GetFileNameWithoutExtension(SolidWorksDoc.GetPathName()));
            }
            else
            {
                processedName = (from Feature feature in SolidWorksDoc.FeatureManager.GetFeatures(true) as IEnumerable<Feature>
                                 where feature.Name.ToUpper() == featureName.ToUpper()
                                 select feature.Name).First();
            }
            
            if (!SolidWorksDoc.Extension.SelectByID(processedName, itemType, 0, 0, 0, false, 0, null))
            {
                throw new SolidWorksActionException(this);
            }

            return processedName;
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
            try
            {
                SolidWorksDoc = SolidWorks.OpenFile(Filename);
            }
            catch (SolidWorksException e)
            {
                throw new SolidWorksActionException(e, this);
            }
        }
    }

    [ActionCategory("SolidWorks", iconName = "solidworks")]
    public class ActivateSolidWorksDoc : SolidworksAction
    {
        public ActivateSolidWorksDoc(Dictionary<String, Object> v)
            : base("Activate SolidWorks Document", v)
        { }
                
        [ActionArgument(displayName="Document Name")]
        public String DocName = "";
    

        protected override void DoAction()
        {
            SolidWorksDoc = SolidWorks.ActivateDoc(DocName);
        }
    }


    [ActionCategory("SolidWorks", displayName="Insert Design Table")]
    public class SWInsertDesignTable : SolidworksAction
    {
        public SWInsertDesignTable(Dictionary<String, Object> v)
            : base("Insert Design Table", v)
        { }


        [ActionArgument(displayName = "Design Table Filename")]
        public String Filename = "";


        protected override void DoAction()
        {
            if (!SolidWorksDoc.InsertFamilyTableOpen(Filename))
            {
                throw new SolidWorksActionException(this);
            }
            
        }
    }

    [ActionCategory("SolidWorks", displayName="Set Active Config")]
    public class SetActiveSolidWorksConfig : SolidworksAction
    {
        public SetActiveSolidWorksConfig(Dictionary<String, Object> v)
            : base("Set Active Config", v)
        { }


        [ActionArgument(displayName = "Configuration name")]
        public String ConfigName = "";

        protected override void DoAction()
        {
            if ((SolidWorksDoc.GetActiveConfiguration() as Configuration).Name != ConfigName)
            {
                if (!SolidWorksDoc.ShowConfiguration2(ConfigName))
                {
                    throw new SolidWorksActionException(this);
                }
            }
                

        }
    }

    [ActionCategory("SolidWorks", displayName="Set Component Config")]
    public class SetSWComponentConfig : SolidworksAction
    {
        public SetSWComponentConfig(Dictionary<String, Object> v)
            : base("Set Component Config", v)
        { }


        [ActionArgument(displayName = "Component Name")]
        public String ComponentName = "";
                
        [ActionArgument(displayName = "Instance ID")]
        public String InstanceID = "";

        [ActionArgument(displayName = "Configuration Name")]
        public String ConfigName = "";


        protected override void DoAction()
        {
            if ((SolidWorksWrapper.FileType)SolidWorksDoc.GetType() != SolidWorksWrapper.FileType.Assembly)
            {
                throw new SolidWorksActionException("Cannot set component config because current document is not an assembly.", this);
            }

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

    [Serializable]
    public class SolidWorksActionException : ActionException
    {
        public SolidWorksActionException(Action action)
            : base(String.Format("Error from Solidworks:\n{0}",
                String.Join("\n", SolidWorksWrapper.Instance.GetErrors())), action)
        { }
        public SolidWorksActionException(string message, Action action)
            : base(message, action)
        { }

        public SolidWorksActionException(Exception inner, Action action)
            : base(String.Format("Error from Solidworks:\n{0}",
                String.Join("\n", SolidWorksWrapper.Instance.GetErrors())), inner, action)
        { }
    }
}

