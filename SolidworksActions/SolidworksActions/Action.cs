using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Centipede;
using SldWorks;
using SwConst;
using Action = Centipede.Action;


// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable ConvertToConstant.Global
// ReSharper disable FieldCanBeMadeReadOnly.Global
// ReSharper disable MemberCanBePrivate.Global
// ReSharper disable ConvertToConstant.Global
// ReSharper disable FieldCanBeMadeReadOnly.Global
namespace SolidworksActions
{   
    public abstract class SolidworksAction : Action
    {
        protected SolidworksAction(String name, Dictionary<String, Object> variables)
            : base(name, variables)
        { }

        
        [ActionArgument(displayName="SolidWorks Document Variable")]
        public String SolidWorksDocVar = "SldWrksDoc";

        protected ModelDoc2 SolidWorksDoc;

        protected SolidWorksWrapper SolidWorks;

        protected override void InitAction()
        {
            
            SolidWorks = SolidWorksWrapper.Instance;

            Object obj;
            Variables.TryGetValue( SolidWorksDocVar, out obj);
            SolidWorksDoc = obj as ModelDoc2;

        }

        protected override void CleanupAction()
        {
            SolidWorksDoc.Rebuild((int)swRebuildOptions_e.swRebuildAll);
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
// ReSharper disable EmptyGeneralCatchClause
                catch (Exception)
                { 
                    //Throwing exceptions from dispose is bad and wrong.
                }
                // ReSharper restore EmptyGeneralCatchClause
                finally
                {
                    SolidWorksDoc = null;
                }
            }
            base.Dispose();

        }

        // ReSharper disable UnusedMember.Global
        

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

    //[ActionCategory("SolidWorks", iconName = "solidworks")]
    //public class ActivateSolidWorksDoc : SolidworksAction
    //{
    //    public ActivateSolidWorksDoc(Dictionary<String, Object> v)
    //        : base("Activate SolidWorks Document", v)
    //    { }
                
    //    [ActionArgument(displayName="Document Name")]
    //    public String DocName = "";
    

    //    protected override void DoAction()
    //    {
    //        SolidWorksDoc = SolidWorks.ActivateDoc(DocName);
    //    }
    //}

    [ActionCategory("SolidWorks", displayName = "Insert Design Table", iconName = "solidworks")]
    public class SwInsertDesignTable : SolidworksAction
    {
        public SwInsertDesignTable(Dictionary<String, Object> v)
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

    [ActionCategory("SolidWorks", displayName = "Set Active Config", iconName = "solidworks")]
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

    [ActionCategory("SolidWorks", displayName = "Set Component Config", iconName = "solidworks")]
    public class SetSwComponentConfig : SolidworksAction
    {
        public SetSwComponentConfig(Dictionary<String, Object> v)
            : base("Set Component Config", v)
        { }

        [ActionArgument(displayName = "Component Name")]
        public String ComponentName = "";
                
        [ActionArgument(displayName = "Instance ID")]
        public String InstanceID = "";

        [ActionArgument(displayName = "Configuration Name")]
        public String ConfigName = "";

        [ActionArgument(displayName = "Supress")]
        public bool SuppressionState = true;

        protected override void DoAction()
        {   
            if (!(SolidWorksDoc is AssemblyDoc))
            {
                throw new SolidWorksActionException("Cannot set component config because current document is not an assembly.", this);
            }
            SolidWorks.SelectFeature(SolidWorksDoc, ComponentName, InstanceID, "COMPONENT");
            var selectionMgr = SolidWorksDoc.SelectionManager as SelectionMgr;
            var component = selectionMgr.GetSelectedObjectsComponent2(1) as Component2;
            component.ReferencedConfiguration = ConfigName;
        }
    }

    [ActionCategory("SolidWorks", displayName = "Set Dimension", iconName = "solidworks")]
    public class SetSwDimension : SolidworksAction
    {
        [ActionArgument(displayName = "Dimension Name")]
        public String DimensionName = "";

        
        [ActionArgument(displayName = "New Value")]
        public String Value = "";


        public SetSwDimension(Dictionary<String, Object> v)
            : base("Set Dimension", v)
        { }

        protected override void DoAction()
        {
            var dimension = SolidWorksDoc.Parameter(DimensionName) as Dimension;
            if (dimension == null)
            {
                throw new SolidWorksActionException(String.Format("Dimension named {0} not found", DimensionName), this);
            }
            if (dimension.ReadOnly)
            {
                Message(String.Format("Changed dimension {0} is read only.", DimensionName));
            }

            Double val;
            if (!Double.TryParse(Value, out val))
            {
                throw new SolidWorksActionException(String.Format("Invalid Value for value: {0}", Value), this);
            }
            if ((swSetValueReturnStatus_e)dimension.SetSystemValue2(val, (int)swSetValueInConfiguration_e.swSetValue_UseCurrentSetting) != swSetValueReturnStatus_e.swSetValue_Successful)
            {
                throw new SolidWorksActionException(this);
            }
            if (dimension.IsAppliedToAllConfigurations())
            {
                Message("Dimension update was limited to active configuration.");
            }
            if (!SolidWorksDoc.EditRebuild3())
            {
                throw new SolidWorksActionException(this);
            }
        }
    }

    //[ActionCategory("SolidWorks", displayName="Rebuild")]
    //public class RebuildSolidWorks : SolidworksAction
    //{
    //    public RebuildSolidWorks(Dictionary<String, Object> v)
    //        : base("Rebuild", v)
    //    { }

    //    protected override void DoAction()
    //    {
    //        throw new NotImplementedException();
    //    }
    //}

    [ActionCategory("SolidWorks", displayName = "Set Suppression State", iconName = "solidworks")]
    public class SetSwSuppressionState : SolidworksAction
    {
        
        [ActionArgument(displayName="Feature Name")]
        public string FeatureName="";

        [ActionArgument(displayName="Instance ID")]
        public string InstanceID="";

        [ActionArgument(displayName="Item Type")]
        public string ItemType = "";

        [ActionArgument]
        public bool Suppress = true;

        public SetSwSuppressionState(Dictionary<String, Object> v)
            : base("Set Suppression State", v)
        { }

        protected override void DoAction()
        {
            SolidWorks.SelectFeature(SolidWorksDoc, FeatureName, InstanceID, ItemType);
            if (Suppress)
            {
                SolidWorksDoc.EditSuppress();
            }
            else
            {
                SolidWorksDoc.EditUnsuppress();
            }
        }
    }

    [ActionCategory("SolidWorks", displayName = "Delete Item", iconName = "solidworks")]
    public class DeleteSwItem : SolidworksAction
    {
        public DeleteSwItem(Dictionary<String, Object> v)
            : base("Delete Item", v)
        { }

        [ActionArgument(displayName = "Feature Name")]
        public string FeatureName = "";

        [ActionArgument(displayName = "Instance ID")]
        public string InstanceID = "";

        [ActionArgument(displayName = "Item Type")]
        public string ItemType = "";

        protected override void DoAction()
        {
            SolidWorks.SelectFeature(SolidWorksDoc, FeatureName, InstanceID, ItemType);
            if (!SolidWorksDoc.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed|(int)swDeleteSelectionOptions_e.swDelete_Children))
            {
                throw new SolidWorksActionException(this);
            }
        }
    }

    [ActionCategory("SolidWorks", displayName = "Delete Inactive Configurations", iconName = "solidworks")]
    public class DeleteInactiveSwConfigurations : SolidworksAction
    {
        public DeleteInactiveSwConfigurations(Dictionary<String, Object> v)
            : base("Delete Inactive Configurations", v)
        { }

        protected override void DoAction()
        {
            IEnumerable<String> configNames = (string[])SolidWorksDoc.GetConfigurationNames();
            string activeConfigName = SolidWorksDoc.ConfigurationManager.ActiveConfiguration.Name;
            foreach (string name in configNames.Where(name => name != activeConfigName))
            {
                SolidWorksDoc.DeleteConfiguration2(name);
            }
        }
    }

    [ActionCategory("SolidWorks", displayName = "Delete Inactive Components", iconName = "solidworks")]
    public class DeleteInactiveSwComponents : SolidworksAction
    {
        public DeleteInactiveSwComponents(Dictionary<String, Object> v)
            : base("Delete Inactive Components", v)
        { }

        protected override void DoAction()
        {
            var assemblyDoc = SolidWorksDoc as AssemblyDoc;
            if (assemblyDoc == null)
            {
                throw new SolidWorksActionException(
                        "Can't delete inactive components because the current document is not an assembly."
                        , this);
            }
            var components = (Component2[])assemblyDoc.GetComponents(true);
            const int options = (int)swDeleteSelectionOptions_e.swDelete_Absorbed |
                                (int)swDeleteSelectionOptions_e.swDelete_Children;
            foreach (var component in components.Where(component => (swComponentSuppressionState_e)component.GetSuppression() ==
                                                                    swComponentSuppressionState_e.swComponentSuppressed))
            {
                component.Select(false);
                SolidWorksDoc.Extension.DeleteSelection2(options);
            }

        }
    }

    [ActionCategory("SolidWorks", displayName = "Save", iconName = "solidworks")]
    public class SaveSolidWorks : SolidworksAction
    {
        public SaveSolidWorks(Dictionary<String, Object> v)
            : base("Save", v)
        { }

        protected override void DoAction()
        {
            const int options = (int)swSaveAsOptions_e.swSaveAsOptions_Silent |
                                (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced;
            int errs=0, warnings=0;
            bool saved = SolidWorksDoc.Save3(options, ref errs, ref warnings);
            if(!saved || errs != 0)
            {
                throw new SolidWorksActionException(this);
            }
        }
    }

    [ActionCategory("SolidWorks", displayName = "Save As", iconName = "solidworks")]
    public class SolidWorksSaveAs : SolidworksAction
    {
        public SolidWorksSaveAs(Dictionary<String, Object> v)
                : base("Save As", v)
        { }

        [ActionArgument]
        public String Filename = "";

        [ActionArgument(displayName = "Save as Copy")]
        public bool SaveAsCopy = false;

        protected override void DoAction()
        {
            string filename = Path.GetFullPath(Filename);

            //tidy appearance
            SolidWorksDoc.ViewDisplayShaded();
            SolidWorksDoc.ShowNamedView2("*Isometric", -1);
            SolidWorksDoc.ViewZoomtofit2();

            var options = (int)swSaveAsOptions_e.swSaveAsOptions_Silent + (SaveAsCopy
                                                ? (int)swSaveAsOptions_e.swSaveAsOptions_Copy
                                                : 0);

            string path = Path.GetDirectoryName(filename);
            if (path != null && !Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }


            int errs = 0, warns = 0;
            bool success = SolidWorksDoc.SaveAs4(filename, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, options,
                                                 ref errs, ref warns);
            if (!success)
            {
                throw new SolidWorksActionException(this);
            }
        }
    }

    [ActionCategory("SolidWorks", displayName = "Close Active Document", iconName = "solidworks")]
    public class CloseActiveSwDoc : SolidworksAction
    {
        public CloseActiveSwDoc(Dictionary<String, Object> v)
            : base("Close Active Document", v)
        { }

        [ActionArgument(displayName = "Save Changes")]
        public bool SaveChanges = false;

        protected override void DoAction()
        {
            if (SaveChanges)
            {
                const int options = (int)swSaveAsOptions_e.swSaveAsOptions_Silent |
                                (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced;
                int errs = 0, warnings = 0;
                bool saved = SolidWorksDoc.Save3(options, ref errs, ref warnings);
                if (!saved || errs != 0)
                {
                    throw new SolidWorksActionException(this);
                }
            }
            if (SolidWorksDoc.GetSaveFlag())
            {
                AskEventEnums.DialogResult result = Ask(string.Format("Closing {0} without saving changes.", SolidWorksDoc.GetPathName()),"Closing document",AskEventEnums.AskType.OKCancel);
                if (result == AskEventEnums.DialogResult.Cancel)
                {
                    return;
                }
            }
            SolidWorksDoc.Quit();
            SolidWorksDoc = null;
        }
    }

    [ActionCategory("SolidWorks", displayName = "Close All", iconName = "solidworks")]
    public class SwCloseAll : SolidworksAction
    {
        public SwCloseAll(Dictionary<String, Object> v)
            : base("Close All", v)
        { }

        [ActionArgument(displayName = "Save Changes")]
        public bool SaveChanges = false;

        protected override void DoAction()
        {
            foreach (ModelDoc2 doc in SolidWorksWrapper.Instance.GetOpenDocuments())
            {
                if (SaveChanges)
                {
                    const int options = (int)swSaveAsOptions_e.swSaveAsOptions_Silent |
                                        (int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced;
                    int errs = 0, warnings = 0;
                    bool saved = doc.Save3(options, ref errs, ref warnings);
                    if (!saved || errs != 0)
                    {
                        throw new SolidWorksActionException(this);
                    }
                }
                if (doc.GetSaveFlag())
                {
                    AskEventEnums.DialogResult result =
                            Ask(string.Format("Closing {0} without saving changes.", SolidWorksDoc.GetPathName()),
                                "Closing document", AskEventEnums.AskType.OKCancel);
                    if (result == AskEventEnums.DialogResult.Cancel)
                    {
                        return;
                    }
                }
                SolidWorksDoc.Quit();
            }
        }
    }

    [ActionCategory("SolidWorks", displayName = "Quit", iconName = "solidworks")]
    public class Quit : SolidworksAction
    {
        public Quit(Dictionary<String, Object> v)
            : base("Quit", v)
        { }

        protected override void DoAction()
        {
            SolidWorksWrapper.Instance.Quit(); 

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

