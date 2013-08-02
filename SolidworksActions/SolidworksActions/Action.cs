using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using CentipedeInterfaces;
using SldWorks;
using SwConst;
using Action = Centipede.Action;

namespace SolidworksActions
{   
	public abstract class SolidworksAction : Action
	{
		protected SolidworksAction(String name, IDictionary<String, Object> variables, ICentipedeCore c)
			: base(name, variables, c)
		{ }

		[ActionArgument(DisplayName="SolidWorks Document Variable", Literal=true)]
		public String SolidWorksDocVar = "SldWrksDoc";

		protected ModelDoc2 SolidWorksDoc;

		protected SolidWorksWrapper SolidWorks;

		protected override void InitAction()
		{
			
			SolidWorks = SolidWorksWrapper.Instance;

			SolidWorks.Visible = true;

			Object obj;
			Variables.TryGetValue(SolidWorksDocVar, out obj);
			SolidWorksDoc = obj as ModelDoc2;

		}

		protected override void CleanupAction()
		{
			//SolidWorksDoc.Rebuild((int)swRebuildOptions_e.swRebuildAll);
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
				catch (Exception e)
				{
					const string source = @"Centipede SolidWorks Action";
					if(!EventLog.SourceExists(source))
					{
						EventLog.CreateEventSource(source, @"Application");
					}
					EventLog.WriteEntry(source, string.Format("Exception in SolidWorksAction.Dispose: {0}", e), EventLogEntryType.Warning);
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

	[ActionCategory("SolidWorks", IconName="solidworks")]
	public class OpenSolidWorksFile : SolidworksAction
	{
		public OpenSolidWorksFile(IDictionary<string, object> v, ICentipedeCore c)
			: base("Open SolidWorks File", v, c)
		{ }


		[ActionArgument]
		public String FileName = "";
	
		protected override void DoAction()
		{
            string myFileName = Path.GetFullPath(ParseStringForVariable(FileName));

            Message("Opening file \"" + myFileName + "\"");

			try
			{
				SolidWorksDoc = SolidWorks.OpenFile(myFileName);
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
	//    public ActivateSolidWorksDoc(IDictionary<String, Object> v)
	//        : base("Activate SolidWorks Document", v)
	//    { }
				
	//    [ActionArgument(displayName="Document Name")]
	//    public String DocName = "";
	

	//    protected override void DoAction()
	//    {
	//        SolidWorksDoc = SolidWorks.ActivateDoc(DocName);
	//    }
	//}

	[ActionCategory("SolidWorks", DisplayName = "Insert Design Table", IconName = "solidworks")]
	public class SwInsertDesignTable : SolidworksAction
	{
		public SwInsertDesignTable(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Insert Design Table", v, c)
		{ }
		
		[ActionArgument(DisplayName = "Design Table Filename")]
		public String Filename = "";
		
		protected override void DoAction()
		{
            string myFileName = ParseStringForVariable(Filename);
            
            Message("Inserting design table \"" + myFileName + "\"");

			if (!SolidWorksDoc.InsertFamilyTableOpen(myFileName))
			{
				throw new SolidWorksActionException(this);
			}
			
		}
	}

	[ActionCategory("SolidWorks", DisplayName = "Set Active Config", IconName = "solidworks", Usage="Sets the active configuration in the active document")]
	public class SetActiveSolidWorksConfig : SolidworksAction
	{
		public SetActiveSolidWorksConfig(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Set Active Config", v, c)
		{ }

		[ActionArgument(DisplayName = "Configuration name")]
		public String ConfigName = "";

		protected override void DoAction()
		{
            string myConfigName = ParseStringForVariable(ConfigName);

            Message("Setting active configuration to \"" + myConfigName + "\"");

			Configuration configuration = (Configuration)SolidWorksDoc.GetActiveConfiguration();
			if (configuration.Name == myConfigName) return;
			if (!SolidWorksDoc.ShowConfiguration2(myConfigName))
			{
				throw new SolidWorksActionException(this);
			}
		}
	}

	[ActionCategory("SolidWorks", DisplayName = "Set Component Config", IconName = "solidworks", Usage="Sets the refernced configuration of a component in the active document")]
	public class SetSwComponentConfig : SolidworksAction
	{
		public SetSwComponentConfig(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Set Component Config", v,c)
		{ }

		[ActionArgument(DisplayName = "Component Name", Usage="Enter the name of the component as shown in the feature manager (Case sensitive)")]
		public String ComponentName = "";
				
		[ActionArgument(DisplayName = "Instance ID", Usage="Enter the instance ID of the component as shown in the feature manager")]
		public String InstanceId = "";

		[ActionArgument(DisplayName = "Configuration Name", Usage="Enter the name of the configuration that this component should reference")]
		public String ConfigName = "";

		protected override void DoAction()
		{
            string myComponentName = ParseStringForVariable(ComponentName);
            string myInstanceId = ParseStringForVariable(InstanceId);
            string myConfigName = ParseStringForVariable(ConfigName);

            Message("Setting component \"" + myComponentName + ":" + myInstanceId + "\" to reference configuration \"" + myConfigName + "\"");

			if (!(SolidWorksDoc is AssemblyDoc))
			{
				throw new SolidWorksActionException("Cannot set component config because current document is not an assembly.", this);
			}
			SolidWorks.SelectFeature(SolidWorksDoc, myComponentName, myInstanceId, "COMPONENT");
			var selectionMgr = SolidWorksDoc.SelectionManager as SelectionMgr;
			if (selectionMgr != null)
			{
				var component = selectionMgr.GetSelectedObjectsComponent2(1) as Component2;
				if (component != null) component.ReferencedConfiguration = myConfigName;
			}
		}
	}

	[ActionCategory("SolidWorks", DisplayName = "Set Dimension", IconName = "solidworks", Usage="Sets the value of a dimension in the active document")]
	public class SetSwDimension : SolidworksAction
	{
		[ActionArgument(DisplayName = "Dimension Name", Usage="Enter the name of the dimension to set, for example: D1@Sketch1")]
		public String DimensionName = "";
        		
		[ActionArgument(DisplayName = "New Value", Usage="Enter the new value: linear dimensions must be specified in metres, angular dimensions in radians")]
		public String Value = "";


		public SetSwDimension(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Set Dimension", v, c)
		{ }

		protected override void DoAction()
		{
            string myDimensionName = ParseStringForVariable(DimensionName);
            string myValue = ParseStringForVariable(Value);

            Message("Setting Dimension \"" + myDimensionName + "\" to value: " + myValue);

			var dimension = SolidWorksDoc.Parameter(myDimensionName) as Dimension;
			if (dimension == null)
			{
				throw new SolidWorksActionException(String.Format("Dimension named {0} not found", myDimensionName), this);
			}
			if (dimension.ReadOnly)
			{
				Warning("Changed dimension {0} is read only.", myDimensionName);
			}

			Double val;

            if (!Double.TryParse(myValue, out val))
			{
				throw new SolidWorksActionException(String.Format("Invalid Value for value: {0}", myValue), this);
			}
			if ((swSetValueReturnStatus_e)dimension.SetSystemValue2(val, (int)swSetValueInConfiguration_e.swSetValue_UseCurrentSetting) != swSetValueReturnStatus_e.swSetValue_Successful)
			{
				throw new SolidWorksActionException(this);
			}
			if (dimension.IsAppliedToAllConfigurations())
			{
				Warning("Dimension update was limited to active configuration.");
			}
			//if (!SolidWorksDoc.EditRebuild3())
			//{
			//	throw new SolidWorksActionException(this);
			//}
		}
	}

    [ActionCategory("SolidWorks", DisplayName = "Rebuild", IconName = "solidworks", Usage = "Performs a forced full rebuild")]
	public class RebuildSolidWorks : SolidworksAction
	{
		public RebuildSolidWorks(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Rebuild", v, c)
		{
		}

		[ActionArgument] 
		public Boolean TopOnly;

		protected override void DoAction()
		{

            Message("Rebuilding");

			if (!SolidWorksDoc.ForceRebuild3(TopOnly))
			{
				throw new SolidWorksActionException("Cannot force rebuild", this);
			}
			
		}
	}

	[ActionCategory("SolidWorks", DisplayName="Set Suppression State", IconName="solidworks", Usage="Suppresses or Unsuppresses features, mates, components, patterns etc. in the active document")]
	public class SetSwSuppressionState : SolidworksAction
	{
		
		[ActionArgument(DisplayName="Feature Name", Usage="Enter the name of the feature as shown in the feature manager (Case sensitive)")]
		public string FeatureName="";

		[ActionArgument(DisplayName="Instance ID", Usage="For components of an assembly enter the instance ID, otherwise leave blank")]
		public string InstanceId="";

		[ActionArgument(DisplayName="Item Type", Usage="See the SolidWorks API help page for \"swSelectType_e\".  This element must contain one of the valid string values")]
		public string ItemType="";

		[ActionArgument]
		public bool Suppress = true;

		public SetSwSuppressionState(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Set Suppression State", v, c)
		{ }

		protected override void DoAction()
		{
            string myFeatureName = ParseStringForVariable(FeatureName);
            string myInstanceId = ParseStringForVariable(InstanceId);
            string myItemType = ParseStringForVariable(ItemType);

            Message("Setting suppression state of \"" + myFeatureName + ":" + myInstanceId + "\" to " + Convert.ToString(Suppress));

			SolidWorks.SelectFeature(SolidWorksDoc, myFeatureName, myInstanceId, myItemType);
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

    [ActionCategory("SolidWorks", DisplayName = "Delete Item", IconName = "solidworks", Usage = "Deletes features, mates, components, patterns etc. in the active document")]
	public class DeleteSwItem : SolidworksAction
	{
		public DeleteSwItem(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Delete Item", v, c)
		{ }

        [ActionArgument(DisplayName = "Feature Name", Usage = "Enter the name of the feature as shown in the feature manager (Case sensitive)")]
        public string FeatureName = "";

        [ActionArgument(DisplayName = "Instance ID", Usage = "For components of an assembly enter the instance ID, otherwise leave blank")]
        public string InstanceId = "";

        [ActionArgument(DisplayName = "Item Type", Usage = "See the SolidWorks API help page for \"swSelectType_e\".  This element must contain one of the valid string values")]
        public string ItemType = "";

		protected override void DoAction()
		{
            string myFeatureName = ParseStringForVariable(FeatureName);
            string myInstanceId = ParseStringForVariable(InstanceId);
            string myItemType = ParseStringForVariable(ItemType);

            Message("Deleting item \"" + myFeatureName + ":" + myInstanceId + "\"");

			SolidWorks.SelectFeature(SolidWorksDoc, myFeatureName, myInstanceId, myItemType);
			if (!SolidWorksDoc.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed|(int)swDeleteSelectionOptions_e.swDelete_Children))
			{
				throw new SolidWorksActionException(this);
			}
		}
	}

	[ActionCategory("SolidWorks", DisplayName = "Delete Inactive Configurations", IconName = "solidworks", Usage="Deletes configurations that are not used in any open documents")]
	public class DeleteInactiveSwConfigurations : SolidworksAction
	{
		public DeleteInactiveSwConfigurations(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Delete Inactive Configurations", v, c)
		{ }

		protected override void DoAction()
		{
    		IEnumerable<String> configNames = (string[])SolidWorksDoc.GetConfigurationNames();
			string activeConfigName = SolidWorksDoc.ConfigurationManager.ActiveConfiguration.Name;
			foreach (string name in configNames.Where(name => name != activeConfigName))
			{
                Message("Deleting inactive configuration: \"" + name + "\"");
				SolidWorksDoc.DeleteConfiguration2(name);
			}
		}
	}

	[ActionCategory("SolidWorks", DisplayName = "Delete Inactive Components", IconName = "solidworks", Usage="Deletes assembly components that are suppressed in the active configuration")]
	public class DeleteInactiveSwComponents : SolidworksAction
	{
		public DeleteInactiveSwComponents(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Delete Inactive Components", v, c)
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
                Message("Deleting inactive component: \"" + component.Name + "\"");
				component.Select(false);
				SolidWorksDoc.Extension.DeleteSelection2(options);
			}

		}
	}

	[ActionCategory("SolidWorks", DisplayName = "Save", IconName = "solidworks", Usage="Saves the active document and any referenced documents")]
	public class SaveSolidWorks : SolidworksAction
	{
		public SaveSolidWorks(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Save", v, c)
		{ }

		protected override void DoAction()
		{
            Message("Saving file: \"" + SolidWorksDoc.GetTitle() + "\"");

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

	[ActionCategory("SolidWorks", DisplayName = "Save As", IconName = "solidworks", Usage="Saves the active document as a new file, you can specify any supported file type")]
	public class SolidWorksSaveAs : SolidworksAction
	{
		public SolidWorksSaveAs(IDictionary<String, Object> v, ICentipedeCore c)
				: base("Save As", v, c)
		{ }

		[ActionArgument (Usage="Include the FileName extension, any file type conversion will be performed if necessary")]
		public String FileName = "";

		[ActionArgument(DisplayName = "Save as Copy")]
		public bool SaveAsCopy;

		protected override void DoAction()
		{
            string myFileName = Path.GetFullPath(ParseStringForVariable(FileName));

            Message("Saving file: \"" + SolidWorksDoc.GetTitle() + "\" as \"" + myFileName + "\"");

			//tidy appearance
			SolidWorksDoc.ViewDisplayShaded();
			SolidWorksDoc.ShowNamedView2("*Isometric", -1);
			SolidWorksDoc.ViewZoomtofit2();

			var options = (int)swSaveAsOptions_e.swSaveAsOptions_Silent + (SaveAsCopy
												? (int)swSaveAsOptions_e.swSaveAsOptions_Copy
												: 0);

			string path = Path.GetDirectoryName(myFileName);
			if (path != null && !Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}


			int errs = 0, warns = 0;
			bool success = SolidWorksDoc.SaveAs4(myFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, options,
												 ref errs, ref warns);
			if (!success)
			{
				throw new SolidWorksActionException(this);
			}
		}
	}

	[ActionCategory("SolidWorks", DisplayName = "Close Active Document", IconName = "solidworks", Usage="Closes the currently active document")]
	public class CloseActiveSwDoc : SolidworksAction
	{
		public CloseActiveSwDoc(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Close Active Document", v, c)
		{ }

		[ActionArgument(DisplayName = "Save Changes")]
		public bool SaveChanges;

		protected override void DoAction()
		{
            Message("Closing file: \"" + SolidWorksDoc.GetTitle() + "\"");

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

	[ActionCategory("SolidWorks", DisplayName = "Close All", IconName = "solidworks", Usage="Closes all open documents")]
	public class SwCloseAll : SolidworksAction
	{
		public SwCloseAll(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Close All", v, c)
		{ }

		[ActionArgument(DisplayName = "Save Changes")]
		public bool SaveChanges;

		protected override void DoAction()
		{
			foreach (ModelDoc2 doc in SolidWorksWrapper.Instance.GetOpenDocuments())
			{
                Message("Closing file: \"" + doc.GetTitle() + "\"");

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

	[ActionCategory("SolidWorks", DisplayName = "Quit", IconName = "solidworks")]
	public class Quit : SolidworksAction
	{
		public Quit(IDictionary<String, Object> v, ICentipedeCore c)
			: base("Quit", v, c)
		{ }

		protected override void DoAction()
		{
			SolidWorksWrapper.Instance.Quit(); 

		}
	}

	[Serializable]
	public class SolidWorksActionException : ActionException
	{
		public SolidWorksActionException(IAction action)
			: base(String.Format("Error from Solidworks:\n{0}",
				String.Join("\n", SolidWorksWrapper.Instance.GetErrors())), action)
		{ }
		public SolidWorksActionException(string message, IAction action)
			: base(message, action)
		{ }

		public SolidWorksActionException(Exception inner, IAction action)
			: base(String.Format("Error from Solidworks:\n{0}",
				String.Join("\n", SolidWorksWrapper.Instance.GetErrors())), inner, action)
		{ }
	}
}

