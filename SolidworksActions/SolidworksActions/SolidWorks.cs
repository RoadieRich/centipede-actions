﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ResharperAnnotations;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using Environment = System.Environment;

namespace SolidworksActions
{
	public class SolidWorksWrapper : IDisposable
	{
		private SolidWorksWrapper()
		{
			_app = new SldWorks();
		}

		private static SldWorks _app;
		public void SelectFeature(ModelDoc2 document, string featureName, string instanceId, string itemType)
		// ReSharper restore UnusedMember.Global
		{
			String processedName = "";
			if (itemType != "COMPONENT")
			{
				if ((FileType)document.GetType() != FileType.Assembly)
				{
					throw new SolidWorksException(String.Format("Cannot select component {0} because current document is not an assembly.", featureName));
				}

				if (!string.IsNullOrEmpty(instanceId))
				{
					processedName = String.Format("{0}-{1}", featureName, instanceId);
				}

				String temp = (from Component2 component in (document as AssemblyDoc).GetComponents(true) as IEnumerable<Component2>
							   where component.Name.ToUpper() == processedName.ToUpper()
							   select component.Name).First();

				processedName = String.Format("{0}@{1}", temp, Path.GetFileNameWithoutExtension(document.GetPathName()));
			}
			else
			{
				processedName = (from Feature feature in document.FeatureManager.GetFeatures(true) as IEnumerable<Feature>
								 where feature.Name.ToUpper() == featureName.ToUpper()
								 select feature.Name).First();
			}

			if (!document.Extension.SelectByID(processedName, itemType, 0, 0, 0, false, 0, null))
			{
				throw new SolidWorksException(String.Format("Selected feature ({0}) does not exist.", featureName));
			}

		}
		void IDisposable.Dispose()
		{
			try
			{
				_app.CloseAllDocuments(true);
				_app.ExitApp();
				_app = null;
			}
// ReSharper disable EmptyGeneralCatchClause
			catch
			{ }
			// ReSharper restore EmptyGeneralCatchClause
		}
		
		public ModelDoc2 OpenFile(String filename)
		{
			int err = 0, warn = 0;
			
			ModelDoc2 doc = _app.OpenDoc6(filename,
				(int)GetDocumentTypeFromFilename(filename, true),
				(int)swOpenDocOptions_e.swOpenDocOptions_Silent,
				"",
				ref err,
				ref warn);


			if (err != 0)
			{

				int count;

				var messages = GetErrors(out count);
				if (count != 0)
				{
					String exceptionMessage;
					if (messages != null)
						exceptionMessage = String.Format("{0} error(s) occurred opening {1}:\n{2}", count, filename,
														String.Join(Environment.NewLine, messages));
					else
						exceptionMessage = "An unknown error occurred.";

					throw new SolidWorksException(exceptionMessage);
				}
			}
			return doc;
		}

		private enum FileType
		{
			Assembly = swDocumentTypes_e.swDocASSEMBLY,
			Drawing = swDocumentTypes_e.swDocDRAWING,
			Part = swDocumentTypes_e.swDocPART,
			Unknown = -1
		}

		private IEnumerable<String> GetErrors(out int count)
		{
			Object messages, msgIds, msgTypes;
			count = _app.GetErrorMessages(out messages, out msgIds, out msgTypes);

			return messages as String[];
		}

		public IEnumerable<String> GetErrors()
		{
			int count;
			return GetErrors(out count);
		}

		private static FileType GetDocumentTypeFromFilename([NotNull] String filename, Boolean throwOnUnknown=false)
		{
			string extension = Path.GetExtension(filename);

			switch ((extension ?? "").ToUpper())
			{
				case ".SLDPRT":
					return FileType.Part;
				case ".SLDASM":
					return FileType.Assembly;
				case ".SLDDRW":
					return FileType.Drawing;
				default:
					//return unknown, or throw??
					if (throwOnUnknown)
						throw new SolidWorksException(String.Format(@"Unknown filetype ""{0}""", filename));
					return FileType.Unknown;
			}
		}

		#region Singleton handling code

		private static readonly Object LockObject = new Object();
		private volatile static SolidWorksWrapper _instance;
		public static SolidWorksWrapper Instance
		{
			get
			{
				if (_instance == null)
				{
					lock (LockObject)
					{
						if (_instance == null)
						{
							_instance = new SolidWorksWrapper();
						}
					}
				}

				return _instance;
			}
		}

		public bool Visible
		{
			get
			{
				return _app.Visible;
			}
			set
			{
				_app.Visible = value;
			}
		}

		#endregion


/*
		internal ModelDoc2 ActivateDoc(string docName)
		{
			int errors = 0;
			ModelDoc2 doc = _app.ActivateDoc2(docName, true, ref errors);

			if (errors > 0)
			{
				throw new SolidWorksException();
			}

			return doc;
		}
*/

		public IEnumerable<ModelDoc2> GetOpenDocuments()
		{
			if (_app.ActiveDoc != null)
			{
				yield return _app.ActiveDoc;
			}
		}

		public void Quit()
		{
			_app.ExitApp();
		}
	}

	[Serializable]
	public class SolidWorksException : Exception
	{
// ReSharper disable UnusedMember.Global
		public SolidWorksException()
			: base(String.Format("Error from Solidworks:\n{0}",
				String.Join("\n", SolidWorksWrapper.Instance.GetErrors())))
		{ }
		public SolidWorksException(string message) : base(message) { }
		public SolidWorksException(string message, Exception inner) : base(message, inner) { }
// ReSharper restore UnusedMember.Global
		protected SolidWorksException(
		  System.Runtime.Serialization.SerializationInfo info,
		  System.Runtime.Serialization.StreamingContext context)
			: base(info, context) { }
	}
}
