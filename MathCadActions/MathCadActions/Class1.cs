using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.XPath;
using CentipedeInterfaces;
using Mathcad;
using PythonEngine;

namespace MathCADActions
{
    public abstract class MathCadAction : Centipede.Action
    {
        protected MathCadAction(string name, IDictionary<string, object> variables, ICentipedeCore core)
            : base(name, variables, core)
        { }

        protected override sealed void InitAction()
        {
            object obj;
            Variables.TryGetValue(SheetVar, out obj);

            Worksheet = obj as Mathcad.IMathcadWorksheet;

            if (this.MathCad == null)
            {
                _mathcad = MathCadWrapper.Instance;
            }

        }

        protected override void CleanupAction()
        {
            Variables[SheetVar] = Worksheet;
        }

        [ActionArgument(Literal = true, DisplayName = "Sheet Variable")]
        public String SheetVar;

        protected Mathcad.IMathcadWorksheet Worksheet;
        private static MathCadWrapper _mathcad;

        protected MathCadWrapper MathCad
        {
            get { return _mathcad; }
        }

        public override void Dispose()
        {
            if (this.MathCad != null)
            {
                ((IDisposable)this.MathCad).Dispose();
            }
            base.Dispose();
        }
    }

    [ActionCategory("MathCad", DisplayName = "Open MathCad File")]
    public class OpenFile : MathCadAction
    {
        public OpenFile(string name, IDictionary<string, object> variables, ICentipedeCore core)
            : base("Open Mathcad File", variables, core)
        { }

        
        [ActionArgument]
        public String Filename = "";


        protected override void DoAction()
        {
            string fname = ParseStringForVariable(Filename);
            Worksheet = MathCad.Application.Worksheets.Open(fname);
        }
    }
    
    [ActionCategory("MathCad", DisplayName = "Save As")]
    public class SaveAs : MathCadAction
    {
        public SaveAs(string name, IDictionary<string, object> variables, ICentipedeCore core)
            : base("Save Mathcad File As", variables, core)
        { }


        [ActionArgument]
        public String Filename = "";


        protected override void DoAction()
        {
            string fname = ParseStringForVariable(Filename);
            Worksheet.SaveAs(fname);
        }
    }
    
    [ActionCategory("MathCad", DisplayName = "Save")]
    public class Save : MathCadAction
    {
        public Save(string name, IDictionary<string, object> variables, ICentipedeCore core)
            : base("Save Mathcad File", variables, core)
        { }

        protected override void DoAction()
        {
            Worksheet.Save();
        }
    }

    [ActionCategory("MathCad", DisplayName = "Close")]
    public class Close : MathCadAction
    {
        public Close(string name, IDictionary<string, object> variables, ICentipedeCore core)
            : base("Close Mathcad File", variables, core)
        { }

        [ActionArgument(DisplayName = "Ignore Changes")]
        public bool Ignore = false;

        protected override void DoAction()
        {
            MCSaveOption options = Ignore ? MCSaveOption.mcDiscardChanges : MCSaveOption.mcPromptToSaveChanges;
            Worksheet.Close(options);
        }
    }

    [ActionCategory("MathCad", DisplayName = "Set Value By Tag")]
    public class SetValueByTag : MathCadAction
    {
        public SetValueByTag(string name, IDictionary<string, object> variables, ICentipedeCore core)
            : base("Set Value By Tag", variables, core)
        { }

        [ActionArgument]
        public string RegionTag = "";

        //[ActionArgument]
        //public string MathcadVarName = "";

        //[ActionArgument]
        //public string MathcadVarSubscript = "";

        [ActionArgument(Literal = true)]
        public string Re = "";

        [ActionArgument(Literal=true)]
        public string Im = "";

        [ActionArgument]
        public string Units = "";


        protected override void DoAction()
        {
            string regionTag = ParseStringForVariable(RegionTag),
                   //varname = ParseStringForVariable(MathcadVarName),
                   //varSubs = ParseStringForVariable(MathcadVarSubscript),
                   //realPart = ParseStringForVariable(Re),
                   //imagPart = ParseStringForVariable(Im),
                   units = ParseStringForVariable(Units);

            IPythonEngine engine = GetCurrentCore().PythonEngine;

            double realpart = engine.Evaluate<double>(Re),
                imagpart = !String.IsNullOrWhiteSpace(Im) ? engine.Evaluate<double>(Im) : 0;

            var region = Worksheet.Regions.Cast<IMathcadRegion2>().SingleOrDefault(r => r.Tag.Equals(regionTag));

            if (region == null)
            {
                throw new ActionException(string.Format("No such region {0}", regionTag), this);
            }

            if (region.Type != MCRegionType.mcMathRegion)
            {
                throw new ActionException(string.Format("Found region matching {0}, but was not math region", regionTag), this);
            }

            XmlDocument regionXml = new XmlDocument();
            regionXml.LoadXml(region.MathInterface.XML);

            XmlElement realElement = regionXml.GetElementsByTagName(@"ml:real").OfType<XmlElement>().Single();
            realElement.Value = realpart.ToString(CultureInfo.InvariantCulture);

            XmlElement imagElement = regionXml.GetElementsByTagName(@"ml:imag").OfType<XmlElement>().Single();
            realElement.Value = imagpart.ToString(CultureInfo.InvariantCulture);

            if (String.IsNullOrEmpty(units))
            {
                XmlElement[] idElements = regionXml.GetElementsByTagName(@"ml:id").OfType<XmlElement>().ToArray();
                if (idElements.Count() > 1)
                {
                    idElements[1].Value = units;
                }
            }

            region.MathInterface.XML = regionXml.InnerText;
        }

        }
    }

}
