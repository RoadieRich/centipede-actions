using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Xml;
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

            Worksheet = obj as WorkSheetWrapper;

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

        protected WorkSheetWrapper Worksheet;
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


        protected void SetValueInRegion(IMathcadRegion2 region, String value, McValueType type, string units)
        {
            Centipede.IPythonEngine engine = GetCurrentCore().PythonEngine;

            dynamic evaluated = engine.Evaluate(value);

            XmlDocument regionXml = new XmlDocument();
            regionXml.LoadXml(region.MathInterface.XML);

            if (!type.HasFlag(McValueType.Numeric))
            {
                SetNumericValue(region, evaluated, type, ParseStringForVariable(units));
            }
            else
            {
                XmlElement textElement = regionXml.GetElementsByTagName(@"ml:str").OfType<XmlElement>().SingleOrDefault();
                textElement.Value = evaluated.ToString();
            }

            region.MathInterface.XML = regionXml.InnerText;
        }

        private void SetNumericValue(IMathcadRegion2 region, dynamic evaluated, McValueType type, string units)
        {
            XmlDocument regionXml = new XmlDocument();
            regionXml.LoadXml(region.MathInterface.XML);

            XmlElement valueElement;
            XmlElement realElement = regionXml.GetElementsByTagName(@"ml:real").OfType<XmlElement>().SingleOrDefault();

            dynamic realpart;
            if (type.HasFlag(McValueType.NumericHasComplexPart))
            {
                XmlElement imagElement = regionXml.GetElementsByTagName(@"ml:imag")
                                                  .OfType<XmlElement>()
                                                  .SingleOrDefault();

                Complex value = evaluated;
                realpart = value.Real;
                if (imagElement != null)
                {
                    imagElement.Value = value.Imaginary.ToString(CultureInfo.InvariantCulture);
                    valueElement = (XmlElement)imagElement.ParentNode.ParentNode;
                }
                else
                {
                    var parent = realElement.ParentNode;
                    XmlElement complexElement = value.ToMathML();
                    parent.ReplaceChild(realElement, complexElement);
                    valueElement = complexElement;
                }
            }
            else
            {
                realpart = evaluated.ToString();
                valueElement = realElement;
            }

            realElement.Value = realpart;

            if (type.HasFlag(McValueType.Dimensioned))
            {
                XmlElement[] idElements = regionXml.GetElementsByTagName(@"ml:id").OfType<XmlElement>().ToArray();
                if (idElements.Count() > 1)
                {
                    idElements.Last().Value = units;
                }
                else
                {
                    valueElement.ParentNode.ReplaceChild(valueElement, this.ApplyUnits(valueElement, units));
                }
            }
        }

        [FlagsAttribute]
        protected enum McValueType
        {
            Numeric = 1,

            [EditorBrowsable(EditorBrowsableState.Never)]
            NumericHasComplexPart = 2,
            Complex = Numeric | NumericHasComplexPart,
            String = 4,
            Dimensioned = 8
        }


        private XmlElement ApplyUnits(XmlElement realElement, string units)
        {
            XmlDocument doc = new XmlDocument();
            var apply = doc.CreateElement("ml:apply");

            var mult = doc.CreateElement("ml:mult");
            mult.SetAttribute("style", "auto-select");

            var unit = doc.CreateElement("ml:id");
            unit.Value = units;

            apply.AppendChild(mult);
            apply.AppendChild(realElement);
            apply.AppendChild(unit);

            return apply;
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
            Worksheet = new WorkSheetWrapper(MathCad.Application.Worksheets.Open(fname));
        }
    }

    [ActionCategory("MathCad", DisplayName = "Save As")]
    public class SaveAs : MathCadAction
    {
        public SaveAs(IDictionary<string, object> variables, ICentipedeCore core)
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
        public Save(IDictionary<string, object> variables, ICentipedeCore core)
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
        public SetValueByTag(IDictionary<string, object> variables, ICentipedeCore core)
            : base("Set Value By Tag", variables, core)
        { }

        [ActionArgument]
        public string RegionTag = "";

        //[ActionArgument]
        //public string MathcadVarName = "";

        //[ActionArgument]
        //public string MathcadVarSubscript = "";

        [ActionArgument(DisplayName = "Value")]
        public string Value = "";

        //[ActionArgument(Literal = true)]
        //public string Im = "";

        [ActionArgument]
        public string Units = "";

        [ActionArgument]
        public Boolean Complex = false;

        [ActionArgument]
        public bool Text = false;

        protected override void DoAction()
        {
            string regionTag = ParseStringForVariable(RegionTag);

            IMathcadRegion2 region = Worksheet.GetRegionByTag(regionTag);

            if (region == null)
            {
                throw new ActionException(string.Format("No such region {0}", regionTag), this);
            }

            if (region.Type != MCRegionType.mcMathRegion)
            {
                throw new ActionException(
                    string.Format("Found region matching {0}, but was not math region", regionTag), this);
            }

            McValueType type = Text ? McValueType.String : Complex ? McValueType.Complex : McValueType.Numeric;

            if (!String.IsNullOrEmpty(Units))
            {
                type |= McValueType.Dimensioned;
            }

            SetValueInRegion(region, Value, type, Units);
        }
    }


    [ActionCategory("MathCad", DisplayName = "Set Value By Name")]
    public class SetValueByName : MathCadAction
    {
        public SetValueByName(IDictionary<string, object> variables, ICentipedeCore core)
            : base("Set Value By Name", variables, core)
        { }

        [ActionArgument]
        public string MathcadVarName = "";

        [ActionArgument]
        public string MathcadVarSubscript = "";

        [ActionArgument(Literal = true)]
        public string Value = "";

        [ActionArgument]
        public string Units = "";

        [ActionArgument]
        public Boolean Complex = false;

        [ActionArgument]
        public bool Text = false;


        protected override void DoAction()
        {
            string varName = ParseStringForVariable(MathcadVarName),
                   varSubscr = ParseStringForVariable(MathcadVarSubscript);

            IMathcadRegion2 region = Worksheet.GetRegionByVar(varName, varSubscr);

            if (region == null)
            {
                throw new ActionException(string.Format("Could not find region for {0}_{1}", varName, varSubscr), this);
            }
            
            McValueType type = Text ? McValueType.String : Complex ? McValueType.Complex : McValueType.Numeric;

            if (!String.IsNullOrEmpty(Units))
            {
                type |= McValueType.Dimensioned;
            }

            SetValueInRegion(region, Value, type, Units);
        }

    }

    [ActionCategory("MathCad", DisplayName = "Get Value")]
    public class GetValue : MathCadAction
    {
        [ActionArgument]
        public string MathCadVarName = "";

        [ActionArgument(Literal = true)]
        public string DestinationVar = "MathCadValue";


        public GetValue(IDictionary<string, object> variables, ICentipedeCore core)
            : base("Get Value", variables, core)
        { }

        protected override void DoAction()
        {
            Variables[MathCadVarName] = Worksheet.GetValue(ParseStringForVariable(MathCadVarName));
        }
    }


    static class Extentions
    {
        public static XmlElement ToMathML(this Complex z, IFormatProvider provider = null)
        {
            if (provider == null)
            {
                provider = CultureInfo.InvariantCulture;
            }

            XmlDocument xmlDoc = new XmlDocument();
            var parens = xmlDoc.CreateElement("ml:parens");

            var apply = xmlDoc.CreateElement("ml:apply");
            parens.AppendChild(apply);

            var plus = xmlDoc.CreateElement("ml:plus");
            var real = xmlDoc.CreateElement("ml:real");
            real.Value = z.Real.ToString(provider);
            var imag = xmlDoc.CreateElement("ml:imag");
            imag.SetAttribute(@"symbol", @"i");
            imag.Value = z.Imaginary.ToString(provider);

            apply.AppendChild(plus);
            apply.AppendChild(real);
            apply.AppendChild(imag);

            return parens;
        }
    }

    
}


