using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.XPath;
using CentipedeInterfaces;


namespace XMLActions
{
    
    public abstract class XmlAction : Centipede.Action
    {
        protected XmlAction(String name, IDictionary<string, object> variables, ICentipedeCore c)
            : base(name, variables, c)
        { }
        
        protected override void InitAction()
        {
            base.InitAction();
            Object obj;
            Variables.TryGetValue(XmlFileVar, out obj);
            XmlNav = obj as XPathNavigator;
            
        }

        protected override void CleanupAction()
        {
            base.CleanupAction();
            Variables[XmlFileVar] = XmlNav;
        }

        protected XPathNavigator XmlNav;
        
        [ActionArgument(displayName = "XML File variable name", usage = "Name of variable used to store the xml file")]
        public String XmlFileVar = "xmlFile";

    }

    [ActionCategory("XML", iconName = "xml", displayName="Open XML File")]
    public class OpenXmlFile : XmlAction
    {
        public OpenXmlFile(IDictionary<string, object> variables, ICentipedeCore c)
            : base("Open XML File", variables, c)
        { }

        [ActionArgument]
        public String Filename = "";
        protected override void DoAction()
        {
            XPathDocument doc;
            string filename = ParseStringForVariable(Filename);
            try
            {
                doc = new XPathDocument(filename);
            }
            catch (IOException e)
            {
                throw new ActionException(string.Format("File {0} not found.", this.Filename),  e, this);
            }
            XmlNav = doc.CreateNavigator();
        }
    }

    [ActionCategory("XML", displayName="Get XPath Node as Number", iconName="xml")]
    public class GetXPathNodeAsNumber : XmlAction
    {
        public GetXPathNodeAsNumber(IDictionary<string, object> v, ICentipedeCore c)
            : base("Get Xpath Node as Number", v, c)
        { }

        [ActionArgument]
        public String XPath = "";

        [ActionArgument(displayName = "Variable to save result")]
        public String ResultVar = "XPathNum";

        protected override void DoAction()
        {
            if (XmlNav == null)
            {
                throw new ActionException("XML file not opened, or incorrect variable name", this);
            }

            String xPath = String.Format("number({0})", ParseStringForVariable(XPath));
            Variables[ParseStringForVariable(ResultVar)] = (Double)XmlNav.Evaluate(ParseStringForVariable(xPath));
        }
    }
    [ActionCategory("XML", displayName = "Get XPath Node as String", iconName = "xml")]
    public class GetXmlNodeAsString : XmlAction
    {
        public GetXmlNodeAsString(IDictionary<string, object> v, ICentipedeCore c)
            : base("Get Xpath Node as String", v, c)
        { }

        [ActionArgument]
        public String XPath = "";

        [ActionArgument(displayName = "Variable to save result")]
        public String ResultVar = "XPathString";

        protected override void DoAction()
        {
            if (XmlNav == null)
            {
                throw new ActionException("XML file not opened, or incorrect variable name", this);
            }

            string parsed = ParseStringForVariable(XPath);
            String xPath = String.Format("string({0})", parsed);
            string varname = ParseStringForVariable(ResultVar);
            Variables[varname] = XmlNav.Evaluate(xPath) as String;
        }
    }

    
    [ActionCategory("XML", displayName="Count matching nodes", iconName="xml")]
    public class CountMatchingXPathNodes : XmlAction
    {
        public CountMatchingXPathNodes(IDictionary<string, object> v, ICentipedeCore c)
            : base("Count Matching nodes",v, c)
        { }

        [ActionArgument]
        public String XPath = "";

        [ActionArgument(displayName = "Variable to save result")]
        public String ResultVar = "XPathMatchCount";

        protected override void DoAction()
        {
            if (XmlNav == null)
            {
                throw new ActionException("XML file not opened, or incorrect variable name", this);
            }

            String xPath = String.Format("number(count({0}))", ParseStringForVariable(XPath));
            var result = XmlNav.Evaluate(xPath);

            String parsedResultVar = ParseStringForVariable(ResultVar);

            Variables[parsedResultVar] = result;
            
        }
    }
}
