using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using Centipede;


namespace Centipede.XMLActions
{
    
    public abstract class XmlAction : Centipede.Action
    {
        protected XmlAction(String name, Dictionary<String, Object> variables)
            : base(name, variables)
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
    public class OpenXMLFile : XmlAction
    {
        public OpenXMLFile(Dictionary<String, Object> variables)
            : base("Open XML File", variables)
        { }

        [ActionArgument]
        public String Filename = "";
        protected override void DoAction()
        {
            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(ParseStringForVariable(Filename));
            }
            catch (IOException e)
            {
                
                throw new ActionException("File not found.",  e, this);
            }
            XmlNav = doc.CreateNavigator();
        }
    }

    [ActionCategory("XML", displayName="Get XPath Node as Number", iconName="xml")]
    public class GetXPathNodeAsNumber : XmlAction
    {
        public GetXPathNodeAsNumber(Dictionary<String, Object> v)
            : base("Get Xpath Node as Number", v)
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
            Variables[ParseStringForVariable(ResultVar)] = (Double)XmlNav.Evaluate(xPath);
        }
    }
    [ActionCategory("XML", displayName = "Get XPath Node as String", iconName = "xml")]
    public class GetXmlNodeAsString : XmlAction
    {
        public GetXmlNodeAsString(Dictionary<String, Object> v)
            : base("Get Xpath Node as String", v)
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

            String xPath = String.Format("string({0})", ParseStringForVariable(XPath));
            Variables[ParseStringForVariable(ResultVar)] = XmlNav.Evaluate(xPath) as String;
        }
    }

    
    [ActionCategory("XML", displayName="Count matching nodes", iconName="xml")]
    public class CountMatchingXPathNodes : XmlAction
    {
        public CountMatchingXPathNodes(Dictionary<String, Object> v)
            : base("Count Matching nodes",v)
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
