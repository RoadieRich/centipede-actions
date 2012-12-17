using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using Centipede.XMLActions;


namespace TestFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<String, Object> variables = new Dictionary<string, object>();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(@"<root><level1><level2 attr1=""3.141"">Tag contents</level2></level1></root>");
            variables["xmlDoc"] = doc.CreateNavigator();



            GetXmlNodeAsString action = new GetXmlNodeAsString(variables);
            action.XmlFileVar = "xmlDoc";
            action.XPath = "number(//level2@attr1)";

            action.Run();

            foreach (var item in variables)
            {
                Console.Out.WriteLine(String.Format("{0}: {1}", item.Key, item.Value));
            }
            Console.ReadKey();

        }
    }
}
