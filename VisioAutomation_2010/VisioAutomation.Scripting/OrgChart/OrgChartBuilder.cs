using System.Collections.Generic;
using VA = VisioAutomation;
using OCMODEL = VisioAutomation.Models.OrgChart;
using SXL = System.Xml.Linq;

namespace VisioAutomation.Scripting.OrgChart
{
    public class OrgChartBuilder
    {
        public static OCMODEL.OrgChartDocument LoadFromXML(Session scriptingsession, string filename)
        {
            var xdoc = SXL.XDocument.Load(filename);
            return LoadFromXML(scriptingsession, xdoc);
        }

        public static OCMODEL.OrgChartDocument LoadFromXML(Session scriptingsession,
                                                             SXL.XDocument xdoc)
        {
            var root = xdoc.Root;

            var dic = new Dictionary<string, OCMODEL.Node>();
            OCMODEL.Node ocroot = null;

            scriptingsession.WriteVerbose("Walking XML");

            foreach (var ev in root.Elements())
            {
                if (ev.Name == "shape")
                {
                    string id = ev.Attribute("id").Value;
                    string parentid = VA.Scripting.XmlUtil.GetAttributeValue(ev, "parentid", null);
                    var name = ev.Attribute("name").Value;

                    scriptingsession.WriteVerbose( "Loading shape: {0} {1} {2}", id, name, parentid);
                    var new_ocnode = new OCMODEL.Node(name);

                    if (ocroot == null)
                    {
                        ocroot = new_ocnode;
                    }

                    dic[id] = new_ocnode;

                    if (parentid != null)
                    {
                        if (dic.ContainsKey(parentid))
                        {
                            var parent = dic[parentid];
                            parent.Children.Add(new_ocnode);
                        }
                    }
                }
            }
            scriptingsession.WriteVerbose( "Finished Walking XML");
            var oc = new OCMODEL.OrgChartDocument();
            oc.OrgCharts.Add(ocroot);
            scriptingsession.WriteVerbose( "Finished Creating OrgChart model");
            return oc;
        }
    }
}