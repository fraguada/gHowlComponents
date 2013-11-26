using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Grasshopper.Kernel.Data;
using System.Text;
using System.IO;
using System.Xml;
using System.Drawing;
using gHowl.Properties;
using Grasshopper;


namespace gHowl.Geo
{
    public class ElevationComponent : GH_Component
    {
        public ElevationComponent() : base("Get Elevation", "E", "Given WGS84 coordinates, this component will return the elevation(s)", "gHowl", "GEO") { }

        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddPointParameter("Geo Points", "P", "WSG84 Decimal Degree formatted points", GH_ParamAccess.list);
            pManager[0].Optional = false;
        }

        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.Register_DoubleParam("Geo Points", "P", "Elevated Points");
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            List<GH_Point> inPts = new List<GH_Point>();
            if (!DA.GetDataList<GH_Point>(0,inPts)) { return; }
            int cnt = 0;
            string preUrl = "http://maps.googleapis.com/maps/api/elevation/xml?locations=";
            string postUrl = "&sensor=false"; //13 Characters
            System.Text.StringBuilder url = new System.Text.StringBuilder();
   
            url.Append(preUrl);
            List<double> outPts = new List<double>();
            int l;

            for (int i = 0; i < inPts.Count; i++)
            {
                l = url.Length;
                l = l + postUrl.Length;


                if (l > 1870 || i == (inPts.Count - 1))
                {
                    if (i == inPts.Count - 1)
                    {
         
                        url.Append(inPts[i].Value.Y);
                        url.Append(",");
                        url.Append(inPts[i].Value.X);
                        url.Append("|");
                    }
                    url.Remove(url.Length - 1, 1);

                    url.Append(postUrl);

                    cnt = 0;
                    try
                    {
                        System.Net.HttpWebRequest HttpWReq = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url.ToString());
                        System.Net.HttpWebResponse HttpWResp = (System.Net.HttpWebResponse)HttpWReq.GetResponse(); // Insert code that uses the response object.
                        Stream receiveStream = HttpWResp.GetResponseStream();
                        System.Text.Encoding encode = System.Text.Encoding.GetEncoding("utf-8");
                        StreamReader readStream = new StreamReader(receiveStream, encode);
                       
                        string xmlData = readStream.ReadToEnd();
                        XmlTextReader reader = new XmlTextReader(new StringReader(xmlData));

                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Name == "status")
                            {
                               // AddRuntimeMessage(GH_RuntimeMessageLevel., reader.ReadElementContentAsString());
                                Instances.DocumentEditor.SetStatusBarEvent(new GH_RuntimeMessage("API Status "+reader.ReadElementContentAsString()));
                            }
                            if (reader.NodeType == XmlNodeType.Element && reader.Name == "elevation")
                            {
                                outPts.Add(reader.ReadElementContentAsDouble());
                            }
                        }

                        reader.Close();
                        HttpWResp.Close();
                        readStream.Close();

                        url.Clear();
                        url.Append(preUrl);
                    }
                    catch (Exception ex)
                    {
                        
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Error,"Exception: " + ex.Message);
                        url.Clear();

                        url.Append(preUrl);
                    }
                }
                else
                {

                }

                url.Append(inPts[i].Value.Y);
                url.Append(",");
                url.Append(inPts[i].Value.X);
                url.Append("|");
                cnt++;
            }
         
            DA.SetDataList(0, outPts);



        }

        public override Guid ComponentGuid
        {
            get { return new Guid("{4c3894ec-8fb3-4a33-910f-5c3cd2614288}"); }
        }

        protected override Bitmap Icon
        {
            get
            {
                return Resources.gHowl_elev;
            }
        }
    }
}
