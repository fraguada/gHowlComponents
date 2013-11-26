using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Rhino.Geometry;
using Google.GData.Spreadsheets;

namespace gHowl.Spreadsheet
{
    public class GSpreadsheetReader : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent1 class.
        /// </summary>
        public GSpreadsheetReader()
            : base("Google Spreadsheet Reader", "#R", "Import spreadsheet data to GH", "gHowl", "#")
        {
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            //string user, string password, string title
            pManager.AddTextParameter("User", "u", "Google Account username", GH_ParamAccess.item);
            pManager.AddTextParameter("Password", "p", "Google Account password", GH_ParamAccess.item);
            pManager.AddTextParameter("Title", "t", "Google Spreadsheet title", GH_ParamAccess.item);

            pManager[0].Optional = false;
            pManager[1].Optional = false;
            pManager[2].Optional = false;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("Data","d","The retrieved data",GH_ParamAccess.tree);
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string user = "";
            string password  = "";
            string title = "";

            //Get Required Inputs
            if ((!DA.GetData(0, ref user)) || user == null ||
                (!DA.GetData(1, ref password)) || password == null ||
                (!DA.GetData(2, ref title)) || title == null
                 )
            {
                return;
            }

            SpreadsheetsService service = new SpreadsheetsService("gHowl");
            service.setUserCredentials(user, password);

            SpreadsheetQuery query = new SpreadsheetQuery();
            query.Title = title;
            SpreadsheetFeed feed = service.Query(query);

            //if we don't find anything, don't do anything
            if (!(feed.Entries.Count > 0))
            {
               // Print("No spreadsheet matches this title.");
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No spreadsheet matches this title.");
                return;
            }


            //TODO: What do we do if we find multiple spreadsheets that match this name?
            SpreadsheetEntry spreadsheet = (SpreadsheetEntry)feed.Entries[0];
            WorksheetFeed wfeed = spreadsheet.Worksheets; //get the worksheets in this spreadsheet

            GH_Structure<GH_String> dt = new GH_Structure<GH_String>();
            //DataTree<object> dt = new DataTree<object>();

            //go through worksheets
            for (int i = 0; i < wfeed.Entries.Count; i++)
            {
                WorksheetEntry worksheet = (WorksheetEntry)wfeed.Entries[i];
                CellQuery cquery = new CellQuery(worksheet.CellFeedLink);
                CellFeed cfeed = service.Query(cquery);

                uint m_rows = worksheet.Rows;
                uint m_columns = worksheet.Cols;

                for (int k = 1; k <= m_rows; k++)
                {
                    int[] args = { i, k };
                    GH_Path p = new GH_Path(args);

                    for (int j = 1; j <= m_columns; j++)
                    {
                        foreach (CellEntry curCell in cfeed.Entries)
                        {
                            if (curCell.Cell.Row == k && curCell.Cell.Column == j)
                            {
                                dt.Insert(new GH_String(curCell.Cell.Value.ToString()), p, j - 1);
                            }

                        }
                    }
                }
            }
           
            dt.Graft(GH_GraftMode.GraftAll);

            DA.SetDataTree(0, dt);
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                //You can add image files to your project resources and access them like this:
                // return Resources.IconForThisComponent;
                return null;
            }
        }

        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("{4febe0c6-c3c0-4517-88da-7ccf55cca325}"); }
        }
    }
}