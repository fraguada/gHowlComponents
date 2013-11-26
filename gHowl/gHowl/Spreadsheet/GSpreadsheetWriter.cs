using System;
using System.Collections.Generic;
using System.Linq;

using Grasshopper.Kernel;
using Rhino.Geometry;
using Google.GData.Spreadsheets;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Google.GData.Client;

namespace gHowl.Spreadsheet
{


     class CellAddress
    {
        public uint Row;
        public uint Col;
        public string IdString;
        public string DataString;

        /**
         * Constructs a CellAddress representing the specified {@code row} and
         * {@code col}. The IdString will be set in 'RnCn' notation.
         */
        public CellAddress(uint row, uint col)
        {
            this.Row = row;
            this.Col = col;
            this.IdString = string.Format("R{0}C{1}", row, col);
        }

        public CellAddress(uint row, uint col, string data)
        {
            this.Row = row;
            this.Col = col;
            this.DataString = data;
            this.IdString = string.Format("R{0}C{1}", row, col);
        }
    }


    public class GSpreadsheetWriter : GH_Component
    {
        /// <summary>
        /// Initializes a new instance of the MyComponent2 class.
        /// </summary>
        public GSpreadsheetWriter()
            : base("Google Spreadsheet Writer", "#W", "Write spreadsheet data to Google", "gHowl", "#")
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
            pManager.AddGenericParameter("Data","d","Data to write",GH_ParamAccess.tree);

            pManager[0].Optional = false;
            pManager[1].Optional = false;
            pManager[2].Optional = false;
            pManager[3].Optional = false;

        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {

            string user = "";
            string password = "";
            string title = "";
            GH_Structure<IGH_Goo> data = new GH_Structure<IGH_Goo>();

            //Get Required Inputs
            if ((!DA.GetData(0, ref user)) || user == null ||
                (!DA.GetData(1, ref password)) || password == null ||
                (!DA.GetData(2, ref title)) || title == null ||
                (!DA.GetDataTree<IGH_Goo>(3, out data)) || data == null || data.IsEmpty
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

            //Check Worksheets, add to the document if they do not exist
            int wsCount = wfeed.Entries.Count;

            //create list of unique roots
            List<int> roots = new List<int>();

            for (int i = 0; i < data.Paths.Count; i++)
            {
                GH_Path p = data.get_Path(i);
                int[] id = p.Indices;
                roots.Add(p[0]);
            }
           
            int[] uniqueRoots = roots.Distinct().ToArray();
          //  Print("Unique Roots {0}", uniqueRoots.Length);

            if (uniqueRoots.Length != wsCount)
            {
               // Print("YES");
                //create the necessary worksheets
                for (int i = 0; i < uniqueRoots.Length - wsCount; i++)
                {
                    // Create a local representation of the new worksheet.
                    WorksheetEntry worksheet = new WorksheetEntry();
                    worksheet.Title.Text = "Sheet" + ((wsCount) + (i + 1)).ToString();
                    worksheet.Cols = 20;
                    worksheet.Rows = 20;

                    // Send the local representation of the worksheet to the API for
                    // creation.  The URL to use here is the worksheet feed URL of our
                    // spreadsheet.
                    service.Insert(wfeed, worksheet);
                }
            }
            else
            {
             //   Print("NOT NECESSARY");
            }

            //update ws count
            wsCount = spreadsheet.Worksheets.Entries.Count;


            for (int i = 0; i < wsCount; i++)
            {
                WorksheetEntry worksheet = (WorksheetEntry)wfeed.Entries[i];

                CellQuery cellQuery = new CellQuery(worksheet.CellFeedLink);
                CellFeed cellFeed = service.Query(cellQuery);

                List<CellAddress> cellAddrs = new List<CellAddress>();

                for (int j = 0; j < data.Paths.Count; j++)
                {
                    GH_Path p = data.get_Path(j);
                    int[] id = p.Indices;

                    if (id[0] == i)
                    {
                        uint row = (uint)id[1] + 1;
                        uint col = (uint)id[2] + 1;

                        cellAddrs.Add(new CellAddress(row, col, data.get_DataItem(p, 0).ToString()));

                    }


                }

                Dictionary<String, CellEntry> cellEntries = GetCellEntryMap(service, cellFeed, cellAddrs);

                CellFeed batchRequest = new CellFeed(cellQuery.Uri, service);
                foreach (CellAddress cellAddr in cellAddrs)
                {
                    CellEntry batchEntry = cellEntries[cellAddr.IdString];
                    batchEntry.InputValue = cellAddr.DataString;
                    batchEntry.BatchData = new GDataBatchEntryData(cellAddr.IdString, GDataBatchOperationType.update);
                    batchRequest.Entries.Add(batchEntry);
                }

                // Submit the update
                CellFeed batchResponse = (CellFeed)service.Batch(batchRequest, new Uri(cellFeed.Batch));

                // Check the results
                bool isSuccess = true;
                foreach (CellEntry entry in batchResponse.Entries)
                {
                    string batchId = entry.BatchData.Id;
                    if (entry.BatchData.Status.Code != 200)
                    {
                        isSuccess = false;
                        GDataBatchStatus status = entry.BatchData.Status;
                      //  Print("{0} failed ({1})", batchId, status.Reason);
                    }
                }

              //  Print(isSuccess ? "Batch operations successful." : "Batch operations failed");
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, isSuccess ? "Batch operations successful." : "Batch operations failed");

            }

        }

        private static Dictionary<String, CellEntry> GetCellEntryMap(SpreadsheetsService service, CellFeed cellFeed, List<CellAddress> cellAddrs)
        {
            CellFeed batchRequest = new CellFeed(new Uri(cellFeed.Self), service);
            foreach (CellAddress cellId in cellAddrs)
            {
                CellEntry batchEntry = new CellEntry(cellId.Row, cellId.Col, cellId.IdString);
                batchEntry.Id = new AtomId(string.Format("{0}/{1}", cellFeed.Self, cellId.IdString));
                batchEntry.BatchData = new GDataBatchEntryData(cellId.IdString, GDataBatchOperationType.query);
                batchRequest.Entries.Add(batchEntry);
            }

            CellFeed queryBatchResponse = (CellFeed)service.Batch(batchRequest, new Uri(cellFeed.Batch));

            Dictionary<String, CellEntry> cellEntryMap = new Dictionary<String, CellEntry>();
            foreach (CellEntry entry in queryBatchResponse.Entries)
            {
                cellEntryMap.Add(entry.BatchData.Id, entry);
                //  Print("batch {0} (CellEntry: id={1} editLink={2} inputValue={3})",
                //    entry.BatchData.Id, entry.Id, entry.EditUri,
                //     entry.InputValue);
            }

            return cellEntryMap;
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
            get { return new Guid("{9de92fda-84d0-49a0-8f10-e721be8fb43e}"); }
        }
    }
}