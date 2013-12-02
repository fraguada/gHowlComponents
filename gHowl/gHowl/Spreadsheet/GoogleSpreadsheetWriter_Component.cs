using System;
using System.Collections.Generic;
using System.Linq;

using Grasshopper.Kernel;
using Rhino.Geometry;
using Grasshopper.Kernel.Parameters;
using gHowl.Properties;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Google.GData.Spreadsheets;
using Google.GData.Client;

namespace gHowl.Spreadsheet
{
    

    public class GoogleSpreadsheetWriter_Component : GH_Component, IGH_VariableParameterComponent
    {
        /// <summary>
        /// Initializes a new instance of the Google SpreadsheetWriter_Component class.
        /// Variable Parameter implementation from http://www.grasshopper3d.com/forum/topics/gha-developers-implementing-variable-parameters
        /// </summary>   
        /// 

        int sheetCnt = 1; //count of sheets on the component
        string _user = "";
        string _pass = "";
        string _key = "";

        public GoogleSpreadsheetWriter_Component()
            : base("Google Spreadsheet Writer", "#W", "Write GH Data to a Google Spreadsheet", "gHowl", "#")
        {
            AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, Params.Input.Count.ToString());
            Params.ParameterNickNameChanged += new GH_ComponentParamServer.ParameterNickNameChangedEventHandler(ParamNameChanged);
            if (Params.Input.Count > 4)
            {
                sheetCnt = Params.Input.Count;
            }
            else 
            {
                sheetCnt = 2;
            }
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            //string user, string password, string title
            pManager.AddTextParameter("User", "U", "Google Account username", GH_ParamAccess.item);
            pManager.AddTextParameter("Password", "P", "Google Account password", GH_ParamAccess.item);
            pManager.AddTextParameter("Key", "K", "The key of the spreadsheet you want to write to.", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheet1", "Sheet1", "Sheet1. If sheet with the same name exists in the document, you can overwite its data.", GH_ParamAccess.tree);
            pManager[0].Optional = false;
            pManager[1].Optional = false;
            pManager[2].Optional = false;
            pManager[3].Optional = true;

            

           // VariableParameterMaintenance();
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
            string key = "";

            GH_Structure<IGH_Goo> data = new GH_Structure<IGH_Goo>();

            //Get Required Inputs
            if ((!DA.GetData(0, ref user)) || user == null ||
                (!DA.GetData(1, ref password)) || password == null ||
                (!DA.GetData(2, ref key)) || key == null             
                 )
            {
                return;
            }

            _user = user;
            _pass = password;
            _key = key;

            SpreadsheetsService service = new SpreadsheetsService("gHowl");
            service.setUserCredentials(user, password);

            //TODO: How to handle "public" worksheets?
            WorksheetQuery wsQuery = new WorksheetQuery(key,"private","full");
            WorksheetFeed wsFeed = service.Query(wsQuery);

            //(!DA.GetDataTree<IGH_Goo>(3, out data)) || data == null || data.IsEmpty



            for (int i = 3; i < Params.Input.Count; i++)
            {
                wsQuery = new WorksheetQuery(key, "private", "full");
                wsQuery.Title = Params.Input[i].NickName;
                wsFeed = service.Query(wsQuery);
                 WorksheetEntry worksheet = new WorksheetEntry();
                if (wsFeed.Entries.Count == 0)
                {

                    //worksheet does not exist
                    // Create a local representation of the new worksheet.
                    worksheet = new WorksheetEntry();
                    worksheet.Title.Text = Params.Input[i].NickName;
                    worksheet.Cols = 20;
                    worksheet.Rows = 20;

                    // Send the local representation of the worksheet to the API for
                    // creation.  The URL to use here is the worksheet feed URL of our
                    // spreadsheet.
                    service.Insert(wsFeed, worksheet);
                }
                else {

                    worksheet = (WorksheetEntry)wsFeed.Entries[0];
                }

                
                
                if (DA.GetDataTree<IGH_Goo>(i, out data))
                {
                    //AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, "Yes"+i.ToString());
                    CellQuery cellQuery = new CellQuery(worksheet.CellFeedLink);
                    CellFeed cellFeed = service.Query(cellQuery);

                    List<CellAddress> cellAddrs = new List<CellAddress>();

                    for (int j = 0; j < data.Paths.Count; j++)
                    {
                        GH_Path p = data.get_Path(j);
                        
                        uint row = (uint)j+1;
                        for(int k = 0; k < data.Branches[j].Count; k++)
                        {
                            uint col = (uint)k+1;
                            cellAddrs.Add(new CellAddress(row, col, data.get_DataItem(p, k).ToString()));
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


        private void ParamNameChanged(object sender, GH_ParamServerEventArgs e)
        { 
            if(((e.ParameterSide!=GH_ParameterSide.Output)&&(e.ParameterIndex!=0))&&(e.OriginalArguments.Type==GH_ObjectEventType.NickNameAccepted))
            {

          
                this.ExpireSolution(true);
            }
        }
        public bool CanInsertParameter(GH_ParameterSide side, int index)
        {
            if (side == GH_ParameterSide.Output)
            {
                return false;
            }
            if (index > 3)
            {
                return true;
            }
            else 
            {
                return false;
            }
            //return true;
        }
        public bool CanRemoveParameter(GH_ParameterSide side, int index)
        {
            if (side == GH_ParameterSide.Output)
            {
                return false;
            }
            if (index <= 3 )
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// Called when creating new inputs.  Keeps a count of all the times an input creation has been requested.
        /// Uses this count as the spreadsheet name index.
        /// </summary>
        public IGH_Param CreateParameter(GH_ParameterSide side, int index)
        {
            
            Param_GenericObject param = new Param_GenericObject();
           
            param.NickName = "Sheet" + sheetCnt;
            param.Name = param.NickName;
            param.Optional = true;
            param.Description = param.NickName + " If sheet with the same name exists in the document, you can overwite its data.";
            sheetCnt++;
            return param;
        }
        /// <summary>
        /// Called when destroying inputs.
        /// </summary>
        public bool DestroyParameter(GH_ParameterSide side, int index)
        {
            return true;
        }
        public void VariableParameterMaintenance()
        {
            for (int i = 0; i < Params.Input.Count; i++)
            {
            IGH_Param p = Params.Input[i];
            
            if (i < 3)
            {
                p.Access = GH_ParamAccess.item;
                p.Optional = false;
            }
            else 
            { 
                p.Access = GH_ParamAccess.tree; 
            }
            
            //p.Description = string.Format("{0}. If sheet with the same name exists in the document, you can overwite its data.",p.NickName);
            }
        }

        /// <summary>
        /// Provides an Icon for the component.
        /// </summary>
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Resources.spreadSheetOut;
            }
        }
        /// <summary>
        /// Gets the unique ID for this component. Do not change this ID after release.
        /// </summary>
        public override Guid ComponentGuid
        {
            get { return new Guid("{bbc501ee-439b-470a-a0d1-7d4b6134ffc8}"); }
        }
    }
}