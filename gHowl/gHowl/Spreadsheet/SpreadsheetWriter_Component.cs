using System;
using System.Collections.Generic;

using Grasshopper.Kernel;
using Rhino.Geometry;
using Grasshopper.Kernel.Parameters;
using gHowl.Properties;

namespace gHowl.Spreadsheet
{
    public class SpreadsheetWriter_Component : GH_Component, IGH_VariableParameterComponent
    {
        /// <summary>
        /// Initializes a new instance of the SpreadsheetWriter_Component class.
        /// Variable Parameter implementation from http://www.grasshopper3d.com/forum/topics/gha-developers-implementing-variable-parameters
        /// </summary>

        public SpreadsheetWriter_Component()
            : base("Spreadsheet Writer", "#W", "Write GH Data to a Spreadsheet", "gHowl", "#")
        {
            Params.ParameterNickNameChanged += new GH_ComponentParamServer.ParameterNickNameChangedEventHandler(ParamNameChanged);
        }

        /// <summary>
        /// Registers all the input parameters for this component.
        /// </summary>
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddTextParameter("Path", "P", "The path and filename for the spreadsheet you want to create", GH_ParamAccess.item);
            pManager.AddGenericParameter("Sheet_1", "Sheet_1", "Sheet_1. If sheet with the same name exists in the document, you can overwite its data.", GH_ParamAccess.tree);
            pManager[0].Optional = false;
            pManager[1].Optional = false;

            VariableParameterMaintenance();
        }

        /// <summary>
        /// Registers all the output parameters for this component.
        /// </summary>
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            //pManager.AddTextParameter("Output", "O", "Messages", GH_ParamAccess.tree); 
        }

        /// <summary>
        /// This is the method that actually does the work.
        /// </summary>
        /// <param name="DA">The DA object is used to retrieve from inputs and store in outputs.</param>
        protected override void SolveInstance(IGH_DataAccess DA)
        {
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
            get { return new Guid("{85d972f4-1d45-4f5a-b0c1-da90a4d94f55}"); }
        }

        private void ParamNameChanged(object sender, GH_ParamServerEventArgs e)
        { 
        if(((e.ParameterSide!=GH_ParameterSide.Output)&&(e.ParameterIndex!=0))&&(e.OriginalArguments.Type==GH_ObjectEventType.NickName))
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
            if (index > 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool CanRemoveParameter(GH_ParameterSide side, int index)
        {
            if (side == GH_ParameterSide.Output)
            {
                return false;
            }
            if (index <= 1)
            {
                return false;
            }
            return true;
        }
        public IGH_Param CreateParameter(GH_ParameterSide side, int index)
        {
            Param_GenericObject param = new Param_GenericObject();
            int id = Params.Input.Count;
            param.Name = "Sheet_" + index;
            param.NickName = "Sheet_" + index;
            param.Description = param.NickName + " If sheet with the same name exists in the document, you can overwite its data.";
            return param;
        }
        public bool DestroyParameter(GH_ParameterSide side, int index)
        {
            return true;
        }
        public void VariableParameterMaintenance()
        {
            for (int i = 0; i < Params.Input.Count; i++)
            {
            IGH_Param p = Params.Input[i];
            p.Optional = true;
            p.Access = GH_ParamAccess.tree;               
            p.Name = string.Format("{0}",p.NickName);
            p.Description = string.Format("{0}. If sheet with the same name exists in the document, you can overwite its data.",p.NickName);
            }
        }

        
    }
}