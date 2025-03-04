//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_HRG_4.cs
//    Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HRG_4
//   Author       :Sasidhar 
//   Creation Date:19-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19-11-2012     Sasidhar  CR-CP-222275 Converted VB HS_HALFEN_PC Project to C#.Net
//   20-03-2013      Vijay    DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class HALFEN_HRG_4 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HRG_4"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputString(3, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(5, "Depth", "Depth", 0.999999)]
        public InputDouble m_dDepth;
        [InputDouble(6, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(7, "Shoe_Th", "Shoe_Th", 0.999999)]
        public InputDouble m_dShoe_Th;
        [InputDouble(8, "Clamp_Th", "Clamp_Th", 0.999999)]
        public InputDouble m_dClamp_Th;
        [InputDouble(9, "Clamp_Depth", "Clamp_Depth", 0.999999)]
        public InputDouble m_dClamp_Depth;
        [InputDouble(10, "Ear_Length", "Ear_Length", 0.999999)]
        public InputDouble m_dEar_Length;
        [InputString(11, "Finish", "Finish", "No Value")]
        public InputString m_oFinish;
        [InputString(12, "Stock_Number", "Stock_Number", "No Value")]
        public InputString m_oStock_Number;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("PIPE_CLAMP_BODY1", "PIPE_CLAMP_BODY1")]
        [SymbolOutput("LEFT_BOTTOM1", "LEFT_BOTTOM1")]
        [SymbolOutput("LEFT_TOP1", "LEFT_TOP1")]
        [SymbolOutput("RIGHT_BOTTOM1", "RIGHT_BOTTOM1")]
        [SymbolOutput("RIGHT_TOP1", "RIGHT_TOP1")]
        [SymbolOutput("PIPE_CLAMP_BODY2", "PIPE_CLAMP_BODY2")]
        [SymbolOutput("LEFT_BOTTOM2", "LEFT_BOTTOM2")]
        [SymbolOutput("LEFT_TOP2", "LEFT_TOP2")]
        [SymbolOutput("RIGHT_BOTTOM2", "RIGHT_BOTTOM2")]
        [SymbolOutput("RIGHT_TOP2", "RIGHT_TOP2")]
        [SymbolOutput("UPRIGHT", "UPRIGHT")]
        [SymbolOutput("SHOE", "SHOE")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double H = m_dH.Value;
                Double depth = m_dDepth.Value;
                Double width = m_dWidth.Value;
                Double shoeThickness = m_dShoe_Th.Value;
                Double clampThickness = m_dClamp_Th.Value;
                Double clampDepth = m_dClamp_Depth.Value;
                Double earLength = m_dEar_Length.Value;
                Double spacer = 0.005;
                Double da = pipeDiameter;

                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrHArguments, "H can not be zero or negative"));
                    return;
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrDepthArguments, "Depth can not be zero or negative"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrWidthArguments, "Width can not be zero or negative"));
                    return;
                }
                if (shoeThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrShoeThickNessArguments, "Shoe_Th can not be zero or negative"));
                    return;
                }
                if (clampThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrClampThickNessArguments, "Clamp_Th can not be zero or negative"));
                    return;
                }
                if (clampDepth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrClampDepthArguments, "Clamp_Depth can not be zero or negative"));
                    return;
                }
                if (earLength <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrEarLengthArguments, "Ear_Length can not be zero or negative"));
                    return;
                }
                if (da < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrDAArguments, "PIPE_DIA can not be negative"));
                    return;
                }

                //ports

                Port port1 = new Port(connection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(connection, part, "Structure", new Position(0, 0, -H), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(depth / 2, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(-1, 0, 0), new Vector(0, 0, 1));
                Projection3d pipeClampBody1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, (da + 2 * clampThickness) / 2.0, clampDepth);
                m_Symbolic.Outputs["PIPE_CLAMP_BODY1"] = pipeClampBody1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(depth / 2.0 - clampDepth / 2.0, (da + 1.5 * clampThickness) / 2 + earLength / 2.0, spacer / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d leftTop1 = (Projection3d)symbolGeometryHelper.CreateBox(null, clampThickness, clampDepth, earLength);
                m_Symbolic.Outputs["LEFT_TOP1"] = leftTop1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(depth / 2.0 - clampDepth / 2.0, (da + 1.5 * clampThickness) / 2.0 + earLength / 2, -spacer / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d leftBottom1 = (Projection3d)symbolGeometryHelper.CreateBox(null, clampThickness, clampDepth, earLength);
                m_Symbolic.Outputs["LEFT_BOTTOM1"] = leftBottom1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(depth / 2.0 - clampDepth / 2.0, -((da + 1.5 * clampThickness) / 2.0 + earLength / 2), spacer / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rightTop1 = (Projection3d)symbolGeometryHelper.CreateBox(null, clampThickness, clampDepth, earLength);
                m_Symbolic.Outputs["RIGHT_TOP1"] = rightTop1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(depth / 2.0 - clampDepth / 2.0, -((da + 1.5 * clampThickness) / 2.0 + earLength / 2), -spacer / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d rightBottom1 = (Projection3d)symbolGeometryHelper.CreateBox(null, clampThickness, clampDepth, earLength);
                m_Symbolic.Outputs["RIGHT_BOTTOM1"] = rightBottom1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-depth / 2, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, 1));
                Projection3d pipeClampBody2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, (da + 2 * clampThickness) / 2, clampDepth);
                m_Symbolic.Outputs["PIPE_CLAMP_BODY2"] = pipeClampBody2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2.0 - clampDepth / 2.0), (da + 1.5 * clampThickness) / 2 + earLength / 2.0, spacer / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d leftTop2 = (Projection3d)symbolGeometryHelper.CreateBox(null, clampThickness, clampDepth, earLength);
                m_Symbolic.Outputs["LEFT_TOP2"] = leftTop2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2.0 - clampDepth / 2.0), (da + 1.5 * clampThickness) / 2.0 + earLength / 2, -spacer / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d leftBottom2 = (Projection3d)symbolGeometryHelper.CreateBox(null, clampThickness, clampDepth, earLength);
                m_Symbolic.Outputs["LEFT_BOTTOM2"] = leftBottom2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2.0 - clampDepth / 2.0), -((da + 1.5 * clampThickness) / 2.0 + earLength / 2), spacer / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rightTop2 = (Projection3d)symbolGeometryHelper.CreateBox(null, clampThickness, clampDepth, earLength);
                m_Symbolic.Outputs["RIGHT_TOP2"] = rightTop2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2.0 - clampDepth / 2.0), -((da + 1.5 * clampThickness) / 2.0 + earLength / 2), -spacer / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d rightBottom2 = (Projection3d)symbolGeometryHelper.CreateBox(null, clampThickness, clampDepth, earLength);
                m_Symbolic.Outputs["RIGHT_BOTTOM2"] = rightBottom2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -((da + 2 * clampThickness) / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d upRight = (Projection3d)symbolGeometryHelper.CreateBox(null, H - da / 2 - clampThickness - shoeThickness, depth, shoeThickness);
                m_Symbolic.Outputs["UPRIGHT"] = upRight;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -(H - shoeThickness));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d shoe = (Projection3d)symbolGeometryHelper.CreateBox(null, shoeThickness, depth, width);
                m_Symbolic.Outputs["SHOE"] = shoe;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HALFEN_HRG_4.cs."));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                String stockNumber;
                String finish;
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                stockNumber = (String)((PropertyValueString)part.GetPropertyValue("IJUAHgrStock_Number", "Stock_Number")).PropValue;
                finish = (String)((PropertyValueString)part.GetPropertyValue("IJUAHgrFinish", "Finish")).PropValue;
                bomString = part.PartDescription + " - " + finish + " - " + stockNumber;
                return bomString;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBOMDescription, "Error in BOMDescription of HALFEN_HRG_4.cs."));
                }
                return "";
            }
        }
        #endregion

    }

}
