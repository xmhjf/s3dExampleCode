//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_HCS_VT63_11_12_13.cs
//    Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_VT63_11_12_13
//   Author       :Sasidhar  
//   Creation Date:17-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-11-2012     Sasidhar CR-CP-222275 Converted VB HS_HALFEN_PC Project to C#.Net
//   20-03-2013      Vijay   DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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
    public class HALFEN_HCS_VT63_11_12_13 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_VT63_11_12_13"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "Stock_Number", "Stock_Number", "No Value")]
        public InputString m_oStock_Number;
        [InputDouble(3, "Beam_Flange_W_Min", "Beam_Flange_W_Min", 0.999999)]
        public InputDouble m_dBeam_Flange_W_Min;
        [InputDouble(4, "Beam_Flange_W_Max", "Beam_Flange_W_Max", 0.999999)]
        public InputDouble m_dBeam_Flange_W_Max;
        [InputString(5, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        [InputDouble(6, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(7, "Depth", "Depth", 0.999999)]
        public InputDouble m_dDepth;
        [InputDouble(8, "Bolt_CC2", "Bolt_CC2", 0.999999)]
        public InputDouble m_dBolt_CC2;
        [InputDouble(9, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(10, "Bolt_CC", "Bolt_CC", 0.999999)]
        public InputDouble m_dBolt_CC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("PLATE", "PLATE")]
        [SymbolOutput("SQUARE_TUBE", "SQUARE_TUBE")]
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

                Double beamFlangeWMin = m_dBeam_Flange_W_Min.Value;
                Double beamFlangeWMax = m_dBeam_Flange_W_Max.Value;
                Double width = m_dWidth.Value;
                Double depth = m_dDepth.Value;
                Double boltCC2 = m_dBolt_CC2.Value;
                Double T = m_dT.Value;
                Double boltCC = m_dBolt_CC.Value;

                const Double value1 = 0.14;
                const Double value2 = 0.073;
                const Double value3 = 0.073;

                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrTArguments, "T can not be zero or negative"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrWidthArguments, "Width can not be zero or negative"));
                    return;
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrDepthArguments, "Depth can not be zero or negative"));
                    return;
                }

                //ports

                Port port1 = new Port(connection, part, "Base", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(connection, part, "Top", new Position(0, 0, T + 0.14), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d plate = (Projection3d)symbolGeometryHelper.CreateBox(null, T, depth, width);
                m_Symbolic.Outputs["PLATE"] = plate;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d squareTube = (Projection3d)symbolGeometryHelper.CreateBox(null, value1, value2, value3);
                m_Symbolic.Outputs["SQUARE_TUBE"] = squareTube;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HALFEN_HCS_VT63_11_12_13.cs."));
                    return;
                }
            }
        }
        #endregion

    }

}
