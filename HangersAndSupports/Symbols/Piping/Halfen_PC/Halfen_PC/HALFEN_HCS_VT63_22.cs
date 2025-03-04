//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_HCS_VT63_22.cs
//    Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_VT63_22
//   Author       :Sasidhar  
//   Creation Date:17-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-11-2012     Sasidhar  CR-CP-222275 Converted VB HS_HALFEN_PC Project to C#.Net
//   20-03-2013      Vijay    DI-CP-228142  Modify the error handling for delivered H&S symbols
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
    public class HALFEN_HCS_VT63_22 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_VT63_22"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("MIDDLE", "MIDDLE")]
        [SymbolOutput("SIDE", "SIDE")]
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

                Double height = 0.165;
                Double L = 0.05;
                Double width = 0.073;
                //ports

                Port port1 = new Port(connection, part, "Base", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(connection, part, "Top", new Position(0, 0, height), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(connection, part, "Middle", new Position(0, 0, height / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                Port port4 = new Port(connection, part, "Right", new Position(0, -width / 2.0, height / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port4"] = port4;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d middle = (Projection3d)symbolGeometryHelper.CreateBox(null, height, width, width);
                m_Symbolic.Outputs["MIDDLE"] = middle;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -width / 2.0, (height / 2.0));
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                Projection3d side = (Projection3d)symbolGeometryHelper.CreateBox(null, L, width, width);
                m_Symbolic.Outputs["SIDE"] = side;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HALFEN_HCS_VT63_22.cs."));
                    return;
                }
            }
        }
        #endregion

    }

}
