//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_HCS_TK.cs
//    Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_TK
//   Author       :Sasidhar  
//   Creation Date:18-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-11-2012     Sasidhar  CR-CP-222275 Converted VB HS_HALFEN_PC Project to C#.Net
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
    public class HALFEN_HCS_TK : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HCS_TK"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "Stock_Number", "Stock_Number", "No Value")]
        public InputString m_oStock_Number;
        [InputString(3, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("CLAMP", "CLAMP")]
        [SymbolOutput("BOLT", "BOLT")]
        [SymbolOutput("LEFT_BOLT", "LEFT_BOLT")]
        [SymbolOutput("RIGHT_BOLT", "RIGHT_BOLT")]
        [SymbolOutput("BOT_BOLT", "BOT_BOLT")]
        [SymbolOutput("TOP_BOX", "TOP_BOX")]
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

                Double chanellW = 0.063;
                Double boltDiameter = 0.013;
                String partSize = m_oSIZE.Value;

                const Double value1=0.049;
                const Double value2 = 0.044;
                const Double value3 = 0.081;
                const Double value4 = 0.01;

                //ports

                Port port1 = new Port(connection, part, "BottomOfClamp", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (partSize == "TK-63")
                {
                    Port port2 = new Port(connection, part, "CenterUBolt", new Position(0, 0, -chanellW / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["port2"] = port2;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d clamp = (Projection3d)symbolGeometryHelper.CreateBox(null, value1, value2, value3);
                m_Symbolic.Outputs["CLAMP"] = clamp;

                if (partSize == "TK - L")
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, -chanellW - value4);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d bolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2, value1 + value4 + chanellW + value4);
                    m_Symbolic.Outputs["BOLT"] = bolt;
                }
                if (partSize == "TK - 63")
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-chanellW / 2, 0, -chanellW - boltDiameter / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d leftBolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2, value1 + value4 + chanellW + (boltDiameter / 2));
                    m_Symbolic.Outputs["LEFT_BOLT"] = leftBolt;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(chanellW / 2, 0, -chanellW - boltDiameter / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d rightBolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2, value1 + value4 + chanellW + (boltDiameter / 2));
                    m_Symbolic.Outputs["RIGHT_BOLT"] = rightBolt;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-chanellW / 2, 0, -chanellW - boltDiameter / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, 1));
                    Projection3d bottomBolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2, chanellW);
                    m_Symbolic.Outputs["BOT_BOLT"] = bottomBolt;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, value1);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d topBox = (Projection3d)symbolGeometryHelper.CreateBox(null, value4 / 2, chanellW + boltDiameter * 2, boltDiameter * 2);
                    m_Symbolic.Outputs["TOP_BOX"] = topBox;
                }
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HALFEN_HCS_TK.cs."));
                    return;
                }
            }
        }
        #endregion
    }
}
