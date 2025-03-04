//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LRParts_FIG295HN.cs
//    LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG295HN
//   Author       :Sasidhar  
//   Creation Date:27-10-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27-10-2012     Sasidhar  Initial Creation
//  26/03/2013     Rajeswari DI-CP-228142  Modify the error handling for delivered H&S symbols
//  30/10/2013     Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
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
    public class LRParts_FIG295HN : CustomSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG295HN"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PipeDiam", "PipeDiam", 0.999999)]
        public InputDouble m_PipeDiam;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble m_B;
        [InputDouble(4, "C", "C", 0.999999)]
        public InputDouble m_C;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(6, "E", "E", 0.999999)]
        public InputDouble m_E;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble m_F;
        [InputDouble(8, "G1", "G1", 0.999999)]
        public InputDouble m_G1;
        [InputDouble(9, "G2", "G2", 0.999999)]
        public InputDouble m_G2;
        [InputDouble(10, "H", "H", 0.999999)]
        public InputDouble m_H;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("BOT1", "BOT1")]
        [SymbolOutput("BOT2", "BOT2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
        [SymbolOutput("BOT_BOLT", "BOT_BOLT")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("MID_BOLT", "MID_BOLT")]
        public AspectDefinition m_oSymbolic;

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
                Double pipeDiameter = m_PipeDiam.Value;
                Double B = m_B.Value;
                Double C = m_C.Value;
                Double D = m_D.Value;
                Double E = m_E.Value;
                Double F = m_F.Value;
                Double G1 = m_G1.Value;
                Double G2 = m_G2.Value;
                Double H = m_H.Value;
                if (G1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidG1, "G1 cannot be zero or negative"));
                    return;
                }
                if (G2 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidG2, "G2 cannot be zero or negative"));
                    return;
                }
                if (D <= 0 && pipeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidDandPipeDiameterGT0, "D and Pipe Diameter cannot be zero or negative"));
                    return;
                }
                if (H <= 0 && pipeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidHandPipeDiameterGT0, "H and Pipe Diameter cannot be zero or negative"));
                    return;
                }
                if (F <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidF, "F cannot be zero or negative"));
                    return;
                }
                if (pipeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidPipeDiameter, "Pipe Diameter cannot be zero or negative"));
                    return;
                }
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, (E - F / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2) - G1 / 2, pipeDiameter / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d top1 = (Projection3d)symbolGeometryHelper.CreateBox(null, D - pipeDiameter / 2, G2, G1);
                m_oSymbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (C / 2) + G1 / 2, pipeDiameter / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d top2 = (Projection3d)symbolGeometryHelper.CreateBox(null, D - pipeDiameter / 2, G2, G1);
                m_oSymbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2) - G1 / 2, -H);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bot1 = (Projection3d)symbolGeometryHelper.CreateBox(null, H - pipeDiameter / 2, G2, G1);
                m_oSymbolic.Outputs["BOT1"] = bot1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (C / 2) + G1 / 2, -H);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bot2 = (Projection3d)symbolGeometryHelper.CreateBox(null, H - pipeDiameter / 2, G2, G1);
                m_oSymbolic.Outputs["BOT2"] = bot2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector midcyl = new Position(0, C / 2 + 2 * G1, B).Subtract(new Position(0, -(C / 2 + 2 * G1), B));
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + 2 * G1), B);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d midBolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, F / 2, midcyl.Length);
                m_oSymbolic.Outputs["MID_BOLT"] = midBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector topboltvector = new Position(0, C / 2 + 2 * G1, E).Subtract(new Position(0, -(C / 2 + 2 * G1), E));
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + 2 * G1), E);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d topBolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, F / 2, topboltvector.Length);
                m_oSymbolic.Outputs["TOP_BOLT"] = topBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector bottomboltvector = new Position(0, C / 2 + 2 * G1, -B).Subtract(new Position(0, -(C / 2 + 2 * G1), -B));
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + 2 * G1), -B);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d botBolt = (Projection3d)symbolGeometryHelper.CreateCylinder(null, F / 2, bottomboltvector.Length);
                m_oSymbolic.Outputs["BOT_BOLT"] = botBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-G2 / 2, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 0, 1));
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + G1, G2);
                m_oSymbolic.Outputs["BODY"] = body;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of LRParts_FIG295HN.cs."));
                return;
            }
        }

        #endregion
    }
}
