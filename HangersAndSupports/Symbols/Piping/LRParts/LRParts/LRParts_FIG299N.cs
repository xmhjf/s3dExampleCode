//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LRParts_FIG299N.cs
//   LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG299N
//   Author       :  Rajeswari
//   Creation Date:  22/10/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  22/10/2012    Rajeswari  Initial Creation
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
    public class LRParts_FIG299N : CustomSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG299N"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;    
        [InputString(2, "CONFIG", "CONFIG", "No Value")]
        public InputString m_CONFIG;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_ROD_DIA;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble m_A;
        [InputDouble(5, "P", "P", 0.999999)]
        public InputDouble m_P;
        [InputDouble(6, "N", "N", 0.999999)]
        public InputDouble m_N;
        [InputDouble(7, "GRIP", "GRIP", 0.999999)]
        public InputDouble m_GRIP;
        [InputDouble(8, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(9, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(10, "W", "W", 0.999999)]
        public InputDouble m_W;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("PIN", "PIN")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT_CYL", "LEFT_CYL")]
        [SymbolOutput("RIGHT_CYL", "RIGHT_CYL")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        /// 
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                Double rodDiameter = m_ROD_DIA.Value;
                Double A = m_A.Value;
                Double pinSize = m_P.Value;
                Double N = m_N.Value;
                Double grip = m_GRIP.Value;
                Double D = m_D.Value;
                Double T = m_T.Value;
                Double W = m_W.Value;
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidT, "T cannot be zero or negative"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidW, "W cannot be zero or negative"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidD, "D cannot be zero or negative"));
                    return;
                }
                if (pinSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidPinSize, "P cannot be zero or negative"));
                    return;
                }
                if (N == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidN, "N cannot be zero."));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                string config = m_CONFIG.Value;

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (config == "Without Pin")
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, rodDiameter / 2 - (A + pinSize / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Port2"] = port2;
                }
                else
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -A + pinSize / 2 + rodDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Port2"] = port2;

                    Vector normal = new Position(0, -(grip / 2 + 2 * T), -A + rodDiameter / 2).Subtract(new Position(0, grip / 2 + 2 * T, -A + rodDiameter / 2));

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, grip / 2 + 2 * T, -A + rodDiameter / 2);
                    symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                    Projection3d projectCyn1 = symbolGeometryHelper.CreateCylinder(null, pinSize / 2, normal.Length);
                    m_Symbolic.Outputs["PIN"] = projectCyn1;
                }
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, rodDiameter / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                Projection3d top = symbolGeometryHelper.CreateCylinder(null, (grip + T) / 2, N);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-W / 2, -grip / 2 - T / 2, -A / 2 + 0.4 * D + rodDiameter / 4);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateBox(null, W, T, A + N / 2 - 0.4 * D);
                m_Symbolic.Outputs["LEFT"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-W / 2, grip / 2 + T / 2, -A / 2 + 0.4 * D + rodDiameter / 4);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateBox(null, W, T, A + N / 2 - 0.4 * D);
                m_Symbolic.Outputs["RIGHT"] = right;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, grip / 2, -A + rodDiameter / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d leftCyl = symbolGeometryHelper.CreateCylinder(null, D / 2, T);
                m_Symbolic.Outputs["LEFT_CYL"] = leftCyl;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -grip / 2 - T, -A + rodDiameter / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d rightCyl = symbolGeometryHelper.CreateCylinder(null, D / 2, T);
                m_Symbolic.Outputs["RIGHT_CYL"] = rightCyl;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of LRParts_FIG299N.cs."));
                return;
            }
        }
        #endregion

        }
}
