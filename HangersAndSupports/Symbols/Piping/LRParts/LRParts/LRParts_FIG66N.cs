//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LRParts_FIG66N.cs
//    LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG66N
//   Author       :  Rajeswari
//   Creation Date:  29/10/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  29/10/2012    Rajeswari  Initial Creation
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
    public class LRParts_FIG66N : CustomSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG66N"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "PLACE_MODE_USER", "PLACE_MODE_USER", "No Value")]
        public InputString m_PLACE_MODE_USER;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_ROD_DIA;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble m_A;
        [InputDouble(5, "E_PRIME", "E_PRIME", 0.999999)]
        public InputDouble m_E_PRIME;
        [InputDouble(6, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(7, "R", "R", 0.999999)]
        public InputDouble m_R;
        [InputDouble(8, "S", "S", 0.999999)]
        public InputDouble m_S;
        [InputDouble(9, "B", "B", 0.999999)]
        public InputDouble m_B;
        [InputDouble(10, "T", "T", 0.999999)]
        public InputDouble m_T;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("SIDE1", "SIDE1")]
        [SymbolOutput("SIDE2", "SIDE2")]
        [SymbolOutput("PIN", "PIN")]
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
                Double rodDiameter = m_ROD_DIA.Value;
                Double A = m_A.Value;
                Double ePrime = m_E_PRIME.Value;
                Double H = m_H.Value;
                Double R = m_R.Value;
                Double S = m_S.Value;
                Double B = m_B.Value;
                Double T = m_T.Value;

                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidT, "T cannot be zero or negative"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidB, "B cannot be zero or negative"));
                    return;
                }
                if (S <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidS, "S cannot be zero or negative"));
                    return;
                }
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidH, "H cannot be zero or negative"));
                    return;
                }
                if (ePrime <= 0 && R <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidEprime, "E_PRIME and R cannot be zero or negative"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;

                string placeMode = m_PLACE_MODE_USER.Value.ToString();

                if (placeMode == "Connect with Hanger Rod")
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, -(R + ePrime));
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d top = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, S + 2 * T);
                    m_Symbolic.Outputs["TOP"] = top;

                    Port port2 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, -A), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Port2"] = port2;
                }
                else
                {
                    Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -(ePrime - H / 2)), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port2"] = port2;

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, -T);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d top = (Projection3d)symbolGeometryHelper.CreateBox(null, T, B, S + 2 * T);
                    m_Symbolic.Outputs["TOP"] = top;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -(S / 2 + T + H / 2), -ePrime);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                    Projection3d pin = symbolGeometryHelper.CreateCylinder(null, H / 2, S + 2 * T + H);
                    m_Symbolic.Outputs["PIN"] = pin;
                }
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, S / 2 + T / 2, -(ePrime + R));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d side1 = (Projection3d)symbolGeometryHelper.CreateBox(null, (ePrime + R), B, T);
                m_Symbolic.Outputs["SIDE1"] = side1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(S / 2 + T / 2), -(ePrime + R));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d side2 = (Projection3d)symbolGeometryHelper.CreateBox(null, (ePrime + R), B, T);
                m_Symbolic.Outputs["SIDE2"] = side2;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of LRParts_FIG66N.cs."));
                return;
            }
        }

        #endregion
    }
}
