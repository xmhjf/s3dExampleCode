//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LRParts_FIG290LN.cs
//    LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG290LN
//   Author       :  Sasidhar
//   Creation Date:  25/10/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//    25/10/2012   Sasidhar   Initial Creation
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
    public class LRParts_FIG290LN : CustomSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG290LN"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_ROD_DIA;
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
        [InputDouble(8, "G", "G", 0.999999)]
        public InputDouble m_G;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("TOP_CYL_1", "TOP_CYL_1")]
        [SymbolOutput("TOP_CYL_2", "TOP_CYL_2")]
        [SymbolOutput("BOT_CYL_1", "BOT_CYL_1")]
        [SymbolOutput("BOT_CYL_2", "BOT_CYL_2")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
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
                Double A = m_ROD_DIA.Value;
                Double B = m_B.Value;
                Double C = m_C.Value;
                Double D = m_D.Value;
                Double E = m_E.Value;
                Double F = m_F.Value;
                Double G = m_G.Value;

                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidD, "D cannot be zero or negative"));
                    return;
                }
                if (F <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidF, "F cannot be zero or negative"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Port port1 = new Port(OccurrenceConnection, part, "InThdLH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, A / 2 - E), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                symbolGeometryHelper.ActivePosition = new Position(0, 0, A / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Ellipse3d ellipse = (Ellipse3d)symbolGeometryHelper.CreateEllipse(null, C / 2 + D, F / 2, 2 * Math.PI);
                Vector normal = new Vector(0, 0, 1);
                Projection3d ellip = new Projection3d(ellipse, normal, G, true);
                m_Symbolic.Outputs["ELLIP"] = ellip;

                Vector normaltop1 = new Position(0, -(B / 2 + D / 2), A / 2 - 0.65 * E).Subtract(new Position(0, -(C / 2 + D / 2), A / 2));
                Vector xaxis = new Vector(1, 0, 0);
                Vector cylinderNormal = normaltop1.Cross(xaxis);
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + D / 2), A / 2);
                symbolGeometryHelper.SetOrientation(normaltop1, cylinderNormal);
                Projection3d topCyl1 = symbolGeometryHelper.CreateCylinder(null, D / 2, normaltop1.Length);
                m_Symbolic.Outputs["TOP_CYL_1"] = topCyl1;

                Vector normaltop2 = new Position(0, (B / 2 + D / 2), A / 2 - 0.65 * E).Subtract(new Position(0, (C / 2 + D / 2), A / 2));
                Vector cylinderNormal2 = normaltop2.Cross(xaxis);
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (C / 2 + D / 2), A / 2);
                symbolGeometryHelper.SetOrientation(normaltop2, cylinderNormal2);
                Projection3d topCyl2 = symbolGeometryHelper.CreateCylinder(null, D / 2, normaltop2.Length);
                m_Symbolic.Outputs["TOP_CYL_2"] = topCyl2;

                Vector normalbottom1 = new Position(0, -0.3 * B, A / 2 - (E + D / 2)).Subtract(new Position(0, -(B / 2 + D / 2), A / 2 - 0.65 * E));
                Vector crossproductbottom1 = normalbottom1.Cross(xaxis);
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(B / 2 + D / 2), A / 2 - 0.65 * E);
                symbolGeometryHelper.SetOrientation(normalbottom1, crossproductbottom1);
                Projection3d botCyl1 = symbolGeometryHelper.CreateCylinder(null, D / 2, normalbottom1.Length);
                m_Symbolic.Outputs["BOT_CYL_1"] = botCyl1;

                Vector normalbottom2 = new Position(0, 0.3 * B, A / 2 - (E + D / 2)).Subtract(new Position(0, (B / 2 + D / 2), A / 2 - 0.65 * E));
                Vector crossproductbottom2 = normalbottom2.Cross(xaxis);
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (B / 2 + D / 2), A / 2 - 0.65 * E);
                symbolGeometryHelper.SetOrientation(normalbottom2, crossproductbottom2);
                Projection3d botCyl2 = symbolGeometryHelper.CreateCylinder(null, D / 2, normalbottom2.Length);
                m_Symbolic.Outputs["BOT_CYL_2"] = botCyl2;

                Vector normalextremebottom = new Position(0, 0.3 * B, A / 2 - (E + D / 2)).Subtract(new Position(0, -0.3 * B, A / 2 - (E + D / 2)));
                Vector crossproductexbottom = normalextremebottom.Cross(xaxis);
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -0.3 * B, A / 2 - (E + D / 2));
                symbolGeometryHelper.SetOrientation(normalextremebottom, crossproductexbottom);
                Projection3d bottom = symbolGeometryHelper.CreateCylinder(null, D / 2, normalextremebottom.Length);
                m_Symbolic.Outputs["BOTTOM"] = bottom;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of LRParts_FIG290LN.cs."));
                return;
            }
        }

        #endregion
    }
}
