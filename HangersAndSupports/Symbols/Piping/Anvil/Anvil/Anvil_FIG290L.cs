//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG290L.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG290L
//   Author       : Vijaya 
//   Creation Date: 3-May-2013  
//   Description: Initial Creation-CR-CP-222292

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   3-May-2013     Vijaya   CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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

    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class Anvil_FIG290L : HangerComponentSymbolDefinition
    {

        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG290L"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(4, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(6, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(8, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(9, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();

                Double A = m_dROD_DIA.Value, B = m_dB.Value, C = m_dC.Value, D = m_dD.Value, E = m_dE.Value, F = m_dF.Value, G = m_dG.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "InThdLH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, A / 2 - E), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (G == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidGNZero, "G value cannot be zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(0, 0, A / 2);
                Ellipse3d curve = (Ellipse3d)symbolGeometryHelper.CreateEllipse(null, F / 2, C / 2 + D, 2 * Math.PI);
                Projection3d top = new Projection3d(curve, new Vector(0, 0, 1), G, true);
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, -(B / 2 + D / 2), A / 2 - 0.65 * E).Subtract(new Position(0, -(C / 2 + D / 2), A / 2));
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + D / 2), A / 2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topCylinder1 = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_CYL_1"] = topCylinder1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, (B / 2 + D / 2), A / 2 - 0.65 * E).Subtract(new Position(0, (C / 2 + D / 2), A / 2));
                symbolGeometryHelper.ActivePosition = new Position(0, (C / 2 + D / 2), A / 2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topCylinder2 = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_CYL_2"] = topCylinder2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, -0.3 * B, A / 2 - (E + D / 2)).Subtract(new Position(0, -(B / 2 + D / 2), A / 2 - 0.65 * E));
                symbolGeometryHelper.ActivePosition = new Position(0, -(B / 2 + D / 2), A / 2 - 0.65 * E);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d bottomCylinder1 = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal.Length);
                m_Symbolic.Outputs["BOT_CYL_1"] = bottomCylinder1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, 0.3 * B, A / 2 - (E + D / 2)).Subtract(new Position(0, (B / 2 + D / 2), A / 2 - 0.65 * E));
                symbolGeometryHelper.ActivePosition = new Position(0, (B / 2 + D / 2), A / 2 - 0.65 * E);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d bottomCylinder2 = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal.Length);
                m_Symbolic.Outputs["BOT_CYL_2"] = bottomCylinder2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, 0.3 * B, A / 2 - (E + D / 2)).Subtract(new Position(0, -0.3 * B, A / 2 - (E + D / 2)));
                symbolGeometryHelper.ActivePosition = new Position(0, -0.3 * B, A / 2 - (E + D / 2));
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d bottom = symbolGeometryHelper.CreateCylinder(null, D / 2.0, normal.Length);
                m_Symbolic.Outputs["BOTTOM"] = bottom;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG290L"));
                    return;
                }
            }
        }
        #endregion
    }

}
