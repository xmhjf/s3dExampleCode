//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG271.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG271
//   Author       : Vijaya 
//   Creation Date: 2-May-2013 
//   Description: Initial Creation-CR-CP-222292

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   2-May-2013    Vijaya   CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    public class Anvil_FIG271 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG271"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(5, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(6, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(7, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(8, "M", "M", 0.999999)]
        public InputDouble m_dM;
        [InputDouble(9, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("R_ROLL", "R_ROLL")]
        [SymbolOutput("L_ROLL", "L_ROLL")]
        [SymbolOutput("R_SIDE", "R_SIDE")]
        [SymbolOutput("L_SIDE", "L_SIDE")]
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

                Double A = m_dA.Value, C = m_dC.Value, D = m_dD.Value, E = m_dE.Value, F = m_dF.Value, G = m_dG.Value, M = m_dM.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -A), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidMGTZero, "M value should be greater than zero"));
                    return;
                }
                if (G <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidGGTZero, "G value should be greater than zero"));
                    return;
                }
                if (E <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidEGTZero, "E value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidCGTZero, "C value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Cone3d rightRoll = symbolGeometryHelper.CreateCone(null, G / 3, G / 2, F / 2);
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                rotateMatrix.Translate(new Vector(0, 0, -(C - A)));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                rightRoll.Transform(rotateMatrix);
                m_Symbolic.Outputs["R_ROLL"] = rightRoll;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Cone3d leftRoll = symbolGeometryHelper.CreateCone(null, G / 3, G / 2, F / 2);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Rotate(-Math.PI / 2, new Vector(0, 1, 0));
                rotateMatrix.Translate(new Vector(0, 0, -(C - A)));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                leftRoll.Transform(rotateMatrix);
                m_Symbolic.Outputs["L_ROLL"] = leftRoll;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(F / 2, -0.25 * G, -(M - A) - C);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                Projection3d rightSide = symbolGeometryHelper.CreateBox(null, (D - F) / 2, 0.5 * G, C, 9);
                rightSide.Transform(rotateMatrix);
                m_Symbolic.Outputs["R_SIDE"] = rightSide;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-D / 2, -0.25 * G, -(M - A) - C);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                Projection3d leftSide = symbolGeometryHelper.CreateBox(null, (D - F) / 2, 0.5 * G, C, 9);
                leftSide.Transform(rotateMatrix);
                m_Symbolic.Outputs["L_SIDE"] = leftSide;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-D / 2, -E / 2, A - M);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                Projection3d bottom = symbolGeometryHelper.CreateBox(null, D, E, M, 9);
                bottom.Transform(rotateMatrix);
                m_Symbolic.Outputs["BOTTOM"] = bottom;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG271"));
                    return;
                }
            }
        }
        #endregion

    }

}
