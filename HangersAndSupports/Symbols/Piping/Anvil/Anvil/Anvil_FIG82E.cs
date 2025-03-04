//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG82E.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG82E
//   Author       :  Hema
//   Creation Date:  14-05-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project
//   
//   Anvil_FIG82E.cs is same for Anvil_FIG98E.cs,Anvil_FIG268E.cs

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   14-05-2013     Hema     CR-CP-233113 Convert HS_Anvil VB Project to C# .Net
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

    public class Anvil_FIG82E : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG82E"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_L", "ROD_L", 0.999999)]
        public InputDouble m_dROD_L;
        [InputDouble(3, "Q", "Q", 0.999999)]
        public InputDouble m_dQ;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(8, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(9, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(10, "DIR", "DIR", 1)]
        public InputDouble m_oDIR;
        [InputDouble(11, "WORKING_TRAV", "WORKING_TRAV", 0.999999)]
        public InputDouble m_dWORKING_TRAV;
        [InputDouble(12, "HOT_LOAD", "HOT_LOAD", 0.999999)]
        public InputDouble m_dHOT_LOAD;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("FLANGE", "FLANGE")]
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

                Double rodL = m_dROD_L.Value;
                Double Q = m_dQ.Value;
                Double A = m_dA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double F = m_dF.Value;
                Double D = m_dD.Value;
                Double G = m_dG.Value;
                Double workingTrav = m_dWORKING_TRAV.Value;
                Double hotLoad = m_dHOT_LOAD.Value;
                Double rodDiameter = m_dA.Value;

                //Initializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "BotInThdRH", new Position(0, 0, -rodL - Q), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "TopInThdRH", new Position(0, 0, B), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;
                Port port4 = new Port(OccurrenceConnection, part, "MidInThdRH", new Position(0, 0, -rodL), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port4"] = port4;


                //Validating Inputs
                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidCGTZero, "C value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRodDiaGTZero, "Rod diameter should be greater than zero"));
                    return;
                }
                if (B == 0 && G == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBvalueAndGGTZero, "B and G values cannot be zero"));
                    return;
                }
                if (Q == 0 && F == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidQvalueAndFGTZero, "Q and F values cannot be zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, C / 2, B - G);
                matrix.Rotate(3 * Math.PI, new Vector(0, 0, 1));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper.ActivePosition = new Position(0, 0, B - G);
                Projection3d flange = symbolGeometryHelper.CreateCylinder(null, D / 2, G);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI, new Vector(0, 0, 1));
                flange.Transform(matrix);
                m_Symbolic.Outputs["FLANGE"] = flange;

                symbolGeometryHelper.ActivePosition = new Position(0, 0, -rodL - F - Q);
                Projection3d bottom = symbolGeometryHelper.CreateCylinder(null, 0.7 * rodDiameter, Q + F * 2);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI, new Vector(0, 0, 1));
                bottom.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bottom;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG82E"));
                    return;
                }
            }
        }
        #endregion
    }
}
