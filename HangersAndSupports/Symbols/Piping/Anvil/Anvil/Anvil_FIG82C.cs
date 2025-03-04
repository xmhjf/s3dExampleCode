//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG82C.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG82C
//   Author       :  Hema
//   Creation Date:  14-05-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project
//   
//   Anvil_FIG82C.cs is same for Anvil_FIG98C.cs,Anvil_FIG268C.cs

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

    public class Anvil_FIG82C : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG82C"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(6, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(7, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(8, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(9, "Z", "Z", 0.999999)]
        public InputDouble m_dZ;
        [InputDouble(10, "J", "J", 0.999999)]
        public InputDouble m_dJ;
        [InputDouble(11, "HOLE_SIZE", "HOLE_SIZE", 0.999999)]
        public InputDouble m_dHOLE_SIZE;
        [InputDouble(12, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(13, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(14, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputDouble(15, "DIR", "DIR", 1)]
        public InputDouble m_oDIR;
        [InputDouble(16, "WORKING_TRAV", "WORKING_TRAV", 0.999999)]
        public InputDouble m_dWORKING_TRAV;
        [InputDouble(17, "HOT_LOAD", "HOT_LOAD", 0.999999)]
        public InputDouble m_dHOT_LOAD;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("FLANGE", "FLANGE")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
        [SymbolOutput("R_LUG", "R_LUG")]
        [SymbolOutput("L_LUG", "L_LUG")]
        [SymbolOutput("L_HOLE", "L_HOLE")]
        [SymbolOutput("R_HOLE", "R_HOLE")]
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

                Double T = m_dT.Value;
                Double A = m_dA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double F = m_dF.Value;
                Double D = m_dD.Value;
                Double G = m_dG.Value;
                Double Z = m_dZ.Value;
                Double J = m_dJ.Value;
                Double holeSize = m_dHOLE_SIZE.Value;
                Double S = m_dS.Value;
                Double H = m_dH.Value;
                Double R = m_dR.Value;
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

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, J + holeSize / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;


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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRodDiaGTZero, "Rod diameter  should be greater than zero"));
                    return;
                }
                if (B == 0 && G == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBvalueAndGGTZero, "B and G values cannot be zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (holeSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHoleSizeGTZero, "HOLE_SIZE should be greater than zero"));
                    return;
                }
                if (Z == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZNZero, "Z value cannot be zero"));
                    return;
                }
                if (H == 0 && R == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHvalueAndRNZero, "H and R values cannot be zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(0, 0, J - H - B);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, C / 2, B - G);
                matrix.Rotate(3 * Math.PI, new Vector(0, 0, 1));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper.ActivePosition = new Position(0, 0, J - H - G);
                Projection3d flange = symbolGeometryHelper.CreateCylinder(null, D / 2, G);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI, new Vector(0, 0, 1));
                flange.Transform(matrix);
                m_Symbolic.Outputs["FLANGE"] = flange;

                symbolGeometryHelper.ActivePosition = new Position(0, 0, -F);
                Projection3d bottom = symbolGeometryHelper.CreateCylinder(null, 0.7 * rodDiameter, Z + (B * 0.75));
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI, new Vector(0, 0, 1));
                bottom.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bottom;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(T + S / 2), -holeSize, J - H);
                Projection3d rLug = symbolGeometryHelper.CreateBox(null, T, 2 * holeSize, H + R, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rLug.Transform(matrix);
                m_Symbolic.Outputs["R_LUG"] = rLug;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(S / 2, -holeSize, J - H);
                Projection3d lLug = symbolGeometryHelper.CreateBox(null, T, 2 * holeSize, H + R, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                lLug.Transform(matrix);
                m_Symbolic.Outputs["L_LUG"] = lLug;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, -S / 2, J).Subtract(new Position(0, -(S / 2 + T), J));
                symbolGeometryHelper.ActivePosition = new Position(0, -(S / 2 + T), J);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d lHole = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, normal.Length);
                m_Symbolic.Outputs["L_HOLE"] = lHole;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(0, S / 2, J).Subtract(new Position(0, (S / 2 + T), J));
                symbolGeometryHelper.ActivePosition = new Position(0, (S / 2 + T), J);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d rHole = symbolGeometryHelper.CreateCylinder(null, holeSize / 2, normal1.Length);
                m_Symbolic.Outputs["R_HOLE"] = rHole;

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG82C"));
                    return;
                }
            }
        }
        #endregion

    }

}
