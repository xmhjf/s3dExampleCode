//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG40.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG40
//   Author       :  Hema
//   Creation Date:  03-05-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   03-05-2013     Hema     CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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

    public class Anvil_FIG40 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG40"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(5, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(6, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(7, "G1", "G1", 0.999999)]
        public InputDouble m_dG1;
        [InputDouble(8, "G2", "G2", 0.999999)]
        public InputDouble m_dG2;
        [InputDouble(9, "S", "S", 0.999999)]
        public InputDouble m_dS;
        [InputDouble(10, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LEFT_FRONT", "LEFT_FRONT")]
        [SymbolOutput("LEFT_BACK", "LEFT_BACK")]
        [SymbolOutput("RIGHT_FRONT", "RIGHT_FRONT")]
        [SymbolOutput("RIGHT_BACK", "RIGHT_BACK")]
        [SymbolOutput("LEFT_OUTER", "LEFT_OUTER")]
        [SymbolOutput("LEFT_INNER", "LEFT_INNER")]
        [SymbolOutput("RIGHT_OUTER", "RIGHT_OUTER")]
        [SymbolOutput("RIGHT_INNER", "RIGHT_INNER")]
        [SymbolOutput("LINE", "LINE")]
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

                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double A = m_dA.Value;
                Double B = m_dB.Value;
                Double E = m_dE.Value;
                Double F = m_dF.Value;
                Double G1 = m_dG1.Value;
                Double G2 = m_dG2.Value;
                Double S = m_dS.Value;

                //Initializing SymbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "LeftPin", new Position(0, E, G2 / 2 - A - F / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "RightPin", new Position(0, -E, G2 / 2 - A - F / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                //Validating Inputs
                if (G1 == 0 && pipeDiameter==0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidG1AndPipeDiaNZero, "G1  and PipeDia values cannot be zero"));
                    return;
                }
                if (G2 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidG2GTZero, "G2 value should be greater than zero"));
                    return;
                }
                if (F <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFGTZero, "F value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(0, 0, -G2 / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + G1, G2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(E + B), -S / 2 - G1, -G2 / 2);
                Projection3d leftFront = symbolGeometryHelper.CreateBox(null, E + B - pipeDiameter / 2, G1, G2, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                leftFront.Transform(matrix);
                m_Symbolic.Outputs["LEFT_FRONT"] = leftFront;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(E + B), S / 2, -G2 / 2);
                Projection3d leftBack = symbolGeometryHelper.CreateBox(null, E + B - pipeDiameter / 2, G1, G2, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                leftBack.Transform(matrix);
                m_Symbolic.Outputs["LEFT_BACK"] = leftBack;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2, -S / 2 - G1, -G2 / 2);
                Projection3d rightfront = symbolGeometryHelper.CreateBox(null, E + B - pipeDiameter / 2, G1, G2, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rightfront.Transform(matrix);
                m_Symbolic.Outputs["RIGHT_FRONT"] = rightfront;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(pipeDiameter / 2, S / 2, -G2 / 2);
                Projection3d rightBack = symbolGeometryHelper.CreateBox(null, E + B - pipeDiameter / 2, G1, G2, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rightBack.Transform(matrix);
                m_Symbolic.Outputs["RIGHT_BACK"] = rightBack;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position((S / 2 + 2 * G1), -E, G2 / 2 - A).Subtract(new Position(-(S / 2 + 2 * G1), -E, G2 / 2 - A));
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2 + 2 * G1), -E, G2 / 2 - A);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d leftOuter = symbolGeometryHelper.CreateCylinder(null, F / 2, normal.Length);
                m_Symbolic.Outputs["LEFT_OUTER"] = leftOuter;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position((S / 2 + 2 * G1), -(pipeDiameter / 2 + G1 + 1.5 * F), 0).Subtract(new Position(-(S / 2 + 2 * G1), -(pipeDiameter / 2 + G1 + 1.5 * F), 0));
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2 + 2 * G1), -(pipeDiameter / 2 + G1 + 1.5 * F), 0);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d leftInner = symbolGeometryHelper.CreateCylinder(null, F / 2, normal1.Length);
                m_Symbolic.Outputs["LEFT_INNER"] = leftInner;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal2 = new Position((S / 2 + 2 * G1), E, G2 / 2 - A).Subtract(new Position(-(S / 2 + 2 * G1), E, G2 / 2 - A));
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2 + 2 * G1), E, G2 / 2 - A);
                symbolGeometryHelper.SetOrientation(normal2, normal2.GetOrthogonalVector());
                Projection3d rightOuter = symbolGeometryHelper.CreateCylinder(null, F / 2, normal2.Length);
                m_Symbolic.Outputs["RIGHT_OUTER"] = rightOuter;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal3 = new Position((S / 2 + 2 * G1), (pipeDiameter / 2 + G1 + 1.5 * F), 0).Subtract(new Position(-(S / 2 + 2 * G1), (pipeDiameter / 2 + G1 + 1.5 * F), 0));
                symbolGeometryHelper.ActivePosition = new Position(-(S / 2 + 2 * G1), (pipeDiameter / 2 + G1 + 1.5 * F), 0);
                symbolGeometryHelper.SetOrientation(normal3, normal3.GetOrthogonalVector());
                Projection3d rightInner = symbolGeometryHelper.CreateCylinder(null, F / 2, normal3.Length);
                m_Symbolic.Outputs["RIGHT_INNER"] = rightInner;

                Line3d line = new Line3d(new Position(0, 0, 0), new Position(0, 0, G2));
                m_Symbolic.Outputs["LINE"] = line;

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG40"));
                    return;
                }
            }
        }
        #endregion
    }
}
