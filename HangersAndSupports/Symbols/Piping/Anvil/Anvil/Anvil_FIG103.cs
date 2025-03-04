//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG103.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG103
//   Author       :  Vijay
//   Creation Date:  30-04-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-04-2013     Vijay    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net   
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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
    public class Anvil_FIG103 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG103"
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
        [InputDouble(5, "C1", "C1", 0.999999)]
        public InputDouble m_dC1;
        [InputDouble(6, "C2", "C2", 0.999999)]
        public InputDouble m_dC2;
        [InputDouble(7, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("R1", "R1")]
        [SymbolOutput("R2", "R2")]
        [SymbolOutput("R3", "R3")]
        [SymbolOutput("L1", "L1")]
        [SymbolOutput("L2", "L2")]
        [SymbolOutput("L3", "L3")]
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
                Double C1 = m_dC1.Value;
                Double C2 = m_dC2.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -B - C1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (pipeDiameter == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidPipeDiaNZero, "Pipe diameter cannot be zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBGTZero, "B value should be greater than zero"));
                    return;
                }
                if (C1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidC1GTZero, "C1 value should be greater than zero"));
                    return;
                }
                if (C2 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidC2GTZero, "C2 value should be greater than zero"));
                    return;
                }

                Matrix4X4 matrix = new Matrix4X4();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Translate(new Vector(0, -C2 / 2, 0));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d bodyCylinder = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + C1, C2);
                bodyCylinder.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = bodyCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(pipeDiameter / 2, -C2 / 2, -C1));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d r1Box = symbolGeometryHelper.CreateBox(null, C1 + C2, C2, 2 * C1, 9);
                r1Box.Transform(matrix);
                m_Symbolic.Outputs["R1"] = r1Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(pipeDiameter / 2 + C1 + C2, -C2 / 2, 0));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d r2Box = symbolGeometryHelper.CreateBox(null, C1, C2, B, 9);
                r2Box.Transform(matrix);
                m_Symbolic.Outputs["R2"] = r2Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector((pipeDiameter / 2 + C1 + C2), -C2 / 2, B));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d r3Box = symbolGeometryHelper.CreateBox(null, (A + C2) / 2 - (pipeDiameter / 2 + C1 + C2), C2, C1, 9);
                r3Box.Transform(matrix);
                m_Symbolic.Outputs["R3"] = r3Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-(pipeDiameter / 2 + C1 + C2), -C2 / 2, -C1));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d l1Box = symbolGeometryHelper.CreateBox(null, C1 + C2, C2, 2 * C1, 9);
                l1Box.Transform(matrix);
                m_Symbolic.Outputs["L1"] = l1Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-pipeDiameter / 2 - 2 * C1 - C2, -C2 / 2, 0));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d l2Box = symbolGeometryHelper.CreateBox(null, C1, C2, B, 9);
                l2Box.Transform(matrix);
                m_Symbolic.Outputs["L2"] = l2Box;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-(A + C2) / 2, -C2 / 2, B));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d l3Box = symbolGeometryHelper.CreateBox(null, (A + C2) / 2 - (pipeDiameter / 2 + C1 + C2), C2, C1, 9);
                l3Box.Transform(matrix);
                m_Symbolic.Outputs["L3"] = l3Box;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG103"));
                    return;
                }
            }
        }
        #endregion
    }
}
