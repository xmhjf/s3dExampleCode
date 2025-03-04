//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG191.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG191
//  //   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
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
    public class Anvil_FIG191 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG191"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(4, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(7, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(8, "F1", "F1", 0.999999)]
        public InputDouble m_dF1;
        [InputDouble(9, "F2", "F2", 0.999999)]
        public InputDouble m_dF2;
        [InputDouble(10, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("ROD", "ROD")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("LEFT_ROD", "LEFT_ROD")]
        [SymbolOutput("RIGHT_ROD", "RIGHT_ROD")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
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
                Double E = m_dE.Value;
                Double rodDiameter = m_dROD_DIA.Value;
                Double B = m_dB.Value;
                Double C = m_dC.Value;
                Double D = m_dD.Value;
                Double F1 = m_dF1.Value;
                Double F2 = m_dF2.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "ExThdRH", new Position(0, 0, -E), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
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
                if (F1 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidF1GTZero, "F1 value should be greater than zero"));
                    return;
                }
                if (F2 <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidF2GTZero, "F2 value should be greater than zero"));
                    return;
                }

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + F1, 126 * (Math.PI / 180));
                matrix.Rotate(207 * (Math.PI / 180), new Vector(0, 0, 1));
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -F2 / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, -0), F2, false);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, -E).Subtract(new Position(0, 0, -E + B));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -E + B);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rod = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2, normal.Length);
                m_Symbolic.Outputs["ROD"] = rod;

                Revolution3d bend = new Revolution3d((new Circle3d(new Position(C / 2, 0, 0), new Vector(0, 0, 1), D / 2)), new Vector(0, -1, 0), new Position(0, 0, 0), Math.PI * 180 / 180, true);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, -(pipeDiameter / 2 - C / 2 + D / 2)));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                bend.Transform(matrix);
                m_Symbolic.Outputs["BEND"] = bend;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, -C / 2, -pipeDiameter / 2).Subtract(new Position(0, -C / 2, -(pipeDiameter / 2 - C / 2 + D / 2)));
                symbolGeometryHelper.ActivePosition = new Position(0, -C / 2, -(pipeDiameter / 2 - C / 2 + D / 2));
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d leftRod = symbolGeometryHelper.CreateCylinder(null, D / 2, normal.Length);
                m_Symbolic.Outputs["LEFT_ROD"] = leftRod;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, C / 2, -pipeDiameter / 2).Subtract(new Position(0, C / 2, -(pipeDiameter / 2 - C / 2 + D / 2)));
                symbolGeometryHelper.ActivePosition = new Position(0, C / 2, -(pipeDiameter / 2 - C / 2 + D / 2));
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rightRod = symbolGeometryHelper.CreateCylinder(null, D / 2, normal.Length);
                m_Symbolic.Outputs["RIGHT_ROD"] = rightRod;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.ActivePosition = new Position(-rodDiameter * 2 - Math.Cos(30 * Math.PI / 180) * pipeDiameter / 2.0, -F2 / 2, (Math.Sin(30 * Math.PI / 180) * pipeDiameter / 2) - F1);
                Projection3d left = symbolGeometryHelper.CreateBox(null, rodDiameter * 2, F2, F1, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                left.Transform(matrix);
                m_Symbolic.Outputs["LEFT"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.ActivePosition = new Position(Math.Cos(30 * Math.PI / 180) * pipeDiameter / 2.0, -F2 / 2, (Math.Sin(30 * Math.PI / 180) * pipeDiameter / 2) - F1);
                Projection3d right = symbolGeometryHelper.CreateBox(null, rodDiameter * 2, F2, F1, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                right.Transform(matrix);
                m_Symbolic.Outputs["RIGHT"] = right;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG191"));
                return;
            }
        }
        #endregion
    }
}
