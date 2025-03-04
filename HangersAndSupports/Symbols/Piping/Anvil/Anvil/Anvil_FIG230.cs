//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG230.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG230
//   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari  CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    public class Anvil_FIG230 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG230"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "OPENING", "OPENING", 0.999999)]
        public InputDouble m_dOPENING;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(4, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT", "LEFT")]
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

                Double opening = m_dOPENING.Value;
                Double rodDiameter = m_dROD_DIA.Value;
                const double CONST_1 = 0.0762;
                const double CONST_2 = 0.0381;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "InThdLH", new Position(0, 0, opening - CONST_1), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRodDiaGTZero, "Rod diameter should be greater than zero"));
                    return;
                }
                if (opening < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidOpeningLTZero, "Opening should not be lessthan zero"));
                    return;
                }

                Vector normal = new Position(0, 0, opening - CONST_2).Subtract(new Position(0, 0, opening - CONST_2 + rodDiameter));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, opening - CONST_2 + rodDiameter);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d top = symbolGeometryHelper.CreateCylinder(null, (rodDiameter * 1.5) / (4 * Math.Cos(30 * Math.PI / 180)) + (rodDiameter * 1.5) * Math.Tan((30 * Math.PI / 180)) / 2, normal.Length);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, 0, -CONST_2 - rodDiameter).Subtract(new Position(0, 0, -CONST_2));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -CONST_2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d bottom = symbolGeometryHelper.CreateCylinder(null, (rodDiameter * 1.5) / (4 * Math.Cos(30 * Math.PI / 180)) + (rodDiameter * 1.5) * Math.Tan((30 * Math.PI / 180)) / 2, normal.Length);
                m_Symbolic.Outputs["BOTTOM"] = bottom;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0.6 * rodDiameter, -0.6 * rodDiameter, -(CONST_2 + rodDiameter / 2));
                Projection3d right = symbolGeometryHelper.CreateBox(null, rodDiameter / 2, 1.2 * rodDiameter, opening + rodDiameter, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                right.Transform(matrix);
                m_Symbolic.Outputs["RIGHT"] = right;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.ActivePosition = new Position(0.6 * rodDiameter, -0.6 * rodDiameter, -(CONST_2 + rodDiameter / 2));
                Projection3d left = symbolGeometryHelper.CreateBox(null, rodDiameter / 2, 1.2 * rodDiameter, opening + rodDiameter, 9);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                left.Transform(matrix);
                m_Symbolic.Outputs["LEFT"] = left;

            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG230"));
                return;
            }
        }
        #endregion

    }

}
