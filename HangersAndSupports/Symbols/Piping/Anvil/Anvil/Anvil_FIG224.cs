//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG224.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG224
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
    public class Anvil_FIG224 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG224"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(4, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(5, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(6, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(7, "M", "M", 0.999999)]
        public InputDouble m_dM;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
        [SymbolOutput("BODY", "BODY")]
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
                Double C = m_dC.Value;
                Double E = m_dE.Value;
                Double F = m_dF.Value;
                Double H = m_dH.Value;
                Double M = m_dM.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, (E - F / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHGTZero, "H value should be greater than zero"));
                    return;
                }
                if (pipeDiameter < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidPipeDiaLTZero, "Pipe diameter should not be less than zero"));
                    return;
                }
                if (F <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFGTZero, "F value should be greater than zero"));
                    return;
                }
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidMGTZero, "M value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(-(C / 2 + H), -M / 2, pipeDiameter / 2 + H / 2);
                Projection3d top1 = symbolGeometryHelper.CreateBox(null, H, M, E - pipeDiameter / 2 + M / 2 - H / 2, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                top1.Transform(matrix);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.ActivePosition = new Position(C / 2, -M / 2, pipeDiameter / 2 + H / 2);
                Projection3d top2 = symbolGeometryHelper.CreateBox(null, H, M, E - pipeDiameter / 2 + M / 2 - H / 2, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                top2.Transform(matrix);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, C / 2 + 2 * H, E).Subtract(new Position(0, -(C / 2 + 2 * H), E));
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + 2 * H), E);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, F / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + H, H);
                matrix = new Matrix4X4();
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -H / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG224"));
                return;
            }
        }
        #endregion

    }

}
