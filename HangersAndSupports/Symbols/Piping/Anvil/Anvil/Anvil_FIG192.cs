//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG192.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG192
//   Author       :  Rajeswari
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
    public class Anvil_FIG192 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG192"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(4, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(6, "C1", "C1", 0.999999)]
        public InputDouble m_dC1;
        [InputDouble(7, "C2", "C2", 0.999999)]
        public InputDouble m_dC2;
        [InputDouble(8, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("ROD", "ROD")]
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
                Double D = m_dD.Value;
                Double rodDiameter = m_dROD_DIA.Value;
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
                Port port2 = new Port(OccurrenceConnection, part, "ExThdRH", new Position(0, 0, -D), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRodDiaGTZero, "Rod diameter should be greater than zero"));
                    return;
                }

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + C1, 120 * (Math.PI / 180));
                matrix.Rotate(210 * (Math.PI / 180), new Vector(0, 0, 1));
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -C2 / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                arc.Transform(matrix);
                curveCollection.Add(arc);

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, -0), C2, false);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, -D).Subtract(new Position(0, 0, -D + B));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -D + B);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rod = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2, normal.Length);
                m_Symbolic.Outputs["ROD"] = rod;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG192"));
                return;
            }
        }
        #endregion

    }

}
