//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG69.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG69
//   Author       :  Manikanth
//   Creation Date:  16-05-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-05-2013    Manikanth CR-CP-233113-Convert HS_Anvil VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
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
    public class Anvil_FIG69 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG69"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_PIPE_DIA;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble m_C;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_B;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble m_A;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("NUT", "NUT")]
        public AspectDefinition m_Symbolic;


        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                double pipeDiameter = m_PIPE_DIA.Value;
                double C = m_C.Value;
                double B = m_B.Value;
                double rodDiameter = m_A.Value;
                double angle = 16;

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, C), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (pipeDiameter < 0.05)
                    angle = 0;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZerorodDiavalue, "Rod diameter should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroBvalue, "B value should be greater than zero"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + (pipeDiameter / 50), (Math.PI + 2 * angle * (Math.PI / 180)));
                matrix.Rotate((Math.PI - angle * (Math.PI / 180)), new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Translate(new Vector(0, -rodDiameter, 0));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                arc1.Transform(matrix);
                curveCollection.Add(arc1);

                curveCollection.Add(new Line3d(new Position(-rodDiameter, pipeDiameter / 2 - (C - pipeDiameter / 2 * Math.Sin((angle * (Math.PI / 180)))) * Math.Tan(((angle * (Math.PI / 180)))), C), new Position(-rodDiameter, (pipeDiameter / 2 + (pipeDiameter / 50)) * Math.Cos((angle * (Math.PI / 180))), (pipeDiameter / 2 + (pipeDiameter / 50)) * Math.Sin((angle * (Math.PI / 180))))));
                curveCollection.Add(new Line3d(new Position(-rodDiameter, pipeDiameter / 2 - (C - pipeDiameter / 2 * Math.Sin((angle * (Math.PI / 180)))) * Math.Tan(((angle * (Math.PI / 180)))), C), new Position(-rodDiameter, -(pipeDiameter / 2 - (C - pipeDiameter / 2 * Math.Sin((angle * (Math.PI / 180)))) * Math.Tan((angle * (Math.PI / 180)))), C)));
                curveCollection.Add(new Line3d(new Position(-rodDiameter, (-pipeDiameter / 2 - (pipeDiameter / 50)) * Math.Cos((angle * (Math.PI / 180))), (pipeDiameter / 2 + (pipeDiameter / 50)) * Math.Sin((angle * (Math.PI / 180)))), new Position(-rodDiameter, -(pipeDiameter / 2 - (C - pipeDiameter / 2 * Math.Sin((angle * (Math.PI / 180)))) * Math.Tan((angle * (Math.PI / 180)))), C)));

                Projection3d pro = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), 2 * rodDiameter, false);
                m_Symbolic.Outputs["BODY"] = pro;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                matrix.Translate(new Vector(0, 0, C));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d nutCylinder = symbolGeometryHelper.CreateCylinder(null, 0.75 * rodDiameter, B - C);
                nutCylinder.Transform(matrix);
                m_Symbolic.Outputs["NUT"] = nutCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG69"));
                    return;
                }
            }
        }
        #endregion

    }

}
