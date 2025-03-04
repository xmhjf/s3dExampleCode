//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG261.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG261
//   Author       : Vijaya 
//   Creation Date: 2-May-2013  
//   Description: Initial Creation-CR-CP-222292

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   2-May-2013     Vijaya   CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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

    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class Anvil_FIG261 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG261"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "G1", "G1", 0.999999)]
        public InputDouble m_dG1;
        [InputDouble(5, "G2", "G2", 0.999999)]
        public InputDouble m_dG2;
        [InputDouble(6, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("CLAMP1", "CLAMP1")]
        [SymbolOutput("CLAMP2", "CLAMP2")]
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
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                Double pipeDiameter = m_dPIPE_DIA.Value, A = m_dA.Value, G1 = m_dG1.Value, G2 = m_dG2.Value;
                double angle1 = Math.Asin((G1 * 1.5) / (pipeDiameter / 2)), angle2 = Math.Asin((G1 * 2.5) / (pipeDiameter / 2 + G1)),
                x1 = Math.Sqrt(((pipeDiameter / 2) * (pipeDiameter / 2) - (G1 * 1.5) * (G1 * 1.5))), x2 = Math.Sqrt(((pipeDiameter / 2 + G1) * (pipeDiameter / 2 + G1) - (G1 * 2.5) * (G1 * 2.5)));

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                //Validating Inputs
                if (G2 == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidG2NZero, "G2 value cannot be zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + G1, Math.PI - 2 * angle2);
                rotateMatrix.Rotate(Math.PI + angle2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -G2 / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc1.Transform(rotateMatrix);
                curveCollection.Add(arc1);

                curveCollection.Add(new Line3d(new Position(-G2 / 2, A / 2, -2.5 * G1), new Position(-G2 / 2, x2, -2.5 * G1)));
                curveCollection.Add(new Line3d(new Position(-G2 / 2, A / 2, -1.5 * G1), new Position(-G2 / 2, A / 2, -2.5 * G1)));
                curveCollection.Add(new Line3d(new Position(-G2 / 2, x1, -1.5 * G1), new Position(-G2 / 2, A / 2, -1.5 * G1)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc2 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, Math.PI - 2 * angle1);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(Math.PI + angle1, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -G2 / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc2.Transform(rotateMatrix);
                curveCollection.Add(arc2);

                curveCollection.Add(new Line3d(new Position(-G2 / 2, -x1, -1.5 * G1), new Position(-G2 / 2, -A / 2, -1.5 * G1)));
                curveCollection.Add(new Line3d(new Position(-G2 / 2, -A / 2, -1.5 * G1), new Position(-G2 / 2, -A / 2, -2.5 * G1)));
                curveCollection.Add(new Line3d(new Position(-G2 / 2, -A / 2, -2.5 * G1), new Position(-G2 / 2, -x2, -2.5 * G1)));

                Projection3d clamp1 = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), G2, true);
                m_Symbolic.Outputs["CLAMP1"] = clamp1;

                //clamp2                
                curveCollection = new Collection<ICurve>();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc3 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + G1, Math.PI - 2 * angle2);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(angle2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -G2 / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc3.Transform(rotateMatrix);
                curveCollection.Add(arc3);

                curveCollection.Add(new Line3d(new Position(-G2 / 2, -A / 2, 2.5 * G1), new Position(-G2 / 2, -x2, 2.5 * G1)));
                curveCollection.Add(new Line3d(new Position(-G2 / 2, -A / 2, 1.5 * G1), new Position(-G2 / 2, -A / 2, 2.5 * G1)));
                curveCollection.Add(new Line3d(new Position(-G2 / 2, -x1, 1.5 * G1), new Position(-G2 / 2, -A / 2, 1.5 * G1)));


                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc4 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, Math.PI - 2 * angle1);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(angle1, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -G2 / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc4.Transform(rotateMatrix);
                curveCollection.Add(arc4);

                curveCollection.Add(new Line3d(new Position(-G2 / 2, x1, 1.5 * G1), new Position(-G2 / 2, A / 2, 1.5 * G1)));
                curveCollection.Add(new Line3d(new Position(-G2 / 2, A / 2, 1.5 * G1), new Position(-G2 / 2, A / 2, 2.5 * G1)));
                curveCollection.Add(new Line3d(new Position(-G2 / 2, A / 2, 2.5 * G1), new Position(-G2 / 2, x2, 2.5 * G1)));

                Projection3d clamp2 = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), G2, true);
                m_Symbolic.Outputs["CLAMP2"] = clamp2;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG261"));
                    return;
                }
            }
        }
        #endregion

    }

}
