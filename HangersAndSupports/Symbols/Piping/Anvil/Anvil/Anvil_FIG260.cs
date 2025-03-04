//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG260.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG260
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
    public class Anvil_FIG260 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG260"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "TAKE_OUT", "TAKE_OUT", 0.999999)]
        public InputDouble m_dTAKE_OUT;
        [InputDouble(4, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(6, "WIDTH", "WIDTH", 0.999999)]
        public InputDouble m_dWIDTH;
        [InputDouble(7, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(8, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("BOLT", "BOLT")]
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

                Double pipeDiameter = m_dPIPE_DIA.Value, takeOut = m_dTAKE_OUT.Value, B = m_dB.Value, rodDiameter = m_dA.Value, width = m_dWIDTH.Value, G = m_dG.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, takeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (G <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidGGTZero, "G value should be greater than zero"));
                    return;
                }
                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidwidthNZero, "Width cannot be zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + (pipeDiameter / 50), Math.PI);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(Math.PI, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -width / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc.Transform(rotateMatrix);
                curveCollection.Add(arc);

                curveCollection.Add(new Line3d(new Position(-width / 2, pipeDiameter / 2 + (pipeDiameter / 50), 0), new Position(-width / 2, pipeDiameter / 2 + (pipeDiameter / 50), pipeDiameter / 2 + 2 * G)));
                curveCollection.Add(new Line3d(new Position(-width / 2, pipeDiameter / 2 + (pipeDiameter / 50), pipeDiameter / 2 + 2 * G), new Position(-width / 2, rodDiameter, B)));
                curveCollection.Add(new Line3d(new Position(-width / 2, rodDiameter, B), new Position(-width / 2, -rodDiameter, B)));
                curveCollection.Add(new Line3d(new Position(-width / 2, -rodDiameter, B), new Position(-width / 2, -pipeDiameter / 2 - (pipeDiameter / 50), pipeDiameter / 2 + 2 * G)));
                curveCollection.Add(new Line3d(new Position(-width / 2, -pipeDiameter / 2 - (pipeDiameter / 50), pipeDiameter / 2 + 2 * G), new Position(-width / 2, -pipeDiameter / 2 - (pipeDiameter / 50), 0)));
                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), width, false);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, -(pipeDiameter / 2 + G), pipeDiameter / 2 + G).Subtract(new Position(0, pipeDiameter / 2 + G, pipeDiameter / 2 + G));
                symbolGeometryHelper.ActivePosition = new Position(0, pipeDiameter / 2 + G, pipeDiameter / 2 + G);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d bolt = symbolGeometryHelper.CreateCylinder(null, G / 2.0, normal.Length);
                m_Symbolic.Outputs["BOLT"] = bolt;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG260"));
                    return;
                }
            }
        }
        #endregion

    }

}
