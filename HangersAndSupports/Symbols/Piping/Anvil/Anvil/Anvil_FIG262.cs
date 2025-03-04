//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG262.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG262
//   Author       : Vijaya 
//   Creation Date: 2-May-2013 
//   Description: Initial Creation-CR-CP-222292

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   2-May-2013     Vijaya    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    public class Anvil_FIG262 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG262"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(3, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(4, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(6, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(7, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                Double D = m_dD.Value, F = m_dF.Value, E = m_dE.Value, B = m_dB.Value, A = m_dA.Value;
               
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, D), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (B == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBNZero, "B value cannot be zero"));
                    return;
                }

                curveCollection.Add(new Line3d(new Position(-B / 2, D - E, 0), new Position(-B / 2, D - E, D - F)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, E - D, Math.PI);
                rotateMatrix.Rotate(Math.PI, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -B / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc1.Transform(rotateMatrix);
                curveCollection.Add(arc1);

                curveCollection.Add(new Line3d(new Position(-B / 2, E - D, 0), new Position(-B / 2, E - D, D - F)));
                curveCollection.Add(new Line3d(new Position(-B / 2, A / 2, D - F), new Position(-B / 2, E - D, D - F)));
                curveCollection.Add(new Line3d(new Position(-B / 2, A / 2, D), new Position(-B / 2, A / 2, D - F)));
                curveCollection.Add(new Line3d(new Position(-B / 2, E - D - F, D), new Position(-B / 2, A / 2, D)));
                curveCollection.Add(new Line3d(new Position(-B / 2, E - D - F, 0), new Position(-B / 2, E - D - F, D)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc2 = symbolGeometryHelper.CreateArc(null, E - D - F, Math.PI);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(Math.PI, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -B / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc2.Transform(rotateMatrix);
                curveCollection.Add(arc2);

                curveCollection.Add(new Line3d(new Position(-B / 2, D - E + F, 0), new Position(-B / 2, D - E + F, D)));
                curveCollection.Add(new Line3d(new Position(-B / 2, D - E + F, D), new Position(-B / 2, -A / 2, D)));
                curveCollection.Add(new Line3d(new Position(-B / 2, -A / 2, D), new Position(-B / 2, -A / 2, D - F)));
                curveCollection.Add(new Line3d(new Position(-B / 2, -A / 2, D - F), new Position(-B / 2, D - E, D - F)));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), B, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG262"));
                    return;
                }
            }
        }
        #endregion

    }

}
