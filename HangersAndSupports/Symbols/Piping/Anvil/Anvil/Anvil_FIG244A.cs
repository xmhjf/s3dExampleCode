//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG244A.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG244A
//   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari  CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    public class Anvil_FIG244A : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG244A"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(4, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble m_dA;
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

                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double W = m_dW.Value;
                Double T = m_dT.Value;
                Double A = m_dA.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -pipeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (T < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTNZero, "T value should not be lessthan zero"));
                    return;
                }
                if (A < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidANZero, "A value should not be lessthan zero"));
                    return;
                }
                if (W == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidWNZero, "W value should not be lessthan zero"));
                    return;
                }

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                curveCollection.Add(new Line3d(new Position(-W / 2, -A / 2 - T, -pipeDiameter / 2), new Position(-W / 2, -A / 2 - T, 0)));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, A / 2 + T, -Math.PI);
                matrix.Rotate(2 * Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -W / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);

                curveCollection.Add(new Line3d(new Position(-W / 2, A / 2 + T, 0), new Position(-W / 2, A / 2 + T, -pipeDiameter / 2)));
                curveCollection.Add(new Line3d(new Position(-W / 2, A / 2 + T, -pipeDiameter / 2), new Position(-W / 2, A / 2, -pipeDiameter / 2)));
                curveCollection.Add(new Line3d(new Position(-W / 2, A / 2, -pipeDiameter / 2), new Position(-W / 2, A / 2, 0)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d innerArc = symbolGeometryHelper.CreateArc(null, A / 2, -Math.PI);
                matrix.Rotate(2 * Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -W / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                innerArc.Transform(matrix);
                curveCollection.Add(innerArc);

                curveCollection.Add(new Line3d(new Position(-W / 2, -A / 2, 0), new Position(-W / 2, -A / 2, -pipeDiameter / 2)));
                curveCollection.Add(new Line3d(new Position(-W / 2, -A / 2, -pipeDiameter / 2), new Position(-W / 2, -A / 2 - T, -pipeDiameter / 2)));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, -0), W, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG244A"));
                return;
            }
        }
        #endregion

    }

}
