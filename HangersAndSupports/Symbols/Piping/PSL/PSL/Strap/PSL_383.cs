//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_383.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_383
//   Author       :  Hema
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Hema    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_383 : HangerComponentSymbolDefinition
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_383"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(4, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(5, "G1", "G1", 0.999999)]
        public InputDouble G1;
        [InputDouble(6, "G2", "G2", 0.999999)]
        public InputDouble G2;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
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
                Part part = (Part)PartInput.Value;
                Double a = A.Value;
                Double c = C.Value;
                Double f = F.Value;
                Double g1 = G1.Value;
                Double g2 = G2.Value;

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, a), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (g1 == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidG1NEZ, "G1 value cannot be zero"));
                    return;
                }
                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, -(f + g2), 0), new Position(-g1 / 2.0, -(f + g2), a - g2)));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, f + g2, Math.PI);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -g1 / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);

                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, f + g2, 0), new Position(-g1 / 2.0, f + g2, a - g2)));
                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, f + g2, a - g2), new Position(-g1 / 2.0, c / 2.0, a - g2)));
                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, c / 2.0, a - g2), new Position(-g1 / 2.0, c / 2.0, a)));
                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, f, a), new Position(-g1 / 2.0, c / 2.0, a)));
                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, f, 0), new Position(-g1 / 2.0, f, a)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d innerArc = symbolGeometryHelper.CreateArc(null, f, Math.PI);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -g1 / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                innerArc.Transform(matrix);
                curveCollection.Add(innerArc);

                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, -f, 0), new Position(-g1 / 2.0, -f, a)));
                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, -f, a), new Position(-g1 / 2.0, -c / 2.0, a)));
                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, -c / 2.0, a), new Position(-g1 / 2.0, -c / 2.0, a - g2)));
                curveCollection.Add(new Line3d(new Position(-g1 / 2.0, -c / 2.0, a - g2), new Position(-g1 / 2.0, -(f + g2), a - g2)));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), g1, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_383."));
                    return;
                }
            }
        }
        #endregion
    }
}
