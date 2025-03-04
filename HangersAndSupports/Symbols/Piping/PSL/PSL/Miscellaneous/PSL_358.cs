//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_358.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_358
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
    public class PSL_358 : HangerComponentSymbolDefinition
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_358"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "RADIUS", "RADIUS", 1)]
        public InputDouble RADIUS;
        [InputDouble(3, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(4, "B_1", "B_1", 0.999999)]
        public InputDouble B_1;
        [InputDouble(5, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(6, "W", "W", 0.999999)]
        public InputDouble W;
        [InputDouble(7, "T", "T", 0.999999)]
        public InputDouble T;
        [InputDouble(8, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(9, "B_3", "B_3", 0.999999)]
        public InputDouble B_3;
        [InputDouble(10, "B_5", "B_5", 0.999999)]
        public InputDouble B_5;
        [InputString(11, "PIPE_NOM_DIA", "PIPE_NOM_DIA", "No Value")]
        public InputString PIPE_NOM_DIA;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("HOLE1", "HOLE1")]
        [SymbolOutput("HOLE2", "HOLE2")]
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
                //-----------------------------------------------------------------------
                // Construction of Physical Aspect 
                //----------------------------------------------------------------------
                Part part = (Part)PartInput.Value;

                Double pipeDiameter = PIPE_DIA.Value;
                Double b_1 = B_1.Value;
                Double d = D.Value;
                Double t = T.Value;
                Double w = W.Value;
                Double a = A.Value;
                Double b_3 = B_3.Value;
                Double b_5 = B_5.Value;
                String pipeNominalDiameter = PIPE_NOM_DIA.Value;
                Double b = b_1;
                Double radius = RADIUS.Value;
                Double holeSize = D.Value;

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                Double elbowRadius = Convert.ToDouble(pipeNominalDiameter) * 1.5 / 1000;
                String actualRadius;
                Double y, x, outerRadius, lAdj, lAdj2, angle1, angle2;

                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualRadius = metadataManager.GetCodelistInfo("PSL_ELBOW2", "UDP").GetCodelistItem((int)radius).ShortDisplayName.Trim();
                else
                    actualRadius = "1.5D";

                if (actualRadius.Equals("5D"))
                {
                    b = b_5;
                    elbowRadius = Convert.ToDouble(pipeNominalDiameter) * 5.0 / 1000;
                }
                if (actualRadius.Equals("3D"))
                {
                    b = b_3;
                    elbowRadius = Convert.ToDouble(pipeNominalDiameter) * 3.0 / 1000;
                }
                y = elbowRadius + w / 2.0;
                x = elbowRadius - w / 2.0;
                outerRadius = elbowRadius + pipeDiameter / 2.0;
                lAdj = Math.Sqrt(Math.Abs(outerRadius * outerRadius - y * y));
                lAdj2 = Math.Sqrt(Math.Abs(outerRadius * outerRadius - x * x));
                angle1 = 45.0 * Math.PI / 180.0;
                angle2 = Math.Acos(x / (elbowRadius + pipeDiameter / 2.0));

                if (pipeDiameter / 2.0 > w / 2.0)
                    angle1 = Math.Acos(y / (elbowRadius + pipeDiameter / 2.0));

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, elbowRadius + b), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (t == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidTNEZ, "T value cannot be zero"));
                    return;
                }
                if (holeSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }

                Collection<ICurve> curveCollection = new Collection<ICurve>();

                curveCollection.Add(new Line3d(new Position(-w / 2.0, -t / 2.0, elbowRadius + b + a), new Position(w / 2.0, -t / 2.0, elbowRadius + b + a)));
                curveCollection.Add(new Line3d(new Position(w / 2.0, -t / 2.0, elbowRadius + b + a), new Position(w / 2.0, -t / 2.0, lAdj2)));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d innerArc = symbolGeometryHelper.CreateArc(null, outerRadius, (angle2 - angle1));
                matrix.Rotate(Math.PI + angle1, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(t / 2.0, elbowRadius, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                innerArc.Transform(matrix);
                curveCollection.Add(innerArc);

                curveCollection.Add(new Line3d(new Position(-w / 2.0, -t / 2.0, lAdj), new Position(-w / 2.0, -t / 2.0, elbowRadius + b + a)));

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(0, 1, 0), t, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d hole1 = symbolGeometryHelper.CreateCylinder(null, holeSize / 2.0, t);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(-(t) / 2.0, 0, elbowRadius + b));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                hole1.Transform(matrix);
                m_Symbolic.Outputs["HOLE1"] = hole1;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_458."));
                    return;
                }
            }
        }
        #endregion
    }
}
