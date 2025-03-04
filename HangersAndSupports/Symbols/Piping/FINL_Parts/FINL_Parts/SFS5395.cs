//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   SFS5395.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5395
//   Author       :  Rajeswari
//   Creation Date:  18-March-2013
//   Description:    CR-CP-222272-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-March-2013 Rajeswari CR-CP-222272-Initial Creation
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
    public class SFS5395 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5395"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "D1", "D1", 0.999999)]
        public InputDouble m_dD1;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(5, "H1", "H1", 0.999999)]
        public InputDouble m_dH1;
        [InputDouble(6, "H2", "H2", 0.999999)]
        public InputDouble m_dH2;
        [InputDouble(7, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(8, "L1", "L1", 0.999999)]
        public InputDouble m_dL1;
        [InputDouble(9, "L2", "L2", 0.999999)]
        public InputDouble m_dL2;
        [InputDouble(10, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(11, "R1", "R1", 0.999999)]
        public InputDouble m_dR1;
        [InputDouble(12, "R2", "R2", 0.999999)]
        public InputDouble m_dR2;
        [InputDouble(13, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(14, "T", "T", 0.999999)]
        public InputDouble m_dT;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("HOLE", "HOLE")]
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
                Double D1 = m_dD1.Value;
                Double D = m_dD.Value;
                Double H1 = m_dH1.Value;
                Double H2 = m_dH2.Value;
                Double H = m_dH.Value;
                Double L1 = m_dL1.Value;
                Double L2 = m_dL2.Value;
                Double F = m_dF.Value;
                Double R1 = m_dR1.Value;
                Double R2 = m_dR2.Value;
                Double B = m_dB.Value;
                Double T = m_dT.Value;

                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                if (B == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidBNZero, "B value cannot be zero"));
                    return;
                }
                if (pipeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidPipeDiameterGTZ, "Pipe Diameter should be greater than zero"));
                    return;
                }
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, pipeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
               
                if (pipeDiameter < 0.073)
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d arc1 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, Math.PI);
                    matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, -B / 2, 0));
                    arc1.Transform(matrix);

                    curveCollection.Add(arc1);
                    curveCollection.Add(new Line3d(new Position(pipeDiameter / 2, -B / 2, 0), new Position(pipeDiameter / 2, -B / 2, -pipeDiameter / 2 + H2)));
                    curveCollection.Add(new Line3d(new Position(pipeDiameter / 2, -B / 2, -pipeDiameter / 2 + H2), new Position(pipeDiameter / 2 + T, -B / 2, -pipeDiameter / 2 + H2)));
                    curveCollection.Add(new Line3d(new Position(pipeDiameter / 2 + T, -B / 2, -pipeDiameter / 2 + H2), new Position(pipeDiameter / 2 + T, -B / 2, 0)));

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d arc2 = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2 + T, Math.PI);
                    matrix = new Matrix4X4();
                    matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                    matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, -B / 2, 0));
                    arc2.Transform(matrix);
                    curveCollection.Add(arc2);

                    curveCollection.Add(new Line3d(new Position(-pipeDiameter / 2 - T, -B / 2, 0), new Position(-pipeDiameter / 2 - T, -B / 2, pipeDiameter / 2 - T)));
                    curveCollection.Add(new Line3d(new Position(-pipeDiameter / 2 - T, -B / 2, pipeDiameter / 2 - T), new Position(-L1, -B / 2, pipeDiameter / 2 - T)));
                    curveCollection.Add(new Line3d(new Position(-L1, -B / 2, pipeDiameter / 2 - T), new Position(-L1, -B / 2, pipeDiameter / 2)));
                    curveCollection.Add(new Line3d(new Position(-L1, -B / 2, pipeDiameter / 2), new Position(-pipeDiameter / 2, -B / 2, pipeDiameter / 2)));
                    curveCollection.Add(new Line3d(new Position(-pipeDiameter / 2, -B / 2, pipeDiameter / 2), new Position(-pipeDiameter / 2, -B / 2, 0)));

                    Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(0, 1, 0), B, true);
                    m_Symbolic.Outputs["BODY"] = body;

                    Vector normal = new Position(-L2, 0, pipeDiameter / 2).Subtract(new Position(-L2, 0, pipeDiameter / 2 - T));
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(-L2, 0, pipeDiameter / 2 - T);
                    symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                    Projection3d hole = symbolGeometryHelper.CreateCylinder(null, D / 2, normal.Length);
                    m_Symbolic.Outputs["HOLE"] = hole;
                }
                else
                {
                    Revolution3d bend = new Revolution3d((new Circle3d(new Position(pipeDiameter / 2 + D / 2, 0, 0), new Vector(0, 0, 1), D / 2)), new Vector(0, -1, 0), new Position(0, 0, 0), Math.PI * 180 / 180, true);
                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                    matrix.Translate(new Vector(0, 0, 0));
                    bend.Transform(matrix);
                    m_Symbolic.Outputs["BEND"] = bend;

                    Vector normal = new Vector(0, 0, 1);
                    
                    symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                    symbolGeometryHelper.SetOrientation(normal,normal.GetOrthogonalVector());
                    Projection3d left = symbolGeometryHelper.CreateCylinder(null, D / 2, H - pipeDiameter / 2 - D);
                    matrix=new Matrix4X4();
                    matrix.Translate(new Vector(-pipeDiameter / 2 - D / 2, 0, 0));
                    left.Transform(matrix);
                    m_Symbolic.Outputs["LEFT"] = left;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                    symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                    Projection3d right = symbolGeometryHelper.CreateCylinder(null, D / 2, H1 - pipeDiameter / 2 - D);
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(pipeDiameter / 2 + D / 2, 0, 0));
                    right.Transform(matrix);
                    m_Symbolic.Outputs["RIGHT"] = right;
                }

            }
            catch //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5395.cs."));
                return;
            }
        }
        #endregion
    }

}
