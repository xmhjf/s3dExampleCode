//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PSL_308.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_308
//   Author       :  Manikanth
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013    Manikanth CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net      
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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class PSL_308 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_308"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "T2", "T2", 0.999999)]
        public InputDouble t2;
        [InputDouble(3, "WIDTH", "WIDTH", 0.999999)]
        public InputDouble WIDTH;
        [InputDouble(4, "BOLT_SIZE", "BOLT_SIZE", 0.999999)]
        public InputDouble BOLT_SIZE;
        [InputDouble(5, "RTO", "RTO", 0.999999)]
        public InputDouble RTO;
        [InputDouble(6, "A", "A", 0.999999)]
        public InputDouble A1;
        [InputDouble(7, "B", "B", 0.999999)]
        public InputDouble B1;
        [InputDouble(8, "C", "C", 0.999999)]
        public InputDouble C1;
        [InputDouble(9, "R", "R", 0.999999)]
        public InputDouble R1;
        [InputDouble(10, "T1", "T1", 0.999999)]
        public InputDouble t1;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
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
                Part part = (Part)PartInput.Value;

                double rto = RTO.Value;
                double A = A1.Value;
                double C = C1.Value;
                double R = R1.Value;
                double T1 = t1.Value;
                double T2 = t2.Value;
                double width = WIDTH.Value;
                double rodDiameter = B1.Value;
                double boltSize = BOLT_SIZE.Value;

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, rto), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, R + T2, (Math.PI));
                matrix.Rotate((Math.PI), new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
                matrix.Translate(new Vector(0, -width / 2, 0));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                arc1.Transform(matrix);
                curveCollection.Add(arc1);

                curveCollection.Add(new Line3d(new Position(-width / 2, R + T2, 0), new Position(-width / 2, R + T2, C - boltSize * 1.5)));
                curveCollection.Add(new Line3d(new Position(-width / 2, R + T2, C - boltSize * 1.5), new Position(-width / 2, R + T2 + T1, C - boltSize * 1.5)));
                curveCollection.Add(new Line3d(new Position(-width / 2, R + T2 + T1, C - boltSize * 1.5), new Position(-width / 2, R + T2 + T1, C + boltSize * 1.5)));
                curveCollection.Add(new Line3d(new Position(-width / 2, R + T2 + T1, C + boltSize * 1.5), new Position(-width / 2, rodDiameter, A)));
                curveCollection.Add(new Line3d(new Position(-width / 2, rodDiameter, A), new Position(-width / 2, -rodDiameter, A)));
                curveCollection.Add(new Line3d(new Position(-width / 2, -rodDiameter, A), new Position(-width / 2, -R - T2 - T1, C + boltSize * 1.5)));
                curveCollection.Add(new Line3d(new Position(-width / 2, -R - T2 - T1, C + boltSize * 1.5), new Position(-width / 2, -R - T2 - T1, C - boltSize * 1.5)));
                curveCollection.Add(new Line3d(new Position(-width / 2, -R - T2 - T1, C - boltSize * 1.5), new Position(-width / 2, -R - T2, C - boltSize * 1.5)));
                curveCollection.Add(new Line3d(new Position(-width / 2, -R - T2, C - boltSize * 1.5), new Position(-width / 2, -R - T2, 0)));

                Projection3d pro = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), width, true);
                m_Symbolic.Outputs["BODY"] = pro;

                if (boltSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBoltSizeGTZ, "BOLT_SIZE should be greater than zero"));
                    return;
                }

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, -(R + T2 + T1 + boltSize * 2), C).Subtract(new Position(0, R + T2 + T1 + boltSize * 2, C));
                symbolGeometryHelper.ActivePosition = new Position(0, R + T2 + T1 + boltSize * 2, C);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d boltCylinder = symbolGeometryHelper.CreateCylinder(null, boltSize / 2, normal.Length);
                m_Symbolic.Outputs["BOLT"] = boltCylinder;


            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_308"));
                    return;
                }
            }
        }

        #endregion

    }
}
