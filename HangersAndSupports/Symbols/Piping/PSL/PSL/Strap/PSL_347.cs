//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PSL_347.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_347
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
    public class PSL_347 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_347"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "G1", "G1", 0.999999)]
        public InputDouble G11;
        [InputDouble(3, "F", "F", 0.999999)]
        public InputDouble F1;
        [InputDouble(4, "G2", "G2", 0.999999)]
        public InputDouble G22;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble A1;
        [InputDouble(6, "B", "B", 0.999999)]
        public InputDouble B1;
        [InputDouble(7, "C", "C", 0.999999)]
        public InputDouble C1;
        [InputDouble(8, "D", "D", 0.999999)]
        public InputDouble D1;
        [InputDouble(9, "E", "E", 0.999999)]
        public InputDouble E1;
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
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                double A = A1.Value;
                double B = B1.Value;
                double C = C1.Value;
                double D = D1.Value;
                double E = E1.Value;
                double F = F1.Value;
                double G1 = G11.Value;
                double G2 = G22.Value;

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, A), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                if (G1 == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidG1NEZ, "G1 value cannot be zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                curveCollection.Add(new Line3d(new Position(-G1 / 2, -(F + G2), 0), new Position(-G1 / 2, -(F + G2), A - G2)));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, F + G2, (Math.PI));
                matrix.Rotate((Math.PI), new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -G1 / 2, 0));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                arc1.Transform(matrix);
                curveCollection.Add(arc1);

                curveCollection.Add(new Line3d(new Position(-G1 / 2, F + G2, 0), new Position(-G1 / 2, F + G2, A - G2)));
                curveCollection.Add(new Line3d(new Position(-G1 / 2, F + G2, A - G2), new Position(-G1 / 2, C / 2, A - G2)));
                curveCollection.Add(new Line3d(new Position(-G1 / 2, C / 2, A - G2), new Position(-G1 / 2, C / 2, A)));
                curveCollection.Add(new Line3d(new Position(-G1 / 2, F, A), new Position(-G1 / 2, C / 2, A)));
                curveCollection.Add(new Line3d(new Position(-G1 / 2, F, 0), new Position(-G1 / 2, F, A)));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc2 = symbolGeometryHelper.CreateArc(null, F, (Math.PI));
                matrix = new Matrix4X4();
                matrix.Rotate((Math.PI), new Vector(0, 0, 1));
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -G1 / 2, 0));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                arc2.Transform(matrix);
                curveCollection.Add(arc2);

                curveCollection.Add(new Line3d(new Position(-G1 / 2, -F, 0), new Position(-G1 / 2, -F, A)));
                curveCollection.Add(new Line3d(new Position(-G1 / 2, -F, A), new Position(-G1 / 2, -C / 2, A)));
                curveCollection.Add(new Line3d(new Position(-G1 / 2, -C / 2, A), new Position(-G1 / 2, -C / 2, A - G2)));
                curveCollection.Add(new Line3d(new Position(-G1 / 2, -C / 2, A - G2), new Position(-G1 / 2, -(F + G2), A - G2)));

                Projection3d pro = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), G1, true);
                m_Symbolic.Outputs["BODY"] = pro;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_347"));
                    return;
                }
            }
        }

        #endregion

    }
}
