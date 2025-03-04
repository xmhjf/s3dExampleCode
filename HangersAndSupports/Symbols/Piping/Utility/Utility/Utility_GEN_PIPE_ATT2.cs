//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GEN_PIPE_ATT2.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_PIPE_ATT2
//   Author       :Sasidhar  
//   Creation Date:5-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   5-11-2012    Sasidhar   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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
    [VariableOutputs]
    public class Utility_GEN_PIPE_ATT2 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_PIPE_ATT2"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_A;
        [InputDouble(4, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(5, "HOLE_DIA", "HOLE_DIA", 0.999999)]
        public InputDouble m_HOLE_DIA;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble m_C;
        [InputDouble(7, "ELBOW_RADIUS", "ELBOW_RADIUS", 0.999999)]
        public InputDouble m_ELBOW_RADIUS;
        [InputDouble(8, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_PIPE_DIA;
        [InputString(9, "BOM_DESC", "BOM_DESC", "No Value")]
        public InputString m_oBOM_DESC;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("HOLE", "HOLE")]
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

                Double D = m_D.Value;
                Double A = m_A.Value;
                Double T = m_T.Value;
                Double holeDiameter = m_HOLE_DIA.Value;
                Double C = m_C.Value;
                Double elbowRadius = m_ELBOW_RADIUS.Value;
                Double pipeDiameter = m_PIPE_DIA.Value;
                Double y, x, acosCheck, outerRadius, angle1, angle2, lAdj, lAdj2;

                if (T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidT, "Thickness T cannot be zero"));
                    return;
                }
                if (holeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrHoleDiameterGTZero, "Hole Diameter should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidwidthD, "Width D should be greater than zero"));
                    return;
                }
               
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                y = elbowRadius + D / 2;
                x = elbowRadius - D / 2;
                outerRadius = elbowRadius + pipeDiameter / 2;
                angle1 = Math.PI / 4;
                acosCheck = (y - D) / (elbowRadius + pipeDiameter / 2);

                if ((acosCheck < -1) || (acosCheck > 1))
                {
                    angle2 = 0;
                }
                else
                {
                    angle2 = Math.Acos((y - D) / (elbowRadius + pipeDiameter / 2));
                }
                lAdj = Math.Sqrt(Math.Abs(outerRadius * outerRadius - y * y));
                lAdj2 = Math.Sqrt(Math.Abs(outerRadius * outerRadius - x * x));

                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, lAdj + A - C), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (pipeDiameter > D)
                {
                    angle1 = (Math.Acos(y / (elbowRadius + pipeDiameter / 2)));
                }
                Collection<ICurve> curvecollection = new Collection<ICurve>();
                curvecollection.Add(new Line3d(new Position(-D / 2, -T / 2, lAdj + A), new Position(D / 2, -T / 2, lAdj + A)));
                curvecollection.Add(new Line3d(new Position(D / 2, -T / 2, lAdj + A), new Position(D / 2, -T / 2, lAdj2)));
                curvecollection.Add(new Line3d(new Position(-D / 2, -T / 2, lAdj), new Position(-D / 2, -T / 2, lAdj + A)));
                SymbolGeometryHelper symbolGeomHlpr = new SymbolGeometryHelper();
                symbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                symbolGeomHlpr.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI + angle1), new Vector(0, 0, 1));

                Arc3d arc2 = symbolGeomHlpr.CreateArc(null, outerRadius, (angle2 - angle1));
                arc2.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                arc2.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 1, 0));
                arc2.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(T / 2, elbowRadius, 0));
                arc2.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                arc2.Transform(matrix);
                curvecollection.Add(arc2);

                Projection3d body = new Projection3d(new ComplexString3d(curvecollection), new Vector(0, 1, 0), T, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeomHlpr = new SymbolGeometryHelper();
                symbolGeomHlpr.ActivePosition = new Position(0, -T / 2 - 0.0001, lAdj + A - C);
                symbolGeomHlpr.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d hole = (Projection3d)symbolGeomHlpr.CreateCylinder(OccurrenceConnection, holeDiameter / 2, T + 0.0002);
                m_Symbolic.Outputs["HOLE"] = hole;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GEN_PIPE_ATT2"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Double T = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "T")).PropValue;
                Double holeDiameter = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "HOLE_DIA")).PropValue;
                String bomDescriptionValue = ((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "BOM_DESC")).PropValue;
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                bomDescription = part.PartDescription + ",Hole Dia=" + holeDiameter + ", Thickness=" + T;
                if (bomDescriptionValue == "")
                {
                    bomDescription = part.PartDescription + ", Hole Dia=" + holeDiameter + ", Thickness=" + T;
                }
                else
                {
                    bomDescription = bomDescriptionValue;
                }

                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_GEN_PIPE_ATT2"));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {

                Double weight, cogX, cogY, cogZ, alpha = 0.0D, arcLength, segmantArea, A, y, x, outerRadius, lAdj, lAdj2;
                const int getSteelDensityKGPerM = 7900;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                Double pipeDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                Double elbowRadius = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "ELBOW_RADIUS")).PropValue;
                Double holeDiameter = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "HOLE_DIA")).PropValue;
                Double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "T")).PropValue;
                Double AValue = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "A")).PropValue;
                Double D = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "D")).PropValue;
                Double C = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT2", "C")).PropValue;

                y = elbowRadius + D / 2;
                x = elbowRadius - D / 2;
                outerRadius = elbowRadius + pipeDiameter / 2;

                lAdj = Math.Sqrt(Math.Abs(outerRadius * outerRadius - y * y));
                lAdj2 = Math.Sqrt(Math.Abs(outerRadius * outerRadius - x * x));

                if (D <= elbowRadius)
                {
                    alpha = Math.Acos((D / 2) / (elbowRadius / 2));
                }

                A = elbowRadius / 2 - (elbowRadius / 2 * Math.Sin(alpha));

                if (elbowRadius / 2 < D / 2)
                {
                    A = elbowRadius / 2;
                    alpha = 0;
                }

                arcLength = elbowRadius / 2 * ((180) - 2 * (180 / Math.PI) * (alpha)) / 180 * Math.PI;
                segmantArea = 0.5 * (elbowRadius / 2 * arcLength - (D * (elbowRadius / 2 - A)));
                weight = (((lAdj + AValue - lAdj2) * D * T) - ((Math.PI) * Math.Pow(holeDiameter / 2, 2) * T) + ((1 / 2.0 * D * (lAdj2 - lAdj) - segmantArea) * T)) * getSteelDensityKGPerM;
                cogX = 0;
                cogY = 0;
                cogZ = 0;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GEN_PIPE_ATT2"));
                }
            }

        }

        #endregion
    }
}
