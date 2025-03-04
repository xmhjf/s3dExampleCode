//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GEN_PIPE_ATT.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_PIPE_ATT
//   Author       :Sasidhar  
//   Creation Date:7-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   7-11-2012    Sasidhar   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
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
    public class Utility_GEN_PIPE_ATT : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GEN_PIPE_ATT"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(3, "R", "R", 0.999999)]
        public InputDouble m_R;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble m_A;
        [InputDouble(5, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(6, "HOLE_DIA", "HOLE_DIA", 0.999999)]
        public InputDouble m_HOLE_DIA;
        [InputDouble(7, "ELBOW_RADIUS", "ELBOW_RADIUS", 0.999999)]
        public InputDouble m_ELBOW_RADIUS;
        [InputDouble(8, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_PIPE_DIA;

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
                Double R = m_R.Value;
                Double A = m_A.Value;
                Double T = m_T.Value;
                Double holeDiameter = m_HOLE_DIA.Value;
                Double elbowRadius = m_ELBOW_RADIUS.Value;
                Double pipeDiameter = m_PIPE_DIA.Value;
                Double y, x, acosCheck, outerRad, angle1, angle2, lAdj, lAdj2;

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
                if (elbowRadius < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidElbowRadius, "Elbow Radius should be greater than zero"));
                    return;
                }
                if (R <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidTopRadius, "Top Radius R should be greater than zero"));
                    return;
                }
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, elbowRadius + A + holeDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                y = elbowRadius + R;
                x = elbowRadius + R - D;
                outerRad = elbowRadius + pipeDiameter / 2;
                angle1 = Math.PI / 4;
                acosCheck = (y - D) / (elbowRadius + pipeDiameter / 2);
                if ((acosCheck < -1) || (acosCheck > 1))
                {
                    angle2 = 0;
                }
                else
                {
                    angle2 = Math.Acos(((y - D) / (elbowRadius + pipeDiameter / 2)));
                }
                lAdj = Math.Sqrt(Math.Abs((outerRad * outerRad - y * y)));
                lAdj2 = Math.Sqrt(Math.Abs((outerRad * outerRad - x * x)));

                if (pipeDiameter / 2 > R)
                {
                    angle1 = Math.Acos(y / (elbowRadius + pipeDiameter / 2));
                }
                Collection<ICurve> curvecollection = new Collection<ICurve>();

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI), new Vector(0, 0, 1));

                Arc3d arc = symbolGeometryHelper.CreateArc(null, R, Math.PI);
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 1, 0));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(T / 2, 0, A + elbowRadius));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                arc.Transform(matrix);
                curvecollection.Add(arc);

                curvecollection.Add(new Line3d(new Position(R, -T / 2, elbowRadius + A), new Position(-R + D, -T / 2, lAdj2)));
                curvecollection.Add(new Line3d(new Position(-R, -T / 2, lAdj), new Position(-R, -T / 2, elbowRadius + A)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI + angle1), new Vector(0, 0, 1));

                arc = symbolGeometryHelper.CreateArc(null, outerRad, (angle2 - angle1));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 1, 0));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(T / 2, elbowRadius, 0));
                arc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                arc.Transform(matrix);
                curvecollection.Add(arc);

                Projection3d body = new Projection3d(new ComplexString3d(curvecollection), new Vector(0, 1, 0), T, true);
                m_Symbolic.Outputs["BODY"] = body;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -T / 2 - 0.0001, elbowRadius + A);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d hole = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeDiameter / 2, T + 0.0002);
                m_Symbolic.Outputs["HOLE"] = hole;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GEN_PIPE_ATT"));
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
                Double TValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT", "T")).PropValue;
                Double holeDiameterValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT", "HOLE_DIA")).PropValue;
                String T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_INCH);
                String holeDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, holeDiameterValue, UnitName.DISTANCE_INCH);

                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                bomDescription = part.PartDescription + ", Hole Dia=" + holeDiameter + ", Thickness=" + T;

                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_GEN_PIPE_ATT"));
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

                Double weight, cogX, cogY, cogZ, angle2, lAdj, lAdj2, alpha = 0.0D, arcLength, segmantArea, A, y, x, outerRadius, angle1, acosCkeck;
                const int getSteelDensityKGPerM = 7900;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                Double pipeDiameter = (Double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                Double elbowRadius = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT", "ELBOW_RADIUS")).PropValue;
                Double holeDiameter = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT", "HOLE_DIA")).PropValue;
                Double thickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT", "T")).PropValue;
                Double AValue = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT", "A")).PropValue;
                Double D = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT", "D")).PropValue;
                Double R = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GEN_PIPE_ATT", "R")).PropValue;
                y = elbowRadius + R;
                x = elbowRadius + R - D;
                outerRadius = elbowRadius + pipeDiameter / 2;
                angle1 = 45;
                acosCkeck = (y - D) / (elbowRadius + pipeDiameter / 2);

                if ((acosCkeck < -1) || (acosCkeck > 1))
                {
                    angle2 = 0;
                }
                else
                {
                    angle2 = (Math.Acos((y - D) / (elbowRadius + pipeDiameter / 2))) * (180 / Math.PI);
                }
                lAdj = Math.Sqrt(Math.Abs(outerRadius * outerRadius - y * y));
                lAdj2 = Math.Sqrt(Math.Abs(outerRadius * outerRadius - x * x));

                if (pipeDiameter / 2 > R)
                {
                    angle1 = (Math.Acos(y / (elbowRadius + pipeDiameter / 2))) * (180 / Math.PI);
                }

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
                arcLength = elbowRadius / 2 * ((180) - 2 * (alpha * (180 / Math.PI))) / 180 * Math.PI;
                segmantArea = 0.5 * (elbowRadius / 2 * arcLength - (D * (elbowRadius / 2 - A)));

                weight = (((((elbowRadius + AValue) - lAdj2) * (D - (D - (R * R)))) * thickness) + (((elbowRadius + AValue) - lAdj2) * ((D - (R * R)) / 2) * thickness) + ((((D * ((lAdj2 - lAdj) / 2)) * thickness) - (segmantArea * thickness))) + (Math.PI * (R * R) * thickness)) * getSteelDensityKGPerM;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GEN_PIPE_ATT"));
                }
            }
        }

        #endregion
    }
}
