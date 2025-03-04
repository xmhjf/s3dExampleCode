//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Util_End_Plate_Metric.cs
//    Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_End_Plate_Metric
//   Author       :  Rajeswari
//   Creation Date:  16/11/2012
//   Description  :  CR-CP-222287-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   16/11/2012   Rajeswari  CR-CP-222287-Initial Creation
//   24/05/2013    Rajeswari Resolved TDL Errors
//  31/10/2013     Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
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
    public class Util_End_Plate_Metric : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_End_Plate_Metric"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(5, "PipeDia", "PipeDia", 0.999999)]
        public InputDouble m_PipeDia;
        [InputString(6, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
        [InputDouble(7, "Radius", "Radius", 0.999999)]
        public InputDouble m_Radius;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
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

                Double T = m_T.Value;
                Double W = m_W.Value;
                Double H = m_H.Value;
                Double pipeDiameter = m_PipeDia.Value;
                Double radius = m_Radius.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                if (T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidTE0, "T cannot be zero"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidW, "W cannot be zero or negative"));
                    return;
                }
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidH, "H cannot be zero or negative"));
                    return;
                }
                if (radius == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidRadiusNE0, "Radius cannot be zero."));
                    return;
                }

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Double alpha = Math.Acos(2.0 / 3.0);
                if (W <= radius * 2.0)
                {
                    alpha = Math.Acos((W / 2.0) / radius) * (180 / Math.PI);
                }
                Double A = radius - (radius * Math.Sin(alpha * Math.PI / 180));
                if (radius < W / 2.0)
                {
                    A = radius;
                    alpha = 0;
                }
                if (W > radius * 2.0)
                {
                    curveCollection.Add(new Line3d(new Position(0, W / 2.0, -radius - H), new Position(0, -W / 2.0, -radius - H)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2.0, -radius - H), new Position(0, -W / 2.0, -radius + A)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2, -radius + A), new Position(0, -radius, -radius + A)));

                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Arc3d topArc = symbolGeometryHelper.CreateArc(null, radius, Math.PI - 2 * alpha);
                    matrix.Rotate((Math.PI + alpha), new Vector(0, 0, 1));
                    matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                    topArc.Transform(matrix);
                    curveCollection.Add(topArc);

                    curveCollection.Add(new Line3d(new Position(0, radius, -radius + A), new Position(0, W / 2.0, -radius + A)));
                    curveCollection.Add(new Line3d(new Position(0, W / 2.0, -radius + A), new Position(0, W / 2.0, -radius - H)));

                    Projection3d body = new Projection3d( new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
                else
                {
                    curveCollection = new Collection<ICurve>();
                    curveCollection.Add(new Line3d(new Position(0, W / 2.0, -radius - H), new Position(0, -W / 2.0, -radius - H)));
                    curveCollection.Add(new Line3d(new Position(0, -W / 2.0, -radius - H), new Position(0, -W / 2.0, -radius + A)));

                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI + alpha * Math.PI / 180), new Vector(0, 0, 1));

                    Arc3d topArc = symbolGeometryHelper.CreateArc(null, radius, Math.PI - 2 * alpha * Math.PI / 180);
                    topArc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                    topArc.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                    topArc.Transform(matrix);
                    curveCollection.Add(topArc);

                    curveCollection.Add(new Line3d(new Position(0, W / 2.0, -radius + A), new Position(0, W / 2.0, -radius - H)));

                    Projection3d body = new Projection3d( new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Util_End_Plate_Metric.cs."));
                return;
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                double TValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricT", "T")).PropValue;
                double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricW", "W")).PropValue;
                double HValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricH", "H")).PropValue;
                double radius = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricRadius", "Radius")).PropValue;
                double alpha = Math.Acos(2.0 / 3.0);

                if (WValue <= radius * 2.0)
                {
                    alpha = Math.Acos((WValue / 2.0) / radius) * (180 / Math.PI);
                }
                double A = radius - (radius * Math.Sin(alpha * Math.PI / 180));

                if (radius < WValue / 2.0)
                {
                    A = radius;
                }
                string T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_MILLIMETER);
                string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_MILLIMETER);
                string H = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, HValue + A, UnitName.DISTANCE_MILLIMETER);

                if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == null)
                    {
                        bomDescription = T + " thk End Plate " + W + " x " + H;
                    }
                    else
                    {
                        bomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue.Trim();
                    }
                }
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrBOMDescription, "Error in BOMDescription of Util_End_Plate_Metric.cs."));
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                const int getSteelDensityKGPerM = 7900;
                double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricT", "T")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricW", "W")).PropValue;
                double H = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricH", "H")).PropValue;
                double radius = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricRadius", "Radius")).PropValue;

                double A, alpha = 0.0, arcLength, segmentArea;
                if (W <= radius * 2.0)
                {
                    alpha = Math.Acos((W / 2) / radius);
                }
                A = radius - (radius * Math.Sin(alpha));
                if (radius < W / 2.0)
                {
                    A = radius;
                    alpha = 0;
                }
                arcLength = radius * ((180) - 2.0 * (alpha * (180 / Math.PI))) / 180 * Math.PI;
                segmentArea = 0.5 * (radius * arcLength - (W * (radius - A)));

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = ((H + A) * W - segmentArea) * T * getSteelDensityKGPerM;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrWeightCG, "Error in WeightCG of Util_End_Plate_Metric.cs."));
            }
        }
        #endregion

    }
}
