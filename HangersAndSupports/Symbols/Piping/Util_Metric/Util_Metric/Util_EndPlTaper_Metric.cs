//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Util_EndPlTaper_Metric.cs
//    Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_EndPlTaper_Metric
//   Author       :  Rajeswari
//   Creation Date:  16/11/2012
//   Description  :  CR-CP-222287-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   16/11/2012    Rajeswari  CR-CP-222287-Initial Creation
//   24/05/2013    Rajeswari  Resolved TDL Errors
//   31/10/2013    Vijaya     CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//   30/12/2014     PVK        TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   12/02/2015     PVK        TR-CP-124751	Model graphics for Generic assembly TypeF do not match Preview
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
    public class Util_EndPlTaper_Metric : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_EndPlTaper_Metric"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "R", "R", 0.999999)]
        public InputDouble m_R;
        [InputDouble(3, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(4, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(5, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(6, "Angle", "Angle", 0.999999)]
        public InputDouble m_Angle;
        [InputString(7, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
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

                Double R = m_R.Value;
                Double T = m_T.Value;
                Double W = m_W.Value;
                Double H = m_H.Value;
                Double angle = m_Angle.Value;

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Double W2 = Math.Sin((angle / 2.0)) * R * 2.0;
                Double Calc1 = Math.Cos((angle / 2.0)) * R;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                if (HgrCompareDoubleService.cmpdbl(T , 0)==true)
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
                if (HgrCompareDoubleService.cmpdbl(R, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidRNE0, "R cannot be zero."));
                    return;
                }

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                curveCollection.Add(new Line3d(new Position(0, W2 / 2.0, -Calc1), new Position(0, W / 2.0, -(H + R))));
                curveCollection.Add(new Line3d(new Position(0, W / 2.0, -(H + R)), new Position(0, -W / 2.0, -(H + R))));
                curveCollection.Add(new Line3d(new Position(0, -W / 2.0, -(H + R)), new Position(0, -W2 / 2.0, -Calc1)));

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d topArc = symbolGeometryHelper.CreateArc(null, R,  angle);
                matrix.Rotate((Math.PI / 2 - angle / 2.0), new Vector(0, 0, 1));
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                topArc.Transform(matrix);
                curveCollection.Add(topArc);

                Projection3d body = new Projection3d( new ComplexString3d(curveCollection), new Vector(1, 0, 0), T, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Util_EndPlTaper_Metric.cs."));
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
                double angleValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricAngle", "Angle")).PropValue;
                double RValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricR", "R")).PropValue;
                double W2Value = Math.Sin((angleValue / 2.0)) * RValue * 2.0;
                double cacl1 = Math.Cos((angleValue / 2.0)) * RValue;

                string T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_MILLIMETER);
                string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, W2Value, UnitName.DISTANCE_MILLIMETER);
                string H = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, HValue + RValue - cacl1, UnitName.DISTANCE_MILLIMETER);

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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrBOMDescription, "Error in BOMDescription of Util_EndPlTaper_Metric.cs."));
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
                double angle = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricAngle", "Angle")).PropValue;
                double R = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricR", "R")).PropValue;

                double W2 = Math.Sin(angle / 2.0) * R * 2.0;
                double arcLength = R * (angle * 180 / Math.PI) / 180 * Math.PI;
                double A = R - (R * Math.Cos(angle / 2.0));
                double segmentArea = 0.5 * (R * arcLength - W2 * (R - A));

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (((H + A) * W2 - segmentArea) * T - ((W2 / 2.0 - W / 2.0) * (H + A) * T)) * getSteelDensityKGPerM;
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrWeightCG, "Error in WeightCG of Util_EndPlTaper_Metric.cs."));
            }
        }
        #endregion
    }
}
