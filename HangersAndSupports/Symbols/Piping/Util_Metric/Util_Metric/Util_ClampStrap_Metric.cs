//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Util_ClampStrap_Metric.cs
//    Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_ClampStrap_Metric
//   Author       :  Rajeswari
//   Creation Date:  19/11/2012
//   Description  :  CR-CP-222287-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19/11/2012   Rajeswari CR-CP-222287-Initial Creation
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
    public class Util_ClampStrap_Metric : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_ClampStrap_Metric"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "R", "R", 0.999999)]
        public InputDouble m_R;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_A;
        [InputDouble(4, "K", "K", 0.999999)]
        public InputDouble m_K;
        [InputDouble(5, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(6, "BoltDia", "BoltDia", 0.999999)]
        public InputDouble m_BoltDia;
        [InputDouble(7, "BoltL", "BoltL", 0.999999)]
        public InputDouble m_BoltL;
        [InputDouble(8, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(9, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(10, "Angle", "Angle", 0.999999)]
        public InputDouble m_Angle;
        [InputDouble(11, "L", "L", 0.999999)]
        public InputDouble m_L;
        [InputDouble(12, "Inset", "Inset", 0.999999)]
        public InputDouble m_Inset;
        [InputString(13, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Bottom", "Bottom")]
        [SymbolOutput("FTop", "FTop")]
        [SymbolOutput("BTop", "BTop")]
        [SymbolOutput("Box1", "Box1")]
        [SymbolOutput("Box2", "Box2")]
        [SymbolOutput("Box3", "Box3")]
        [SymbolOutput("Box4", "Box4")]
        [SymbolOutput("Pin1", "Pin1")]
        [SymbolOutput("Pin2", "Pin2")]
        [SymbolOutput("Pin3", "Pin3")]
        [SymbolOutput("Pin4", "Pin4")]
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
                Double A = m_A.Value;
                Double K = m_K.Value;
                Double H = m_H.Value;
                Double boltDiameter = m_BoltDia.Value;
                Double boltLength = m_BoltL.Value;
                Double T = m_T.Value;
                Double W = m_W.Value;
                Double angle = m_Angle.Value;
                Double L = m_L.Value;
                Double inset = m_Inset.Value;

                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidW, "W cannot be zero or negative"));
                    return;
                }
                if (boltDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidBoltDiameter, "Bolt Diameter cannot be zero or negative"));
                    return;
                }
                if (boltLength == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidBoltLength, "Bolt Length cannot be zero"));
                    return;
                }
                if (R == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidRNE0, "R cannot be zero."));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidTNLT0, "T cannot be zero or negative"));
                    return;
                }
                if (L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidLNE0, "L cannot be zero."));
                    return;
                }
                if (H < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidHLT0, "H cannot be negative."));
                    return;
                }
                Double calc1, calc2, calc3, calc4;

                calc1 = Math.Sin(angle / 2) * R;
                calc2 = Math.Sqrt((R * R) - (calc1 * calc1));
                calc3 = Math.Sin(angle / 2) * (R + T);

                if (((R + T) * (R + T)) - (calc3 * calc3) > 0)
                    calc4 = Math.Sqrt(((R + T) * (R + T)) - (calc3 * calc3));
                else
                    calc4 = 0;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, -R - T), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if ((angle) * 180 / Math.PI < 180)
                {
                    calc2 = -calc2;
                    calc4 = -calc4;
                }
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc1 = symbolGeometryHelper.CreateArc(null, R + T, angle);
                matrix.Rotate(Math.PI / 2 - (angle / 2), new Vector(0, 0, 1));
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -L / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc1.Transform(matrix);
                curveCollection.Add(arc1);

                curveCollection.Add(new Line3d(new Position(-L / 2, calc3, calc4), new Position(-L / 2, calc1, calc2)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc2 = symbolGeometryHelper.CreateArc(null, R, angle);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2 - (angle / 2), new Vector(0, 0, 1));
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -L / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc2.Transform(matrix);
                curveCollection.Add(arc2);

                curveCollection.Add(new Line3d(new Position(-L / 2, -calc1, calc2), new Position(-L / 2, -calc3, calc4)));

                Projection3d bottom = new Projection3d( new ComplexString3d(curveCollection), new Vector(1, 0, 0), L, true);
                m_Symbolic.Outputs["Bottom"] = bottom;

                curveCollection = new Collection<ICurve>();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc3 = symbolGeometryHelper.CreateArc(null, R + T, -angle);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2 + (angle / 2), new Vector(0, 0, 1));
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -L / 2 + inset + W / 2.0, 0));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc3.Transform(matrix);
                curveCollection.Add(arc3);
                curveCollection.Add(new Line3d(new Position(L / 2 - inset - W / 2.0, calc3, calc4), new Position(L / 2 - inset - W / 2.0, calc1, calc2)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc4 = symbolGeometryHelper.CreateArc(null, R, -angle);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2 + (angle / 2), new Vector(0, 0, 1));
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -L / 2 + inset + W / 2.0, 0));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc4.Transform(matrix);
                curveCollection.Add(arc4);

                curveCollection.Add(new Line3d(new Position(L / 2 - inset - W / 2.0, -calc1, calc2), new Position(L / 2 - inset - W / 2.0, -calc3, calc4)));

                Projection3d ftop = new Projection3d( new ComplexString3d(curveCollection), new Vector(1, 0, 0), W, true);
                m_Symbolic.Outputs["FTop"] = ftop;

                curveCollection = new Collection<ICurve>();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc5 = symbolGeometryHelper.CreateArc(null, R + T, -angle);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2 + (angle / 2), new Vector(0, 0, 1));
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -L / 2 + inset - W / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc5.Transform(matrix);
                curveCollection.Add(arc5);

                curveCollection.Add(new Line3d(new Position(-L / 2 + inset - W / 2.0, calc3, calc4), new Position(-L / 2 + inset - W / 2.0, calc1, calc2)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc6 = symbolGeometryHelper.CreateArc(null, R, -angle);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI / 2 + (angle / 2), new Vector(0, 0, 1));
                matrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -L / 2 + inset - W / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc6.Transform(matrix);
                curveCollection.Add(arc6);

                curveCollection.Add(new Line3d(new Position(-L / 2 + inset - W / 2.0, -calc1, calc2), new Position(-L / 2 + inset - W / 2.0, -calc3, calc4)));

                Projection3d btop = new Projection3d( new ComplexString3d(curveCollection), new Vector(1, 0, 0), W, true);
                m_Symbolic.Outputs["BTop"] = btop;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-L / 2.0 + inset + W / 2, R / 2.0 + A / 4.0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(-1, 0, 0), new Vector(0, 1, 0));
                Projection3d box1 = (Projection3d)symbolGeometryHelper.CreateBox(null, W, A / 2.0 - R, H + T * 2.0);
                m_Symbolic.Outputs["Box1"] = box1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(L / 2.0 - inset - W / 2, R / 2.0 + A / 4.0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d box2 = (Projection3d)symbolGeometryHelper.CreateBox(null, W, A / 2.0 - R, H + T * 2.0);
                m_Symbolic.Outputs["Box2"] = box2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-L / 2.0 + inset + W / 2, -(R / 2.0 + A / 4.0), 0);
                symbolGeometryHelper.SetOrientation(new Vector(-1, 0, 0), new Vector(0, 1, 0));
                Projection3d box3 = (Projection3d)symbolGeometryHelper.CreateBox(null, W, A / 2.0 - R, H + T * 2.0);
                m_Symbolic.Outputs["Box3"] = box3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(L / 2.0 - inset - W / 2, -(R / 2.0 + A / 4.0), 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d box4 = (Projection3d)symbolGeometryHelper.CreateBox(null, W, A / 2.0 - R, H + T * 2.0);
                m_Symbolic.Outputs["Box4"] = box4;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(L / 2.0 - inset, R + T + boltDiameter / 2.0, boltLength / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d pin1 = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2.0, boltLength);
                m_Symbolic.Outputs["Pin1"] = pin1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-L / 2.0 + inset, R + T + boltDiameter / 2.0, boltLength / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d pin2 = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2.0, boltLength);
                m_Symbolic.Outputs["Pin2"] = pin2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(L / 2.0 - inset, -(R + T + boltDiameter / 2.0), boltLength / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d pin3 = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2.0, boltLength);
                m_Symbolic.Outputs["Pin3"] = pin3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-L / 2.0 + inset, -(R + T + boltDiameter / 2.0), boltLength / 2.0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d pin4 = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2.0, boltLength);
                m_Symbolic.Outputs["Pin4"] = pin4;

            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Util_ClampStrap_Metric.cs."));
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
                if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == null)
                    {
                        bomDescription = "Clamp Strap";
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrBOMDescription, "Error in BOMDescription of Util_ClampStrap_Metric.cs."));
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
                double R = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricR", "R")).PropValue;
                double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricT", "T")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricW", "W")).PropValue;
                double A = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricA", "A")).PropValue;
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricL", "L")).PropValue;
                double angle = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricAngle", "Angle")).PropValue;
                double boltDiameter = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricBoltDia", "BoltDia")).PropValue;
                double boltLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricBoltL", "BoltL")).PropValue;

                double plateWt = (((Math.PI) * (R + T) * (R + T) * L) - ((Math.PI) * R * R * L)) * (angle / 360) * getSteelDensityKGPerM;
                double strapWt = (((Math.PI) * (R + T) * (R + T) * W) - ((Math.PI) * R * R * W)) * ((360 - angle) / 360) * getSteelDensityKGPerM * 2.0;
                double earWt = (((L * (A / 2.0 - R - T) * (T * 2.0)) * 2.0 + (W * (A / 2.0 - R - T) * (T * 2.0)) * 2.0)) * getSteelDensityKGPerM;
                double boltWt = (boltDiameter / 2.0 * boltDiameter / 2.0 * boltLength * (Math.PI)) * getSteelDensityKGPerM * 4.0;

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = plateWt + strapWt + earWt + boltWt;
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrWeightCG, "Error in WeightCG of Util_ClampStrap_Metric.cs."));
            }
        }
        #endregion
    }
}
