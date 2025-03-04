//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Util_NotchPl_Metric.cs
//    Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_NotchPl_Metric
//   Author       :  Rajeswari
//   Creation Date:  16/11/2012
//   Description  :  CR-CP-222287-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   16/11/2012    Rajeswari CR-CP-222287-Initial Creation
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
    public class Util_NotchPl_Metric : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_NotchPl_Metric"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(3, "Width", "Width", 0.999999)]
        public InputDouble m_Width;
        [InputDouble(4, "Depth", "Depth", 0.999999)]
        public InputDouble m_Depth;
        [InputDouble(5, "X", "X", 0.999999)]
        public InputDouble m_X;
        [InputDouble(6, "Z", "Z", 0.999999)]
        public InputDouble m_Z;
        [InputString(7, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
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
                Part part = (Part)m_PartInput.Value;

                Double T = m_T.Value;
                Double width = m_Width.Value;
                Double depth = m_Depth.Value;
                Double X = m_X.Value;
                Double Z = m_Z.Value;

                Port port1 = new Port(OccurrenceConnection, part, "MidStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "CornerStructure", new Position(-width / 2.0, depth / 2.0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidTNLT0, "T cannot be zero or negative"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidWidth, "width cannot be zero or negative"));
                    return;
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidDepth, "depth cannot be zero or negative"));
                    return;
                }

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-width / 2.0 + X, -depth / 2.0 + Z, -T / 2.0));
                pointCollection.Add(new Position(-width / 2.0, -depth / 2.0 + Z, -T / 2.0));
                pointCollection.Add(new Position(-width / 2.0, depth / 2.0, -T / 2.0));
                pointCollection.Add(new Position(width / 2.0, depth / 2.0, -T / 2.0));
                pointCollection.Add(new Position(width / 2.0, -depth / 2.0, -T / 2.0));
                pointCollection.Add(new Position(-width / 2.0 + X, -depth / 2.0, -T / 2.0));
                pointCollection.Add(new Position(-width / 2.0 + X, -depth / 2.0 + Z, -T / 2.0));

                Projection3d body = new Projection3d( new LineString3d(pointCollection), new Vector(0, 0, 1), T, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Util_NotchPl_Metric.cs."));
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
                Double TValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricT", "T")).PropValue;
                Double widthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricWidth", "Width")).PropValue;
                Double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricDepth", "Depth")).PropValue;
                Double XValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricX", "X")).PropValue;
                Double ZValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricZ", "Z")).PropValue;

                String T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_MILLIMETER);
                String width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_MILLIMETER);
                String depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_MILLIMETER);
                String X = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, ZValue, UnitName.DISTANCE_MILLIMETER);
                String Z = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, XValue, UnitName.DISTANCE_MILLIMETER);

                if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == null)
                    {
                        bomDescription = T + " Plate Steel, " + width + " x " + depth + " w/ " + X + " x " + Z + " notch";
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrBOMDescription, "Error in BOMDescription of Util_NotchPl_Metric.cs."));
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
                Double T = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricT", "T")).PropValue;
                Double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricWidth", "Width")).PropValue;
                Double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricDepth", "Depth")).PropValue;
                Double X = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricX", "X")).PropValue;
                Double Z = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricZ", "Z")).PropValue;

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = ((depth * width * T) - (X * Z * T)) * getSteelDensityKGPerM;
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrWeightCG, "Error in WeightCG of Util_NotchPl_Metric.cs."));
            }
        }
        #endregion
    }
}
