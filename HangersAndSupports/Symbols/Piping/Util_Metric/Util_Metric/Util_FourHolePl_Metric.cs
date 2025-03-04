//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Util_FourHolePl_Metric.cs
//    Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_FourHolePl_Metric
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
    public class Util_FourHolePl_Metric : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_FourHolePl_Metric"
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
        [InputString(5, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
        [InputDouble(6, "C", "C", 0.999999)]
        public InputDouble m_C;
        [InputDouble(7, "HoleSize", "HoleSize", 0.999999)]
        public InputDouble m_HoleSize;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("LINE", "LINE")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("R_F_HOLE", "R_F_HOLE")]
        [SymbolOutput("R_B_HOLE", "R_B_HOLE")]
        [SymbolOutput("L_F_HOLE", "L_F_HOLE")]
        [SymbolOutput("L_B_HOLE", "L_B_HOLE")]
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
                Double C = m_C.Value;
                Double holeSize = m_HoleSize.Value;

                if (C <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidC, "C cannot be zero or negative"));
                    return;
                }
                if (holeSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidHoleSize, "Hole Size cannot be zero or negative"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidTNLT0, "T cannot be zero or negative"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidWidth, "Width cannot be zero or negative"));
                    return;
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidDepth, "Depth cannot be zero or negative"));
                    return;
                }

                Port port1 = new Port(OccurrenceConnection, part, "TopStructure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "BotStructure", new Position(0, 0, T), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d body = (Projection3d)symbolGeometryHelper.CreateBox(null, T, depth, width);
                m_Symbolic.Outputs["BODY"] = body;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-depth / 2, 0, T / 2));
                pointCollection.Add(new Position(depth / 2, 0, T / 2));
                LineString3d line = new LineString3d(pointCollection);
                m_Symbolic.Outputs["LINE"] = line;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2 - C), -(width / 2 - C), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bottomright = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeSize / 2, T + 0.00002);
                m_Symbolic.Outputs["R_F_HOLE"] = bottomright;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-(depth / 2 - C), (width / 2 - C), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d topright = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeSize / 2, T + 0.00002);
                m_Symbolic.Outputs["R_B_HOLE"] = topright;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position((depth / 2 - C), -(width / 2 - C), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bottomleft = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeSize / 2, T + 0.00002);
                m_Symbolic.Outputs["L_F_HOLE"] = bottomleft;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position((depth / 2 - C), (width / 2 - C), -0.00001);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d topleft = (Projection3d)symbolGeometryHelper.CreateCylinder(null, holeSize / 2, T + 0.00002);
                m_Symbolic.Outputs["L_B_HOLE"] = topleft;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Util_FourHolePl_Metric.cs."));
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
                Double CValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricC", "C")).PropValue;
                Double holeSizeValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricHoleSize", "HoleSize")).PropValue;

                String T = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, TValue, UnitName.DISTANCE_MILLIMETER);
                String width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_MILLIMETER);
                String depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_MILLIMETER);
                String holeSize = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, holeSizeValue, UnitName.DISTANCE_MILLIMETER);
                String C = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, CValue, UnitName.DISTANCE_MILLIMETER);

                if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == null)
                    {
                        bomDescription = T + " Plate Steel, " + width + " x " + depth + " w/Four " + holeSize + " Holes " + C + " from edge";
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrBOMDescription, "Error in BOMDescription of Util_FourHolePl_Metric.cs."));
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
                Double holeSize = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricHoleSize", "HoleSize")).PropValue;

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = ((depth * width * T) - (4.0 * (holeSize / 2.0 * holeSize / 2.0 * T * Math.PI))) * getSteelDensityKGPerM;
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrWeightCG, "Error in WeightCG of Util_FourHolePl_Metric.cs."));
            }
        }
        #endregion
    }
}
