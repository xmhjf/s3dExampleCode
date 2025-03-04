//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Util_ClevisPl_Metric.cs
//    Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_ClevisPl_Metric
//   Author       :  Rajeswari
//   Creation Date:  15/11/2012
//   Description  :  CR-CP-222287-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15/11/2012   Rajeswari  CR-CP-222287-Initial Creation
//   24/05/2013    Rajeswari Resolved TDL Errors
//  31/10/2013     Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    public class Util_ClevisPl_Metric : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_ClevisPl_Metric"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Width", 0.999999)]
        public InputDouble m_Width;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble m_C;
        [InputDouble(4, "HoleSize", "HoleSize", 0.999999)]
        public InputDouble m_HoleSize;
        [InputDouble(5, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(6, "R", "R", 0.999999)]
        public InputDouble m_R;
        [InputDouble(7, "S", "S", 0.999999)]
        public InputDouble m_S;
        [InputDouble(8, "T", "T", 0.999999)]
        public InputDouble m_T;
        [InputDouble(9, "U", "U", 0.999999)]
        public InputDouble m_U;
        [InputDouble(10, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(11, "G", "G", 0.999999)]
        public InputDouble m_G;
        [InputString(12, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
        [InputDouble(13, "InputWeight", "InputWeight", 0.999999)]
        public InputDouble m_InputWeight;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Box1", "Box1")]
        [SymbolOutput("Box2", "Box2")]
        [SymbolOutput("Box3", "Box3")]
        [SymbolOutput("Box4", "Box4")]
        [SymbolOutput("Pin", "Pin")]
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

                Double width = m_Width.Value;
                Double C = m_C.Value;
                Double holeSize = m_HoleSize.Value;
                Double W = m_W.Value;
                Double R = m_R.Value;
                Double S = m_S.Value;
                Double T = m_T.Value;
                Double U = m_U.Value;
                Double H = m_H.Value;
                Double G = m_G.Value;
                Double inputWeight = m_InputWeight.Value;

                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidW, "W cannot be zero or negative"));
                    return;
                }
                if (G <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidG, "G cannot be zero or negative"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidWidth, "Width cannot be zero or negative"));
                    return;
                }
                if (S <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidS, "S cannot be zero or negative"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidTNLT0, "T cannot be zero or negative"));
                    return;
                }
                if (U <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidU, "U cannot be zero or negative"));
                    return;
                }
                if (H <= 0 && R <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidHandRNLT0, "H and R cannot be zero or negative"));
                    return;
                }
                if (R < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidRNE0, "R cannot be zero."));
                    return;
                }
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -H - G), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(0, 1, 0));
                Projection3d box1 = (Projection3d)symbolGeometryHelper.CreateBox(null, G, width, width);
                m_Symbolic.Outputs["Box1"] = box1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, S / 2.0, -(G + (H + R) / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d box2 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, W, H + R);
                m_Symbolic.Outputs["Box2"] = box2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -S / 2.0 - T, -(G + (H + R) / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d box3 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, W, H + R);
                m_Symbolic.Outputs["Box3"] = box3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -G);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(0, 1, 0));
                Projection3d box4 = (Projection3d)symbolGeometryHelper.CreateBox(null, T, S, W);
                m_Symbolic.Outputs["Box4"] = box4;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -S / 2.0 - T * 2.0, -G - H);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                Projection3d pin = symbolGeometryHelper.CreateCylinder(null, U / 2.0, S + T * 4.0);
                m_Symbolic.Outputs["Pin"] = pin;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Util_ClevisPl_Metric.cs."));
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
                string inputBomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue;

                if (inputBomDescription == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if (inputBomDescription == null)
                        bomDescription = "Concrete Clevis Plate";
                    else
                        bomDescription = inputBomDescription.Trim();
                }
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrBOMDescription, "Error in BOMDescription of Util_ClevisPl_Metric.cs."));
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

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (Double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricInWt", "InputWeight")).PropValue;
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrWeightCG, "Error in WeightCG of Util_ClevisPl_Metric.cs."));
            }
        }
        #endregion
    }
}
