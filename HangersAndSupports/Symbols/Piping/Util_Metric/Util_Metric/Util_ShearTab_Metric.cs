//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Util_ShearTab_Metric.cs
//    Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_ShearTab_Metric
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
    public class Util_ShearTab_Metric : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Util_Metric,Ingr.SP3D.Content.Support.Symbols.Util_ShearTab_Metric"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "NoTabs", "NoTabs", 1)]
        public InputDouble m_NoTabs;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(5, "L", "L", 0.999999)]
        public InputDouble m_L;
        [InputDouble(6, "Radius", "Radius", 0.999999)]
        public InputDouble m_Radius;
        [InputDouble(7, "Angle", "Angle", 0.999999)]
        public InputDouble m_Angle;
        [InputString(8, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("TAB2", "TAB2")]
        [SymbolOutput("TAB2", "TAB2")]
        [SymbolOutput("TAB3", "TAB3")]
        [SymbolOutput("TAB4", "TAB4")]
        [SymbolOutput("TAB1", "TAB1")]
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

                Double W = m_W.Value;
                Double D = m_D.Value;
                Double L = m_L.Value;
                Double radius = m_Radius.Value;
                Double angle = m_Angle.Value;
                Double numTabs = m_NoTabs.Value;
                String actualNumTabs;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                PropertyValueCodelist numCodeList = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrUtilMetricNoTabs", "NoTabs");
                CodelistItem codelist = numCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)numTabs);

                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidD, "D cannot be zero or negative"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidW, "W cannot be zero or negative"));
                    return;
                }
                if (L <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidL, "L cannot be zero or negative"));
                    return;
                }

                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                if (codelist != null)
                {
                    actualNumTabs = codelist.ShortDisplayName.Trim();
                    if (codelist.Value > 1 && angle != 0)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidAngle, "Angle offset can only be used on single tabs."));
                    }
                }
                else
                    actualNumTabs = "2";

                if (actualNumTabs == "2")
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, radius);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d tab2 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, L, W);
                    m_Symbolic.Outputs["TAB2"] = tab2;
                }
                else if (actualNumTabs == "4")
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, radius);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d tab2 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, L, W);
                    m_Symbolic.Outputs["TAB2"] = tab2;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, radius, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                    Projection3d tab3 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, W, L);
                    m_Symbolic.Outputs["TAB3"] = tab3;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -radius, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(0, 0, 1));
                    Projection3d tab4 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, W, L);
                    m_Symbolic.Outputs["TAB4"] = tab4;
                }

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -radius);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d tab1 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, L, W);
                m_Symbolic.Outputs["TAB1"] = tab1;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Util_ShearTab_Metric.cs."));
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
                PropertyValueCodelist tabsCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricNoTabs", "NoTabs");
                long numTabsValue = tabsCodelist.PropValue;
                if (numTabsValue < 1 || numTabsValue > 3)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidCodeListValue, "NoTabs code list value should be between 1 and 3."));
                    numTabsValue = 1;
                }

                double D = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricD", "D")).PropValue;
                double W = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricW", "W")).PropValue;
                double L = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricL", "L")).PropValue;

                if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if ((string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtilMetricBomDesc", "InputBomDesc")).PropValue == null)
                    {
                        bomDescription = "Shear Tab with " + numTabsValue + " " + L + " X " + W + " X " + D + " Plate Steel";
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrBOMDescription, "Error in BOMDescription of Util_ShearTab_Metric.cs."));
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
                PropertyValueCodelist tabsCodelist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricNoTabs", "NoTabs");
                if (tabsCodelist.PropValue < 1 || tabsCodelist.PropValue > 3)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrInvalidCodeListValue, "NoTabs code list value should be between 1 and 3."));
                    tabsCodelist.PropValue = 1;
                }
                double D = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricD", "D")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricW", "W")).PropValue;
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtilMetricL", "L")).PropValue;

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = D * W * L * tabsCodelist.PropValue * getSteelDensityKGPerM;
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
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Util_MetricLocalizer.GetString(Util_MetricSymbolsResourceIDs.ErrWeightCG, "Error in WeightCG of Util_ShearTab_Metric.cs."));
            }
        }
        #endregion
    }
}
