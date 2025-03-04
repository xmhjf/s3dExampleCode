//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_SHEAR_TAB.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_SHEAR_TAB
//   Author       :Sasidhar  
//   Creation Date:3-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   3-11-2012     Sasidhar  CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   30/Dec/2014    PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    public class Utility_SHEAR_TAB : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_SHEAR_TAB"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "NO_TABS", "NO_TABS", 1)]
        public InputDouble m_NO_TABS;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(5, "L", "L", 0.999999)]
        public InputDouble m_L;
        [InputDouble(6, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_PIPE_DIA;
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
                Part part =(Part)m_PartInput.Value ;

                Double W = m_W.Value;
                Double D = m_D.Value;
                Double L = m_L.Value;
                Double pipeDiameter = m_PIPE_DIA.Value;
                int actualNoofTabs =(int) m_NO_TABS.Value;

                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrTabDepthGTZero, "Tab Depth should be greater than zero"));
                    return;
                }
                if (L <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrTabLengthGTZero, "Tab Length should be greater than zero"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrTabWidthGTZero, "Tab Width should be greater than zero"));
                    return;
                }
                if (actualNoofTabs < 1 || actualNoofTabs > 3)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidTabs, "Number of Tabs should be between 1 to 3"));
                    return;
                }
                

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (actualNoofTabs == 2)
                {
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, pipeDiameter / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d tab2 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, L, W);
                    m_Symbolic.Outputs["TAB2"] = tab2;
                }

                if (actualNoofTabs == 3)//The short description of 3 is 4
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, pipeDiameter / 2);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                    Projection3d tab2 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, L, W);
                    m_Symbolic.Outputs["TAB2"] = tab2;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, pipeDiameter / 2, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(0, 0, 1));
                    Projection3d tab3 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, W, L);
                    m_Symbolic.Outputs["TAB3"] = tab3;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, -pipeDiameter / 2, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(0, 0, 1));
                    Projection3d tab4 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, W, L);
                    m_Symbolic.Outputs["TAB4"] = tab4;
                }

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -pipeDiameter / 2);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, -1), new Vector(1, 0, 0));
                Projection3d tab1 = (Projection3d)symbolGeometryHelper.CreateBox(null, D, L, W);
                m_Symbolic.Outputs["TAB1"] = tab1;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_SHEAR_TAB"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_SHEAR_TAB", "NO_TABS");
                long noTabsValue = finishCodelist.PropValue;
                string noTabs = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)noTabsValue).DisplayName;
                double DValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_SHEAR_TAB", "D")).PropValue;
                double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_SHEAR_TAB", "W")).PropValue;
                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_SHEAR_TAB", "L")).PropValue;

                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_INCH);
                string W = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_INCH);
                string D = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, DValue, UnitName.DISTANCE_INCH);

                bomString = "Shear Tab with " + noTabs + " " + L + " X " + W + " X " + D;

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_SHEAR_TAB"));
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

                Double weight, cogX, cogY, cogZ;
                const int getSteelDensityKGPerM = 7900;
                double D = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_SHEAR_TAB", "D")).PropValue;
                double W = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_SHEAR_TAB", "W")).PropValue;
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_SHEAR_TAB", "L")).PropValue;

                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrUtility_SHEAR_TAB", "NO_TABS");
                long noTabs = finishCodelist.PropValue;
                long numTabs = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)noTabs).Value;

                weight = D * W * L * numTabs * getSteelDensityKGPerM;

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_SHEAR_TAB"));
                }
            }
        }

        #endregion
    }
}
