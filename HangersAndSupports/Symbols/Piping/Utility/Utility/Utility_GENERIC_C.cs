//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GENERIC_C.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GENERIC_C
//   Author       :  Rajeswari
//   Creation Date:  30/10/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 30/10/2012     Rajeswari   CR-CP-222288  Converted HS_Utility VB Project to C# .Net   
//	 27/03/2013		Hema 	  DI-CP-228142  Modify the error handling for delivered H&S symbols
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
    public class Utility_GENERIC_C : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GENERIC_C"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_L;
        [InputDouble(3, "Width", "Width", 0.999999)]
        public InputDouble m_Width;
        [InputDouble(4, "Depth", "Depth", 0.999999)]
        public InputDouble m_Depth;
        [InputDouble(5, "FlangeTh", "FlangeTh", 0.999999)]
        public InputDouble m_FlangeTh;
        [InputDouble(6, "WebTh", "WebTh", 0.999999)]
        public InputDouble m_WebTh;
        [InputString(7, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_InputBomDesc;
        [InputDouble(8, "BomUnits", "BomUnits", 1)]
        public InputDouble m_BomUnits;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
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
                Part part =(Part)m_PartInput.Value ;

                Double L = m_L.Value;
                Double width = m_Width.Value;
                Double depth = m_Depth.Value;
                Double flangeThickness = m_FlangeTh.Value;
                Double webThickness = m_WebTh.Value;

                Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, depth / 2, width / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                if (L <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrLGTZero, "Length should be greater than zero"));
                    return;
                }
                if (depth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrDepthGTZero, "Depth should be greater than zero"));
                    return;
                }
                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWidthGTZero, "Width should be greater than zero"));
                    return;
                }
                if (flangeThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrFlangeThicknessGTZero, "Flange Thickness should be greater than zero"));
                    return;
                }
                if (webThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWebThicknessGTZero, "Web Thickness should be greater than zero"));
                    return;
                }

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, 0, 0));
                pointCollection.Add(new Position(0, 0, width));
                pointCollection.Add(new Position(0, flangeThickness, width));
                pointCollection.Add(new Position(0, flangeThickness, webThickness));
                pointCollection.Add(new Position(0, depth - flangeThickness, webThickness));
                pointCollection.Add(new Position(0, depth - flangeThickness, width));
                pointCollection.Add(new Position(0, depth, width));
                pointCollection.Add(new Position(0, depth, 0));
                pointCollection.Add(new Position(0, 0, 0));
                               
                Vector projectionVector = new Vector(L, 0, 0);
                Projection3d body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GENERIC_C"));
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


                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_L", "L")).PropValue;
                double widthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_Width", "Width")).PropValue;
                double flangeThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_FlangeTh", "FlangeTh")).PropValue;
                double webThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_WebTh", "WebTh")).PropValue;
                double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_Depth", "Depth")).PropValue;

                string inputBomDesc = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_BomDesc", "InputBomDesc")).PropValue;

                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_BomUnits", "BomUnits");

                string bomUnits = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                string width, flangeThickness, webThickness, depth, L;
                //Set the BOM units as per the users selection
                if (bomUnits.ToUpper() == "METRIC")
                {
                    width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_MILLIMETER);
                    flangeThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, flangeThicknessValue, UnitName.DISTANCE_MILLIMETER);
                    webThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, webThicknessValue, UnitName.DISTANCE_MILLIMETER);
                    depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_MILLIMETER);
                    L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_MILLIMETER);
                }
                else
                {
                    width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_INCH);
                    flangeThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, flangeThicknessValue, UnitName.DISTANCE_INCH);
                    webThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, webThicknessValue, UnitName.DISTANCE_INCH);
                    depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_INCH);
                    L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_INCH);
                }
                if (inputBomDesc.ToUpper() == "NONE")
                {
                    bomString = "";
                }
                else if (inputBomDesc == "")
                {
                    bomString = "C - " + width + "x" + depth + ",Flange =" + flangeThickness + ", Web =" + webThickness + ", Length = " + L;
                }
                else
                {
                    return inputBomDesc;
                }
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_GENERIC_C"));
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

                Double weight, cogX, cogY, cogZ, area;
                const int getSteelDensityKGPerM = 7900;
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_L", "L")).PropValue;
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_Width", "Width")).PropValue;
                double flangeThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_FlangeTh", "FlangeTh")).PropValue;
                double webThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_WebTh", "WebTh")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_Depth", "Depth")).PropValue;
                area = 2 * flangeThickness * width + webThickness * (depth - 2 * flangeThickness);
                weight = area * L * getSteelDensityKGPerM;
                cogX = L / 2;
                cogY = depth / 2;
                cogZ = width / 2;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GENERIC_C"));
                }
            }
        }

        #endregion
    }
}
