//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_GenericBrace.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GenericBrace
//   Author       :  Rajeswari
//   Creation Date:  01/11/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01/11/2012    Rajeswari CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   30/Dec/2014    PVK       TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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

    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class Utility_GenericBrace : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_GenericBrace"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_L;
        [InputDouble(3, "W", "W", 0.999999)]
        public InputDouble m_W;
        [InputDouble(4, "Depth", "Depth", 0.999999)]
        public InputDouble m_Depth;
        [InputDouble(5, "FlangeT", "FlangeT", 0.999999)]
        public InputDouble m_FlangeT;
        [InputDouble(6, "WebT", "WebT", 0.999999)]
        public InputDouble m_WebT;
        [InputString(7, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_oInputBomDesc;
        [InputString(8, "BraceType", "BraceType", "No Value")]
        public InputString m_BraceType;
        [InputDouble(9, "Angle", "Angle", 0.999999)]
        public InputDouble m_Angle;
        [InputDouble(10, "BraceOrient", "BraceOrient", 1)]
        public InputDouble m_BraceOrient;

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
                Part part = (Part)m_PartInput.Value;

                Double L = m_L.Value;
                Double W = m_W.Value;
                Double depth = m_Depth.Value;
                Double flangeThickness = m_FlangeT.Value;
                Double webThickness = m_WebT.Value;
                Double angle = m_Angle.Value;
                int braceOrient = (int)m_BraceOrient.Value;
                string braceType = m_BraceType.Value;

                Double xDirX1, xDirY1, xDirZ1;
                Double zDirX1, zDirY1, zDirZ1;
                Double xDirX2, xDirY2, xDirZ2;
                Double zDirX2, zDirY2, zDirZ2;

                if (HgrCompareDoubleService.cmpdbl(L, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidLength, "Length cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(flangeThickness, 0) == true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidFlangeThickness, "Flange Thickness cannot be zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(angle * 180 / Math.PI), 0) == false)
                {
                    if (HgrCompareDoubleService.cmpdbl(Math.Abs(angle * 180 / Math.PI) % 180, 0) == true)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidangle, "Angle cannot be zero"));
                        return;
                    }
                }
                if (HgrCompareDoubleService.cmpdbl(angle, 0) == true)
                {
                    xDirX1 = 1;
                    xDirY1 = 0;
                    xDirZ1 = 0;
                    zDirX1 = 0;
                    zDirY1 = 0;
                    zDirZ1 = 1;
                    xDirX2 = 1;
                    xDirY2 = 0;
                    xDirZ2 = 0;
                    zDirX2 = 0;
                    zDirY2 = 0;
                    zDirZ2 = 1;
                }
                else if (braceOrient == 1)
                {
                    xDirX1 = Math.Cos(angle) / (Math.Cos(angle) * Math.Cos(angle) + Math.Sin(angle) * Math.Sin(angle));
                    xDirY1 = 0;
                    xDirZ1 = Math.Sin(angle) / (Math.Cos(angle) * Math.Cos(angle) + Math.Sin(angle) * Math.Sin(angle));
                    zDirX1 = Math.Cos(Math.PI * 90 / ((float)180) + angle) / (Math.Cos(Math.PI * 90 / ((float)180) + angle) * Math.Cos(Math.PI * 90 / ((float)180) + angle)) + (Math.Sin(Math.PI * 90 / ((float)180) + angle) * Math.Sin(Math.PI * 90 / ((float)180) + angle));
                    zDirY1 = 0;
                    zDirZ1 = Math.Sin(Math.PI * 90 / ((float)180) + angle) / (Math.Cos(Math.PI * 90 / ((float)180) + angle) * Math.Cos(Math.PI * 90 / ((float)180) + angle)) + (Math.Sin(Math.PI * 90 / ((float)180) + angle) * Math.Sin(Math.PI * 90 / ((float)180) + angle));
                    xDirX2 = Math.Cos(-angle) / (Math.Cos(-angle) * Math.Cos(-angle) + Math.Sin(-angle) * Math.Sin(-angle));
                    xDirY2 = 0;
                    xDirZ2 = Math.Sin(-angle) / (Math.Cos(-angle) * Math.Cos(-angle) + Math.Sin(-angle) * Math.Sin(-angle));
                    zDirX2 = Math.Cos(Math.PI * 90 / ((float)180) - angle) / (Math.Cos(Math.PI * 90 / ((float)180) - angle) * Math.Cos(Math.PI * 90 / ((float)180) - angle)) + (Math.Sin(Math.PI * 90 / ((float)180) - angle) * Math.Sin(Math.PI * 90 / ((float)180) - angle));
                    zDirY2 = 0;
                    zDirZ2 = Math.Sin(Math.PI * 90 / ((float)180) - angle) / (Math.Cos(Math.PI * 90 / ((float)180) - angle) * Math.Cos(Math.PI * 90 / ((float)180) - angle)) + (Math.Sin(Math.PI * 90 / ((float)180) - angle) * Math.Sin(Math.PI * 90 / ((float)180) - angle));
                }
                else
                {
                    xDirX1 = Math.Sin(angle) / (Math.Cos(angle) * Math.Cos(angle) + Math.Sin(angle) * Math.Sin(angle));
                    xDirY1 = Math.Cos(angle) / (Math.Cos(angle) * Math.Cos(angle) + Math.Sin(angle) * Math.Sin(angle));
                    xDirZ1 = 0;
                    zDirX1 = 0;
                    zDirY1 = 0;
                    zDirZ1 = 1;
                    xDirX2 = Math.Sin(-angle) / (Math.Cos(-angle) * Math.Cos(-angle) + Math.Sin(-angle) * Math.Sin(-angle));
                    xDirY2 = Math.Cos(-angle) / (Math.Cos(-angle) * Math.Cos(-angle) + Math.Sin(-angle) * Math.Sin(-angle));
                    xDirZ2 = 0;
                    zDirX2 = 0;
                    zDirY2 = 0;
                    zDirZ2 = 1;
                }
                Collection<Position> pointCollection = new Collection<Position>();
                Projection3d body;
                Vector projectionVector;
                if (braceType.ToUpper() == "W SECTION")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                    m_Symbolic.Outputs["Port1"] = port1;
                    Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, 0, depth / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Port2"] = port2;
                    Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                    m_Symbolic.Outputs["Port3"] = port3;

                    pointCollection.Add(new Position(0, -W / 2, depth / 2));
                    pointCollection.Add(new Position(0, W / 2, depth / 2));
                    pointCollection.Add(new Position(0, W / 2, depth / 2 - flangeThickness));
                    pointCollection.Add(new Position(0, webThickness / 2, depth / 2 - flangeThickness));
                    pointCollection.Add(new Position(0, webThickness / 2, -depth / 2 + flangeThickness));
                    pointCollection.Add(new Position(0, W / 2, -depth / 2 + flangeThickness));
                    pointCollection.Add(new Position(0, W / 2, -depth / 2));
                    pointCollection.Add(new Position(0, -W / 2, -depth / 2));
                    pointCollection.Add(new Position(0, -W / 2, -depth / 2 + flangeThickness));
                    pointCollection.Add(new Position(0, -webThickness / 2, -depth / 2 + flangeThickness));
                    pointCollection.Add(new Position(0, -webThickness / 2, depth / 2 - flangeThickness));
                    pointCollection.Add(new Position(0, -W / 2, depth / 2 - flangeThickness));
                    pointCollection.Add(new Position(0, -W / 2, depth / 2));

                    projectionVector = new Vector(L, 0, 0);
                    body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
                else if (braceType.ToUpper() == "C SECTION")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                    m_Symbolic.Outputs["Port1"] = port1;
                    Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, W / 2, depth / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Port2"] = port2;
                    Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                    m_Symbolic.Outputs["Port3"] = port3;

                    pointCollection = new Collection<Position>();
                    pointCollection.Add(new Position(0, 0, 0));
                    pointCollection.Add(new Position(0, 0, depth));
                    pointCollection.Add(new Position(0, flangeThickness, depth));
                    pointCollection.Add(new Position(0, flangeThickness, webThickness));
                    pointCollection.Add(new Position(0, W - flangeThickness, webThickness));
                    pointCollection.Add(new Position(0, W - flangeThickness, depth));
                    pointCollection.Add(new Position(0, W, depth));
                    pointCollection.Add(new Position(0, W, 0));
                    pointCollection.Add(new Position(0, 0, 0));

                    projectionVector = new Vector(L, 0, 0);
                    body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
                else if (braceType.ToUpper() == "HSS SECTION")
                {
                    Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                    m_Symbolic.Outputs["Port1"] = port1;
                    Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, W / 2, depth / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Port2"] = port2;
                    Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                    m_Symbolic.Outputs["Port3"] = port3;

                    pointCollection = new Collection<Position>();
                    pointCollection.Add(new Position(0, 0, 0));
                    pointCollection.Add(new Position(0, 0, depth));
                    pointCollection.Add(new Position(0, W, depth));
                    pointCollection.Add(new Position(0, W, 0));
                    pointCollection.Add(new Position(0, 0, 0));

                    projectionVector = new Vector(L, 0, 0);
                    body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
                else
                {
                    Port port1 = new Port(OccurrenceConnection, part, "BeginCap", new Position(0, 0, 0), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(zDirX1, zDirY1, zDirZ1));
                    m_Symbolic.Outputs["Port1"] = port1;
                    Port port2 = new Port(OccurrenceConnection, part, "Neutral", new Position(L / 2, W, depth), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_Symbolic.Outputs["Port2"] = port2;
                    Port port3 = new Port(OccurrenceConnection, part, "EndCap", new Position(L, 0, 0), new Vector(xDirX2, xDirY2, xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                    m_Symbolic.Outputs["Port3"] = port3;

                    pointCollection = new Collection<Position>();
                    pointCollection.Add(new Position(0, 0, 0));
                    pointCollection.Add(new Position(0, 0, depth));
                    pointCollection.Add(new Position(0, flangeThickness, depth));
                    pointCollection.Add(new Position(0, flangeThickness, flangeThickness));
                    pointCollection.Add(new Position(0, W, flangeThickness));
                    pointCollection.Add(new Position(0, W, 0));
                    pointCollection.Add(new Position(0, 0, 0));

                    projectionVector = new Vector(L, 0, 0);
                    body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                    m_Symbolic.Outputs["BODY"] = body;
                }
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_GenericBrace"));
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
                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GenericBrace", "L")).PropValue;
                double WValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GenericBrace", "W")).PropValue;
                double flangeThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GenericBrace", "FlangeT")).PropValue;
                double webThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GenericBrace", "WebT")).PropValue;
                double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GenericBrace", "Depth")).PropValue;
                string braceType = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GenericBrace", "BraceType")).PropValue;
                string inputBomDesc = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GenericBrace", "InputBomDesc")).PropValue;

                string width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WValue, UnitName.DISTANCE_MILLIMETER);
                string flangeThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, flangeThicknessValue, UnitName.DISTANCE_MILLIMETER);
                string webThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, webThicknessValue, UnitName.DISTANCE_MILLIMETER);
                string depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_MILLIMETER);
                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_MILLIMETER);

                if (inputBomDesc.Trim() == "")
                {
                    if (braceType.ToUpper() == "W SECTION")
                    {
                        bomString = "W " + width + " x " + depth + ", Flange =" + flangeThickness + ", Web =" + webThickness + ", Length = " + L;
                    }
                    else if (braceType.ToUpper() == "C SECTION")
                    {
                        bomString = "C " + width + " x " + depth + ", Flange =" + flangeThickness + ", Web =" + webThickness + ", Length = " + L;
                    }
                    else
                    {
                        bomString = "L" + width + " x " + depth + " x " + flangeThickness + ", Length = " + L;
                    }
                }
                else
                {
                    bomString = inputBomDesc.Trim();
                }
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_GenericBrace"));
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
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GenericBrace", "L")).PropValue;
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GenericBrace", "W")).PropValue;
                double flangeThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GenericBrace", "FlangeT")).PropValue;
                double webThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GenericBrace", "WebT")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GenericBrace", "Depth")).PropValue;
                string braceType = (string)((PropertyValueString)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GenericBrace", "BraceType")).PropValue;

                if (braceType.ToUpper() == "W SECTION")
                {
                    weight = (((L * width * flangeThickness) * 2) + (L * (depth - flangeThickness * 2) * webThickness)) * getSteelDensityKGPerM;
                }
                else if (braceType.ToUpper() == "C SECTION")
                {
                    weight = 2 * (flangeThickness * depth) + (webThickness * width) * L * getSteelDensityKGPerM;
                }
                else
                {
                    weight = ((L * (width - flangeThickness) * flangeThickness) + (L * depth * flangeThickness)) * getSteelDensityKGPerM;
                }
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_GenericBrace"));
                }
            }
        }

        #endregion
    }
}
