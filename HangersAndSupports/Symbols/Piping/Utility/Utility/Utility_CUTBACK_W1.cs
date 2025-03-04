//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_CUTBACK_W1.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_CUTBACK_W1
//   Author       :  Hema
//   Creation Date:  03.11.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   03.11.2012     Hema     CR-CP-222288  Converted HS_Utility VB Project to C# .Net 
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols 
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
    public class Utility_CUTBACK_W1 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_CUTBACK_W1"
        //--------------------------------------------------------------------------------------------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(3, "WIDTH", "WIDTH", 0.999999)]
        public InputDouble m_dWIDTH;
        [InputDouble(4, "T_FLANGE", "T_FLANGE", 0.999999)]
        public InputDouble m_dT_FLANGE;
        [InputDouble(5, "DEPTH", "DEPTH", 0.999999)]
        public InputDouble m_dDEPTH;
        [InputDouble(6, "T_WEB", "T_WEB", 0.999999)]
        public InputDouble m_dT_WEB;
        [InputDouble(7, "ANGLE", "ANGLE", 0.999999)]
        public InputDouble m_dANGLE;
        [InputDouble(8, "ANGLE2", "ANGLE2", 0.999999)]
        public InputDouble m_dANGLE2;
        [InputString(9, "BOM_DESC", "BOM_DESC", "No Value")]
        public InputString m_oBOM_DESC;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY1", "BODY1")]
        [SymbolOutput("BODY2", "BODY2")]
        [SymbolOutput("BODY3", "BODY3")]
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

                Double L = m_dL.Value;
                Double width = m_dWIDTH.Value;
                Double flangeThickness = m_dT_FLANGE.Value;
                Double depth = m_dDEPTH.Value;
                Double webThickness = m_dT_WEB.Value;
                Double angle = m_dANGLE.Value;
                Double angle2 = m_dANGLE2.Value;

                Double Z1 = 0;
                Double Z2 = 0;
                Double Z3 = 0;
                Double Z4 = 0;
                Double xDirX1 = 0;
                Double xDirY1 = 0;
                Double xDirZ1 = 0;
                Double zDirX1 = 0;
                Double zDirY1 = 0;
                Double zDirZ1 = 0;
                Double xDirX2 = 0;
                Double xDirY2 = 0;
                Double xDirZ2 = 0;
                Double zDirX2 = 0;
                Double zDirY2 = 0;
                Double zDirZ2 = 0;

                if (flangeThickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidFlangeThickness, "Flange Thickness cannot be zero"));
                    return;
                }
                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                if (webThickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidWebThickness, "Web Thickness cannot be zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                if (angle > 0 && angle2 > 0)
                {
                    Z1 = (width / 2.0) / (Math.Tan(angle));
                    Z2 = (width / 2.0) / (Math.Tan(angle2));
                    Z3 = (webThickness / 2.0) / (Math.Tan(angle));
                    Z4 = (webThickness / 2.0) / (Math.Tan(angle2));
                }
                else if (angle > -0.0001 && angle2 < 0.0001)
                {
                    Z1 = 0;
                    if (angle2 > -0.0001 && angle2 < 0.0001)
                    {
                        Z2 = 0;
                        Z4 = 0;
                    }
                    else
                    {
                        Z4 = (webThickness / 2.0) / Math.Tan(angle2);
                        Z2 = (width / 2.0) / Math.Tan(angle2);
                    }
                    Z3 = 0;
                }
                else if (angle2 > -0.0001 && angle2 < 0.0001)
                {
                    if (angle > -0.0001 && angle < 0.0001)
                    {
                        Z1 = 0;
                        Z3 = 0;
                    }
                    else
                    {
                        Z1 = (width / 2.0) / Math.Tan(angle);
                        Z3 = (webThickness / 2.0) / Math.Tan(angle);
                    }
                    Z4 = 0;
                    Z2 = 0;
                }
                if ((angle == 0))
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
                else
                {
                    if (Math.Abs(angle * 180 / Math.PI) < 5)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrAngleGTFive, "Angle must be greater than 5 deg"));
                    }
                    xDirX1 = Math.Sin(angle2) / ((Math.Cos(angle2) * Math.Cos(angle2)) + (Math.Sin(angle2) * Math.Sin(angle2)));
                    xDirY1 = 0;
                    xDirZ1 = Math.Cos(angle2) / ((Math.Cos(angle2) * Math.Cos(angle2)) + (Math.Sin(angle2) * Math.Sin(angle2)));
                    zDirX1 = Math.Sin(Math.PI * (90 / ((float)180)) + angle2) / ((Math.Cos(Math.PI * (90 / ((float)180)) + angle2) * Math.Cos(Math.PI * (90 / ((float)180)) + angle2)) + (Math.Sin(Math.PI * (90 / ((float)180)) + angle2) * Math.Sin(Math.PI * (90 / ((float)180)) + angle2)));
                    zDirY1 = 0;
                    zDirZ1 = Math.Cos(Math.PI * (90 / ((float)180)) + angle2) / ((Math.Cos(Math.PI * (90 / ((float)180)) + angle2) * Math.Cos(Math.PI * (90 / ((float)180)) + angle2)) + (Math.Sin(Math.PI * (90 / ((float)180)) + angle2) * Math.Sin(Math.PI * (90 / ((float)180)) + angle2)));
                    xDirX2 = Math.Sin(-angle) / ((Math.Cos(-angle) * Math.Cos(-angle)) + (Math.Sin(-angle) * Math.Sin(-angle)));
                    xDirY2 = 0;
                    xDirZ2 = Math.Cos(-angle) / ((Math.Cos(-angle) * Math.Cos(-angle)) + (Math.Sin(-angle) * Math.Sin(-angle)));
                    zDirX2 = Math.Sin(Math.PI * (90 / ((float)180)) - angle) / ((Math.Cos(Math.PI * (90 / ((float)180)) + angle) * Math.Cos(Math.PI * (90 / ((float)180)) + angle)) + (Math.Sin(Math.PI * (90 / ((float)180)) - angle) * Math.Sin(Math.PI * (90 / ((float)180)) - angle)));
                    zDirY2 = 0;
                    zDirZ2 = Math.Cos(Math.PI * (90 / ((float)180)) - angle) / ((Math.Cos(Math.PI * (90 / ((float)180)) - angle) * Math.Cos(Math.PI * (90 / ((float)180)) - angle)) + (Math.Sin(Math.PI * (90 / ((float)180)) - angle) * Math.Sin(Math.PI * (90 / ((float)180)) - angle)));
                }

                Port port1 = new Port(OccurrenceConnection, part, "StartStructure", new Position(-width / 2, 0, -Z1), new Vector(-zDirX1, -zDirY1, -zDirZ1), new Vector(-xDirX1, -xDirY1, -xDirZ1));
                m_Symbolic.Outputs["Port1"] = port1;

                //Ports
                Port port2 = new Port(OccurrenceConnection, part, "MidStructure");

                Port port3 = new Port(OccurrenceConnection, part, "EndStructure");

                if ((port2 != null))
                {
                    port2.Origin = new Position(0, 0, L / 2);
                    port2.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                    m_Symbolic.Outputs["Port2"] = port2;

                    port3.Origin = new Position(-width / 2.0, 0, L + Z2);
                    port3.SetOrientation(new Vector(zDirX2, zDirY2, zDirZ2), new Vector(xDirX2, xDirY2, xDirZ2));
                    m_Symbolic.Outputs["Port3"] = port3;
                }
                else
                {
                    port3 = new Port(OccurrenceConnection, part, "EndStructure", new Position(-width / 2.0, 0, L + Z2), new Vector(zDirX2, zDirY2, zDirZ2), new Vector(xDirX2, xDirY2, xDirZ2));
                    m_Symbolic.Outputs["Port2"] = port3;
                }

                Collection<Position> pointCollection = new Collection<Position>();

                pointCollection.Add(new Position(-width / 2.0, depth / 2.0, -Z1));
                pointCollection.Add(new Position(-width / 2.0, depth / 2.0, L + Z2));
                pointCollection.Add(new Position(width / 2.0, depth / 2.0, L - Z2));
                pointCollection.Add(new Position(width / 2.0, depth / 2.0, Z1));
                pointCollection.Add(new Position(-width / 2.0, depth / 2.0, -Z1));

                Vector projectionVector = new Vector(0, -flangeThickness, 0);
                Projection3d body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY1"] = body;

                pointCollection = new Collection<Position>();

                pointCollection.Add(new Position(-width / 2.0, -depth / 2.0, -Z1));
                pointCollection.Add(new Position(-width / 2.0, -depth / 2.0, L + Z2));
                pointCollection.Add(new Position(width / 2.0, -depth / 2.0, L - Z2));
                pointCollection.Add(new Position(width / 2.0, -depth / 2.0, Z1));
                pointCollection.Add(new Position(-width / 2.0, -depth / 2.0, -Z1));

                Collection<ICurve> curveCollection2 = new Collection<ICurve>();
                curveCollection2.Add(new LineString3d(pointCollection));

                projectionVector = new Vector(0, flangeThickness, 0);
                body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY2"] = body;

                pointCollection = new Collection<Position>();

                pointCollection.Add(new Position(-webThickness / 2.0, -depth / 2.0 + flangeThickness, -Z3));
                pointCollection.Add(new Position(-webThickness / 2.0, -depth / 2.0 + flangeThickness, L + Z4));
                pointCollection.Add(new Position(webThickness / 2.0, -depth / 2.0 + flangeThickness, L - Z4));
                pointCollection.Add(new Position(webThickness / 2.0, -depth / 2.0 + flangeThickness, Z3));
                pointCollection.Add(new Position(-webThickness / 2.0, -depth / 2.0 + flangeThickness, -Z3));
                                
                projectionVector = new Vector(0, depth - flangeThickness * 2.0, 0);
                body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY3"] = body;

            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_CUTBACK_W1"));
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
                double flangeThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "T_FLANGE")).PropValue;
                double webThicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "T_WEB")).PropValue;
                double widthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "WIDTH")).PropValue;
                double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "DEPTH")).PropValue;
                double angleValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_CUTBACK", "ANGLE")).PropValue;
                double angle2Value = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_CUTBACK", "ANGLE2")).PropValue;
                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "L")).PropValue;
                string bomDescription = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "BOM_DESC")).PropValue;
                string flangeThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, flangeThicknessValue, UnitName.DISTANCE_INCH);
                string webThickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, webThicknessValue, UnitName.DISTANCE_INCH);
                string width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_INCH);
                string depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_INCH);
                string angle = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, angleValue, UnitName.ANGLE_DEGREE);
                string angle2 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, angle2Value, UnitName.ANGLE_DEGREE);
                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_INCH);

                double bomZ1, bomZ2,bomOrderLength;

                bomZ1 = (widthValue / 2) / Math.Tan((angleValue));
                bomZ2 = (widthValue / 2) / Math.Tan((angle2Value));

                bomOrderLength = LValue + bomZ1 + bomZ2;

                if (bomDescription.Trim() == "None")
                {
                    bomString = " ";
                }
                else
                {
                    if (bomDescription.Trim() == "")
                    {
                        bomString = "Starting Angle = " + angle + ", Ending Angle = " + angle2 + ", Order Length = " + width;
                    }
                    else
                    {
                        bomString = bomDescription.Trim();
                    }
                }
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_CUTBACK_W1"));
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
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "L")).PropValue;
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "WIDTH")).PropValue;
                double flangeThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "T_FLANGE")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "DEPTH")).PropValue;
                double webThickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_W", "T_WEB")).PropValue;

                area = 2 * (flangeThickness * width) + webThickness * (depth - 2 * flangeThickness);
                weight = area * L * getSteelDensityKGPerM;

                cogX = 0;
                cogY = 0;
                cogZ = L / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_CUTBACK_W1"));
                }
            }
        }
    }
   #endregion
}

