//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_CUTBACK_L2.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_CUTBACK_L2
//   Author       :  Vijay
//   Creation Date:  02.11.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   02.11.2012     Vijay    CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
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
    public class Utility_CUTBACK_L2 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG 
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_CUTBACK_L2"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(3, "WIDTH", "WIDTH", 0.999999)]
        public InputDouble m_dWIDTH;
        [InputDouble(4, "THICKNESS", "THICKNESS", 0.999999)]
        public InputDouble m_dTHICKNESS;
        [InputDouble(5, "DEPTH", "DEPTH", 0.999999)]
        public InputDouble m_dDEPTH;
        [InputDouble(6, "ANGLE", "ANGLE", 0.999999)]
        public InputDouble m_dANGLE;
        [InputDouble(7, "ANGLE2", "ANGLE2", 0.999999)]
        public InputDouble m_dANGLE2;
        [InputString(8, "BOM_DESC", "BOM_DESC", "No Value")]
        public InputString m_oBOM_DESC;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("BODY1", "BODY1")]
        [SymbolOutput("BODY2", "BODY2")]
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
                Double thickness = m_dTHICKNESS.Value;
                Double depth = m_dDEPTH.Value;
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

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }
                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                if (angle > 0 && angle2 > 0)
                {
                    Z1 = thickness / Math.Tan(angle);
                    Z4 = depth / Math.Tan(angle2);
                    Z2 = thickness / Math.Tan(angle2);
                    Z3 = depth / Math.Tan(angle);
                }
                else if (angle > -0.0001 && angle < 0.0001)
                {
                    Z1 = 0;
                    if (angle2 > -0.0001 && angle2 < 0.0001)
                    {
                        Z4 = 0;
                        Z2 = 0;
                    }
                    else
                    {
                        Z4 = depth / Math.Tan(angle2);
                        Z4 = thickness / Math.Tan(angle2);
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
                        Z1 = thickness / Math.Tan(angle);
                        Z3 = depth / Math.Tan(angle);
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
                    xDirX1 = 0;
                    xDirY1 = Math.Cos(angle) / ((Math.Cos(angle) * Math.Cos(angle)) + (Math.Sin(angle) * Math.Sin(angle)));
                    xDirZ1 = Math.Sin(angle) / ((Math.Cos(angle) * Math.Cos(angle)) + (Math.Sin(angle) * Math.Sin(angle)));
                    zDirX1 = 0;
                    zDirY1 = Math.Cos(Math.PI * (90 / ((float)180)) + angle) / ((Math.Cos(Math.PI * (90 / ((float)180)) + angle) * Math.Cos(Math.PI * (90 / ((float)180)) + angle)) + (Math.Sin(Math.PI * (90 / ((float)180)) + angle) * Math.Sin(Math.PI * (90 / ((float)180)) + angle)));
                    zDirZ1 = Math.Sin(Math.PI * (90 / ((float)180)) + angle) / ((Math.Cos(Math.PI * (90 / ((float)180)) + angle) * Math.Cos(Math.PI * (90 / ((float)180)) + angle)) + (Math.Sin(Math.PI * (90 / ((float)180)) + angle) * Math.Sin(Math.PI * (90 / ((float)180)) + angle)));
                    xDirX2 = 0;
                    xDirY2 = Math.Cos(-angle2) / ((Math.Cos(-angle2) * Math.Cos(-angle2)) + (Math.Sin(-angle2) * Math.Sin(-angle2)));
                    xDirZ2 = Math.Sin(-angle2) / ((Math.Cos(-angle2) * Math.Cos(-angle2)) + (Math.Sin(-angle2) * Math.Sin(-angle2)));
                    zDirX2 = 0;
                    zDirY2 = Math.Cos(Math.PI * (90 / ((float)180)) - angle2) / ((Math.Cos(Math.PI * (90 / ((float)180)) - angle2) * Math.Cos(Math.PI * (90 / ((float)180)) - angle2)) + (Math.Sin(Math.PI * (90 / ((float)180)) + angle2) * Math.Sin(Math.PI * (90 / ((float)180)) + angle2)));
                    zDirZ2 = Math.Sin(Math.PI * (90 / ((float)180)) - angle2) / ((Math.Cos(Math.PI * (90 / ((float)180)) - angle2) * Math.Cos(Math.PI * (90 / ((float)180)) - angle2)) + (Math.Sin(Math.PI * (90 / ((float)180)) - angle2) * Math.Sin(Math.PI * (90 / ((float)180)) - angle2)));
                }

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "StartStructure", new Position(0, 0, -Z3 / 2), new Vector(xDirX1, xDirY1, xDirZ1), new Vector(-zDirX1, -zDirY1, -zDirZ1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "MidStructure");

                Port port3 = new Port(OccurrenceConnection, part, "EndStructure");

                if ((port2 != null))
                {
                    port2.Origin = new Position(0, 0, L / 2);
                    port2.SetOrientation(new Vector(0, 0, 1), new Vector(0, 1, 0));
                    m_Symbolic.Outputs["Port2"] = port2;

                    port3.Origin = new Position(0, 0, L + Z4 / 2);
                    port3.SetOrientation(new Vector(-xDirX2, -xDirY2, -xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                    m_Symbolic.Outputs["Port3"] = port3;
                }
                else
                {
                    port3 = new Port(OccurrenceConnection, part, "EndStructure", new Position(0, 0, L + Z4 / 2), new Vector(-xDirX2, -xDirY2, -xDirZ2), new Vector(zDirX2, zDirY2, zDirZ2));
                    m_Symbolic.Outputs["Port2"] = port3;
                }

                Collection<Position> pointCollection = new Collection<Position>();

                pointCollection.Add(new Position(0, -thickness, -Z3 / 2 + Z1));
                pointCollection.Add(new Position(0, -thickness, L + (Z4 / 2) - Z2));
                pointCollection.Add(new Position(0, 0, L + (Z4 / 2)));
                pointCollection.Add(new Position(0, 0, -Z3 / 2));
                pointCollection.Add(new Position(0, -thickness, -Z3 / 2 + Z1));

                Vector projectionVector = new Vector(width, 0, 0);
                Projection3d body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY1"] = body;

                pointCollection = new Collection<Position>();

                pointCollection.Add(new Position(0, -thickness, -Z3 / 2 + Z1));
                pointCollection.Add(new Position(0, -thickness, L + (Z4 / 2) - Z2));
                pointCollection.Add(new Position(0, -depth, L - (Z4 / 2)));
                pointCollection.Add(new Position(0, -depth, Z3 / 2));
                pointCollection.Add(new Position(0, -thickness, -Z3 / 2 + Z1));

                projectionVector = new Vector(thickness, 0, 0);
                body = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY2"] = body;

            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_CUTBACK_L2"));
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
                double thicknessValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "THICKNESS")).PropValue;
                double widthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "WIDTH")).PropValue;
                double depthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "DEPTH")).PropValue;
                double angleValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_CUTBACK", "ANGLE")).PropValue;
                double angle2Value = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_CUTBACK", "ANGLE2")).PropValue;
                double LValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "L")).PropValue;
                string thickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, thicknessValue, UnitName.DISTANCE_INCH);
                string width = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, widthValue, UnitName.DISTANCE_INCH);
                string depth = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depthValue, UnitName.DISTANCE_INCH);
                string angle = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, angleValue, UnitName.ANGLE_DEGREE);
                string angle2 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Angle, angle2Value, UnitName.ANGLE_DEGREE);
                string L = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, LValue, UnitName.DISTANCE_INCH);
                String bomDescription = (String)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "BOM_DESC")).PropValue;

                double bomZ3, bomZ4, bomOrderLength;

                bomZ3 = depthValue / Math.Tan(angleValue);
                bomZ4 = depthValue / Math.Tan(angle2Value);

                bomOrderLength = LValue + Math.Abs(bomZ3 / 2) + Math.Abs(bomZ4 / 2);

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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_CUTBACK_L2"));
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
                Double weight, cogX, cogY, cogZ, area, volume, extraWebVolumeEnd, extraWebVolumeStart;
                const int getSteelDensityKGPerM = 7900;
                double thickness = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "THICKNESS")).PropValue;
                double width = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "WIDTH")).PropValue;
                double depth = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "DEPTH")).PropValue;
                double angle = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_CUTBACK", "ANGLE")).PropValue;
                double angle2 = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_CUTBACK", "ANGLE2")).PropValue;
                double L = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_GENERIC_L", "L")).PropValue;

                area = thickness * width + thickness * (depth - thickness);
                volume = area * L;

                if ((angle * 180 / Math.PI) > 0)
                {
                    extraWebVolumeStart = ((thickness * thickness) / (2 * Math.Tan(angle)) + thickness * ((depth / 2 - thickness) / Math.Tan(angle))) * (width - thickness);
                }
                else
                {
                    extraWebVolumeStart = 0;
                }
                if ((angle2 * 180 / Math.PI) > 0)
                {
                    extraWebVolumeEnd = ((thickness * thickness) / (2 * Math.Tan(angle2)) + thickness * ((depth / 2 - thickness) / Math.Tan(angle2))) * (width - thickness);
                }
                else
                {
                    extraWebVolumeEnd = 0;
                }

                volume = volume + extraWebVolumeStart + extraWebVolumeEnd;

                weight = volume * getSteelDensityKGPerM;

                cogX = width / 2;
                cogY = -depth / 2;
                cogZ = L / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_CUTBACK_L2"));
                }
            }
        }
    }
        #endregion
}
