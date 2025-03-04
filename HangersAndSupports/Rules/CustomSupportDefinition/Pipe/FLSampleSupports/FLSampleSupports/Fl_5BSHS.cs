//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Fl_5BSHS.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.Fl_5BSHS
//   Author       : Hema
//   Creation Date: 29-Aug-013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   29-Aug-2013  Hema CR-CP-224478 Convert FlSample_Supports to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Exceptions;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class Fl_5BSHS : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string PLATE1 = "PLATE1";
        private const string PLATE2 = "PLATE2";
        private const string PLATEBASE = "PLATEBASE";
        private const string FOOTING = "FOOTING";
        private const string PLATE3 = "PLATE3";
        private const string SCREWEDFLANGE = "SCREWEDFLANGE";

        private const double twoinSched40ThicknessInM = 0.00391;
        private const double threeInSched40ThicknessInM = 0.00549;
        private const double fourInSchedSTDThicknessInM = 0.00602;
        Double pipeDiameterInch;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    pipeDiameterInch = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                    string plate1 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "Plate_1")).PropValue;
                    string plate2 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "Plate_2")).PropValue;
                    string footing = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Cylinder", "Cylinder_1")).PropValue;
                    string screwedFlange = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Cylinder", "Cylinder_1")).PropValue;

                    if (3 <= pipeDiameterInch && pipeDiameterInch <= 6)
                    {
                        parts.Add(new PartInfo(PLATE1, plate1));
                        parts.Add(new PartInfo(PLATE2, plate1));
                        parts.Add(new PartInfo(PLATEBASE, plate2));
                        parts.Add(new PartInfo(SCREWEDFLANGE, screwedFlange));
                        parts.Add(new PartInfo(FOOTING, footing));
                    }
                    else if (8 <= pipeDiameterInch && pipeDiameterInch <= 20)
                    {
                        parts.Add(new PartInfo(PLATE1, plate1));
                        parts.Add(new PartInfo(PLATE2, plate1));
                        parts.Add(new PartInfo(PLATEBASE, plate2));
                        parts.Add(new PartInfo(SCREWEDFLANGE, screwedFlange));
                        parts.Add(new PartInfo(FOOTING, footing));
                        parts.Add(new PartInfo(PLATE3, plate1));
                    }
                    else
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR:" + "5BSHS can only be placed on 3in to 20in pipes.", "", "Fl_5BSHS", 85);
                        return null;
                    }
                    // Return the collection of Catalog Parts
                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double ndpPipeDiameter = pipeInfo.OutsideDiameter;

                Double plate1Length, plate1Height, plate1Thickness, plate2Length, plate2Height, plate2Thickness, screwedFlangeL = 0.0254, pipeHeight, padHeight;

                plate1Length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_1")).PropValue;
                plate1Height = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_1")).PropValue;
                plate1Thickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_1")).PropValue;
                plate2Length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_2")).PropValue;
                plate2Height = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_2")).PropValue;
                plate2Thickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_2")).PropValue;
                pipeHeight = (Double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrPipeHieght", "PipeHieght")).PropValue;
                padHeight = (Double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrPadHieght", "PadHieght")).PropValue;
                int pipeMaterialCodelist = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrPipeMaterial", "PipeMaterial")).PropValue;
                string plate1BOM = string.Empty, plate2BOM = string.Empty;
                PropertyValueCodelist bomList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrPipeMatPlateXY_5BSHS", "PlateXY_5BSHS");
                plate1BOM = bomList.PropertyInfo.CodeListInfo.GetCodelistItem(bomList.PropValue).DisplayName;
                plate2BOM = bomList.PropertyInfo.CodeListInfo.GetCodelistItem(bomList.PropValue).ShortDisplayName;
                Double longElbow = pipeDiameterInch * 1.5 * 0.0254;
                Double plate12L = 0;
                try
                {
                    if (RefPortHelper.PortLCS("TurnRef") != null)
                    {
                        plate12L = longElbow + ndpPipeDiameter / 2 - 0.04;
                        plate12L = plate12L + pipeHeight + pipeInfo.InsulationThickness;
                        plate12L = plate12L - plate2Thickness;
                    }
                }
                catch (SupportInvalidArgumentException)
                {
                    plate12L = ndpPipeDiameter / 2 - 0.04;
                    plate12L = plate12L + pipeHeight + pipeInfo.InsulationThickness;
                    plate12L = plate12L - plate2Thickness;
                }
                double stanchionWeight, fullVolume, hollowVolume, roundPlateFullVolume, roundPlateHoleVolume, roundPlateWeight, stanchionThickness = 0, stanctionRadius = 0, standL = 0, routeStructureDis, footingRadius, screwedFlangeRadius;
                const double getSteelDensityKGPerM = 7900;

                routeStructureDis = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                footingRadius = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrSPL_Cylinder", "FootingRadius")).PropValue;
                screwedFlangeRadius = 2 * footingRadius;

                IPart part = support.SupportDefinition;
                string partNumber = part.PartNumber;
                if (pipeDiameterInch <= 6 && partNumber.Equals("5BSHS_2"))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "For 5BSHS_2, Pipe Size should be at least 8in.", "", "Fl_5BSHS", 155);
                    return;
                }
                if (pipeDiameterInch >= 8 && partNumber.Equals("5BSHS_1"))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "For 5BSHS_1, Pipe Size should be at least 6in.", "", "Fl_5BSHS", 160);
                    return;
                }
                if (3 <= pipeDiameterInch && pipeDiameterInch <= 6)
                {
                    componentDictionary[PLATE1].SetPropertyValue(plate1Thickness, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE1].SetPropertyValue(ndpPipeDiameter, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE1].SetPropertyValue(plate12L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE1].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    componentDictionary[PLATE2].SetPropertyValue(plate1Thickness, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE2].SetPropertyValue(ndpPipeDiameter, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE2].SetPropertyValue(plate12L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE2].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    componentDictionary[PLATEBASE].SetPropertyValue(plate2Length, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2Height, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2Thickness, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");
                    if (3 <= pipeDiameterInch && pipeDiameterInch <= 4)
                    {
                        stanchionThickness = twoinSched40ThicknessInM;
                        stanctionRadius = 2 * 0.0254;
                    }
                    else
                    {
                        stanchionThickness = threeInSched40ThicknessInM;
                        stanctionRadius = 3 * 0.0254;
                    }
                    if (3 <= pipeDiameterInch && pipeDiameterInch <= 4)
                    {
                        componentDictionary[FOOTING].SetPropertyValue(stanctionRadius / 2, "IJOAHgrUtility_Cylinder", "RADIUS");
                        componentDictionary[FOOTING].SetPropertyValue("2in SCH. 80 C.S. SMLS. PIPE", "IJOAHgrUtilMMBomDesc", "InputBomDesc");
                        componentDictionary[SCREWEDFLANGE].SetPropertyValue(stanctionRadius, "IJOAHgrUtility_Cylinder", "RADIUS");
                        componentDictionary[SCREWEDFLANGE].SetPropertyValue("150# Screwed Flange for 2in SCH. 80 Pipe", "IJOAHgrUtilMMBomDesc", "InputBomDesc");
                    }
                    else
                    {
                        componentDictionary[FOOTING].SetPropertyValue(stanctionRadius / 2, "IJOAHgrUtility_Cylinder", "RADIUS");
                        componentDictionary[FOOTING].SetPropertyValue("3in SCH. 40 C.S. SMLS. PIPE", "IJOAHgrUtilMMBomDesc", "InputBomDesc");
                        componentDictionary[SCREWEDFLANGE].SetPropertyValue(stanctionRadius, "IJOAHgrUtility_Cylinder", "RADIUS");
                        componentDictionary[SCREWEDFLANGE].SetPropertyValue("150# Screwed Flange for 3in SCH. 40 Pipe", "IJOAHgrUtilMMBomDesc", "InputBomDesc");
                    }
                    try
                    {
                        if (RefPortHelper.PortLCS("TurnRef") != null)
                        {
                            standL = routeStructureDis - padHeight + longElbow - 0.04 - plate12L - plate2Thickness;
                            if (Configuration == 1)
                                JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, -longElbow + 0.04, 0, -longElbow);
                            else
                                JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, -longElbow + 0.04, 0, longElbow);
                        }
                    }
                    catch (SupportInvalidArgumentException)
                    {
                        standL = routeStructureDis - padHeight - 0.04 - plate12L - plate2Thickness;
                        JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.04, 0, 0);
                    }
                    componentDictionary[FOOTING].SetPropertyValue(standL, "IJOAHgrUtility_Cylinder", "L");
                    componentDictionary[SCREWEDFLANGE].SetPropertyValue(screwedFlangeL, "IJOAHgrUtility_Cylinder", "L");

                    JointHelper.CreateRigidJoint(PLATE1, "TopStructure", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                    JointHelper.CreateRigidJoint(PLATE1, "BotStructure", PLATEBASE, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(PLATEBASE, "BotStructure", FOOTING, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(FOOTING, "EndOther", SCREWEDFLANGE, "EndOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else if (8 <= pipeDiameterInch && pipeDiameterInch <= 20)
                {
                    componentDictionary[PLATE1].SetPropertyValue(plate1Thickness, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE1].SetPropertyValue(ndpPipeDiameter / 2, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE1].SetPropertyValue(plate12L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE1].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    componentDictionary[PLATE2].SetPropertyValue(plate1Thickness, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE2].SetPropertyValue(ndpPipeDiameter / 2, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE2].SetPropertyValue(plate12L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE2].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    componentDictionary[PLATEBASE].SetPropertyValue(plate2Length, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2Height, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate2Thickness, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATEBASE].SetPropertyValue(plate1BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    stanchionThickness = fourInSchedSTDThicknessInM;
                    stanctionRadius = 4 * 0.0254;

                    componentDictionary[FOOTING].SetPropertyValue(stanctionRadius / 2, "IJOAHgrUtility_Cylinder", "RADIUS");
                    componentDictionary[SCREWEDFLANGE].SetPropertyValue(stanctionRadius, "IJOAHgrUtility_Cylinder", "RADIUS");
                    componentDictionary[FOOTING].SetPropertyValue("4in - STD", "IJOAHgrUtilMMBomDesc", "InputBomDesc");
                    componentDictionary[SCREWEDFLANGE].SetPropertyValue("150# Screwed Flange for 4in - STD", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                    double plate3L = 0;
                    try
                    {
                        if (RefPortHelper.PortLCS("TurnRef") != null)
                        {
                            plate3L = longElbow + ndpPipeDiameter / 2 - Math.Sqrt((longElbow + ndpPipeDiameter / 2) * (longElbow + ndpPipeDiameter / 2) - (longElbow + ndpPipeDiameter / 2 - ndpPipeDiameter / 4) * (longElbow + ndpPipeDiameter / 2 - ndpPipeDiameter / 4));
                            plate3L = plate3L + pipeHeight + pipeInfo.InsulationThickness;
                            plate3L = plate3L - 0.0254 / 2;
                            standL = routeStructureDis - padHeight + longElbow - 0.04 - plate12L - plate2Thickness;
                        }
                    }
                    catch (SupportInvalidArgumentException)
                    {
                        plate3L = plate12L - 0.045; //this is so the plates do not dig into the pipe
                        componentDictionary[PLATE2].SetPropertyValue(plate12L - 0.045, "IJOAHgrUtility_PLATE", "THICKNESS");
                        componentDictionary[PLATE1].SetPropertyValue(plate12L - 0.045, "IJOAHgrUtility_PLATE", "THICKNESS");
                        standL = routeStructureDis - padHeight - 0.04 - plate12L - plate2Thickness;
                    }
                    componentDictionary[FOOTING].SetPropertyValue(standL, "IJOAHgrUtility_Cylinder", "L");
                    componentDictionary[SCREWEDFLANGE].SetPropertyValue(screwedFlangeL, "IJOAHgrUtility_Cylinder", "L");

                    componentDictionary[PLATE3].SetPropertyValue(plate1Thickness, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[PLATE3].SetPropertyValue(0.13 - plate1Thickness * 2, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[PLATE3].SetPropertyValue(plate3L, "IJOAHgrUtility_PLATE", "THICKNESS");
                    componentDictionary[PLATE3].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");
                    try
                    {
                        if (RefPortHelper.PortLCS("TurnRef") != null)
                        {
                            if (Configuration == 1)
                                JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, -longElbow + 0.04, 0.065 - plate1Thickness / 2, -longElbow - ndpPipeDiameter * 1 / 4);
                            else
                                JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, -longElbow + 0.04, 0.065 - plate1Thickness / 2, longElbow + ndpPipeDiameter * 1 / 4);
                        }
                    }
                    catch (SupportInvalidArgumentException)
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0.04 + 0.045, 0.065 - plate1Thickness / 2, 0);
                    }
                    JointHelper.CreateRigidJoint(PLATE1, "TopStructure", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -(0.13 - plate1Thickness));
                    JointHelper.CreateRigidJoint(PLATE1, "BotStructure", PLATEBASE, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -(0.13 - plate1Thickness) / 2);
                    JointHelper.CreateRigidJoint(PLATEBASE, "TopStructure", PLATE3, "BotStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                    JointHelper.CreateRigidJoint(PLATEBASE, "BotStructure", FOOTING, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(FOOTING, "EndOther", SCREWEDFLANGE, "EndOther", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                fullVolume = Math.Atan(1) * 4.0 * (stanctionRadius / 2 * stanctionRadius / 2) * standL;
                hollowVolume = Math.Atan(1) * 4.0 * Math.Pow(((stanctionRadius / 2) - stanchionThickness), 2) * standL;
                stanchionWeight = (fullVolume - hollowVolume) * getSteelDensityKGPerM;
                componentDictionary[FOOTING].SetPropertyValue(stanchionWeight, "IJOAHgrFlSampleWeight", "InputWeight");

                roundPlateFullVolume = Math.Atan(1) * 4.0 * (stanctionRadius * stanctionRadius) * screwedFlangeL;
                roundPlateHoleVolume = Math.Atan(1) * 4.0 * (stanctionRadius / 2 * stanctionRadius / 2) * screwedFlangeL;
                roundPlateWeight = (roundPlateFullVolume - roundPlateHoleVolume) * getSteelDensityKGPerM;
                componentDictionary[SCREWEDFLANGE].SetPropertyValue(roundPlateWeight, "IJOAHgrFlSampleWeight", "InputWeight");
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 2;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    routeConnections.Add(new ConnectionInfo(PLATE2, 1));// partindex, routeindex

                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(PLATE2, 1));// partindex, Structureindex    

                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = string.Empty;
            try
            {
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Double pipeDiameterInch = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);
                if (3 <= pipeDiameterInch && pipeDiameterInch <= 6)
                    bomDescription = "5BSHS1";
                else
                    bomDescription = "5BSHS2";
                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOMDescription." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}
