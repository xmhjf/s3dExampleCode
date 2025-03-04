//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Fl_5BSL1.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.Fl_5BSL1
//   Author       : Hema
//   Creation Date: 29-Aug-013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//     29-Aug-2013  Hema CR-CP-224478 Convert FlSample_Supports to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;

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

    public class Fl_5BSL1 : CustomSupportDefinition
    {
        private const double threeInSched40ThicknessInM = 0.00549;
        private const double sixInSched40ThicknessInM = 0.00711;
        Double pipeDiameterInch;
        private const string PLATE1 = "PLATE1";
        private const string PLATE2 = "PLATE2";
        private const string CYLINDER = "CYLINDER";
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
                    string cylinder = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Cylinder", "Cylinder_1")).PropValue;

                    if (3 <= pipeDiameterInch && pipeDiameterInch <= 10)
                    {
                        parts.Add(new PartInfo(PLATE1, plate1));
                        parts.Add(new PartInfo(PLATE2, plate2));
                        parts.Add(new PartInfo(CYLINDER, cylinder));
                    }
                    else
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR:" + "5BSL1 can only be placed from 3in to 10in pipe.", "", "Fl_5BSL1", 67);
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

                Double plate1Length, plate1Height, plate1Thickness, plate2Length, plate2Height, plate2Thickness, plate1M, cylinderLength, cylinderRadius;
                string plate1BOM, plate2BOM, cylinderBOM;
                plate1Length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_1")).PropValue;
                plate1Height = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_1")).PropValue;
                plate1Thickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_1")).PropValue;
                plate1M = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateM_1")).PropValue;
                plate1BOM = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateBomDesc_1")).PropValue;
                plate2Length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_2")).PropValue;
                plate2Height = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_2")).PropValue;
                plate2Thickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_2")).PropValue;
                plate2BOM = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateBomDesc_2")).PropValue;
                cylinderLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Cylinder", "L_1")).PropValue;
                cylinderRadius = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Cylinder", "RADIUS_1")).PropValue;
                cylinderRadius = cylinderRadius / 2;
                cylinderBOM = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Cylinder", "CylinderBomDesc_1")).PropValue;

                double flangeBoltCircle, flangeInsideRadius, curveAngle1, curveAngle2;
                flangeBoltCircle = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrFlange", "BoltCircle")).PropValue;
                flangeInsideRadius = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrFlange", "InsideRadius")).PropValue;
                curveAngle1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "CurveAngle_1")).PropValue;
                curveAngle2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "CurveAngle_2")).PropValue;

                double plate1LOffset = flangeBoltCircle / 2 * Math.Cos((curveAngle1 / 2) * Math.PI / 180) - flangeInsideRadius * Math.Cos((curveAngle1 / 2) * Math.PI / 180 - flangeInsideRadius * Math.Cos(curveAngle1 / 2) * Math.PI / 180 + (curveAngle2) * Math.PI / 180);
                Double longElbow = pipeDiameterInch * 1.5 * 0.0254;

                double stanchionWeight, fullVolume, hollowVolume, stanchionThickness = 0, routeStructureDis;
                const double getSteelDensityKGPerM = 7900;

                routeStructureDis = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                componentDictionary[PLATE1].SetPropertyValue(plate1Thickness, "IJOAHgrUtility_PLATE", "DEPTH");
                componentDictionary[PLATE1].SetPropertyValue(plate1Height, "IJOAHgrUtility_PLATE", "WIDTH");
                componentDictionary[PLATE1].SetPropertyValue(plate1Length + plate1LOffset, "IJOAHgrUtility_PLATE", "THICKNESS");
                componentDictionary[PLATE1].SetPropertyValue(plate1BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                componentDictionary[PLATE2].SetPropertyValue(plate2Thickness, "IJOAHgrUtility_PLATE", "DEPTH");
                componentDictionary[PLATE2].SetPropertyValue(plate2Height, "IJOAHgrUtility_PLATE", "WIDTH");
                componentDictionary[PLATE2].SetPropertyValue(plate2Length, "IJOAHgrUtility_PLATE", "THICKNESS");
                componentDictionary[PLATE2].SetPropertyValue(plate2BOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                cylinderLength = routeStructureDis - plate1Thickness - plate2Thickness;
                if (cylinderLength > (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Cylinder", "L_1")).PropValue)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING:" + "STUB LENGTH CAN NOT BE LARGER THAN 1200mm.", "", "Fl_5BSL1", 136);
                if (ndpPipeDiameter * 1000 < 4)
                    stanchionThickness = threeInSched40ThicknessInM;
                else
                    stanchionThickness = sixInSched40ThicknessInM;

                fullVolume = Math.Atan(1) * 4.0 * cylinderRadius * cylinderRadius * cylinderLength;
                hollowVolume = Math.Atan(1) * 4.0 * Math.Pow((cylinderRadius - stanchionThickness), 2) * cylinderLength;
                stanchionWeight = (fullVolume - hollowVolume) * getSteelDensityKGPerM;

                componentDictionary[CYLINDER].SetPropertyValue(cylinderLength, "IJOAHgrUtility_Cylinder", "L");
                componentDictionary[CYLINDER].SetPropertyValue(cylinderRadius, "IJOAHgrUtility_Cylinder", "RADIUS");
                componentDictionary[CYLINDER].SetPropertyValue(cylinderBOM, "IJOAHgrUtilMMBomDesc", "InputBomDesc");
                componentDictionary[CYLINDER].SetPropertyValue(stanchionWeight, "IJOAHgrFlSampleWeight", "InputWeight");

                if (Configuration == 1)
                    JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY, plate1Thickness / 2, flangeInsideRadius * Math.Cos((curveAngle2) * Math.PI / 180), 0);
                else if (Configuration == 2)
                    JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, plate1Thickness / 2, -flangeInsideRadius * Math.Cos((curveAngle2) * Math.PI / 180), 0);
                else if (Configuration == 3)
                    JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, plate1Thickness / 2, 0, flangeInsideRadius * Math.Cos((curveAngle2) * Math.PI / 180));
                else if (Configuration == 4)
                    JointHelper.CreateRigidJoint("-1", "Route", PLATE1, "TopStructure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, plate1Thickness / 2, 0, -flangeInsideRadius * Math.Cos((curveAngle2) * Math.PI / 180));

                JointHelper.CreateRigidJoint(PLATE1, "BotStructure", CYLINDER, "StartOther", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, -plate1M, 0, -plate1Thickness / 2);
                JointHelper.CreateRigidJoint(CYLINDER, "EndOther", PLATE2, "TopStructure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY, plate2Thickness / 2, -plate2Height / 2, 0);
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
                return 4;
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

                    routeConnections.Add(new ConnectionInfo(PLATE1, 1));// partindex, routeindex

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
    }
}
