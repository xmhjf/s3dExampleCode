﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   G4G_1400_U.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.VesselGuide
//   Author       : Hema
//   Creation Date: 04.April.2013
//   Description:   Converted HS_Assembly VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04.April.2013   Hema     Converted HS_Assembly VB Project to C# .Net 
//   22-Jan-2015     PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report 
//   05-Oct-2015     PVK      TR-CP-278849  Unable to place vessel guide support using ‘place by ref’ command  
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Content.Support.Symbols;

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
   
    public class VesselGuide : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.VesselGuide"
        //----------------------------------------------------------------------------------
        //Constants
        private const string CSECTION1 = "nCSection1";
        private const string CSECTION2 = "nCSection2";
        private const string LSECTION1 = "nLSection1";
        private const string LSECTION2 = "nLSection2";

        private string cSectionSize, lSectionSize;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    //Gets SupportHelper
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                    //Create a new collection to hold the caltalog parts
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    cSectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecSize")).PropValue;
                    lSectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecSize2")).PropValue;

                    parts.Add(new PartInfo(CSECTION1, "Utility_CUTBACK_C1_1")); //C Section
                    parts.Add(new PartInfo(CSECTION2, "Utility_CUTBACK_C1_1")); //C Section
                    parts.Add(new PartInfo(LSECTION1, lSectionSize)); //L Section
                    parts.Add(new PartInfo(LSECTION2, lSectionSize)); //L Section

                    //Return the collection of Catalog Parts
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                //Get interface for accessing items on the collection of Part Occurences
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                double  pipeRadius, L, cWidth, plateW, gap, vesselDiameter = 0, cDepth, cWebT, cFlangeT, lWidth, lFlangeT, lWebT, lDepth, lSectionL, Z1, extraLengthOutside;
                double Z2, extraLengthInside, cutBackAngle, cSectionL, curvedPlateAngle, curvedPlateLocAngle, plateLocator = 0.1, curvedPlateLoc1, curvedPlateLoc2;
                string cSectionStandard, cSectionSize, lSectionStandard,lSectionSize, cBom;

                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                pipeRadius = (routeInfo.OutsideDiameter) / 2;

                //Get the attributes from the assy
                L = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                cSectionStandard = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecStd")).PropValue;
                cSectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecSize")).PropValue;
                lSectionStandard = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecStd2")).PropValue;
                lSectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrVesselGuide", "SecSize2")).PropValue;
                plateW = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrVesselGuide", "PlateW")).PropValue;
                gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrVesselGuide", "Gap")).PropValue;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByReference && SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.GenericSurface)
                    {
                        vesselDiameter = SupportingHelper.SupportingObjectInfo(1).Diameter;
                    }
                }

                if (HgrCompareDoubleService.cmpdbl(vesselDiameter , 0)==true)
                    vesselDiameter = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrVesselDia", "VesselDia")).PropValue;

                support.SetPropertyValue(vesselDiameter, "IJOAHgrVesselDia", "VesselDia");

                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================

                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();

                //Get the steel Data
                CrossSection crossSection = catalogStructHelper.GetCrossSection(cSectionStandard, cSectionSize.Substring(10));
                cWidth = crossSection.Width;
                cFlangeT = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                cWebT = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tw")).PropValue;
                cDepth = crossSection.Depth;

                //Get the L Section Data
                CrossSection crossSection1 = catalogStructHelper.GetCrossSection(lSectionStandard, lSectionSize.Substring(10));
                lWidth = crossSection1.Width;
                lFlangeT = (double)((PropertyValueDouble)crossSection1.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                lWebT = (double)((PropertyValueDouble)crossSection1.GetPropertyValue("IStructFlangedSectionDimensions", "tw")).PropValue;
                lDepth = crossSection1.Depth;

                lSectionL = pipeRadius * 2 + 2 * cWidth + 2 * gap;
                Z1 = Math.Sqrt((vesselDiameter / 2 * vesselDiameter / 2) - (lSectionL / 2 * lSectionL / 2));
                extraLengthOutside = vesselDiameter / 2 - Z1;
                Z2 = Math.Sqrt((vesselDiameter / 2 * vesselDiameter / 2) - ((lSectionL / 2 - cWidth) * (lSectionL / 2 - cWidth)));

                extraLengthInside = vesselDiameter / 2 - Z2;
                cutBackAngle = Math.Atan((extraLengthOutside - extraLengthInside) / cWidth);
                cSectionL = L + pipeRadius + gap + lWidth + extraLengthInside + (extraLengthOutside - extraLengthInside) / 2;  //original

                curvedPlateAngle = (plateW / (Math.Atan(1) * 4 * vesselDiameter)) * 360;
                curvedPlateLocAngle = Math.Asin((gap + pipeRadius + cWidth / 2) / (vesselDiameter / 2));
                curvedPlateLoc1 = Math.Cos(curvedPlateLocAngle) * plateLocator;
                curvedPlateLoc2 = Math.Sin(curvedPlateLocAngle) * plateLocator;

                cBom = "AISC " + (cSectionSize.Substring(10)).Trim() + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(support.GetPropertyValue("IJOAHgrVesselDia", "VesselDia").PropertyInfo.UOMType, cSectionL + (extraLengthOutside - extraLengthInside) / 2, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.DISTANCE_INCH_SYMBOL);

                //Set our C Props
                (componentDictionary[CSECTION1]).SetPropertyValue(cSectionL, "IJOAHgrUtility_GENERIC_W", "L");
                (componentDictionary[CSECTION1]).SetPropertyValue(cWidth, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                (componentDictionary[CSECTION1]).SetPropertyValue(cDepth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                (componentDictionary[CSECTION1]).SetPropertyValue(cFlangeT, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                (componentDictionary[CSECTION1]).SetPropertyValue(cWebT, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                (componentDictionary[CSECTION1]).SetPropertyValue(90 * Math.Atan(1) * 4 / 180, "IJOAHgrUtility_CUTBACK", "ANGLE");
                (componentDictionary[CSECTION1]).SetPropertyValue(90 * Math.Atan(1) * 4 / 180 + cutBackAngle, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                (componentDictionary[CSECTION1]).SetPropertyValue(cBom, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");

                (componentDictionary[CSECTION2]).SetPropertyValue(cSectionL, "IJOAHgrUtility_GENERIC_W", "L");
                (componentDictionary[CSECTION2]).SetPropertyValue(cWidth, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                (componentDictionary[CSECTION2]).SetPropertyValue(cDepth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                (componentDictionary[CSECTION2]).SetPropertyValue(cFlangeT, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                (componentDictionary[CSECTION2]).SetPropertyValue(cWebT, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                (componentDictionary[CSECTION2]).SetPropertyValue(90 * Math.Atan(1) * 4 / 180, "IJOAHgrUtility_CUTBACK", "ANGLE");
                (componentDictionary[CSECTION2]).SetPropertyValue(90 * Math.Atan(1) * 4 / 180 + cutBackAngle, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                (componentDictionary[LSECTION1]).SetPropertyValue(lSectionL, "IJUAHgrOccLength", "Length");
                (componentDictionary[LSECTION2]).SetPropertyValue(lSectionL, "IJUAHgrOccLength", "Length");

                if (Configuration == 1)
                {
                    //Add a Joint between the Second L and the second C
                    JointHelper.CreateRigidJoint(LSECTION2, "BeginCap", CSECTION2, "StartStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, -cWidth + lSectionL, -cDepth / 2, lWidth);
                    //Add a Joint between the First L and the second C
                    JointHelper.CreateRigidJoint(LSECTION1, "BeginCap", CSECTION2, "StartStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, cWidth, -cDepth / 2, -pipeRadius * 2 - gap * 2 - lWidth);
                }
                else if (Configuration == 2)
                {
                    //Add a Joint between the Second L and the second C
                    JointHelper.CreateRigidJoint(LSECTION2, "BeginCap", CSECTION2, "StartStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -cWidth + lSectionL, -cDepth / 2, -pipeRadius * 2 - gap * 2 - lWidth);
                    //Add a Joint between the First L and the second C
                    JointHelper.CreateRigidJoint(LSECTION1, "BeginCap", CSECTION2, "StartStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, cWidth, -cDepth / 2, lWidth);
                }

                //Add a Joint between the First C and the route
                JointHelper.CreateRigidJoint(CSECTION1, "StartStructure", "-1", "Route", Plane.XY, Plane.NegativeZX, Axis.X, Axis.Z, pipeRadius + gap, 0, pipeRadius + gap + lWidth);

                //Add a Joint between the Second C and the route
                JointHelper.CreateRigidJoint(CSECTION2, "StartStructure", "-1", "Route", Plane.XY, Plane.ZX, Axis.X, Axis.Z, pipeRadius + gap, 0, pipeRadius + gap + lWidth);
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
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

                    routeConnections.Add(new ConnectionInfo(CSECTION1, 1));  //partindex, routeindex

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
                //We are not connecting to any structure so we have nothing to return
                Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                return structConnections;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get MirroredConfiguration
        //-----------------------------------------------------------------------------------
        public override int MirroredConfiguration(int CurrentMirrorToggleValue, MirrorPlane eMirrorPlane)
        {
            try
            {
                //Initialize the MirrorToggle to CurrentToggle
                return CurrentMirrorToggleValue;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Mirrored Configuration." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}



