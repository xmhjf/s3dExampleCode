﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   TypeU.cs
//   PipeHgrAssemblies2,Ingr.SP3D.Content.Support.Rules.BeamU_Bracket
//   Author       :Vinay
//   Creation Date:27-11-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  21.Mar.2016     PVK      TR-CP-288920	Issues found in HS_Assembly_V2
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle.Hidden;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class BeamU_Bracket : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM1 = "HGRBEAM1";
        private const string HGRBEAM2 = "HGRBEAM2";
        private const string HGRBEAM3 = "HGRBEAM3";
        private const string EXTERNALBRACKET1 = "EXTERNALBRACKET1";
        private const string EXTERNALBRACKET2 = "EXTERNALBRACKET2";
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
        string[] bracketPartKeys;
        bool leftPad, rightPad;
        double d, gap;
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
                    object[] attributeCollection = SupportAssemblyServices.GetPipeStructuralASMAttributes(this);

                    //Get the attributes from assembly
                    d = (double)attributeCollection[0];
                    gap = (double)attributeCollection[1];
                    leftPad = (bool)attributeCollection[2];
                    rightPad = (bool)attributeCollection[3];

                    int noOfPipes = SupportHelper.SupportedObjects.Count;
                    bracketPartKeys = new string[3 * noOfPipes];

                    string sectionPartClass = (string)((PropertyValueString)support.SupportDefinition.GetPropertyValue("IJUAHSA_SecPartClass", "SecPartClass")).PropValue;

                    parts.Add(new PartInfo(HGRBEAM1, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));
                    parts.Add(new PartInfo(HGRBEAM2, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));
                    parts.Add(new PartInfo(HGRBEAM3, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));
                    parts.Add(new PartInfo(EXTERNALBRACKET1, "S3Dhs_TriangularBracket_1"));
                    parts.Add(new PartInfo(EXTERNALBRACKET2, "S3Dhs_TriangularBracket_1"));

                    for (int index = 0; index < 3 * noOfPipes; index++)
                    {
                        bracketPartKeys[index] = "INTERNALBRACKET" + (index + 1);
                        parts.Add(new PartInfo(bracketPartKeys[index], "S3Dhs_TriangularBracket_1"));
                    }
                    if (rightPad)
                        parts.Add(new PartInfo(RIGHTPAD, SupportAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));
                    if (leftPad)
                        parts.Add(new PartInfo(LEFTPAD, SupportAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));
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
                return 1;
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
                //load standard bounding box.
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;
                string refPlane = string.Empty;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;
                double[] boxOffset = SupportAssemblyServices.GetBoundaryObjectDimension(this, boundingBox), lugOffset = new double[2];
                //get the offset value based on LoadFactor
                lugOffset[0] = d ;
                lugOffset[1] = d ;
                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = SupportAssemblyServices.GetIsLugEndOffsetApplied(this);

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                
                string[] structPort = new string[2];
                structPort = SupportAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0], rightStructPort = structPort[1];
                double rightPadLength = 0, rightPadWidth = 0, leftPadLength = 0, leftPadWidth = 0;
                Part padPart = null;
                string rightPadPartnumber = "", leftPadPartnumber = "";
                string rightPadtype = "", leftPadtype = "";
                if (rightPad)
                {
                    padPart = (Part)componentDictionary[RIGHTPAD].GetRelationship("madeFrom", "part").TargetObjects[0];
                    rightPadPartnumber = padPart.PartNumber;
                    rightPadtype = SupportAssemblyServices.GetPadPartType(this, rightPadPartnumber);
                    if (rightPadtype == "Triangular")
                    {
                        rightPadLength = (double)((PropertyValueDouble)componentDictionary[RIGHTPAD].GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                        rightPadWidth = (double)((PropertyValueDouble)componentDictionary[RIGHTPAD].GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                        componentDictionary[RIGHTPAD].SetPropertyValue(-rightPadLength / 6.0, "IJUAhsStandardPort1", "P1xOffset");
                        componentDictionary[RIGHTPAD].SetPropertyValue(-rightPadWidth / 6.0, "IJUAhsStandardPort1", "P1yOffset");
                        componentDictionary[RIGHTPAD].SetPropertyValue(-rightPadLength / 6.0, "IJUAhsStandardPort2", "P2xOffset");
                        componentDictionary[RIGHTPAD].SetPropertyValue(-rightPadWidth / 6.0, "IJUAhsStandardPort2", "P2yOffset");
                    }

                }
                if (leftPad)
                {
                    padPart = (Part)componentDictionary[LEFTPAD].GetRelationship("madeFrom", "part").TargetObjects[0];
                    leftPadPartnumber = padPart.PartNumber;
                    leftPadtype = SupportAssemblyServices.GetPadPartType(this, leftPadPartnumber);
                    if (leftPadtype == "Triangular")
                    {
                        leftPadLength = (double)((PropertyValueDouble)componentDictionary[LEFTPAD].GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                        leftPadWidth = (double)((PropertyValueDouble)componentDictionary[LEFTPAD].GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                        componentDictionary[LEFTPAD].SetPropertyValue(-leftPadLength / 6.0, "IJUAhsStandardPort1", "P1xOffset");
                        componentDictionary[LEFTPAD].SetPropertyValue(-leftPadWidth / 6.0, "IJUAhsStandardPort1", "P1yOffset");
                        componentDictionary[LEFTPAD].SetPropertyValue(-leftPadLength / 6.0, "IJUAhsStandardPort2", "P2xOffset");
                        componentDictionary[LEFTPAD].SetPropertyValue(-leftPadWidth / 6.0, "IJUAhsStandardPort2", "P2yOffset");
                    }
                }

                //Create Joints               
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + gap, -lugOffset[0]);
                else
                    if (isOffsetApplied[0])
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + gap, -lugOffset[0], 0);
                    else
                        JointHelper.CreatePrismaticJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, boundingBoxHeight + gap, 0);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[1]);
                else
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[1]);

                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0, 0, 0);
                JointHelper.CreateRigidJoint(HGRBEAM3, "BeginCap", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, 0);
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                JointHelper.CreatePrismaticJoint(HGRBEAM3, "BeginCap", HGRBEAM3, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    if (leftPad)
                    {
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        JointHelper.CreatePlanarJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "Port2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                    else
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                else
                    if (leftPad)
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, 0, 0);

                        JointHelper.CreatePlanarJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "Port2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }

                    else
                        if (isOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, 0, 0);

                if (rightPad)
                {
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", rightStructPort, RIGHTPAD, "Port2", Plane.XY, Plane.NegativeXY, 0);
                    else
                        JointHelper.CreatePlanarSlotJoint("-1", rightStructPort, RIGHTPAD, "Port1", Plane.XY, Plane.NegativeXY, Axis.X, 0, 0);

                    JointHelper.CreatePlanarJoint(RIGHTPAD, "Port1", HGRBEAM3, "EndCap", Plane.XY, Plane.XY, 0);
                    JointHelper.CreatePrismaticJoint(RIGHTPAD, "Port1", HGRBEAM3, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                else
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", rightStructPort, HGRBEAM3, "EndCap", Plane.XY, Plane.NegativeXY, 0);
                    else
                        JointHelper.CreatePlanarSlotJoint("-1", rightStructPort, HGRBEAM3, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, 0, 0);

                double sectionWidth = 0.0, sectionDepth = 0.0, flangeThickness = 0.0;
                BusinessObject sectionPart = componentDictionary[HGRBEAM1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth               
                sectionWidth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                sectionDepth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                flangeThickness = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                //externBracket
                componentDictionary[EXTERNALBRACKET1].SetPropertyValue(sectionWidth, "IJUAhsWidth1", "Width1");
                componentDictionary[EXTERNALBRACKET1].SetPropertyValue(sectionWidth, "IJUAhsLength1", "Length1");
                componentDictionary[EXTERNALBRACKET1].SetPropertyValue(flangeThickness, "IJUAhsThickness1", "Thickness1");
                componentDictionary[EXTERNALBRACKET2].SetPropertyValue(sectionWidth, "IJUAhsWidth1", "Width1");
                componentDictionary[EXTERNALBRACKET2].SetPropertyValue(sectionWidth, "IJUAhsLength1", "Length1");
                componentDictionary[EXTERNALBRACKET2].SetPropertyValue(flangeThickness, "IJUAhsThickness1", "Thickness1");

                componentDictionary[EXTERNALBRACKET1].SetPropertyValue(-sectionWidth/ 2.0, "IJUAhsStandardPort1", "P1xOffset");
                componentDictionary[EXTERNALBRACKET1].SetPropertyValue(-sectionWidth / 2.0, "IJUAhsStandardPort1", "P1yOffset");
                componentDictionary[EXTERNALBRACKET2].SetPropertyValue(-sectionWidth / 2.0, "IJUAhsStandardPort1", "P1xOffset");
                componentDictionary[EXTERNALBRACKET2].SetPropertyValue(-sectionWidth / 2.0, "IJUAhsStandardPort1", "P1yOffset");

                JointHelper.CreatePrismaticJoint(HGRBEAM1, "EndCap", EXTERNALBRACKET1, "Port1", Plane.YZ, Plane.ZX, Axis.Z, Axis.NegativeX, sectionWidth, flangeThickness);
                JointHelper.CreatePlanarJoint(HGRBEAM2, "BeginCap", EXTERNALBRACKET1, "Port1", Plane.YZ, Plane.YZ, sectionWidth);
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "EndCap", EXTERNALBRACKET2, "Port1", Plane.YZ, Plane.ZX, Axis.Z, Axis.NegativeX, sectionWidth, flangeThickness);
                JointHelper.CreatePlanarJoint(HGRBEAM3, "BeginCap", EXTERNALBRACKET2, "Port1", Plane.YZ, Plane.YZ, sectionWidth);

                //InternBracket
                int noOfPipes = SupportHelper.SupportedObjects.Count;
                for (int i = 0; i < 3 * noOfPipes; i++)
                {
                    componentDictionary[bracketPartKeys[i]].SetPropertyValue(sectionWidth / 2.0, "IJUAhsWidth1", "Width1");
                    componentDictionary[bracketPartKeys[i]].SetPropertyValue(sectionDepth / 2.0, "IJUAhsLength1", "Length1");
                    componentDictionary[bracketPartKeys[i]].SetPropertyValue(flangeThickness, "IJUAhsThickness1", "Thickness1");
                    componentDictionary[bracketPartKeys[i]].SetPropertyValue(-sectionWidth / 4.0, "IJUAhsStandardPort1", "P1xOffset");
                    componentDictionary[bracketPartKeys[i]].SetPropertyValue(-sectionDepth / 4.0, "IJUAhsStandardPort1", "P1yOffset");
                }

                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", bracketPartKeys[0], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, flangeThickness, flangeThickness);
                JointHelper.CreatePointOnPlaneJoint(bracketPartKeys[0], "Port1", "-1", "Route", Plane.ZX);
                JointHelper.CreateRigidJoint(bracketPartKeys[0], "Port1", bracketPartKeys[1], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0, -3 * flangeThickness);
                JointHelper.CreateRigidJoint(bracketPartKeys[0], "Port1", bracketPartKeys[2], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0, 3 * flangeThickness);

                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", bracketPartKeys[3], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, flangeThickness, flangeThickness);
                JointHelper.CreatePointOnPlaneJoint(bracketPartKeys[3], "Port1", "-1", "Route_2", Plane.ZX);
                JointHelper.CreateRigidJoint(bracketPartKeys[3], "Port1", bracketPartKeys[4], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0, -3 * flangeThickness);
                JointHelper.CreateRigidJoint(bracketPartKeys[3], "Port1", bracketPartKeys[5], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0, 3 * flangeThickness);

                if (noOfPipes > 2)
                {
                    JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", bracketPartKeys[6], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, flangeThickness, flangeThickness);
                    JointHelper.CreatePointOnPlaneJoint(bracketPartKeys[6], "Port1", "-1", "Route_3", Plane.ZX);
                    JointHelper.CreateRigidJoint(bracketPartKeys[6], "Port1", bracketPartKeys[7], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0, -3 * flangeThickness);
                    JointHelper.CreateRigidJoint(bracketPartKeys[6], "Port1", bracketPartKeys[8], "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0, 3 * flangeThickness);
                }

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

                    for (int index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                        routeConnections.Add(new ConnectionInfo(HGRBEAM2, index)); // partindex, routeindex

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

                    if (leftPad)
                        structConnections.Add(new ConnectionInfo(LEFTPAD, 1)); // partindex, Structureindex                       
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM1, 1));
                    if (rightPad)
                        structConnections.Add(new ConnectionInfo(RIGHTPAD, 1)); // partindex, Structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM3, 1)); // partindex, Structureindex
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

