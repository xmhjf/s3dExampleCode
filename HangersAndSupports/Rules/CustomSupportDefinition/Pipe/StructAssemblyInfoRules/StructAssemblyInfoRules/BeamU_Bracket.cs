//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   TypeU.cs
//   SupportStructAssemblyInfoRules,Ingr.SP3D.Content.Support.Rules.BeamU_Bracket
//   Author       :Vijaya
//   Creation Date:5.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  5.Aug.2013     Vijaya   CR-CP-224488  Convert HgrSupStructAssmInfoRules to C# .Net  
// 06.Jun.2016     Vinay    TR-CP-296065	Fix new coverity issues found in H&S Content 
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
                    object[] attributeCollection = SupportStructAssemblyServices.GetPipeStructuralASMAttributes(this);

                    //Get the attributes from assembly
                    d = (double)attributeCollection[0];
                    gap = (double)attributeCollection[1];
                    leftPad = (bool)attributeCollection[2];
                    rightPad = (bool)attributeCollection[3];

                    int noOfPipes = SupportHelper.SupportedObjects.Count;
                    bracketPartKeys = new string[noOfPipes];

                    //Use the default selection rule to get a catalog part for each part class
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("HgrSupExternalBracket");
                    string partSelectionRule = partClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();

                    parts.Add(new PartInfo(HGRBEAM1, "HgrBeam", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));
                    parts.Add(new PartInfo(HGRBEAM2, "HgrBeam", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));
                    parts.Add(new PartInfo(HGRBEAM3, "HgrBeam", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));
                    parts.Add(new PartInfo(EXTERNALBRACKET1, "HgrSupExternalBracket", partSelectionRule));
                    parts.Add(new PartInfo(EXTERNALBRACKET2, "HgrSupExternalBracket", partSelectionRule));

                    partClass = (PartClass)catalogBaseHelper.GetPartClass("HgrSupInternalBracket");

                    SupportComponentUtils supportComponentUtils = new SupportComponentUtils();
                    Part internalBracketPart = null, beam = supportComponentUtils.GetPartFromPartClass("HgrBeam", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor", support);                   
                    CrossSection crossSection = (CrossSection)beam.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                    foreach (BusinessObject bracketPart in partClass.Parts)
                    {
                        if (((string)((PropertyValueString)bracketPart.GetPropertyValue("IJUAHgrBracketType", "BracketType")).PropValue) == crossSection.CrossSectionClass.Name)
                        {
                            internalBracketPart = (Part)bracketPart;
                            break;
                        }
                    }
                    for (int index = 0; index < noOfPipes; index++)
                    {
                        bracketPartKeys[index] = "INTERNALBRACKET" + (index + 1);
                        if (internalBracketPart != null)
                            parts.Add(new PartInfo(bracketPartKeys[index], internalBracketPart.ToString()));
                    }
                    if (rightPad)
                        parts.Add(new PartInfo(RIGHTPAD, SupportStructAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));
                    if (leftPad)
                        parts.Add(new PartInfo(LEFTPAD, SupportStructAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));
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
                double[] boxOffset = SupportStructAssemblyServices.GetBoundaryObjectDimension(this, boundingBox), lugOffset = new double[2];
                //get the offset value based on LoadFactor
                lugOffset[0] = d / 2 - boxOffset[0] / 2;
                lugOffset[1] = d / 2 - boxOffset[3] / 2;
                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = SupportStructAssemblyServices.GetIsLugEndOffsetApplied(this);

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                
                string[] structPort = new string[2];
                structPort = SupportStructAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0], rightStructPort = structPort[1];

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
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        JointHelper.CreatePlanarJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                    else
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                else
                    if (leftPad)
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, 0, 0);

                        JointHelper.CreatePlanarJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }

                    else
                        if (isOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, 0, 0);

                if (rightPad)
                {
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", rightStructPort, RIGHTPAD, "HgrPort_2", Plane.XY, Plane.NegativeXY, 0);
                    else
                        JointHelper.CreatePlanarSlotJoint("-1", rightStructPort, RIGHTPAD, "HgrPort_1", Plane.XY, Plane.NegativeXY, Axis.X, 0, 0);

                    JointHelper.CreatePlanarJoint(RIGHTPAD, "HgrPort_1", HGRBEAM3, "EndCap", Plane.XY, Plane.XY, 0);
                    JointHelper.CreatePrismaticJoint(RIGHTPAD, "HgrPort_1", HGRBEAM3, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
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
                componentDictionary[EXTERNALBRACKET1].SetPropertyValue(sectionWidth, "IJUAHgrBracketOcc", "AngleHeight");
                componentDictionary[EXTERNALBRACKET2].SetPropertyValue(sectionWidth, "IJUAHgrBracketOcc", "AngleHeight");
                componentDictionary[EXTERNALBRACKET1].SetPropertyValue(sectionWidth, "IJUAHgrBracketOcc", "AngleWidth");
                componentDictionary[EXTERNALBRACKET2].SetPropertyValue(sectionWidth, "IJUAHgrBracketOcc", "AngleWidth");
                componentDictionary[EXTERNALBRACKET1].SetPropertyValue(flangeThickness, "IJUAHgrBracketOcc", "BracketThickness");
                componentDictionary[EXTERNALBRACKET2].SetPropertyValue(flangeThickness, "IJUAHgrBracketOcc", "BracketThickness");

                JointHelper.CreatePrismaticJoint(HGRBEAM1, "EndCap", EXTERNALBRACKET1, "HgrPort_1", Plane.YZ, Plane.ZX, Axis.Z, Axis.NegativeX, sectionWidth, 0);
                JointHelper.CreatePlanarJoint(HGRBEAM2, "BeginCap", EXTERNALBRACKET1, "HgrPort_1", Plane.YZ, Plane.YZ, sectionWidth);
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "EndCap", EXTERNALBRACKET2, "HgrPort_1", Plane.YZ, Plane.ZX, Axis.Z, Axis.NegativeX, sectionWidth, 0);
                JointHelper.CreatePlanarJoint(HGRBEAM3, "BeginCap", EXTERNALBRACKET2, "HgrPort_1", Plane.YZ, Plane.YZ, sectionWidth);

                //InternBracket
                componentDictionary[bracketPartKeys[0]].SetPropertyValue(sectionDepth, "IJUAHgrBracketOcc", "AngleHeight");
                componentDictionary[bracketPartKeys[1]].SetPropertyValue(sectionDepth, "IJUAHgrBracketOcc", "AngleHeight");
                componentDictionary[bracketPartKeys[0]].SetPropertyValue(sectionWidth, "IJUAHgrBracketOcc", "AngleWidth");
                componentDictionary[bracketPartKeys[1]].SetPropertyValue(sectionWidth, "IJUAHgrBracketOcc", "AngleWidth");
                componentDictionary[bracketPartKeys[0]].SetPropertyValue(flangeThickness, "IJUAHgrBracketOcc", "BracketThickness");
                componentDictionary[bracketPartKeys[1]].SetPropertyValue(flangeThickness, "IJUAHgrBracketOcc", "BracketThickness");
                componentDictionary[bracketPartKeys[0]].SetPropertyValue(3 * flangeThickness, "IJUAHgrBracketOcc", "GapBTWBrackets");
                componentDictionary[bracketPartKeys[1]].SetPropertyValue(3 * flangeThickness, "IJUAHgrBracketOcc", "GapBTWBrackets");

                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", bracketPartKeys[0], "HgrPort_1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, flangeThickness, flangeThickness);
                JointHelper.CreatePointOnPlaneJoint(bracketPartKeys[0], "HgrPort_1", "-1", "Route", Plane.ZX);
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", bracketPartKeys[1], "HgrPort_1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, flangeThickness, flangeThickness);
                JointHelper.CreatePointOnPlaneJoint(bracketPartKeys[1], "HgrPort_1", "-1", "Route_2", Plane.ZX);
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

