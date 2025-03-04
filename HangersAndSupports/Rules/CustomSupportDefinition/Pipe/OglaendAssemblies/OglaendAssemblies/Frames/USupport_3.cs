//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   USupport_3.cs
//   OglaendAssemblies,Ingr.SP3D.Content.Support.Rules.USupport_3
//   Author       :  PVK
//   Creation Date:  07-11-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET  
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
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
    public class USupport_3 : CustomSupportDefinition, ICustomHgrBOMDescription
    {

        private const string SECTION = "SECTION";
        private const string SECTION_2 = "SECTION_2";
        private const string SECTION_3 = "SECTION_3";
        private const string LEG_1 = "LEG_1";
        private const string LEG_2 = "LEG_2";
        private const string WELDSTAT_1 = "WELDSTAT_1";
        private const string WELDSTAT_2 = "WELDSTAT_2";

        //Steel Members - Stores Data related to steel member (Part Number, Depth, Width etc...)
        private OglaendAssemblyServices.HSSteelMember section;
        private OglaendAssemblyServices.HSSteelMember leg1;
        private OglaendAssemblyServices.HSSteelMember leg2;

        //Steel Configuration - Stores Data related to steel configuration
        private OglaendAssemblyServices.HSSteelConfig sectionData;
        private OglaendAssemblyServices.HSSteelConfig leg1Data;
        private OglaendAssemblyServices.HSSteelConfig leg2Data;

        int[] structureIndex = new int[2];
        Boolean isMember1, isMember2, isMember3,isLeg1, isLeg2, isweldstat1, isweldstat2;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    string strRule = "";
                    BusinessObject ruleBO = null;
                    // Add the Steel Section for the Frame Support
                    string member1Part = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsMember1", "Member1Part")).PropValue;
                    isMember1 = OglaendAssemblyServices.AddPart(this, SECTION, member1Part, strRule, parts, ruleBO);

                    string member2Part = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsHorSection1", "HorSection1Part")).PropValue;
                    isMember2 = OglaendAssemblyServices.AddPart(this, SECTION_2, member2Part, strRule, parts, ruleBO);

                    string member3Part = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsHorSection2", "HorSection2Part")).PropValue;
                    isMember3 = OglaendAssemblyServices.AddPart(this, SECTION_3, member3Part, strRule, parts, ruleBO);

                    string leg1Part = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLeg1", "Leg1Part")).PropValue;
                    isLeg1 = OglaendAssemblyServices.AddPart(this, LEG_1, leg1Part, null, parts);

                    string leg2Part = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLeg2", "Leg2Part")).PropValue;
                    isLeg2 = OglaendAssemblyServices.AddPart(this, LEG_2, leg2Part, null, parts);

                    // Add the Plates
                    string Weldstat1 = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsWeldStarter1", "WeldStarter1Part")).PropValue;
                    isweldstat1 = OglaendAssemblyServices.AddPart(this, WELDSTAT_1, Weldstat1, null, parts);

                    string Weldstat2 = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsWeldStarter2", "WeldStarter2Part")).PropValue;
                    isweldstat2 = OglaendAssemblyServices.AddPart(this, WELDSTAT_2, Weldstat2, null, parts);

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
        //Get Assembly Implied Parts
        //-----------------------------------------------------------------------------------
        public override ReadOnlyCollection<PartInfo> ImpliedParts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> impliedParts = new Collection<PartInfo>();
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    string supportpartnumber = support.SupportDefinition.PartNumber;
                    OglaendAssemblyServices.AddImpliedPartbyInterface(this, impliedParts, "IJUAhsFrImpPart", supportpartnumber, 10005);
                    ReadOnlyCollection<PartInfo> rImpliedParts = new ReadOnlyCollection<PartInfo>(impliedParts);
                    return rImpliedParts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Implied Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                SP3D.Support.Middle.Support support = SupportHelper.Support;

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                // ==========================
                // Get Required Information about the Steel Parts
                // ==========================
                // Get the Steel Cross Section Data
                section = OglaendAssemblyServices.GetSectionDataFromPartIndex(this, SECTION);
                leg1 = OglaendAssemblyServices.GetSectionDataFromPartIndex(this, LEG_1);
                leg2 = OglaendAssemblyServices.GetSectionDataFromPartIndex(this, LEG_2);

                // Get the Steel Configuration Data
                sectionData.Orient = (OglaendAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue);
                leg1Data.Orient = (OglaendAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue);
                leg2Data.Orient = (OglaendAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsLeg2Ang", "Leg2OrientationAngle")).PropValue);

                // Get the Width and Depth of each steel part according to the orientation angle
                double leg1Depth = 0, leg1Width = 0, leg2Depth = 0, leg2Width = 0, sectionDepth = 0, sectionWidth = 0;
                if (sectionData.Orient == (OglaendAssemblyServices.SteelOrientationAngle)0 || sectionData.Orient == (OglaendAssemblyServices.SteelOrientationAngle)180)
                {
                    sectionDepth = section.depth;
                    sectionWidth = section.width;
                }
                else
                {
                    sectionDepth = section.width;
                    sectionWidth = section.depth;
                }
                if (leg1Data.Orient == (OglaendAssemblyServices.SteelOrientationAngle)0 || leg1Data.Orient == (OglaendAssemblyServices.SteelOrientationAngle)180)
                {
                    leg1Depth = leg1.depth;
                    leg1Width = leg1.width;
                }
                else
                {
                    leg1Depth = leg1.width;
                    leg1Width = leg1.depth;
                }
                if (leg2Data.Orient == (OglaendAssemblyServices.SteelOrientationAngle)0 || leg2Data.Orient == (OglaendAssemblyServices.SteelOrientationAngle)180)
                {
                    leg2Depth = leg2.depth;
                    leg2Width = leg2.depth;
                }
                else
                {
                    leg2Depth = leg2.width;
                    leg2Width = leg2.depth;
                }
                // ==========================
                // Create the Frame Bounding Box
                // ==========================
                double boundingBoxWidth = 0, boundingBoxDepth = 0;
                string boundingBoxPort = "BBFrame_Low", boundingBoxName = "BBFrame";
                int frameOrientation = 0;

                frameOrientation = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsFrameOrientation", "FrameOrientation")).PropValue;

                Boolean includeInsulation = false;

                PropertyValue omirrorFrame = OglaendAssemblyServices.GetAttributeValue(support, "IJUAhsMirrorFrame", "MirrorFrame", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, true);

                Boolean mirrorFrame;
                try
                {
                    mirrorFrame = ((bool)((PropertyValueBoolean)omirrorFrame).PropValue);
                }
                catch
                {
                    mirrorFrame = false;
                }

                OglaendAssemblyServices.CreateFrameBoundingBox(this, boundingBoxName, (OglaendAssemblyServices.FrameBBXOrientation)frameOrientation, includeInsulation, mirrorFrame, SupportedHelper.IsSupportedObjectVertical(1, 45));

                boundingBoxWidth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Width;
                boundingBoxDepth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Height;
                // ==========================
                // Determine the number of Supporting Objects
                // ==========================
                int supportingCount = 0;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    // For Place-By-Reference Command No Supporting Objects are Selected
                    // However, we want to treat it like 1 was selected
                    supportingCount = 1;
                else
                    supportingCount = SupportHelper.SupportingObjects.Count;

                // ==========================
                // Organize the Structure Reference Ports based on BBX Y-Axis Direction
                // ==========================
                string[] structurePort = new string[2];
                if (supportingCount > 1)
                {
                    structurePort[0] = "Structure";
                    structureIndex[0] = 1;
                    structurePort[1] = "Struct_2";
                    structureIndex[1] = 2;
                }
                else
                {
                    structurePort[0] = "StructAlt";
                    structureIndex[0] = 1;
                    structurePort[1] = "StructAlt";
                    structureIndex[1] = 1;
                }


                //'==========================
                // 'SPAN
                // '==========================

                Double SpanWidth=0, MaxSpanWidth = 0, MinSpanWidth = 0, BBXWidth = 0, BBXWidth2 = 0, MaxHeight = 0;
                String sSupportNumber;
                Boolean lvalue = true;
                PropertyValue olvalue1 = OglaendAssemblyServices.GetAttributeValue(support, "IJUAhsAdjustSupport", "AutoConnect", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, true);

                lvalue = ((bool)((PropertyValueBoolean)olvalue1).PropValue);
                sSupportNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsSupportNumber", "SupportNumber")).PropValue;

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass auxTable = (PartClass)cataloghelper.GetPartClass("HS_OGAssy_MaxValues");
                ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in classItems)
                {
                    string svalue = ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGMaxValue", "SupportNumber")).PropValue);
                    bool isEqual = String.Equals(sSupportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGMaxValue", "SupportNumber")).PropValue), StringComparison.Ordinal);
                    if (isEqual == true)
                    {
                        MaxHeight = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGMaxValue", "MaximumHeight")).PropValue);
                        MaxSpanWidth = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGMaxValue", "MaximumWidth")).PropValue);
                        break;
                    }

                }

                // Span//
                if (supportingCount > 1)
                {
                    if (lvalue == true)
                    {
                        int supportingface = 0;
                        if ((SupportHelper.SupportingObjects.Count != 0))
                            supportingface = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                        switch (supportingface)
                        {
                            case 257:
                                SpanWidth = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortAxisType.Z);
                                break;
                            case 258:
                                SpanWidth = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortAxisType.Z);
                                break;
                            case 513:
                                SpanWidth = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortAxisType.Y);
                                break;
                            case 514:
                                SpanWidth = RefPortHelper.DistanceBetweenPorts("Structure", "Struct_2", PortAxisType.Y);
                                break;
                        }
                        if (SpanWidth < 0)
                        {
                            SpanWidth = -SpanWidth;
                        }
                        SpanWidth = SpanWidth + leg1Depth / 2 + leg2Depth / 2;

                        if (SpanWidth > MaxSpanWidth)
                        {
                            SpanWidth = MaxSpanWidth;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMaxWidth,"Width Exceeds Maximum Allowable Width. Resetting it to Maximum Width."), "", "USupport_1", 346);
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                        else
                        {
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                    }
                    else
                    {
                        SpanWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGSpanWidth", "SpanWidth")).PropValue;
                        if (SpanWidth > MaxSpanWidth)
                        {
                            SpanWidth = MaxSpanWidth;

                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMaxWidth, "Width Exceeds Maximum Allowable Width. Resetting it to Maximum Width."), "", "USupport_1", 346);
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                        else if (HgrCompareDoubleService.cmpdbl(SpanWidth, 0) == true)
                        {
                            SpanWidth = boundingBoxWidth + leg1Depth + leg2Depth;
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                        else if (SpanWidth < MinSpanWidth)
                        {
                            SpanWidth = MinSpanWidth;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinWidth,"Width is less than Minimum Allowable Width. Resetting it to Minimum Width."), "", "USupport_1", 346);
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                    }

                }
                else
                {
                    try
                    {
                        BBXWidth2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyValue2", "DummyValue2")).PropValue;
                    }
                    catch
                    {
                        BBXWidth2 = 0;
                    }
                    BBXWidth = boundingBoxWidth;
                    support.SetPropertyValue(BBXWidth, "IJUAhsOGDummyValue2", "DummyValue2");
                    if (HgrCompareDoubleService.cmpdbl(Math.Round(BBXWidth2 * 1000, 0), Math.Round(BBXWidth * 1000, 0)) == false)
                    {
                        if ((boundingBoxWidth + leg1Depth / 2 + leg2Depth / 2) > MaxSpanWidth)
                        {
                            SpanWidth = MaxSpanWidth;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMaxWidth,"Width Exceeds Maximum Allowable Width. Resetting it to Maximum Width."), "", "USupport_1", 346);
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                        else
                        {
                            SpanWidth = boundingBoxWidth + leg1Depth + leg2Depth;
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");

                        }
                    }
                    else
                    {
                        SpanWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGSpanWidth", "SpanWidth")).PropValue;
                        if (SpanWidth > MaxSpanWidth)
                        {
                            SpanWidth = MaxSpanWidth;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMaxWidth,"Width Exceeds Maximum Allowable Width. Resetting it to Maximum Width."), "", "USupport_1", 346);
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                        else if (HgrCompareDoubleService.cmpdbl(SpanWidth, 0) == true)
                        {
                            SpanWidth = boundingBoxWidth + leg1Depth + leg2Depth;
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                        else if (SpanWidth < MinSpanWidth)
                        {
                            SpanWidth = MinSpanWidth;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinWidth,"Width is less than Minimum Allowable Width. Resetting it to Minimum Width."), "", "USupport_1", 346);
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }

                    }

                }

                //Height//
                double ShoeHeight, RSDistance, InputHeight, SupportHeight, RSdistance2;
                string Refport;
                if (supportingCount > 1)
                {
                    if ((RefPortHelper.DistanceBetweenPorts("Structure", boundingBoxPort, PortAxisType.Z)) > (RefPortHelper.DistanceBetweenPorts("Struct_2", boundingBoxPort, PortAxisType.Z)))
                        Refport = "Structure";
                    else
                        Refport = "Struct_2";
                }
                else
                    Refport = "StructAlt";

                ShoeHeight = 0;
                try
                {
                    InputHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight3", "Height3")).PropValue;
                }
                catch
                {
                    InputHeight = 0;
                }
                try
                {
                    RSdistance2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGDummyValue1", "DummyValue1")).PropValue;
                }
                catch
                {
                    RSdistance2 = 0;
                }
                RSDistance = RefPortHelper.DistanceBetweenPorts(Refport, boundingBoxPort, PortAxisType.Z);
                SupportHeight = RSDistance + sectionDepth;
                support.SetPropertyValue(RSDistance, "IJUAhsOGDummyValue1", "DummyValue1");

                if ((HgrCompareDoubleService.cmpdbl(Math.Round(RSdistance2 * 1000, 0), Math.Round(RSDistance * 1000, 0)) == false))
                {
                    if (((Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsLengthOverride", "LengthOverride")).PropValue) == true)
                    {
                        if (RSDistance > MaxHeight)
                        {
                            ShoeHeight = ShoeHeight - (SupportHeight - MaxHeight);
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxHeight,"Height Exceeds Maximum Allowable Height. Resetting it to Maximum Height."), "", "USupport_1", 346);
                            support.SetPropertyValue(MaxHeight, "IJUAhsOGHeight3", "Height3");
                        }
                        else
                        {
                            ShoeHeight = 0;
                            support.SetPropertyValue(SupportHeight, "IJUAhsOGHeight3", "Height3");
                        }
                    }
                    else
                    {
                        if (InputHeight > MaxHeight)
                        {
                            ShoeHeight = ShoeHeight - (SupportHeight - MaxHeight);
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxHeight,"Height Exceeds Maximum Allowable Height. Resetting it to Maximum Height."), "", "USupport_1", 346);
                            support.SetPropertyValue(MaxHeight, "IJUAhsOGHeight3", "Height3");
                        }
                        else
                        {
                            ShoeHeight = ShoeHeight - (SupportHeight - InputHeight);
                            support.SetPropertyValue(InputHeight, "IJUAhsOGHeight3", "Height3");
                        }
                    }
                }
                else
                {
                    if (InputHeight > MaxHeight)
                    {
                        ShoeHeight = ShoeHeight - (SupportHeight - MaxHeight);
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxHeight,"Height Exceeds Maximum Allowable Height. Resetting it to Maximum Height."), "", "USupport_1", 346);
                        support.SetPropertyValue(MaxHeight, "IJUAhsOGHeight3", "Height3");
                    }
                    else if (HgrCompareDoubleService.cmpdbl(InputHeight, 0) == true)
                    {
                        ShoeHeight = 0;
                        support.SetPropertyValue(SupportHeight, "IJUAhsOGHeight3", "Height3");
                    }
                    else
                    {
                        ShoeHeight = ShoeHeight - (SupportHeight - InputHeight);
                        support.SetPropertyValue(InputHeight, "IJUAhsOGHeight3", "Height3");
                    }

                }
                int connection1Type = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsCornerConn1Type", "Connection1Type")).PropValue; ;
                int connection2Type = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsCornerConn2Type", "Connection2Type")).PropValue;

                //'==========================
                //' Leg 1 and Leg 2 Begin Overhang
                //'==========================
                double leg1BeginOverHang = 0;
                double leg2BeginOverHang = 0;

                //'==========================
                //' Member 1  Overhangs
                //'==========================
                double member1BeginOverhang = 0;
                double member1EndOverhang = 0;

                //==========================
                // Set the Frame Connections
                //==========================
                //Set Connection between Leg1 and Section

                Boolean connectionSwap = (bool)((PropertyValueBoolean)part.GetPropertyValue("IJUAhsCornerConn1Swap", "Connection1Swap")).PropValue;
                Boolean connection2Swap = (bool)((PropertyValueBoolean)part.GetPropertyValue("IJUAhsCornerConn2Swap", "Connection2Swap")).PropValue;
                Boolean connection1Mirror = (bool)((PropertyValueBoolean)part.GetPropertyValue("IJUAhsCornerConn1Mirror", "Connection1Mirror")).PropValue;
                Boolean connection2Mirror = (Boolean)((PropertyValueBoolean)part.GetPropertyValue("IJUAhsCornerConn2Mirror", "Connection2Mirror")).PropValue;

                OglaendAssemblyServices.SetSteelConnection(this, LEG_1, "BeginCap", ref leg1Data, leg1BeginOverHang, SECTION, "BeginCap", ref sectionData, 0, (OglaendAssemblyServices.SteelConnectionAngle)90, member1BeginOverhang + leg1Depth / 2, (OglaendAssemblyServices.SteelConnection)connection1Type, connection1Mirror, OglaendAssemblyServices.SteelJointType.SteelJoint_RIGID);

                //Set Connection between Leg1 and Section
                OglaendAssemblyServices.SetSteelConnection(this, LEG_2, "EndCap", ref leg2Data, leg2BeginOverHang, SECTION, "EndCap", ref sectionData, 0, (OglaendAssemblyServices.SteelConnectionAngle)270, -(member1EndOverhang + leg2Depth / 2), (OglaendAssemblyServices.SteelConnection)connection2Type, connection2Mirror, OglaendAssemblyServices.SteelJointType.SteelJoint_RIGID);
                //==========================
                // Set Steel CPs
                //==========================
                // Set the CP for the SECTION Flex Ports (Used to connect the main section to the BBX)
                PropertyValueCodelist cardinalPoint66Section = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint67Section = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCPFlexPort", "CP7");                
                
                PropertyValueCodelist cardinalPoint6Leg1 = (PropertyValueCodelist)componentDictionary[LEG_1].GetPropertyValue("IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7Leg1 = (PropertyValueCodelist)componentDictionary[LEG_1].GetPropertyValue("IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist cardinalPoint2Leg1 = (PropertyValueCodelist)componentDictionary[LEG_1].GetPropertyValue("IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Leg1 = (PropertyValueCodelist)componentDictionary[LEG_1].GetPropertyValue("IJOAhsSteelCP", "CP1");
                
                PropertyValueCodelist cardinalPoint6Leg2 = (PropertyValueCodelist)componentDictionary[LEG_2].GetPropertyValue("IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7Leg2 = (PropertyValueCodelist)componentDictionary[LEG_2].GetPropertyValue("IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist cardinalPoint2Leg2 = (PropertyValueCodelist)componentDictionary[LEG_2].GetPropertyValue("IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Leg2 = (PropertyValueCodelist)componentDictionary[LEG_2].GetPropertyValue("IJOAhsSteelCP", "CP1");
                switch (sectionData.Orient)
                {
                    case 0:
                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 2, "IJOAhsSteelCP", "CP6");
                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                        break;
                    case (OglaendAssemblyServices.SteelOrientationAngle)90:
                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 6, "IJOAhsSteelCP", "CP6");
                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                        break;
                    case (OglaendAssemblyServices.SteelOrientationAngle)180:
                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 8, "IJOAhsSteelCP", "CP6");
                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                        break;
                    case (OglaendAssemblyServices.SteelOrientationAngle)270:
                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 4, "IJOAhsSteelCP", "CP6");
                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                        break;
                }
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsFlexPort", "FlexPortRotZ");
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsEndFlexPort", "EndFlexPortRotZ");

                //Set the CP for the Leg Ports that connect to the supporting Structure
                componentDictionary[LEG_1].SetPropertyValue(cardinalPoint2Leg1.PropValue = leg1Data.CardinalPoint, "IJOAhsSteelCP", "CP2");
                componentDictionary[LEG_1].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(leg1Data.Orient) * Math.PI / 180), "IJOAhsEndCap", "EndCapRotZ");
                componentDictionary[LEG_2].SetPropertyValue(cardinalPoint1Leg2.PropValue = leg2Data.CardinalPoint, "IJOAhsSteelCP", "CP1");
                componentDictionary[LEG_2].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(leg2Data.Orient) * Math.PI / 180), "IJOAhsBeginCap", "BeginCapRotZ");

                //=====================
                // Set Span Length
                //=====================
                componentDictionary[SECTION_2].SetPropertyValue(SpanWidth, "IJUAHgrOccLength", "Length");
                componentDictionary[SECTION].SetPropertyValue(SpanWidth, "IJUAHgrOccLength", "Length");
                componentDictionary[SECTION_3].SetPropertyValue(SpanWidth, "IJUAHgrOccLength", "Length");

                //'==========================
                //' Offset 
                //'==========================
                double offset1 = -RefPortHelper.DistanceBetweenPorts("BBFrame_Low", structurePort[0], PortAxisType.Y);
                if (connection1Type == 0)
                {
                    if (offset1 < 0)
                        offset1 = offset1 + SpanWidth;
                }
                else
                {
                    if (offset1 < 0)
                        offset1 = offset1 + SpanWidth - leg1Width / 2;
                    else
                        offset1 = offset1 + leg1Width / 2;

                }
                if (supportingCount > 1)
                {
                    double offsetY1, offsetY2;
                    offsetY1 = RefPortHelper.DistanceBetweenPorts("BBFrame_Low", structurePort[0], PortAxisType.Y);
                    offsetY2 = RefPortHelper.DistanceBetweenPorts("BBFrame_Low", structurePort[1], PortAxisType.Y);
                    if (offsetY1 > 0 && offsetY2 > 0)
                    {
                        if (Math.Abs(offsetY1) < Math.Abs(offsetY2))
                            offset1 = offset1 - SpanWidth + leg1Width;
                    }
                    else if (offsetY1 < 0 && offsetY2 < 0)
                    {
                        if (Math.Abs(offsetY1) < Math.Abs(offsetY2))
                            offset1 = offset1 + SpanWidth - leg1Width;
                    }
                }
                //==========================
                // Joints To Connect the Main Steel Section to the BBX
                //==========================
                double FrameOffset;
                if (supportingCount > 1)
                    FrameOffset = offset1;
                else
                    FrameOffset = SpanWidth / 2 - boundingBoxWidth / 2;

                componentDictionary[SECTION].SetPropertyValue(FrameOffset, "IJOAhsFlexPort", "FlexPortZOffset");
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsEndFlexPort", "EndFlexPortZOffset");

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreatePrismaticJoint(SECTION, "BeginFlex", "-1", boundingBoxPort, Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, -ShoeHeight);
                    JointHelper.CreatePointOnAxisJoint(LEG_1, "EndFlex", "-1", structurePort[0], Axis.X);
                }
                else
                    JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -ShoeHeight, 0, 0);

                //'================================================
                //' Joints To Connect Leg 1 To Supporting Structure
                //'================================================
                double planeangle;
                planeangle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[0], PortAxisType.Z, OrientationAlong.Direct);
                componentDictionary[LEG_1].SetPropertyValue(cardinalPoint6Leg1.PropValue = 2, "IJOAhsSteelCP", "CP6");
                componentDictionary[LEG_1].SetPropertyValue(cardinalPoint7Leg1.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");

                if (HgrCompareDoubleService.cmpdbl(Math.Round(planeangle * 180 / Math.PI), 180) == true)
                {
                    JointHelper.CreatePrismaticJoint(LEG_1, "BeginFlex", LEG_1, "EndFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    JointHelper.CreatePointOnPlaneJoint(LEG_1, "EndFlex", "-1", structurePort[0], Plane.NegativeXY);
                }
                else
                {
                    JointHelper.CreatePrismaticJoint(LEG_1, "BeginFlex", LEG_1, "EndFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    JointHelper.CreatePointOnPlaneJoint(LEG_1, "EndFlex", "-1", structurePort[0], Plane.NegativeZX);

                }
                //'================================================
                //' Joints To Connect Leg 2 To Supporting Structure
                //'================================================
                planeangle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[1], PortAxisType.Z, OrientationAlong.Direct);
                componentDictionary[LEG_2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 2, "IJOAhsSteelCP", "CP6");
                componentDictionary[LEG_2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");

                if (HgrCompareDoubleService.cmpdbl(Math.Round(planeangle * 180 / Math.PI), 180) == true)
                {
                    JointHelper.CreatePrismaticJoint(LEG_2, "EndFlex", LEG_2, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    JointHelper.CreatePointOnPlaneJoint(LEG_2, "BeginFlex", "-1", structurePort[1], Plane.NegativeXY);
                }
                else
                {
                    JointHelper.CreatePrismaticJoint(LEG_2, "EndFlex", LEG_2, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    JointHelper.CreatePointOnPlaneJoint(LEG_2, "BeginFlex", "-1", structurePort[1], Plane.NegativeZX);
                }

                //'========================
                //'Joints for additional sections
                //'========================

                double SectionHeight1 = 0;
                double SectionHeight2 = 0;

                InputHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight3", "Height3")).PropValue;
                try
                {
                    SectionHeight2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight2", "Height2")).PropValue;
                }
                catch
                {
                    SectionHeight2 = 0;
                }
                try
                {
                    SectionHeight1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight1", "Height1")).PropValue;
                }
                catch
                {
                    SectionHeight1 = 0;
                }

                if ((HgrCompareDoubleService.cmpdbl(Math.Round(SectionHeight1), 0) == true) || (HgrCompareDoubleService.cmpdbl(Math.Round(SectionHeight2), 0) == true))
                {
                    SectionHeight1 = InputHeight / 3;
                    SectionHeight2 = (2* InputHeight / 3);
                    support.SetPropertyValue(SectionHeight1, "IJUAhsOGHeight1", "Height1");
                    support.SetPropertyValue(SectionHeight2, "IJUAhsOGHeight2", "Height2");
                }
                else
                    SectionHeight2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight2", "Height2")).PropValue;
                SectionHeight1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGHeight1", "Height1")).PropValue;

                JointHelper.CreateAngularRigidJoint(SECTION_2, "Neutral", SECTION, "Neutral", new Vector(0, SectionHeight1 - InputHeight, 0), new Vector(0, 0, 0));
                JointHelper.CreateAngularRigidJoint(SECTION_3, "Neutral", SECTION, "Neutral", new Vector(0, SectionHeight2 - InputHeight, 0), new Vector(0, 0, 0));

                //'========================
                //'Joints for Wels Starters
                //'========================
                double Wlength1 = 0, StructureIntPosZ = 0, Yoffset = 0, Zoffset = 0;

                PropertyValue oPropertyValue = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDSTAT_1], "IJUAHgrOGLength", "Length", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);

                Wlength1 = (double)(((PropertyValueDouble)oPropertyValue).PropValue);

                //MessageBox.Show("Wlength1");
                StructureIntPosZ = (double)((PropertyValueDouble)componentDictionary[WELDSTAT_1].GetPropertyValue("IJUAHgrOGStructureIntPort", "StuctureIntPosZ")).PropValue;
                Yoffset = leg1Depth / 2;
                Zoffset = Wlength1 / 2;

                JointHelper.CreateAngularRigidJoint(WELDSTAT_1, "StructureInt", LEG_1, "EndCap", new Vector(-0, Yoffset, -Zoffset + StructureIntPosZ), new Vector(0, 0, Math.PI));
                JointHelper.CreateAngularRigidJoint(WELDSTAT_2, "StructureInt", LEG_2, "BeginCap", new Vector(-0, Yoffset, Zoffset + StructureIntPosZ), new Vector(0, 0, Math.PI));

                componentDictionary[LEG_1].SetPropertyValue(-Wlength1 / 2, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[LEG_2].SetPropertyValue(-Wlength1 / 2, "IJUAHgrOccOverLength", "BeginOverLength");

                //'Query For Loads
                Double CapacityPy = 0, CapacityPx = 0;
                SpanWidth = Math.Round(SpanWidth * 1000, 0);

                PartClass auxTable1 = (PartClass)cataloghelper.GetPartClass("HS_OGAssy_LoadTable");
                ReadOnlyCollection<BusinessObject> classItems1 = auxTable1.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                double minwidth = 0; double maxwidth = 0;
                foreach (BusinessObject classItem in classItems1)
                {
                    minwidth = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "MinWidth")).PropValue);
                    maxwidth = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "MaxWidth")).PropValue);

                    bool isEqual = String.Equals(sSupportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGLoadTable", "SupportNumber")).PropValue), StringComparison.Ordinal);
                    if ((isEqual == true) && (SpanWidth > minwidth && SpanWidth < maxwidth + 0.00001))
                    {
                        CapacityPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "CapacityPy")).PropValue);
                        CapacityPx = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "CapacityPx")).PropValue);
                    }

                }
                //set load properties
                if (support.SupportsInterface("IJUAhsOGCapacityPy"))
                    support.SetPropertyValue(CapacityPy, "IJUAhsOGCapacityPy", "CapacityPy");
                if (support.SupportsInterface("IJUAhsOGCapacityPx"))
                    support.SetPropertyValue(CapacityPx, "IJUAhsOGCapacityPx", "CapacityPx");

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
                return 1;
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
                    int numberOfRoutes = SupportHelper.SupportedObjects.Count;
                    for (int routeIndex = 1; routeIndex <= numberOfRoutes; routeIndex++)
                    {
                        routeConnections.Add(new ConnectionInfo(SECTION, routeIndex)); // partindex, routeindex     
                    }

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
                    if (isweldstat1)
                        structConnections.Add(new ConnectionInfo(WELDSTAT_1, 1)); // partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(LEG_1, 1));

                    if (isweldstat2)
                        structConnections.Add(new ConnectionInfo(WELDSTAT_2, 1)); // partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(LEG_2, 1));
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

        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDesciption = "";
            try
            {
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject catalogpart = SupportOrComponent.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                string supportNumber = (string)((PropertyValueString)catalogpart.GetPropertyValue("IJUAhsSupportNumber", "SupportNumber")).PropValue;
                bomDesciption = supportNumber.Substring(0, 2) + "-" + supportNumber.Substring(2, 1) + ": " + supportNumber.Substring(3, 2);

                return bomDesciption;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}
