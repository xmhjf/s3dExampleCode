//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   TFrame.cs
//   Frame_Assy,Ingr.SP3D.Content.Support.Rules.UFrame
//   Author       :  Hema
//   Creation Date:  26-Jul-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   26-Jul-2013     Hema    CR-CP-224474 Convert HS_S3DFrame to C# .Net 
//   30-Mar-2014     PVK     CR-CP-245789	Modify the exsisting .Net Frame_Assy to behave like URS Frame supports
//   27-APr-2015     PVK     TR-CP-253033	Elevation CP not shown by default for frame supports.
//   06-May-2015     PVK     CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
//   17/12/2015      Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   22/02/2016      PVK     Corrected Shoe Height for U Frame
//   25/05/2016      PVK     TR-CP-292743	Corrected the VerticalOffset attribute
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Linq;
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

    public class UFrame : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Steel Sections
        private const string SECTION = "SECTION"; //Main Member
        private const string LEG1 = "LEG1"; //1st Leg of U-Frame
        private const string LEG2 = "LEG2"; //2nd Leg of U-Frame

        //Other Optional Parts
        private const string CAPPLATE1 = "CAPPLATE1"; //Plate on the end of Leg 1
        private const string CAPPLATE2 = "CAPPLATE2"; //Plate on the end of Leg 2
        private const string CAPPLATE3 = "CAPPLATE3"; //Plate on the end of the Main Member
        private const string CAPPLATE4 = "CAPPLATE4"; //Plate on the end of the Main Member
        private const string BASEPLATE1 = "BASEPLATE1"; //Base Plate for Leg 1
        private const string BASEPLATE2 = "BASEPLATE2"; //Base Plate for Leg 2

        private const string STRUCT_CONN = "STRUCT_CONN"; //Logical Structural Connection

        //Collections for the Weld Data and Weld Part Index's
        Collection<FrameAssemblyServices.WeldData> weldCollection = new Collection<FrameAssemblyServices.WeldData>();  //Collection of Welds (hsWeldData Type)

        //Steel Members - Stores Data related to steel member (Part Number, Depth, Width etc...)
        private FrameAssemblyServices.HSSteelMember section;
        private FrameAssemblyServices.HSSteelMember leg1;
        private FrameAssemblyServices.HSSteelMember leg2;

        //Steel Configuration - Stores Data related to steel configuration
        private FrameAssemblyServices.HSSteelConfig sectionData;
        private FrameAssemblyServices.HSSteelConfig leg1Data;
        private FrameAssemblyServices.HSSteelConfig leg2Data;

        int[] structureIndex = new int[2];
        Boolean isMember, isLeg1, isLeg2, isCapPlate1, isCapPlate2, isCapPlate3, isCapPlate4, isBasePlate1, isBasePlate2, isBolt1Part, isBolt2Part;
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
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    if (part.SupportsInterface("IJUAHgrURSCommon"))
                    {
                        string family = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support, "IJUAHgrURSCommon", "Family")).PropValue;
                        string type = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support, "IJUAHgrURSCommon", "Type")).PropValue;
                        if (family != "" && type != null)
                            FrameAssemblyServices.CheckSupportWithFamilyAndType(this, family, type);
                    }
                    // Add the Steel Section for the Frame Support
                    string member1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1", "Member1Part")).PropValue;
                    string member1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1Rl", "Member1Rule")).PropValue;
                    isMember = FrameAssemblyServices.AddPart(this, SECTION, member1Part, member1Rule, parts);

                    string leg1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1", "Leg1Part")).PropValue;
                    string leg1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1Rl", "Leg1Rule")).PropValue;
                    isLeg1 = FrameAssemblyServices.AddPart(this, LEG1, leg1Part, leg1Rule, parts);

                    string leg2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg2", "Leg2Part")).PropValue;
                    string leg2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg2Rl", "Leg2Rule")).PropValue;
                    isLeg2 = FrameAssemblyServices.AddPart(this, LEG2, leg2Part, leg2Rule, parts);
                    // Add the Plates
                    string capPlate1Part1 = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate1", "CapPlate1Part")).PropValue;
                    string capPlate1Rule1 = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate1Rl", "CapPlate1Rule")).PropValue;
                    isCapPlate1 = FrameAssemblyServices.AddPart(this, CAPPLATE1, capPlate1Part1, capPlate1Rule1, parts);

                    string capPlate1Part2 = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate2", "CapPlate2Part")).PropValue;
                    string capPlate1Rule2 = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate2Rl", "CapPlate2Rule")).PropValue;
                    isCapPlate2 = FrameAssemblyServices.AddPart(this, CAPPLATE2, capPlate1Part2, capPlate1Rule2, parts);

                    string capPlate1Part3 = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate3", "CapPlate3Part")).PropValue;
                    string capPlate1Rule3 = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate3Rl", "CapPlate3Rule")).PropValue;
                    isCapPlate3 = FrameAssemblyServices.AddPart(this, CAPPLATE3, capPlate1Part3, capPlate1Rule3, parts);

                    string capPlate1Part4 = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate4", "CapPlate4Part")).PropValue;
                    string capPlate1Rule4 = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate4Rl", "CapPlate4Rule")).PropValue;
                    isCapPlate4 = FrameAssemblyServices.AddPart(this, CAPPLATE4, capPlate1Part4, capPlate1Rule4, parts);

                    string basePlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1", "BasePlate1Part")).PropValue;
                    string basePlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1Rl", "BasePlate1Rule")).PropValue;
                    isBasePlate1 = FrameAssemblyServices.AddPart(this, BASEPLATE1, basePlate1Part, basePlate1Rule, parts);

                    string basePlate2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate2", "BasePlate2Part")).PropValue;
                    string basePlate2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate2Rl", "BasePlate2Rule")).PropValue;
                    isBasePlate2 = FrameAssemblyServices.AddPart(this, BASEPLATE2, basePlate2Part, basePlate2Rule, parts);

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        parts.Add(new PartInfo(STRUCT_CONN, "Log_Conn_Part_1"));

                    //Add the Weld Objects from the Weld Sheet
                    string weldServiceClassName = string.Empty;
                    if (part.SupportsInterface("IJUAhsWeldServClass"))
                        weldServiceClassName = (String)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsWeldServClass", "WeldServiceClassName")).PropValue;
                    if (weldServiceClassName == null)
                        weldServiceClassName = string.Empty;
                    weldCollection = FrameAssemblyServices.AddWeldsFromCatalog(this, parts, "IJUAhsUFrameWelds", ((IPart)part).PartNumber, weldServiceClassName);

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
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> impliedParts = new Collection<PartInfo>();
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                    int bolt1Quantity = 0, bolt2Quantity = 0;
                    string ImpliedPartServiceClass = null;

                    // Add the Optional Bolts as Implied Parts
                    if (part.SupportsInterface("IJUAhsImpServClass"))
                        ImpliedPartServiceClass = (String)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsImpServClass", "ImpServiceClassName")).PropValue;
                    if (ImpliedPartServiceClass == string.Empty)
                        ImpliedPartServiceClass = null;

                    if (ImpliedPartServiceClass != null)
                        FrameAssemblyServices.AddImpliedPartFromCatalog(this, impliedParts, ImpliedPartServiceClass);

                    if (isBasePlate1)
                    {
                        string bolt1Part, bolt1Rule;
                        bolt1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1", "Bolt1Part")).PropValue;
                        bolt1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1Rl", "Bolt1Rule")).PropValue;
                        bolt1Quantity = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1Qty", "Bolt1Quantity")).PropValue;
                        isBolt1Part = FrameAssemblyServices.AddImpliedPart(this, bolt1Part, bolt1Rule, impliedParts, null, bolt1Quantity);
                    }
                    if (isBasePlate2)
                    {
                        string bolt2Part, bolt2Rule;
                        bolt2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt2", "Bolt2Part")).PropValue;
                        bolt2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt2Rl", "Bolt2Rule")).PropValue;
                        bolt2Quantity = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt2Qty", "Bolt2Quantity")).PropValue;
                        isBolt2Part = FrameAssemblyServices.AddImpliedPart(this, bolt2Part, bolt2Rule, impliedParts, null, bolt2Quantity);
                    }
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
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                Collection<FrameAssemblyServices.WeldData> frameConnection1Welds = new Collection<FrameAssemblyServices.WeldData>(); // Welds at the connection between Leg and Main Member
                Collection<FrameAssemblyServices.WeldData> frameConnection2Welds = new Collection<FrameAssemblyServices.WeldData>();
                Collection<FrameAssemblyServices.WeldData> otherWelds = new Collection<FrameAssemblyServices.WeldData>(); // Other welds in the support

                // ==========================
                // Get Required Information about the Steel Parts
                // ==========================
                // Get the Steel Cross Section Data
                section = FrameAssemblyServices.GetSectionDataFromPartIndex(this, SECTION);
                leg1 = FrameAssemblyServices.GetSectionDataFromPartIndex(this, LEG1);
                leg2 = FrameAssemblyServices.GetSectionDataFromPartIndex(this, LEG2);

                // Get the Steel Configuration Data
                sectionData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue);
                leg1Data.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue);
                leg2Data.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg2Ang", "Leg2OrientationAngle")).PropValue);

                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                // Get the Width and Depth of each steel part according to the orientation angle
                double leg1Depth = 0, leg1Width = 0, leg2Depth = 0, leg2Width = 0, sectionDepth = 0, sectionWidth = 0;
                if (sectionData.Orient == (FrameAssemblyServices.SteelOrientationAngle)0 || sectionData.Orient == (FrameAssemblyServices.SteelOrientationAngle)180)
                {
                    sectionDepth = section.depth;
                    sectionWidth = section.width;
                }
                else
                {
                    sectionDepth = section.width;
                    sectionWidth = section.depth;
                }
                if (leg1Data.Orient == (FrameAssemblyServices.SteelOrientationAngle)0 || leg1Data.Orient == (FrameAssemblyServices.SteelOrientationAngle)180)
                {
                    leg1Depth = leg1.depth;
                    leg1Width = leg1.width;
                }
                else
                {
                    leg1Depth = leg1.width;
                    leg1Width = leg1.depth;
                }
                if (leg2Data.Orient == (FrameAssemblyServices.SteelOrientationAngle)0 || leg2Data.Orient == (FrameAssemblyServices.SteelOrientationAngle)180)
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
                int frameOrientation = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOrientation", "FrameOrientation")).PropValue;
                Boolean includeInsulation = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue;
                Boolean mirrorFrame = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsMirrorFrame", "MirrorFrame")).PropValue;

                FrameAssemblyServices.CreateFrameBoundingBox(this, boundingBoxName, (FrameAssemblyServices.FrameBBXOrientation)frameOrientation, includeInsulation, mirrorFrame, SupportedHelper.IsSupportedObjectVertical(1, 45));

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
                    if ((RefPortHelper.DistanceBetweenPorts("BBFrame_Low", "Structure", PortAxisType.Y)) < (RefPortHelper.DistanceBetweenPorts("BBFrame_Low", "Struct_2", PortAxisType.Y)))
                    {
                        // Corner Frame, Orientation will be towards the primary supporting object
                        structurePort[0] = "Structure";
                        structureIndex[0] = 1;
                        structurePort[1] = "Struct_2";
                        structureIndex[1] = 2;
                    }
                    else
                    {
                        structurePort[0] = "Struct_2";
                        structureIndex[0] = 1;
                        structurePort[1] = "Structure";
                        structureIndex[1] = 1;
                    }
                }
                else
                {
                    structurePort[0] = "Structure";
                    structureIndex[0] = 1;
                    structurePort[1] = "Structure";
                    structureIndex[1] = 1;
                }
                // Check if the Support is placed on a Surface, such as Equipment
                Boolean isPlacedOnSurface = false;
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByReference && SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.GenericSurface)
                        isPlacedOnSurface = true;
                }
                else
                    isPlacedOnSurface = false;
                // ==========================
                // Organize the Welds into three Collections
                // ==========================
                FrameAssemblyServices.WeldData weld = new FrameAssemblyServices.WeldData();
                for (int weldCount = 0; weldCount < weldCollection.Count; weldCount++)
                {
                    weld = weldCollection[weldCount];
                    if ((weld.connection).ToUpper().Equals("C"))
                        frameConnection1Welds.Add(weld);
                    else if ((weld.connection).ToUpper().Equals("F"))
                        frameConnection2Welds.Add(weld);
                    else
                        otherWelds.Add(weld);
                }
                // ==========================
                // Determine the Connections to the Structural Steel
                // ==========================
                FrameAssemblyServices.HSSteelMember supportingSection = new FrameAssemblyServices.HSSteelMember();
                int supportingFace = 0;
                double structWidthOffset = 0;
                int structureConnection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsStructureConn", "StructureConnection")).PropValue;
                if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel"))
                {
                    supportingFace = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                    supportingSection = FrameAssemblyServices.GetSupportingSectionData(this, 1);
                }
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    switch (structureConnection)
                    {
                        case 1: //Normal Connection
                            structWidthOffset = 0;
                            break;
                        case 2:
                            {
                                // Lapped Connection
                                switch (supportingFace)
                                {
                                    case 513:
                                    case 514:
                                        structWidthOffset = -supportingSection.width / 2 - leg1Width / 2;
                                        break;
                                    case 257:
                                    case 258:
                                        structWidthOffset = -supportingSection.depth / 2 - leg1Width / 2;
                                        break;
                                    default:
                                        structWidthOffset = 0;
                                        break;
                                }
                                break;
                            }
                        case 3:
                            {
                                // Reverse Lapped Connection
                                switch (supportingFace)
                                {
                                    case 513:
                                    case 514:
                                        structWidthOffset = supportingSection.width / 2 + leg1Width / 2;
                                        break;
                                    case 257:
                                    case 258:
                                        structWidthOffset = supportingSection.depth / 2 + leg1Width / 2;
                                        break;
                                    default:
                                        structWidthOffset = 0;
                                        break;
                                }
                                break;
                            }
                    }
                    //Attach the Structure Conn to the Structure
                    if ((RefPortHelper.AngleBetweenPorts(structurePort[0], PortAxisType.Y, boundingBoxPort, PortAxisType.X, OrientationAlong.Direct) < (Math.Atan(1) * 4.0) / 2))
                        JointHelper.CreateRigidJoint("-1", structurePort[0], STRUCT_CONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, structWidthOffset, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", structurePort[0], STRUCT_CONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -structWidthOffset, 0);
                }
                else
                {
                    //Offset to be used if Pipe is parallel to steel
                    switch (structureConnection)
                    {
                        case 1: //Normal Connection
                            structWidthOffset = 0;
                            break;
                        case 2:
                            {
                                // Lapped Connection
                                switch (supportingFace)
                                {
                                    case 513:
                                    case 514:
                                        structWidthOffset = supportingSection.width / 2 + leg1Depth / 2;
                                        break;
                                    case 257:
                                    case 258:
                                        structWidthOffset = -supportingSection.depth / 2 + leg1Depth / 2;
                                        break;
                                    default:
                                        structWidthOffset = 0;
                                        break;
                                }
                                break;
                            }
                        case 3:
                            {
                                // Reverse Lapped Connection
                                switch (supportingFace)
                                {
                                    case 513:
                                    case 514:
                                        structWidthOffset = -supportingSection.width / 2 - leg1Depth / 2;
                                        break;
                                    case 257:
                                    case 258:
                                        structWidthOffset = -supportingSection.depth / 2 - leg1Depth / 2;
                                        break;
                                    default:
                                        structWidthOffset = 0;
                                        break;
                                }
                                break;
                            }
                    }
                }
                FrameAssemblyServices.HSSteelMember supportingSection2;
                int supportingFace2, structureConnection2 = 1;
                if (part.SupportsInterface("IJUAhsStructureConn2"))
                    structureConnection2 = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsStructureConn2", "StructureConnection2")).PropValue; ;
                double structWidthOffset2 = 0;
                if (supportingCount == 2)
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByReference && SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Member && SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {
                        supportingFace2 = SupportingHelper.SupportingObjectInfo(2).FaceNumber;
                        supportingSection2 = FrameAssemblyServices.GetSupportingSectionData(this, 2);
                        // Offset to be used if Pipe is parallel to steel
                        switch (structureConnection2)
                        {
                            case 1: //Normal Connection
                                structWidthOffset2 = 0;
                                break;
                            case 2:
                                {
                                    // Lapped Connection
                                    switch (supportingFace2)
                                    {
                                        case 513:
                                            structWidthOffset2 = supportingSection2.width / 2 + leg2Depth / 2;
                                            break;
                                        case 514:
                                            structWidthOffset2 = supportingSection2.width / 2 + leg2Depth / 2;
                                            break;
                                        case 257:
                                            structWidthOffset2 = supportingSection2.depth / 2 + leg2Depth / 2;
                                            break;
                                        case 258:
                                            structWidthOffset2 = supportingSection2.depth / 2 + leg2Depth / 2;
                                            break;
                                        default:
                                            structWidthOffset2 = 0;
                                            break;
                                    }
                                    break;
                                }
                            case 3:
                                {
                                    // Reverse Lapped Connection
                                    switch (supportingFace2)
                                    {
                                        case 513:
                                            structWidthOffset2 = -supportingSection2.width / 2 - leg2Depth / 2;
                                            break;
                                        case 514:
                                            structWidthOffset2 = -supportingSection2.width / 2 - leg2Depth / 2;
                                            break;
                                        case 257:
                                            structWidthOffset2 = -supportingSection2.depth / 2 - leg2Depth / 2;
                                            break;
                                        case 258:
                                            structWidthOffset2 = -supportingSection2.depth / 2 - leg2Depth / 2;
                                            break;
                                        default:
                                            structWidthOffset2 = 0;
                                            break;
                                    }
                                    break;
                                }
                        }
                    }
                }

                // ==========================
                //  Offset 1 
                // ==========================
                double offset1, pipeDiameter, offset2;
                int offset1Definition, offset1Selection, offset2Definition, offset2Selection;
                int routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Left);
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                pipeDiameter = pipeInfo.OutsideDiameter;
                double insulationThickness = pipeInfo.InsulationThickness;

                offset1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset1Def", "Offset1Definition")).PropValue;
                offset1Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset1Sel", "Offset1Selection")).PropValue;
                string offset1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset1Rl", "Offset1Rule")).PropValue;
                if (supportingCount == 1 || offset1Definition == 4 || offset1Definition == 8 || SupportingHelper.SupportingObjectInfo(structureIndex[0]).SupportingObjectType != SupportingObjectType.Member)
                {
                    //Get the Offsets From the Input Attributes
                    if (offset1Selection == 1)
                    {
                        GenericHelper.GetDataByRule(offset1Rule, (BusinessObject)support, out offset1);
                        support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                    }
                    else
                        offset1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOffset1", "Offset1Value")).PropValue;
                    switch (offset1Definition)
                    {
                        case 1: //Inside Steel to Edge of Pipe
                            offset1 = offset1 + leg1Depth / 2;
                            break;
                        case 2:
                            //Center Steel to Edge of Pipe
                            break;
                        case 3:
                            offset1 = offset1 - leg1Depth / 2;
                            break;
                        case 4: //End of Steel to Edge of Pipe
                            break;
                        case 5: //Inside Steel to CL of Pipe
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset1 = offset1 + leg1Depth / 2 - pipeDiameter / 2 - insulationThickness;
                            else
                                offset1 = offset1 + leg1Depth / 2 - pipeDiameter / 2;
                            break;
                        case 6: //Center Steel to CL of Pipe
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset1 = offset1 - pipeDiameter / 2 - insulationThickness;
                            else
                                offset1 = offset1 - pipeDiameter / 2;
                            break;
                        case 7:
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset1 = offset1 - leg1Depth / 2 - pipeDiameter / 2 - insulationThickness;
                            else
                                offset1 = offset1 - leg1Depth / 2 - pipeDiameter / 2;
                            break;
                        case 8:
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset1 = offset1 - pipeDiameter / 2 - insulationThickness;
                            else
                                offset1 = offset1 - pipeDiameter / 2;
                            break;
                        default:
                            offset1 = offset1 + leg1Depth / 2;
                            break;
                    }

                }
                else
                {
                    //Get the Offset From the locations of the two supporting objects
                    //This makes the Offsets an Output Dimension instead of an Input   
                    offset1 = -RefPortHelper.DistanceBetweenPorts("BBFrame_Low", structurePort[0], PortAxisType.Y) + structWidthOffset;
                    switch (offset1Definition)
                    {
                        case 1: //Inside Steel to Edge of Pipe
                            support.SetPropertyValue(offset1 - leg1Depth / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 2: //Center Steel to Edge of Pipe
                            support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 3: //Outside Steel to Edge of Pipe
                            support.SetPropertyValue(offset1 + leg1Depth / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 5:
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                support.SetPropertyValue(offset1 - leg1Depth / 2 + pipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                            else
                                support.SetPropertyValue(offset1 - leg1Depth / 2 + pipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 6:
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                support.SetPropertyValue(offset1 + pipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                            else
                                support.SetPropertyValue(offset1 + pipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 7:
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                support.SetPropertyValue(offset1 + leg1Depth / 2 + pipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                            else
                                support.SetPropertyValue(offset1 + leg1Depth / 2 + pipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        default:
                            support.SetPropertyValue(offset1 - leg1Depth / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                    }
                }
                // ==========================
                //  Offset 2
                // ==========================
                routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Right);
                pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                pipeDiameter = pipeInfo.OutsideDiameter;
                insulationThickness = pipeInfo.InsulationThickness;

                offset2Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Def", "Offset2Definition")).PropValue;
                offset2Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Sel", "Offset2Selection")).PropValue;
                string offset2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Rl", "Offset2Rule")).PropValue;
                if (supportingCount == 1 || offset2Definition == 4 || offset2Definition == 8 || SupportingHelper.SupportingObjectInfo(structureIndex[1]).SupportingObjectType != SupportingObjectType.Member)
                {
                    //Get the Offsets From the Input Attributes
                    if (offset2Selection == 1)
                    {
                        GenericHelper.GetDataByRule(offset2Rule, (BusinessObject)support, out offset2);
                        support.SetPropertyValue(offset2, "IJUAhsFrameOffset2", "Offset2Value");
                    }
                    else
                        offset2 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOffset2", "Offset2Value")).PropValue;
                    switch (offset2Definition)
                    {
                        case 1: //Inside Steel to Edge of Pipe
                            offset2 = offset2 + leg2Depth / 2;
                            break;
                        case 2:
                            //Center Steel to Edge of Pipe
                            break;
                        case 3:
                            offset2 = offset2 - leg2Depth / 2;
                            break;
                        case 4: //End of Steel to Edge of Pipe
                            break;
                        case 5: //Inside Steel to CL of Pipe
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset2 = offset2 + leg2Depth / 2 - pipeDiameter / 2 - insulationThickness;
                            else
                                offset2 = offset2 + leg2Depth / 2 - pipeDiameter / 2;
                            break;
                        case 6: //Center Steel to CL of Pipe
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset2 = offset2 - pipeDiameter / 2 - insulationThickness;
                            else
                                offset2 = offset2 - pipeDiameter / 2;
                            break;
                        case 7:
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset2 = offset2 - leg2Depth / 2 - pipeDiameter / 2 - insulationThickness;
                            else
                                offset2 = offset2 - leg2Depth / 2 - pipeDiameter / 2;
                            break;
                        case 8:
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset2 = offset2 - pipeDiameter / 2 - insulationThickness;
                            else
                                offset2 = offset2 - pipeDiameter / 2;
                            break;
                        default:
                            offset2 = offset2 + leg2Depth / 2;
                            break;
                    }

                }
                else
                {
                    //Get the Offset From the locations of the two supporting objects
                    //This makes the Offsets an Output Dimension instead of an Input   
                    offset2 = RefPortHelper.DistanceBetweenPorts("BBFrame_High", structurePort[1], PortAxisType.Y) + structWidthOffset2;
                    switch (offset2Definition)
                    {
                        case 1: //Inside Steel to Edge of Pipe
                            support.SetPropertyValue(offset2 - leg2Depth / 2, "IJUAhsFrameOffset2", "Offset2Value");
                            break;
                        case 2: //Center Steel to Edge of Pipe
                            support.SetPropertyValue(offset2, "IJUAhsFrameOffset2", "Offset2Value");
                            break;
                        case 3: //Outside Steel to Edge of Pipe
                            support.SetPropertyValue(offset2 + leg2Depth / 2, "IJUAhsFrameOffset2", "Offset2Value");
                            break;
                        case 5: //Inside Steel to CL of Pipe
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                support.SetPropertyValue(offset2 - leg2Depth / 2 + pipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset2", "Offset2Value");
                            else
                                support.SetPropertyValue(offset2 - leg2Depth / 2 + pipeDiameter / 2, "IJUAhsFrameOffset2", "Offset2Value");
                            break;
                        case 6:
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                support.SetPropertyValue(offset2 + pipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset2", "Offset2Value");
                            else
                                support.SetPropertyValue(offset2 + pipeDiameter / 2, "IJUAhsFrameOffset2", "Offset2Value");
                            break;
                        case 7: //Outside Steel to CL of Pipe
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                support.SetPropertyValue(offset2 + leg2Depth / 2 + pipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset2", "Offset2Value");
                            else
                                support.SetPropertyValue(offset2 + leg2Depth / 2 + pipeDiameter / 2, "IJUAhsFrameOffset2", "Offset2Value");
                            break;
                        default: //Inside Steel to Edge of Pipe
                            support.SetPropertyValue(offset2 - leg2Depth / 2, "IJUAhsFrameOffset2", "Offset2Value");
                            break;
                    }
                }

                //==========================
                // Shoe Height
                //==========================
                double shoeHeight = 0;
                int shoeHeightDefinition, shoeHeightSelection;
                int frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                shoeHeightDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameShoeHeightDef", "ShoeHeightDefinition")).PropValue;
                shoeHeightSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameShoeHeightSel", "ShoeHeightSelection")).PropValue;
                string shoeHeightRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue;

                //Get the Shoe Height From the Input Attributes
                if (shoeHeightSelection == 1)
                {
                    GenericHelper.GetDataByRule(shoeHeightRule, (BusinessObject)support, out shoeHeight);
                    support.SetPropertyValue(shoeHeight, "IJUAhsFrameShoeHeight", "ShoeHeightValue");
                }
                else
                    shoeHeight = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameShoeHeight", "ShoeHeightValue")).PropValue;

                switch (shoeHeightDefinition)
                {
                    case 1: //Edge of Bounding Box
                        break;
                    case 2: //Centerline of Primary Pipe
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            shoeHeight = shoeHeight - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z));
                        else
                            shoeHeight = shoeHeight + (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) - boundingBoxDepth;
                        break;
                    default: //Edge of Bounding Box
                        break;
                }

                //==========================
                //Leg 1 Begin Overhang
                //========================== 
                double leg1BeginOverHang;
                int leg1BeginOverHangDefinition, leg1BeginOverHangSelection;
                leg1BeginOverHangDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1BeginOHDef", "Leg1BeginOverhangDefinition")).PropValue;
                leg1BeginOverHangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1BeginOHSel", "Leg1BeginOverhangSelection")).PropValue;
                string leg1BeginOverHangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1BeginOHRl", "Leg1BeginOverhangRule")).PropValue;
                if (leg1BeginOverHangSelection == 1)
                {
                    GenericHelper.GetDataByRule(leg1BeginOverHangRule, (BusinessObject)support, out leg1BeginOverHang);
                    support.SetPropertyValue(leg1BeginOverHang, "IJUAhsLeg1BeginOH", "Leg1BeginOverhangValue");
                }
                else
                    leg1BeginOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsLeg1BeginOH", "Leg1BeginOverhangValue")).PropValue;
                switch (leg1BeginOverHangDefinition)
                {
                    case 1: //External Edge of Steel
                        break;
                    case 2: //External Edge of Steel
                        leg1BeginOverHang = leg1BeginOverHang - sectionDepth / 2;
                        break;
                    case 3:
                        leg1BeginOverHang = leg1BeginOverHang - sectionDepth;
                        break;
                    case 4: //Primary Pipe Center Line (DEPENDS ON SHOE HEIGHT)
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            leg1BeginOverHang = leg1BeginOverHang - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) - shoeHeight - sectionDepth;
                        else
                            leg1BeginOverHang = leg1BeginOverHang - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) + boundingBoxDepth + shoeHeight;
                        break;
                    default: //Default to External Edge of Steel
                        break;
                }

                //==========================
                //Leg 2 Begin Overhang
                //========================== 
                double leg2BeginOverHang;
                int leg2BeginOverHangDefinition, leg2BeginOverHangSelection;
                leg2BeginOverHangDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg2BeginOHDef", "Leg2BeginOverhangDefinition")).PropValue;
                leg2BeginOverHangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg2BeginOHSel", "Leg2BeginOverhangSelection")).PropValue;
                string leg2BeginOverHangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg2BeginOHRl", "Leg2BeginOverhangRule")).PropValue;
                if (leg2BeginOverHangSelection == 1)
                {
                    GenericHelper.GetDataByRule(leg2BeginOverHangRule, (BusinessObject)support, out leg2BeginOverHang);
                    support.SetPropertyValue(leg2BeginOverHang, "IJUAhsLeg2BeginOH", "Leg2BeginOverhangValue");
                }
                else
                    leg2BeginOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsLeg2BeginOH", "Leg2BeginOverhangValue")).PropValue;
                switch (leg2BeginOverHangDefinition)
                {
                    case 1: //External Edge of Steel
                        break;
                    case 2: //External Edge of Steel
                        leg2BeginOverHang = leg2BeginOverHang - sectionDepth / 2;
                        break;
                    case 3:
                        leg2BeginOverHang = leg2BeginOverHang - sectionDepth;
                        break;
                    case 4: //Primary Pipe Center Line (DEPENDS ON SHOE HEIGHT)
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            leg2BeginOverHang = leg2BeginOverHang - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) - shoeHeight - sectionDepth;
                        else
                            leg2BeginOverHang = leg2BeginOverHang - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) + boundingBoxDepth + shoeHeight;
                        break;
                    default: //Default to External Edge of Steel
                        break;
                }

                //==========================
                // Leg 1 End Overhang
                //==========================
                double leg1EndOverHang;
                int leg1EndOverHangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1EndOHSel", "Leg1EndOverhangSelection")).PropValue;
                string leg1EndOverHangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1EndOHRl", "Leg1EndOverhangRule")).PropValue;
                //!!! Need Face Position Info !!!   NOT DONE IN THIS ITERATION
                if (leg1EndOverHangSelection == 1)
                {
                    GenericHelper.GetDataByRule(leg1EndOverHangRule, (BusinessObject)support, out leg1EndOverHang);
                    support.SetPropertyValue(leg1EndOverHang, "IJUAhsLeg1EndOH", "Leg1EndOverhangValue");
                }
                else
                    leg1EndOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsLeg1EndOH", "Leg1EndOverhangValue")).PropValue;

                //==========================
                // Leg 2 End Overhang
                //==========================
                double leg2EndOverHang;
                int leg2EndOverHangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg2EndOHSel", "Leg2EndOverhangSelection")).PropValue;
                string leg2EndOverHangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg2EndOHRl", "Leg2EndOverhangRule")).PropValue;
                //!!! Need Face Position Info !!!   NOT DONE IN THIS ITERATION
                if (leg2EndOverHangSelection == 1)
                {
                    GenericHelper.GetDataByRule(leg2EndOverHangRule, (BusinessObject)support, out leg2EndOverHang);
                    support.SetPropertyValue(leg2EndOverHang, "IJUAhsLeg2EndOH", "Leg2EndOverhangValue");
                }
                else
                    leg2EndOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsLeg2EndOH", "Leg2EndOverhangValue")).PropValue;

                //==========================
                // Member 1 Begin Overhang
                //==========================
                double member1BeginOverhang = 0;
                int member1BeginOverhangDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1BeginOHDef", "Member1BeginOverhangDefinition")).PropValue;
                int member1BeginOverhangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1BeginOHSel", "Member1BeginOverhangSelection")).PropValue;
                string member1BeginOverhangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1BeginOHRl", "Member1BeginOverhangRule")).PropValue;
                if (member1BeginOverhangSelection == 1)
                {
                    GenericHelper.GetDataByRule(member1BeginOverhangRule, (BusinessObject)support, out member1BeginOverhang);
                    support.SetPropertyValue(member1BeginOverhang, "IJUAhsMember1BeginOH", "Member1BeginOverhangValue");
                }
                else
                    member1BeginOverhang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsMember1BeginOH", "Member1BeginOverhangValue")).PropValue;

                switch (member1BeginOverhangDefinition)
                {
                    case 1:
                        // External Edge of Steel
                        break;
                    case 2:
                        // Center Line of Steel
                        member1BeginOverhang = member1BeginOverhang - leg1Depth / 2;
                        break;
                    case 3:
                        // Internal Edge of Steel
                        member1BeginOverhang = member1BeginOverhang - leg1Depth;
                        break;
                    default:
                        //External Edge of Steel
                        break;
                }

                //==========================
                // Member 1 End Overhang
                //==========================
                double member1EndOverhang = 0;
                int member1EndOverhangDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1EndOHDef", "Member1EndOverhangDefinition")).PropValue;
                int member1EndOverhangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1EndOHSel", "Member1EndOverhangSelection")).PropValue;
                string member1EndOverhangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1EndOHRl", "Member1EndOverhangRule")).PropValue;
                if (member1EndOverhangSelection == 1)
                {
                    GenericHelper.GetDataByRule(member1EndOverhangRule, (BusinessObject)support, out member1EndOverhang);
                    support.SetPropertyValue(member1EndOverhang, "IJUAhsMember1EndOH", "Member1EndOverhangValue");
                }
                else
                    member1EndOverhang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsMember1EndOH", "Member1EndOverhangValue")).PropValue;
                switch (member1EndOverhangDefinition)
                {
                    case 1://External Edge of Steel
                        break;
                    case 2: //Center Line of Steel
                        member1EndOverhang = member1EndOverhang - leg2Depth / 2;
                        break;
                    case 3: //Internal Edge of Steel
                        member1EndOverhang = member1EndOverhang - leg2Depth;
                        break;
                    default: //External Edge of Steel
                        break;
                }

                //==========================
                //Handle all the Frame Configurations and Toggles
                //==========================
                PropertyValueCodelist connection1TypeList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn1Type", "Connection1Type");
                PropertyValueCodelist connection2TypeList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn2Type", "Connection2Type");
                double sectionZOffset;

                frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                    sectionZOffset = 0 - shoeHeight;
                else
                    sectionZOffset = boundingBoxDepth + sectionDepth + shoeHeight;

                Boolean reflectMember, reflectLeg1, reflectLeg2;
                if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                {
                    reflectMember = false;
                    reflectLeg1 = false;
                    reflectLeg2 = false;
                }
                else
                {
                    reflectMember = true;
                    if (connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                        reflectLeg1 = true;
                    else
                        reflectLeg1 = false;
                    if (connection2TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                        reflectLeg2 = true;
                    else
                        reflectLeg2 = false;
                }
                //==========================
                // Set the Frame Outputs as per their Definitions
                //==========================
                // SPAN
                int spanDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameSpanDef", "SpanDefinition")).PropValue;
                switch (spanDefinition)
                {
                    case 1: //Inside Edge of Steel
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 - leg1Depth / 2 - leg2Depth / 2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    case 2: //Center Line of Steel
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    case 3: //Outside Edge of Steel
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 + leg1Depth / 2 + leg2Depth / 2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    case 4: // Steel Length
                        if (offset1Definition == 4 || offset1Definition == 8)
                            support.SetPropertyValue(boundingBoxWidth + offset1 + offset2, "IJUAhsFrameSpan", "SpanValue");
                        else
                            support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 + member1BeginOverhang + member1EndOverhang, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    default: // Inside Edge of Steel
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 - leg1Depth / 2 - leg2Depth / 2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                }

                //LENGTH1 will be set in the BOM, because they are determined by the DCM Constraint Solver
                //==========================
                // Set the Frame Connections
                //==========================
                //Set Connection between Leg1 and Section
                Boolean connectionSwap = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn1Swap", "Connection1Swap")).PropValue;
                Boolean connection2Swap = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn2Swap", "Connection2Swap")).PropValue;
                int connection1Type = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn1Type", "Connection1Type")).PropValue;
                Boolean connection1Mirror = (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn1Mirror", "Connection1Mirror")).PropValue;

                int connection2Type = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn2Type", "Connection2Type")).PropValue;
                Boolean connection2Mirror = (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn2Mirror", "Connection2Mirror")).PropValue;

                if (reflectMember == false)
                {
                    if (offset1Definition == 4 || offset1Definition == 8)
                        if (connectionSwap)
                            FrameAssemblyServices.SetSteelConnection(this, LEG1, "BeginCap", ref leg1Data, leg1BeginOverHang, SECTION, "BeginCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)90, member1BeginOverhang + leg1Depth / 2, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        else
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, LEG1, "BeginCap", ref leg1Data, 0, (FrameAssemblyServices.SteelConnectionAngle)270, leg1BeginOverHang + sectionDepth / 2, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                    else
                    {
                        if (connectionSwap)
                            FrameAssemblyServices.SetSteelConnection(this, LEG1, "BeginCap", ref leg1Data, leg1BeginOverHang, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        else
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, LEG1, "BeginCap", ref leg1Data, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                    }
                    //Set Connection between Leg2 and Section
                    if (offset1Definition == 4 || offset1Definition == 8)
                    {
                        if (connectionSwap)
                            FrameAssemblyServices.SetSteelConnection(this, LEG2, "EndCap", ref leg2Data, leg2BeginOverHang, SECTION, "EndCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)270, -(member1EndOverhang + leg2Depth / 2), (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                        else
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1EndOverhang, LEG2, "EndCap", ref leg2Data, 0, (FrameAssemblyServices.SteelConnectionAngle)90, -(leg2BeginOverHang + sectionDepth / 2), (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                    }
                    else
                    {
                        if (connection2Swap)
                            FrameAssemblyServices.SetSteelConnection(this, LEG2, "EndCap", ref leg2Data, leg2BeginOverHang, SECTION, "EndCap", ref sectionData, member1EndOverhang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                        else
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1EndOverhang, LEG2, "EndCap", ref leg2Data, leg2BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                    }
                }
                else
                {
                    //Set Connection between Leg1 and Section
                    if (offset1Definition == 4 || offset1Definition == 8)
                    {
                        if (connectionSwap)
                        {
                            if (reflectLeg1 == false)
                                FrameAssemblyServices.SetSteelConnection(this, LEG1, "BeginCap", ref leg1Data, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)270, -(member1BeginOverhang + leg1Depth / 2), (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, LEG1, "EndCap", ref leg1Data, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)90, member1BeginOverhang + leg1Depth / 2, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        }
                        else
                        {
                            if (reflectLeg1 == false)
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG1, "BeginCap", ref leg1Data, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG1, "EndCap", ref leg1Data, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        }
                    }
                    else
                    {
                        if (connectionSwap)
                        {
                            if (reflectLeg1 == false)
                                FrameAssemblyServices.SetSteelConnection(this, LEG1, "BeginCap", ref leg1Data, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, LEG1, "EndCap", ref leg1Data, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        }
                        else
                        {
                            if (reflectLeg1 == false)
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG1, "BeginCap", ref leg1Data, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG1, "EndCap", ref leg1Data, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        }
                    }
                    //Set Connection between Leg2 and Section
                    if (offset1Definition == 4 || offset1Definition == 8)
                    {
                        if (connectionSwap)
                        {
                            if (reflectLeg2 == false)
                                FrameAssemblyServices.SetSteelConnection(this, LEG2, "EndCap", ref leg2Data, leg2BeginOverHang, SECTION, "BeginCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)90, (member1EndOverhang + leg2Depth / 2), (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, LEG2, "BeginCap", ref leg2Data, leg2BeginOverHang, SECTION, "BeginCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)270, -(member1EndOverhang + leg2Depth / 2), (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                        }
                        else
                        {
                            if (reflectLeg2 == false)
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1EndOverhang, LEG2, "EndCap", ref leg2Data, leg2BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1EndOverhang, LEG2, "BeginCap", ref leg2Data, leg2BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                        }

                    }
                    else
                    {
                        if (connectionSwap)
                        {
                            if (reflectLeg2 == false)
                                FrameAssemblyServices.SetSteelConnection(this, LEG2, "EndCap", ref leg2Data, leg2BeginOverHang, SECTION, "BeginCap", ref sectionData, member1EndOverhang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, LEG2, "BeginCap", ref leg2Data, leg2BeginOverHang, SECTION, "BeginCap", ref sectionData, member1EndOverhang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                        }
                        else
                        {
                            if (reflectLeg2 == false)
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1EndOverhang, LEG2, "EndCap", ref leg2Data, leg2BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1EndOverhang, LEG2, "BeginCap", ref leg2Data, leg2BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection2Type, connection2Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection2Welds);
                        }
                    }

                }
                //==========================
                // Set Steel CPs
                //==========================
                // Set the CP for the SECTION Flex Ports (Used to connect the main section to the BBX)
                PropertyValueCodelist cardinalPoint66Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint67Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist beginCutbackAnchorPointLeg1 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG1], "IJOAhsCutback", "BeginCutbackAnchorPoint");

                PropertyValueCodelist endCutbackAnchorPoint = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG1], "IJOAhsCutback", "EndCutbackAnchorPoint");
                PropertyValueCodelist cardinalPoint6Leg1 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG1], "IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7Leg1 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG1], "IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist cardinalPoint2Leg1 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG1], "IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Leg1 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG1], "IJOAhsSteelCP", "CP1");

                PropertyValueCodelist beginCutbackAnchorPointLeg2 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG2], "IJOAhsCutback", "BeginCutbackAnchorPoint");
                PropertyValueCodelist endCutbackAnchorPointLeg2 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG2], "IJOAhsCutback", "EndCutbackAnchorPoint");
                PropertyValueCodelist cardinalPoint6Leg2 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG2], "IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7Leg2 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG2], "IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist cardinalPoint2Leg2 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG2], "IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Leg2 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG2], "IJOAhsSteelCP", "CP1");

                switch (sectionData.Orient)
                {
                    case 0:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 2, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 8, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)90:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 6, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 4, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)180:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 8, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 2, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)270:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 4, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint66Section.PropValue = 6, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint67Section.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                }
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsFlexPort", "FlexPortRotZ");
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsEndFlexPort", "EndFlexPortRotZ");

                //Set the CP for the Leg Ports that connect to the supporting Structure


                if (reflectLeg1 == false)
                {
                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint2Leg1.PropValue = leg1Data.CardinalPoint, "IJOAhsSteelCP", "CP2");
                    componentDictionary[LEG1].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(leg1Data.Orient) * Math.PI / 180), "IJOAhsEndCap", "EndCapRotZ");
                }
                else
                {
                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint1Leg1.PropValue = leg1Data.CardinalPoint, "IJOAhsSteelCP", "CP1");
                    componentDictionary[LEG1].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(leg1Data.Orient) * Math.PI / 180), "IJOAhsBeginCap", "BeginCapRotZ");
                }
                if (reflectLeg2 == false)
                {
                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint1Leg2.PropValue = leg2Data.CardinalPoint, "IJOAhsSteelCP", "CP1");
                    componentDictionary[LEG2].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(leg2Data.Orient) * Math.PI / 180), "IJOAhsBeginCap", "BeginCapRotZ");
                }
                else
                {
                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint2Leg2.PropValue = leg2Data.CardinalPoint, "IJOAhsSteelCP", "CP2");
                    componentDictionary[LEG2].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(leg2Data.Orient) * Math.PI / 180), "IJOAhsEndCap", "EndCapRotZ");
                }

                //==========================
                // Joints To Connect the Main Steel Section to the BBX
                //==========================

                componentDictionary[SECTION].SetPropertyValue(boundingBoxWidth + offset1 + offset2, "IJUAHgrOccLength", "Length");

                if (reflectMember == false)
                {
                    componentDictionary[SECTION].SetPropertyValue(offset1, "IJOAhsFlexPort", "FlexPortZOffset");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsEndFlexPort", "EndFlexPortZOffset");

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreatePrismaticJoint(SECTION, "BeginFlex", "-1", boundingBoxPort, Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, sectionZOffset);
                        JointHelper.CreatePointOnPlaneJoint(LEG1, "EndFlex", STRUCT_CONN, "Connection", Plane.ZX);
                    }
                    else
                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                }
                else
                {
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsFlexPort", "FlexPortZOffset");
                    componentDictionary[SECTION].SetPropertyValue(-offset1, "IJOAhsEndFlexPort", "EndFlexPortZOffset");
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreatePrismaticJoint(SECTION, "EndFlex", "-1", boundingBoxPort, Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, 0, -sectionZOffset);
                        if (reflectLeg1 == false)
                            JointHelper.CreatePointOnPlaneJoint(LEG1, "EndFlex", STRUCT_CONN, "Connection", Plane.ZX);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG1, "BeginFlex", STRUCT_CONN, "Connection", Plane.ZX);
                    }
                    else
                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.ZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                }

                //==========================
                // Joints To Connect Leg 1 To Supporting Structure
                //==========================
                //If it is Equipment or a Shape, we will use GetProjectedPointOnSurface
                //For Steel, Slab and Wall, we will use a PointOnJoint

                Vector boundingBox_X = new Vector(0, 0, 0), boundingBox_Y = new Vector(0, 0, 0), boundingBox_Z = new Vector(0, 0, 0), structZ = new Vector(0, 0, 0), structX = new Vector(0, 0, 0), structY = new Vector(0, 0, 0);
                Vector projectXZ = new Vector(0, 0, 0), projectZY = new Vector(0, 0, 0), leg1YOffset = new Vector(0, 0, 0), leg2YOffset = new Vector(0, 0, 0), memberProjectedNormal = new Vector(0, 0, 0), leg1ProjectedNormal = new Vector(0, 0, 0), leg2ProjectedNormal = new Vector(0, 0, 0);
                Position boundingBox_Position = new Position(0, 0, 0), leg2Start = new Position(0, 0, 0), leg1Start = new Position(0, 0, 0), leg1ProjectedPoint = new Position(0, 0, 0), leg2ProjectedPoint = new Position(0, 0, 0);
                Matrix4X4 port = new Matrix4X4();
                double planeAngle = 0, leg1CutbackAngle = 0, leg2CutbackAngle = 0, leg1Length = 0, leg2Length = 0;

                if (isPlacedOnSurface == true)
                {
                    // Get Projection from calculated point
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    boundingBox_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    boundingBox_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    boundingBox_Y = boundingBox_Z.Cross(boundingBox_X);
                    boundingBox_Position.Set(port.Origin.X, port.Origin.Y, port.Origin.Z);

                    leg1YOffset.Set(-boundingBox_Y.X, -boundingBox_Y.Y, -boundingBox_Y.Z);
                    if (offset1Definition == 4 || offset1Definition == 8)
                        leg1YOffset.Length = offset1 - member1BeginOverhang - leg1Depth / 2;
                    else
                        leg1YOffset.Length = offset1;

                    leg1Start = boundingBox_Position.Offset(leg1YOffset);
                    try
                    {
                        BusinessObject SupportingFace = (BusinessObject)support.SupportingFaces.First();
                        SupportingHelper.GetProjectedPointOnSurface(leg1Start, boundingBox_Z, SupportingFace, out leg1ProjectedPoint, out leg1ProjectedNormal);
                        // Get Projection from calculated point
                        frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            leg1Length = leg1Start.DistanceToPoint(leg1ProjectedPoint) + shoeHeight + sectionDepth / 2;
                        else
                            leg1Length = leg1Start.DistanceToPoint(leg1ProjectedPoint) - boundingBoxDepth - shoeHeight - sectionDepth / 2;
                    }
                    catch { }
                    componentDictionary[LEG1].SetPropertyValue(leg1Length, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[0], PortAxisType.Z, OrientationAlong.Direct);
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    boundingBox_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    boundingBox_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    boundingBox_Y = boundingBox_Z.Cross(boundingBox_X);

                    port = new Matrix4X4();
                    port = RefPortHelper.PortLCS(structurePort[0]);
                    structX.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    structZ.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    structY = structZ.Cross(structX);

                    if ((planeAngle * 180 / Math.PI) > 45 && (planeAngle * 180 / Math.PI) < 135)
                    {
                        projectXZ = FrameAssemblyServices.ProjectVectorIntoPlane(structY, boundingBox_Y);
                        if (FrameAssemblyServices.GetVectorProjection(structY, boundingBox_Z) < 0)
                            projectXZ.Set(-projectXZ.X, -projectXZ.Y, -projectXZ.Z);
                        projectZY = FrameAssemblyServices.ProjectVectorIntoPlane(structY, boundingBox_X);
                        if (FrameAssemblyServices.GetVectorProjection(structY, boundingBox_Z) < 0)
                            projectZY.Set(-projectZY.X, -projectZY.Y, -projectZY.Z);
                    }
                    else
                    {
                        projectXZ = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, boundingBox_Y);
                        if (FrameAssemblyServices.GetVectorProjection(structZ, boundingBox_Z) < 0)
                            projectXZ.Set(-projectXZ.X, -projectXZ.Y, -projectXZ.Z);
                        projectZY = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, boundingBox_X);
                        if (FrameAssemblyServices.GetVectorProjection(structZ, boundingBox_Z) < 0)
                            projectZY.Set(-projectZY.X, -projectZY.Y, -projectZY.Z);
                    }

                    if (FrameAssemblyServices.AngleBetweenVectors(boundingBox_Z, projectXZ) > FrameAssemblyServices.AngleBetweenVectors(boundingBox_Z, projectZY))
                    {
                        // Cut Across BBX Plane
                        switch (leg1Data.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint6Leg1.PropValue = 6, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint7Leg1.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg1 == false)
                                        componentDictionary[LEG1].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG1].SetPropertyValue(beginCutbackAnchorPointLeg1.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint6Leg1.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint7Leg1.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg1 == false)
                                        componentDictionary[LEG1].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG1].SetPropertyValue(beginCutbackAnchorPointLeg1.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint6Leg1.PropValue = 4, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint7Leg1.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg1 == false)
                                        componentDictionary[LEG1].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG1].SetPropertyValue(beginCutbackAnchorPointLeg1.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint6Leg1.PropValue = 2, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint7Leg1.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg1 == false)
                                        componentDictionary[LEG1].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG1].SetPropertyValue(beginCutbackAnchorPointLeg1.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                        }
                        if (reflectLeg1 == false)
                            componentDictionary[LEG1].SetPropertyValue(leg1CutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                        else
                            componentDictionary[LEG1].SetPropertyValue(leg1CutbackAngle, "IJOAhsCutback", "CutbackBeginAngle");
                    }
                    else
                    {
                        // Cut in BBX Plane
                        switch (leg1Data.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint6Leg1.PropValue = 8, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint7Leg1.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg1 == false)
                                        componentDictionary[LEG1].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG1].SetPropertyValue(beginCutbackAnchorPointLeg1.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint6Leg1.PropValue = 4, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint7Leg1.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg1 == false)
                                        componentDictionary[LEG1].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG1].SetPropertyValue(beginCutbackAnchorPointLeg1.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint6Leg1.PropValue = 2, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint7Leg1.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg1 == false)
                                        componentDictionary[LEG1].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG1].SetPropertyValue(beginCutbackAnchorPointLeg1.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint6Leg1.PropValue = 6, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG1].SetPropertyValue(cardinalPoint7Leg1.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg1 == false)
                                        componentDictionary[LEG1].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG1].SetPropertyValue(beginCutbackAnchorPointLeg1.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                        }
                        if (reflectLeg1 == false)
                            componentDictionary[LEG1].SetPropertyValue(leg1CutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                        else
                            componentDictionary[LEG1].SetPropertyValue(leg1CutbackAngle, "IJOAhsCutback", "CutbackBeginAngle");
                    }
                    if (reflectLeg1 == false)
                        JointHelper.CreatePrismaticJoint(LEG1, "BeginFlex", LEG1, "EndFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    else
                        JointHelper.CreatePrismaticJoint(LEG1, "EndFlex", LEG1, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                double basePlate1Th = 0;
                if (isBasePlate1)
                    basePlate1Th = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1], "IJUAhsThickness1", "Thickness1")).PropValue;
                else
                    basePlate1Th = 0;

                if (reflectLeg1 == false)
                {
                    componentDictionary[LEG1].SetPropertyValue(leg1EndOverHang - basePlate1Th, "IJUAHgrOccOverLength", "EndOverLength");
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[0], PortAxisType.Z, OrientationAlong.Direct);
                    if (isPlacedOnSurface == false)
                    {
                        if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                            JointHelper.CreatePointOnPlaneJoint(LEG1, "EndFlex", "-1", structurePort[0], Plane.NegativeXY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG1, "EndFlex", "-1", structurePort[0], Plane.NegativeZX);
                    }
                }
                else
                {
                    componentDictionary[LEG1].SetPropertyValue(leg1EndOverHang - basePlate1Th, "IJUAHgrOccOverLength", "BeginOverLength");
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[0], PortAxisType.Z, OrientationAlong.Direct);
                    if (isPlacedOnSurface == false)
                    {
                        if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                            JointHelper.CreatePointOnPlaneJoint(LEG1, "BeginFlex", "-1", structurePort[0], Plane.NegativeXY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG1, "BeginFlex", "-1", structurePort[0], Plane.NegativeZX);
                    }
                }
                //==========================
                // Joints To Connect Leg 2 To Supporting Structure
                //==========================
                //If it is Equipment or a Shape, we will use GetProjectedPointOnSurface
                //For Steel, Slab and Wall, we will use a PointOnJoint
                if (isPlacedOnSurface)
                {
                    //Get Projection from calculated point
                    leg2YOffset.Set(boundingBox_Y.X, boundingBox_Y.Y, boundingBox_Z.Z);
                    if (offset2Definition == 4 || offset2Definition == 8)
                        leg2YOffset.Length = boundingBoxWidth + offset2 - member1EndOverhang - leg2Depth / 2;
                    else
                        leg2YOffset.Length = boundingBoxWidth + offset2;
                    leg2Start = boundingBox_Position.Offset(leg2YOffset);
                    try
                    {
                        BusinessObject SupportingFace = (BusinessObject)support.SupportingFaces.First();
                        SupportingHelper.GetProjectedPointOnSurface(leg2Start, boundingBox_Z, SupportingFace, out leg2ProjectedPoint, out leg2ProjectedNormal);

                        //Get Projection from calculated point
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            leg2Length = leg2Start.DistanceToPoint(leg2ProjectedPoint) + shoeHeight + sectionDepth / 2;
                        else
                            leg2Length = leg2Start.DistanceToPoint(leg2ProjectedPoint) - boundingBoxDepth - shoeHeight - sectionDepth / 2;
                    }
                    catch { }
                    componentDictionary[LEG2].SetPropertyValue(leg2Length, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    port = new Matrix4X4();
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    boundingBox_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    boundingBox_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    boundingBox_Y = boundingBox_Z.Cross(boundingBox_X);

                    port = new Matrix4X4();
                    port = RefPortHelper.PortLCS(structurePort[1]);
                    structX.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    structZ.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    structY = structZ.Cross(structX);

                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[1], PortAxisType.Z, OrientationAlong.Direct);

                    if ((planeAngle * 180 / Math.PI) > 45 && (planeAngle * 180 / Math.PI) < 135)
                    {
                        projectXZ = FrameAssemblyServices.ProjectVectorIntoPlane(structY, boundingBox_Y);
                        if (FrameAssemblyServices.GetVectorProjection(structY, boundingBox_Z) < 0)
                            projectXZ.Set(-projectXZ.X, -projectXZ.Y, -projectXZ.Z);
                        projectZY = FrameAssemblyServices.ProjectVectorIntoPlane(structY, boundingBox_X);
                        if (FrameAssemblyServices.GetVectorProjection(structY, boundingBox_Z) < 0)
                            projectZY.Set(-projectZY.X, -projectZY.Y, -projectZY.Z);
                    }
                    else
                    {
                        projectXZ = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, boundingBox_Y);
                        if (FrameAssemblyServices.GetVectorProjection(structZ, boundingBox_Z) < 0)
                            projectXZ.Set(-projectXZ.X, -projectXZ.Y, -projectXZ.Z);
                        projectZY = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, boundingBox_X);
                        if (FrameAssemblyServices.GetVectorProjection(structZ, boundingBox_Z) < 0)
                            projectZY.Set(-projectZY.X, -projectZY.Y, -projectZY.Z);
                    }

                    if (FrameAssemblyServices.AngleBetweenVectors(boundingBox_Z, projectXZ) > FrameAssemblyServices.AngleBetweenVectors(boundingBox_Z, projectZY))
                    {
                        // Cut Across BBX Plane
                        switch (leg1Data.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 6, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg2 == false)
                                        componentDictionary[LEG2].SetPropertyValue(beginCutbackAnchorPointLeg2.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG2].SetPropertyValue(endCutbackAnchorPointLeg2.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");

                                    leg2CutbackAngle = boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 8, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg2 == false)
                                        componentDictionary[LEG2].SetPropertyValue(beginCutbackAnchorPointLeg2.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG2].SetPropertyValue(endCutbackAnchorPointLeg2.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");

                                    leg2CutbackAngle = -boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 4, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg2 == false)
                                        componentDictionary[LEG2].SetPropertyValue(beginCutbackAnchorPointLeg2.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG2].SetPropertyValue(endCutbackAnchorPointLeg2.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");

                                    leg2CutbackAngle = -boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 2, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg2 == false)
                                        componentDictionary[LEG2].SetPropertyValue(beginCutbackAnchorPointLeg2.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG2].SetPropertyValue(endCutbackAnchorPointLeg2.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");

                                    leg2CutbackAngle = boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                        }

                        if (reflectLeg2 == false)
                            componentDictionary[LEG2].SetPropertyValue(leg2CutbackAngle, "IJOAhsCutback", "CutbackBeginAngle");
                        else
                            componentDictionary[LEG2].SetPropertyValue(leg2CutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                    }
                    else
                    {
                        // Cut in BBX Plane
                        switch (leg2Data.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 8, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg2 == false)
                                        componentDictionary[LEG2].SetPropertyValue(beginCutbackAnchorPointLeg2.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG2].SetPropertyValue(endCutbackAnchorPointLeg2.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");

                                    leg2CutbackAngle = -boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 4, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg2 == false)
                                        componentDictionary[LEG2].SetPropertyValue(beginCutbackAnchorPointLeg2.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG2].SetPropertyValue(endCutbackAnchorPointLeg2.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");

                                    leg2CutbackAngle = -boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 2, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg2 == false)
                                        componentDictionary[LEG2].SetPropertyValue(beginCutbackAnchorPointLeg2.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG2].SetPropertyValue(endCutbackAnchorPointLeg2.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");

                                    leg2CutbackAngle = boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint6Leg2.PropValue = 6, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG2].SetPropertyValue(cardinalPoint7Leg2.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg2 == false)
                                        componentDictionary[LEG2].SetPropertyValue(beginCutbackAnchorPointLeg2.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG2].SetPropertyValue(endCutbackAnchorPointLeg2.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");

                                    leg2CutbackAngle = boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                        }
                        if (reflectLeg2 == false)
                            componentDictionary[LEG2].SetPropertyValue(leg2CutbackAngle, "IJOAhsCutback", "CutbackBeginAngle");
                        else
                            componentDictionary[LEG2].SetPropertyValue(leg2CutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                    }
                    if (reflectLeg2 == false)
                        JointHelper.CreatePrismaticJoint(LEG2, "EndFlex", LEG2, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    else
                        JointHelper.CreatePrismaticJoint(LEG2, "BeginFlex", LEG2, "EndFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                double basePlate2Thickness;
                if (isBasePlate2)
                    basePlate2Thickness = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE2], "IJUAhsThickness1", "Thickness1")).PropValue;
                else
                    basePlate2Thickness = 0;

                if (reflectLeg2 == false)
                {
                    componentDictionary[LEG2].SetPropertyValue(Convert.ToDouble(leg2EndOverHang - basePlate2Thickness), "IJUAHgrOccOverLength", "BeginOverLength");
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[1], PortAxisType.Z, OrientationAlong.Direct);
                    if (isPlacedOnSurface == false)
                    {
                        if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                            JointHelper.CreatePointOnPlaneJoint(LEG2, "BeginFlex", "-1", structurePort[1], Plane.XY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG2, "BeginFlex", "-1", structurePort[1], Plane.ZX);
                    }
                }
                else
                {
                    componentDictionary[LEG2].SetPropertyValue(leg2EndOverHang - basePlate2Thickness, "IJUAHgrOccOverLength", "EndOverLength");
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[1], PortAxisType.Z, OrientationAlong.Direct);
                    if (isPlacedOnSurface == false)
                    {
                        if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                            JointHelper.CreatePointOnPlaneJoint(LEG2, "EndFlex", "-1", structurePort[1], Plane.XY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG2, "EndFlex", "-1", structurePort[1], Plane.ZX);
                    }
                }

                // ==========================
                // Joints For Remaining Optional Plates
                // ==========================
                // Joints for End Plates
                double capPlate1Angle = 0, capPlate2Angle = 0, capPlate3Angle = 0, capPlate4Angle = 0, basePlate1Angle = 0;
                if (isCapPlate1)
                    capPlate1Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate1Ang", "CapPlate1Angle")).PropValue;
                if (isCapPlate2)
                    capPlate2Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate2Ang", "CapPlate2Angle")).PropValue;
                if (isCapPlate3)
                    capPlate3Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate3Ang", "CapPlate3Angle")).PropValue;
                if (isCapPlate4)
                    capPlate4Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate4Ang", "CapPlate4Angle")).PropValue;
                if (isBasePlate1)
                    basePlate1Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1Ang", "BasePlate1Angle")).PropValue;
                double capHorizontalOffset = 0, capVerticalOffset = 0, horizontalOffset = 0, verticalOffset = 0;
                string caplate = "";
                if (isCapPlate3)
                    caplate = CAPPLATE3;
                else
                    caplate = CAPPLATE4;
                if (part.SupportsInterface("IJUAhsCapHorOffset"))
                    capHorizontalOffset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapHorOffset", "CapHorOffset")).PropValue;
                if (part.SupportsInterface("IJUAhsCapVerOffset"))
                    capVerticalOffset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapVerOffset", "CapVerOffset")).PropValue;

                if (isCapPlate3 || isCapPlate4)
                {
                    horizontalOffset = ((double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[caplate], "IJUAhsWidth1", "Width1")).PropValue) / 2 - (sectionWidth / 2 + capHorizontalOffset);
                    verticalOffset = ((double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[caplate], "IJUAhsLength1", "Length1")).PropValue) / 2 - (sectionDepth / 2 + capVerticalOffset);
                }
                if (isCapPlate1)
                {
                    if (reflectLeg1 == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE1, "Port2", LEG1, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate1Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE1, "Port1", LEG1, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate1Angle));
                }
                if (isCapPlate2)
                {
                    if (reflectLeg2 == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE2, "Port1", LEG2, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate2Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE2, "Port2", LEG2, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate2Angle));
                }
                if (isCapPlate3)
                {
                    if (reflectMember == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE3, "Port2", SECTION, "BeginFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate3Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE3, "Port1", SECTION, "EndFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate3Angle));
                }
                if (isCapPlate4)
                {
                    if (reflectMember == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE4, "Port1", SECTION, "EndFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate4Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE4, "Port2", SECTION, "BeginFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate4Angle));
                }

                //Joints for the Base Plates
                if (isBasePlate1)
                {
                    if (reflectLeg2 == false)
                        JointHelper.CreateAngularRigidJoint(BASEPLATE1, "Port1", LEG1, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(BASEPLATE1, "Port2", LEG1, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));
                }
                if (isBasePlate2)
                {
                    if (reflectLeg2 == false)
                        JointHelper.CreateAngularRigidJoint(BASEPLATE2, "Port2", LEG2, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(BASEPLATE2, "Port1", LEG2, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));
                }

                //==========================
                // Joints For the Remaining Weld Objects
                //==========================
                for (int weldCount = 0; weldCount < otherWelds.Count; weldCount++)
                {
                    weld = otherWelds[weldCount];
                    string leg1WeldPort = string.Empty, leg2WeldPort = string.Empty;
                    if (reflectLeg1 == false)
                        leg1WeldPort = "EndFace";
                    else
                        leg1WeldPort = "BeginFace";
                    if (reflectLeg2 == false)
                        leg2WeldPort = "BeginFace";
                    else
                        leg2WeldPort = "EndFace";
                    double length1, width1;
                    switch (weld.connection)
                    {
                        case "A":
                            {
                                if (isBasePlate1)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1], "IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE1, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE1, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE1, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE1, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                        case "B":
                            {
                                switch (weld.location)
                                {
                                    case 2:
                                        {
                                            JointHelper.CreateRigidJoint(LEG1, leg1WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - leg1.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                    case 4:
                                        {
                                            JointHelper.CreateRigidJoint(LEG1, leg1WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - leg2.width / 2);
                                            break;
                                        }
                                    case 6:
                                        {
                                            JointHelper.CreateRigidJoint(LEG1, leg1WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + leg2.width / 2);
                                            break;
                                        }
                                    case 8:
                                        {
                                            JointHelper.CreateRigidJoint(LEG1, leg1WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + leg1.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                }
                                break;
                            }
                        case "D":
                            {
                                if (isCapPlate3)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3], "IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                        case "E":
                            {
                                if (isCapPlate1)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE1], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE1], "IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE1, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE1, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE1, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE1, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                        case "G":
                            {
                                if (isCapPlate4)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE4], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE4], "IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE4, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE4, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE4, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE4, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                        case "H":
                            {
                                if (isCapPlate2)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE2], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE2], "IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE2, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE2, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE2, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE2, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                        case "I":
                            {
                                if (isBasePlate2)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE2], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE2], "IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE2, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE2, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE2, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE2, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                        case "J":
                            {
                                switch (weld.location)
                                {
                                    case 2:
                                        {
                                            JointHelper.CreateRigidJoint(LEG2, leg2WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - leg2.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                    case 4:
                                        {
                                            JointHelper.CreateRigidJoint(LEG2, leg2WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - leg2.width / 2);
                                            break;
                                        }
                                    case 6:
                                        {
                                            JointHelper.CreateRigidJoint(LEG2, leg2WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + leg2.width / 2);
                                            break;
                                        }
                                    case 8:
                                        {
                                            JointHelper.CreateRigidJoint(LEG2, leg2WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + leg2.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                }
                                break;
                            }
                    }
                }

                //==========================
                // Drawing Dimensions Notes and Labels
                //==========================
                double spanOffset1, spanOffset2, lengthOffset1, lengthOffset2;
                //SPAN
                switch (spanDefinition)
                {
                    case 1: //Inside Edge of Steel
                        spanOffset1 = leg1Depth / 2;
                        spanOffset2 = leg2Depth / 2;
                        break;
                    case 2: //Center Line of Steel
                        spanOffset1 = 0;
                        spanOffset2 = 0;
                        break;
                    case 3: //Outside Edge of Steel
                        spanOffset1 = -leg1Depth / 2;
                        spanOffset2 = -leg2Depth / 2;
                        break;
                    case 4: //Steel Length
                        spanOffset1 = -leg1Depth / 2 - member1BeginOverhang;
                        spanOffset2 = -leg2Depth / 2 - member1EndOverhang;
                        break;
                    default: //Inside Edge of Steel
                        spanOffset1 = leg1Depth / 2;
                        spanOffset2 = leg2Depth / 2;
                        break;
                }
                //L1
                int length1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameLength1Def", "Length1Definition")).PropValue;
                switch (length1Definition)
                {
                    case 1: //Internal edge Steel
                        lengthOffset1 = -sectionDepth / 2;
                        break;
                    case 2: //External Edge of Steel
                        lengthOffset1 = sectionDepth / 2;
                        break;
                    case 3: //End of Steel
                        lengthOffset1 = sectionDepth / 2 + leg1BeginOverHang;
                        break;
                    case 4:
                        //Pipe CL
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            lengthOffset1 = -sectionDepth / 2 - shoeHeight - (RefPortHelper.DistanceBetweenPorts("BBFrame_Low", "Route", PortAxisType.Z));
                        else
                            lengthOffset1 = sectionDepth / 2 + shoeHeight + Math.Abs(RefPortHelper.DistanceBetweenPorts("BBFrame_High", "Route", PortAxisType.Z));
                        break;
                    default:
                        lengthOffset1 = -sectionDepth / 2;
                        break;
                }
                int length2Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameLength2Def", "Length2Definition")).PropValue;
                switch (length2Definition)
                {
                    case 1: //Internal edge Steel
                        lengthOffset2 = -sectionDepth / 2;
                        break;
                    case 2: //External Edge of Steel
                        lengthOffset2 = sectionDepth / 2;
                        break;
                    case 3: //End of Steel
                        lengthOffset2 = sectionDepth / 2 + leg2BeginOverHang;
                        break;
                    case 4:
                        //Pipe CL
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            lengthOffset2 = -sectionDepth / 2 - shoeHeight - (RefPortHelper.DistanceBetweenPorts("BBFrame_Low", "Route", PortAxisType.Z));
                        else
                            lengthOffset2 = sectionDepth / 2 + shoeHeight + Math.Abs(RefPortHelper.DistanceBetweenPorts("BBFrame_High", "Route", PortAxisType.Z));
                        break;
                    default:
                        lengthOffset2 = -sectionDepth / 2;
                        break;
                }

                //If lapped offset dim points
                double lapped1Offset = 0;
                double lapped2Offset = 0;
                double length1Offset = 0;
                double length2Offset = 0;
                if ((connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Lapped))
                {
                    lapped1Offset = -(leg1Width / 2 + sectionWidth / 2);
                    lapped2Offset = -(leg2Width / 2 + sectionWidth / 2);
                    length1Offset = lapped1Offset;
                    length2Offset = lapped2Offset;

                    if (connectionSwap == true)
                    {
                        lapped1Offset = 0;
                        lapped2Offset = 0;
                        length1Offset = sectionWidth / 2;
                        length2Offset = sectionWidth / 2;
                    }

                    if (connection1Mirror == true)
                    {
                        lapped1Offset = -lapped1Offset;
                        lapped2Offset = -lapped2Offset;
                        length1Offset = -length1Offset;
                        length2Offset = -length2Offset;
                    }
                }
                else if ((connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Nested))
                {
                    lapped1Offset = leg1.flangeThickness;
                    lapped2Offset = leg2.flangeThickness;
                    length1Offset = lapped1Offset;
                    length2Offset = lapped2Offset;

                    if (connectionSwap == true)
                    {
                        lapped1Offset = 0;
                        lapped2Offset = 0;
                        length1Offset = sectionWidth / 2;
                        length2Offset = sectionWidth / 2;
                    }

                    if (connection1Mirror == true)
                    {
                        lapped1Offset = -lapped1Offset;
                        lapped2Offset = -lapped2Offset;
                        length1Offset = -length1Offset;
                        length2Offset = -length2Offset;
                    }
                }

                //Set Dimension Points
                ControlPoint controlPoint;
                Note note;
                Boolean excludeNotes = (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                if (offset1Definition == 4 || offset1Definition == 8)
                {
                    if (reflectLeg1 == false)
                        note = CreateNote("SpanStart", LEG1, "BeginCap", new Position(lapped1Offset, spanOffset1, -lengthOffset1), " ", true, 2, 1, out controlPoint);
                    else
                        note = CreateNote("SpanStart", LEG1, "EndCap", new Position(lapped1Offset, -spanOffset1, -lengthOffset1), " ", true, 2, 1, out controlPoint);
                }
                else
                {
                    if (reflectMember == false)
                        note = CreateNote("SpanStart", SECTION, "BeginCap", new Position(lapped1Offset, lengthOffset1, spanOffset1), " ", true, 2, 1, out controlPoint);
                    else
                        note = CreateNote("SpanStart", SECTION, "EndCap", new Position(lapped1Offset, -lengthOffset1, spanOffset1), " ", true, 2, 1, out controlPoint);
                }
                if (offset2Definition == 4 || offset2Definition == 8)
                {
                    if (reflectLeg2 == false)
                        note = CreateNote("SpanEnd", LEG2, "EndCap", new Position(lapped2Offset, spanOffset2, lengthOffset2), " ", true, 2, 1, out controlPoint);
                    else
                        note = CreateNote("SpanEnd", LEG2, "BeginCap", new Position(lapped2Offset, -spanOffset2, lengthOffset2), " ", true, 2, 1, out controlPoint);
                }
                else
                {
                    if (reflectMember == false)
                        note = CreateNote("SpanEnd", SECTION, "EndCap", new Position(lapped2Offset, lengthOffset2, -spanOffset2), " ", true, 2, 1, out controlPoint);
                    else
                        note = CreateNote("SpanEnd", SECTION, "BeginCap", new Position(lapped2Offset, -lengthOffset2, -spanOffset2), " ", true, 2, 1, out controlPoint);
                }

                if (reflectLeg1 == false)
                    note = CreateNote("Length1End", LEG1, "EndCap", new Position(length1Offset, spanOffset1, 0), " ", true, 2, 1, out controlPoint);
                else
                    note = CreateNote("Length1End", LEG1, "BeginCap", new Position(length1Offset, -spanOffset1, 0), " ", true, 2, 1, out controlPoint);

                if (reflectLeg2 == false)
                    note = CreateNote("Length2End", LEG2, "BeginCap", new Position(length2Offset, spanOffset2, 0), " ", true, 2, 1, out controlPoint);
                else
                    note = CreateNote("Length1End", LEG2, "EndCap", new Position(length2Offset, -spanOffset2, 0), " ", true, 2, 1, out controlPoint);

                //Offset 1 & 2
                double offsetToLeftPipeCL, offsetToRightPipeCL;
                int leftPipeIndex, rightPipeIndex;
                string leftPipePort, rightPipePort;
                leftPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Left);
                if (leftPipeIndex == 1)
                    leftPipePort = "Route";
                else
                    leftPipePort = "Route_" + leftPipeIndex.ToString();

                offsetToLeftPipeCL = RefPortHelper.DistanceBetweenPorts(boundingBoxPort, leftPipePort, PortAxisType.Y);
                rightPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Right);
                if (rightPipeIndex == 1)
                    rightPipePort = "Route";
                else
                    rightPipePort = "Route_" + rightPipeIndex.ToString();

                offsetToRightPipeCL = RefPortHelper.DistanceBetweenPorts(boundingBoxPort, rightPipePort, PortAxisType.Y);

                if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                {
                    note = CreateNote("Pipe1", "-1", boundingBoxPort, new Position(0, offsetToLeftPipeCL, -shoeHeight), " ", true, 2, 1, out controlPoint);
                    if (SupportHelper.SupportedObjects.Count > 1)
                        note = CreateNote("Pipe2", "-1", boundingBoxPort, new Position(0, offsetToRightPipeCL, -shoeHeight), " ", true, 2, 1, out controlPoint);
                    else
                        DeleteNoteIfExists("Pipe2");
                }
                else
                {
                    note = CreateNote("Pipe1", "-1", boundingBoxPort, new Position(0, offsetToLeftPipeCL, boundingBoxDepth + shoeHeight), " ", true, 2, 1, out controlPoint);
                    if (SupportHelper.SupportedObjects.Count > 1)
                        note = CreateNote("Pipe2", "-1", boundingBoxPort, new Position(0, offsetToRightPipeCL, boundingBoxDepth + shoeHeight), " ", true, 2, 1, out controlPoint);
                    else
                        DeleteNoteIfExists("Pipe2");
                }

                //Elevation
                if (excludeNotes == false)
                {
                    if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                    {
                        if (reflectMember == false)
                            note = CreateNote("Elevation", SECTION, "BeginFlex", new Position(0, 0, 0), "Elevation", true, 2, 51, out controlPoint);
                        else
                            note = CreateNote("Elevation", SECTION, "EndFlex", new Position(0, 0, 0), "Elevation", true, 2, 51, out controlPoint);
                    }
                    else
                    {
                        if (reflectMember == false)
                            note = CreateNote("Elevation", SECTION, "BeginFlex", new Position(0, sectionDepth, 0), "Elevation", true, 2, 51, out controlPoint);
                        else
                            note = CreateNote("Elevation", SECTION, "EndFlex", new Position(0, -sectionDepth, 0), "Elevation", true, 2, 51, out controlPoint);
                    }
                }
                else
                    DeleteNoteIfExists("Elevation");

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
                    if (isBasePlate1)
                        structConnections.Add(new ConnectionInfo(BASEPLATE1, 1)); // partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(LEG1, 1));

                    if (isBasePlate2)
                        structConnections.Add(new ConnectionInfo(BASEPLATE2, 1)); // partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(LEG2, 1));
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
                // Get Section Information of each steel part
                IPart part;
                double sectionDepth = 0, sectionWidth = 0;
                part = (IPart)componentDictionary[SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                section = FrameAssemblyServices.GetSectionDataFromPart(part.PartNumber);
                if ((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue == 0 || (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue == 180)
                {
                    sectionDepth = section.depth;
                    sectionWidth = section.width;
                }
                else
                {
                    sectionDepth = section.width;
                    sectionWidth = section.depth;
                }
                string supportNumber = string.Empty, L1 = string.Empty, L2 = string.Empty, span = string.Empty;
                double length1Value = 0, length2Value = 0, spanValue = 0, shoeHeight = 0;
                // ==========================
                // Shoe Height
                // ==========================
                int shoeHeightDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsFrameShoeHeightDef", "ShoeHeightDefinition")).PropValue;
                string shoeHeightRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue;
                if (!string.IsNullOrEmpty(shoeHeightRule))
                {
                    GenericHelper.GetDataByRule((string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue, null, out shoeHeight);
                    SupportOrComponent.SetPropertyValue(shoeHeight, "IJUAhsFrameShoeHeight", "ShoeHeightValue");
                }
                else
                    shoeHeight = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAhsFrameShoeHeight", "ShoeHeightValue")).PropValue;

                int frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                switch (shoeHeightDefinition)
                {
                    case 1:
                        // Edge of Bounding Box
                        break;
                    case 2:
                        {
                            // Centerline of Primary Pipe
                            if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                                shoeHeight = shoeHeight - RefPortHelper.DistanceBetweenPorts("BBFrame_Low", "Route", PortAxisType.Z);
                            else
                                shoeHeight = shoeHeight - Math.Abs(RefPortHelper.DistanceBetweenPorts("BBFrame_High", "Route", PortAxisType.Z));
                        }
                        break;
                }

                // ==========================
                // Span
                // ==========================
                // If it is a Corner Frame, then the Span is determined by the length of the member

                spanValue = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAhsFrameSpan", "SpanValue")).PropValue;
                supportNumber = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsSupportNumber", "SupportNumber")).PropValue;
                // ==========================
                // Length 1
                // ==========================
                length1Value = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG1], "IJUAHgrOccLength", "Length")).PropValue;
                int length1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsFrameLength1Def", "Length1Definition")).PropValue;
                switch (length1Definition)
                {
                    case 1:
                        length1Value = length1Value - sectionDepth / 2;
                        break;
                    case 2:
                        length1Value = length1Value + sectionDepth / 2;
                        break;
                    case 3:
                        {
                            double beginOverLength = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                            length1Value = length1Value + Convert.ToDouble(beginOverLength);
                        }
                        break;
                    case 4:
                        {
                            if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                                length1Value = length1Value - sectionDepth / 2 - shoeHeight - RefPortHelper.DistanceBetweenPorts("BBFrame_Low", "Route", PortAxisType.Z);
                            else
                                length1Value = length1Value + sectionDepth / 2 + shoeHeight + Math.Abs(RefPortHelper.DistanceBetweenPorts("BBFrame_High", "Route", PortAxisType.Z));
                        }
                        break;
                    default:
                        length1Value = length1Value - sectionDepth / 2;
                        break;
                }
                SupportOrComponent.SetPropertyValue(length1Value, "IJUAhsFrameLength1", "Length1Value");

                // ==========================
                // Length 2
                // ==========================
                length2Value = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG2], "IJUAHgrOccLength", "Length")).PropValue;
                int length2Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsFrameLength2Def", "Length2Definition")).PropValue;
                switch (length2Definition)
                {
                    case 1:
                        length2Value = length2Value - sectionDepth / 2;
                        break;
                    case 2:
                        length2Value = length2Value + sectionDepth / 2;
                        break;
                    case 3:
                        {
                            double beginOverLength = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG2], "IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                            length2Value = length2Value + Convert.ToDouble(beginOverLength);
                        }
                        break;
                    case 4:
                        {
                            if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                                length2Value = length2Value - sectionDepth / 2 - shoeHeight - RefPortHelper.DistanceBetweenPorts("BBFrame_Low", "Route", PortAxisType.Z);
                            else
                                length2Value = length2Value + sectionDepth / 2 + shoeHeight + Math.Abs(RefPortHelper.DistanceBetweenPorts("BBFrame_High", "Route", PortAxisType.Z));
                        }
                        break;
                    default:
                        length2Value = length2Value - sectionDepth / 2;
                        break;
                }
                SupportOrComponent.SetPropertyValue(length2Value, "IJUAhsFrameLength2", "Length2Value");

                // Populate the BOM String
                Collection<object> collection = new Collection<object>();
                GenericHelper.GetDataByRule("FrameBOMPrimaryDistanceUnits", SupportOrComponent, out collection);
                UnitName primaryUnits = (UnitName)collection[0];
                GenericHelper.GetDataByRule("FrameBOMSecondaryDistanceUnits", SupportOrComponent, out collection);
                UnitName secondaryUnits = (UnitName)collection[0];
                GenericHelper.GetDataByRule("FrameBOMDistancePrecisionType", SupportOrComponent, out collection);
                PrecisionType precisionType = (PrecisionType)collection[0];
                GenericHelper.GetDataByRule("FrameBOMDistancePrecision", SupportOrComponent, out collection);
                int precision = (int)collection[0];

                span = FrameAssemblyServices.FormatValueWithUnits(SupportOrComponent, "IJUAhsFrameSpan", "SpanValue", spanValue, primaryUnits, secondaryUnits, precisionType, precision);
                L1 = FrameAssemblyServices.FormatValueWithUnits(SupportOrComponent, "IJUAhsFrameLength1", "Length1Value", length1Value, primaryUnits, secondaryUnits, precisionType, precision);
                L2 = FrameAssemblyServices.FormatValueWithUnits(SupportOrComponent, "IJUAhsFrameLength2", "Length2Value", length2Value, primaryUnits, secondaryUnits, precisionType, precision);
                PropertyValueCodelist steelStandardList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsSteelStandard", "SteelStandard");

                bomDesciption = supportNumber + "," + steelStandardList.PropertyInfo.CodeListInfo.GetCodelistItem(steelStandardList.PropValue).ShortDisplayName + "," + section.sectionName + ", Span =" + span + ", L1=" + L1 + ",L2=" + L2;

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
