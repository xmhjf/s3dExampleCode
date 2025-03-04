//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AK1.cs
//   Frame_Assy,Ingr.SP3D.Content.Support.Rules.LFrame
//   Author       :  Rajeswari
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 31-Jul-2013  Rajeswari  CR-CP-224474- Convert HS_S3DFrame to C# .Net 
// 30-Mar-2014     PVK     CR-CP-245789	Modify the exsisting .Net Frame_Assy to behave like URS Frame supports
// 27-Apr-2015     PVK     TR-CP-253033	Elevation CP not shown by default for frame supports.
// 06-May-2015     PVK     CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
// 16-Jul-2015     PVK     Resolve coverity issues found in July 2015 report
// 17/12/2015      Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
// 25/05/2016      PVK     TR-CP-292743	Corrected the VerticalOffset attribute
// 07-Jun-2016     PVK     TR-CP-293408	Delivered HS_S3DAssy L Frame and Corner Frames Fail with Cap Plate 3
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

    public class LFrame : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string SECTION = "SECTION";
        private const string LEG = "LEG";
        private const string CAPPLATE1 = "CAPPLATE1";
        private const string CAPPLATE2 = "CAPPLATE2";
        private const string CAPPLATE3 = "CAPPLATE3";
        private const string BASEPLATE1 = "BASEPLATE1";
        private const string STRUCTCONN = "STRUCTCONN";

        FrameAssemblyServices.HSSteelMember section = new FrameAssemblyServices.HSSteelMember();
        FrameAssemblyServices.HSSteelMember leg = new FrameAssemblyServices.HSSteelMember();
        Collection<FrameAssemblyServices.WeldData> weldCollection = new Collection<FrameAssemblyServices.WeldData>();
        int[] structureIndex = new int[2];
        bool isCapPlate1, isCapPlate2, isCapPlate3, isBasePlate1, isbolt1, isbolt2;
        BusinessObject part = null;
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
                    part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

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
                    FrameAssemblyServices.AddPart(this, SECTION, member1Part, member1Rule, parts);

                    string leg1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1", "Leg1Part")).PropValue;
                    string leg1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1Rl", "Leg1Rule")).PropValue;
                    FrameAssemblyServices.AddPart(this, LEG, leg1Part, leg1Rule, parts);

                    // Add the Plates
                    string capPlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate1", "CapPlate1Part")).PropValue;
                    string capPlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate1Rl", "CapPlate1Rule")).PropValue;
                    isCapPlate1 = FrameAssemblyServices.AddPart(this, CAPPLATE1, capPlate1Part, capPlate1Rule, parts);
                    string capPlate2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate2", "CapPlate2Part")).PropValue;
                    string capPlate2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate2Rl", "CapPlate2Rule")).PropValue;
                    isCapPlate2 = FrameAssemblyServices.AddPart(this, CAPPLATE2, capPlate2Part, capPlate2Rule, parts);
                    string capPlate3Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate3", "CapPlate3Part")).PropValue;
                    string capPlate3Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate3Rl", "CapPlate3Rule")).PropValue;
                    isCapPlate3 = FrameAssemblyServices.AddPart(this, CAPPLATE3, capPlate3Part, capPlate3Rule, parts);

                    string basePlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1", "BasePlate1Part")).PropValue;
                    string basePlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1Rl", "BasePlate1Rule")).PropValue;
                    isBasePlate1 = FrameAssemblyServices.AddPart(this, BASEPLATE1, basePlate1Part, basePlate1Rule, parts);

                    // Add the Structure Connection Object if it is Place-By-Structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        parts.Add(new PartInfo(STRUCTCONN, "Log_Conn_Part_1"));

                    // Add the Weld Objects from the Weld Sheet
                    string weldServiceClassName = string.Empty;
                    if (part.SupportsInterface("IJUAhsWeldServClass"))
                        weldServiceClassName = (String)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsWeldServClass", "WeldServiceClassName")).PropValue;
                    if (weldServiceClassName == null)
                        weldServiceClassName = string.Empty;
                    weldCollection = FrameAssemblyServices.AddWeldsFromCatalog(this, parts, "IJUAhsLFrameWelds", ((IPart)part).PartNumber, weldServiceClassName);

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
        public override ReadOnlyCollection<PartInfo> ImpliedParts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> impliedParts = new Collection<PartInfo>();
                    part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
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
                        string bolt1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1", "Bolt1Part")).PropValue;
                        string bolt1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1Rl", "Bolt1Rule")).PropValue;
                        bolt1Quantity = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1Qty", "Bolt1Quantity")).PropValue;
                        isbolt1 = FrameAssemblyServices.AddImpliedPart(this, bolt1Part, bolt1Rule, impliedParts, null, bolt1Quantity);
                    }
                    if (isCapPlate3)
                    {
                        string bolt2Part = "", bolt2Rule = "";
                        if (part.SupportsInterface("IJUAhsBolt2"))
                            bolt2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt2", "Bolt2Part")).PropValue;
                        if (part.SupportsInterface("IJUAhsBolt2Rl"))
                            bolt2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt2Rl", "Bolt2Rule")).PropValue;
                        if (part.SupportsInterface("IJUAhsBolt2Qty"))
                            bolt2Quantity = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt2Qty", "Bolt2Quantity")).PropValue;
                        isbolt2 = FrameAssemblyServices.AddImpliedPart(this, bolt2Part, bolt2Rule, impliedParts, null, bolt2Quantity);
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
                part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                FrameAssemblyServices.HSSteelConfig sectionData = new FrameAssemblyServices.HSSteelConfig();
                FrameAssemblyServices.HSSteelConfig legData = new FrameAssemblyServices.HSSteelConfig();

                Collection<FrameAssemblyServices.WeldData> frameConnection1Welds = new Collection<FrameAssemblyServices.WeldData>(); // Welds at the connection between Leg and Main Member
                Collection<FrameAssemblyServices.WeldData> otherWelds = new Collection<FrameAssemblyServices.WeldData>(); // Other welds in the support
                // ==========================
                // Get Required Information about the Steel Parts
                // ==========================
                // Get the Steel Cross Section Data
                section = FrameAssemblyServices.GetSectionDataFromPartIndex(this, SECTION);
                leg = FrameAssemblyServices.GetSectionDataFromPartIndex(this, LEG);
                // Get the Steel Configuration Data
                sectionData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue);
                legData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue);

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
                double legDepth = 0, legWidth = 0, sectionDepth = 0, sectionWidth = 0;
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
                if (legData.Orient == (FrameAssemblyServices.SteelOrientationAngle)0 || legData.Orient == (FrameAssemblyServices.SteelOrientationAngle)180)
                {
                    legDepth = leg.depth;
                    legWidth = leg.width;
                }
                else
                {
                    legDepth = leg.width;
                    legWidth = leg.depth;
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
                    // Corner Frame, Orientation will be towards the primary supporting object
                    structurePort[0] = "Structure";
                    structureIndex[0] = 1;
                    structurePort[1] = "Struct_2";
                    structureIndex[1] = 2;
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
                // Organize the Welds into two Collections
                // ==========================
                FrameAssemblyServices.WeldData weld = new FrameAssemblyServices.WeldData();
                int weldCount = 0;
                for (weldCount = 0; weldCount < weldCollection.Count; weldCount++)
                {
                    weld = weldCollection[weldCount];
                    if ((weld.connection).ToUpper() == "C")
                        frameConnection1Welds.Add(weld);
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
                        case 1:
                            structWidthOffset = 0;
                            break;
                        case 2:
                            {
                                // Lapped Connection
                                switch (supportingFace)
                                {
                                    case 513:
                                        structWidthOffset = -supportingSection.width / 2 - legWidth / 2;
                                        break;
                                    case 514:
                                        structWidthOffset = -supportingSection.width / 2 - legWidth / 2;
                                        break;
                                    case 257:
                                        structWidthOffset = -supportingSection.depth / 2 - legWidth / 2;
                                        break;
                                    case 258:
                                        structWidthOffset = -supportingSection.depth / 2 - legWidth / 2;
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
                                        structWidthOffset = supportingSection.width / 2 + legWidth / 2;
                                        break;
                                    case 514:
                                        structWidthOffset = supportingSection.width / 2 + legWidth / 2;
                                        break;
                                    case 257:
                                        structWidthOffset = supportingSection.depth / 2 + legWidth / 2;
                                        break;
                                    case 258:
                                        structWidthOffset = supportingSection.depth / 2 + legWidth / 2;
                                        break;
                                    default:
                                        structWidthOffset = 0;
                                        break;
                                }
                                break;
                            }
                    }
                    // Attach the Structure Conn to the Structure
                    if (RefPortHelper.AngleBetweenPorts(structurePort[0], PortAxisType.Y, boundingBoxPort, PortAxisType.X, OrientationAlong.Direct) < Math.Atan(1) * 4.0 / 2)
                        JointHelper.CreateRigidJoint("-1", structurePort[0], STRUCTCONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, structWidthOffset, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", structurePort[0], STRUCTCONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -structWidthOffset, 0);
                }
                else
                {
                    // Offset to be used if Pipe is parallel to steel
                    switch (structureConnection)
                    {
                        case 1:
                            structWidthOffset = 0;
                            break;
                        case 2:
                            {
                                // Lapped Connection
                                switch (supportingFace)
                                {
                                    case 513:
                                        structWidthOffset = supportingSection.width / 2 + legDepth / 2;
                                        break;
                                    case 514:
                                        structWidthOffset = supportingSection.width / 2 + legDepth / 2;
                                        break;
                                    case 257:
                                        structWidthOffset = supportingSection.depth / 2 + legDepth / 2;
                                        break;
                                    case 258:
                                        structWidthOffset = supportingSection.depth / 2 + legDepth / 2;
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
                                        structWidthOffset = -supportingSection.width / 2 - legDepth / 2;
                                        break;
                                    case 514:
                                        structWidthOffset = -supportingSection.width / 2 - legDepth / 2;
                                        break;
                                    case 257:
                                        structWidthOffset = -supportingSection.depth / 2 - legDepth / 2;
                                        break;
                                    case 258:
                                        structWidthOffset = -supportingSection.depth / 2 - legDepth / 2;
                                        break;
                                    default:
                                        structWidthOffset = 0;
                                        break;
                                }
                                break;
                            }
                    }
                }

                // ==========================
                //  Offset 1 (The offset to the Leg)
                // ==========================
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                double offset1, structAngle, leftPipeDiameter, rightPipeDiameter, offset2;
                int offset1Definition, offset1Selection, offset2Definition, offset2Selection;
                int routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Left);
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                leftPipeDiameter = pipeInfo.OutsideDiameter;
                double insulationThickness = pipeInfo.InsulationThickness;

                offset1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset1Def", "Offset1Definition")).PropValue;
                //lOffset1Selection = HH.GetAttr("Offset1Selection", , , False)
                offset1Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset1Sel", "Offset1Selection")).PropValue;
                Boolean flipFrame = false;
                structAngle = (RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.X, structurePort[0], PortAxisType.X, OrientationAlong.Direct) * 180 / Math.PI);
                if (SupportHelper.PlacementType != PlacementType.PlaceByReference && supportingCount == 1 && (supportingType == "Steel") && (structAngle < 45 || structAngle > 135))
                {
                    if (RefPortHelper.DistanceBetweenPorts("BBFrame_Low", structurePort[0], PortAxisType.Y) < 0)
                        offset1 = -RefPortHelper.DistanceBetweenPorts("BBFrame_Low", structurePort[0], PortAxisType.Y) + structWidthOffset;
                    else
                    {
                        flipFrame = true;
                        offset1 = RefPortHelper.DistanceBetweenPorts("BBFrame_High", structurePort[0], PortAxisType.Y) + structWidthOffset;
                    }
                    switch (offset1Definition)
                    {
                        case 1:
                            support.SetPropertyValue(offset1 - legDepth / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 2:
                            support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 3:
                            support.SetPropertyValue(offset1 + legDepth / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 4:
                            support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                        case 5:
                            {
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    support.SetPropertyValue(offset1 - legDepth / 2 + leftPipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                                else
                                    support.SetPropertyValue(offset1 - legDepth / 2 + leftPipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            }
                            break;
                        case 6:
                            {
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    support.SetPropertyValue(offset1 + leftPipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                                else
                                    support.SetPropertyValue(offset1 + leftPipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            }
                            break;
                        case 7:
                            {
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    support.SetPropertyValue(offset1 + legDepth / 2 + leftPipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                                else
                                    support.SetPropertyValue(offset1 + legDepth / 2 + leftPipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            }
                            break;
                        case 8:
                            {
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    support.SetPropertyValue(offset1 + leftPipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                                else
                                    support.SetPropertyValue(offset1 + leftPipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            }
                            break;
                        default:
                            support.SetPropertyValue(offset1 - legDepth / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            break;
                    }

                }
                else
                {
                    if (offset1Selection == 1)
                    {
                        string offset1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOffset1Rl", "Offset1Rule")).PropValue;
                        GenericHelper.GetDataByRule(offset1Rule, (BusinessObject)support, out offset1);
                        support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                    }
                    else
                        offset1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOffset1", "Offset1Value")).PropValue;
                    switch (offset1Definition)
                    {
                        case 1:
                            // Inside Steel to Edge of Pipe
                            offset1 = offset1 + legDepth / 2;
                            break;
                        case 2:
                            // Center Steel to Edge of Pipe
                            break;
                        case 3:
                            /// Outside Steel to Edge of Pipe
                            offset1 = offset1 - legDepth / 2;
                            break;
                        case 4:
                            // End of Steel to Edge of Pipe
                            break;
                        case 5:
                            {
                                // Inside Steel to CL of Pipe
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    offset1 = offset1 + legDepth / 2 - leftPipeDiameter / 2 - insulationThickness;
                                else
                                    offset1 = offset1 + legDepth / 2 - leftPipeDiameter / 2;
                            }
                            break;
                        case 6:
                            {
                                // Center Steel to CL of Pipe
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    offset1 = offset1 - leftPipeDiameter / 2 - insulationThickness;
                                else
                                    offset1 = offset1 - leftPipeDiameter / 2;
                            }
                            break;
                        case 7:
                            {
                                // Outside Steel to CL of Pipe
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    offset1 = offset1 - legDepth / 2 - leftPipeDiameter / 2 - insulationThickness;
                                else
                                    offset1 = offset1 - legDepth / 2 - leftPipeDiameter / 2;
                            }
                            break;
                        case 8:
                            {
                                // End of Steel to CL of Pipe
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    offset1 = offset1 - leftPipeDiameter / 2 - insulationThickness;
                                else
                                    offset1 = offset1 - leftPipeDiameter / 2;
                            }
                            break;
                        default:
                            // Inside Steel to Edge of Pipe
                            offset1 = offset1 + legDepth / 2;
                            break;
                    }
                }
                // ==========================
                //  Offset 2
                // ==========================
                routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Right);
                pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                rightPipeDiameter = pipeInfo.OutsideDiameter;
                insulationThickness = pipeInfo.InsulationThickness;

                offset2Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Def", "Offset2Definition")).PropValue;
                offset2Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Sel", "Offset2Selection")).PropValue;

                //Get the Offset From the Input Attributes
                string offset2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Rl", "Offset2Rule")).PropValue;
                if (supportingCount == 2)
                {
                    if (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, structurePort[1], PortAxisType.Y) > 0)
                        offset2 = Math.Abs(RefPortHelper.DistanceBetweenPorts("BBFrame_High", structurePort[1], PortAxisType.Y));
                    else
                    {
                        flipFrame = true;
                        offset2 = Math.Abs(RefPortHelper.DistanceBetweenPorts("BBFrame_Low", structurePort[1], PortAxisType.Y));
                    }

                    if (offset2Definition == 2)
                    {
                        if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                            support.SetPropertyValue(offset2 + rightPipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset2", "Offset2Value");
                        else
                            support.SetPropertyValue(offset2 + rightPipeDiameter / 2, "IJUAhsFrameOffset2", "Offset2Value");
                    }
                    else
                        support.SetPropertyValue(offset2, "IJUAhsFrameOffset2", "Offset2Value");
                }
                else
                {
                    if (offset2Selection == 1)
                    {
                        GenericHelper.GetDataByRule(offset2Rule, (BusinessObject)support, out offset2);
                        support.SetPropertyValue(offset2, "IJUAhsFrameOffset2", "Offset2Value");
                    }
                    else
                        offset2 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOffset2", "Offset2Value")).PropValue;

                    switch (offset2Definition)
                    {
                        case 1:
                            // Edge of Pipe
                            break;
                        case 2:
                            {
                                // Pipe Center Line
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    offset2 = offset2 - rightPipeDiameter / 2 - insulationThickness;
                                else
                                    offset2 = offset2 - rightPipeDiameter / 2;
                            }
                            break;
                        default:
                            // Edge of Pipe
                            break;
                    }
                }
                //==========================
                // Shoe Height
                //==========================
                double shoeHeight;
                int shoeHeightDefinition, shoeHeightSelection, frameConfiguration;
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
                        {
                            frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                            if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                                shoeHeight = shoeHeight - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z));
                            else
                                shoeHeight = shoeHeight + (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) - boundingBoxDepth;
                            break;
                        }
                }

                //==========================
                //Leg 1 Begin Overhang
                //========================== 
                double leg1BeginOverHang;
                int leg1BeginOverHangDefinition, leg1BeginOverHangSelection;
                leg1BeginOverHangDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1BeginOHDef", "Leg1BeginOverhangDefinition")).PropValue;
                leg1BeginOverHangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1BeginOHSel", "Leg1BeginOverhangSelection")).PropValue;
                string legBeginOverHangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1BeginOHRl", "Leg1BeginOverhangRule")).PropValue;
                if (leg1BeginOverHangSelection == 1)
                {
                    GenericHelper.GetDataByRule(legBeginOverHangRule, (BusinessObject)support, out leg1BeginOverHang);
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
                        {
                            frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                            if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                                leg1BeginOverHang = leg1BeginOverHang - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) - shoeHeight - sectionDepth;
                            else
                                leg1BeginOverHang = leg1BeginOverHang - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) + boundingBoxDepth + shoeHeight;
                            break;
                        }
                }
                //==========================
                // Leg 1 End Overhang
                //==========================
                double leg1EndOverHang;
                int leg1EndOverHangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1EndOHSel", "Leg1EndOverhangSelection")).PropValue;
                string leg1EndOverHangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1EndOHRl", "Leg1EndOverhangRule")).PropValue;
                if (structureConnection != 1)
                {
                    if (leg1EndOverHangSelection == 1)
                    {
                        GenericHelper.GetDataByRule(leg1EndOverHangRule, (BusinessObject)support, out leg1EndOverHang);
                        support.SetPropertyValue(leg1EndOverHang, "IJUAhsLeg1EndOH", "Leg1EndOverhangValue");
                    }
                    else
                        leg1EndOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsLeg1EndOH", "Leg1EndOverhangValue")).PropValue;

                    switch (supportingFace)
                    {
                        case 513:
                            leg1EndOverHang = supportingSection.depth + leg1EndOverHang;
                            break;
                        case 514:
                            leg1EndOverHang = supportingSection.depth + leg1EndOverHang;
                            break;
                        case 257:
                            leg1EndOverHang = supportingSection.webThickness / 2 + supportingSection.width / 2 + leg1EndOverHang;
                            break;
                        case 258:
                            leg1EndOverHang = supportingSection.webThickness / 2 + supportingSection.width / 2 + leg1EndOverHang;
                            break;
                    }
                }
                else
                {
                    if (leg1EndOverHangSelection == 1)
                    {
                        GenericHelper.GetDataByRule(leg1EndOverHangRule, (BusinessObject)support, out leg1EndOverHang);
                        support.SetPropertyValue(leg1EndOverHang, "IJUAhsLeg1EndOH", "Leg1EndOverhangValue");
                    }
                    else
                    {
                        leg1EndOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsLeg1EndOH", "Leg1EndOverhangValue")).PropValue;
                    }
                }

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
                        member1BeginOverhang = member1BeginOverhang - legDepth / 2;
                        break;
                    case 3:
                        // Internal Edge of Steel
                        member1BeginOverhang = member1BeginOverhang - legDepth;
                        break;
                    default:
                        //External Edge of Steel
                        break;
                }

                //==========================
                // Member 1 End Overhang
                //==========================
                double member1EndOverhang = 0;
                int member1EndOverhangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1EndOHSel", "Member1EndOverhangSelection")).PropValue;
                string member1EndOverhangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1EndOHRl", "Member1EndOverhangRule")).PropValue;
                if (member1EndOverhangSelection == 1)
                {
                    GenericHelper.GetDataByRule(member1EndOverhangRule, (BusinessObject)support, out member1EndOverhang);
                    support.SetPropertyValue(member1EndOverhang, "IJUAhsMember1EndOH", "Member1EndOverhangValue");
                }
                else
                    member1EndOverhang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsMember1EndOH", "Member1EndOverhangValue")).PropValue;

                //==========================
                //Handle all the Frame Configurations and Toggles
                //==========================
                PropertyValueCodelist connection1TypeList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn1Type", "Connection1Type");
                double sectionZOffset;
                Boolean reflectMember, reflectLeg;
                frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                {
                    sectionZOffset = 0 - shoeHeight;
                    reflectMember = false;
                    reflectLeg = false;
                }
                else
                {
                    sectionZOffset = boundingBoxDepth + sectionDepth + shoeHeight;
                    reflectMember = true;
                    if (connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                        reflectLeg = true;
                    else
                        reflectLeg = false;
                }
                //==========================
                // Set the Frame Outputs as per their Definitions
                //==========================
                // SPAN
                int spanDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameSpanDef", "SpanDefinition")).PropValue;
                if (supportingCount == 1)
                {
                    switch (spanDefinition)
                    {
                        case 1: //Inside Edge of Steel
                            if (offset1Definition == 4 || offset1Definition == 8)
                                support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 - legDepth - member1BeginOverhang, "IJUAhsFrameSpan", "SpanValue");
                            else
                                support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 - legDepth / 2, "IJUAhsFrameSpan", "SpanValue");
                            break;
                        case 2: //Center Line of Steel
                            support.SetPropertyValue(boundingBoxWidth + offset1 + offset2, "IJUAhsFrameSpan", "SpanValue");
                            break;
                        case 3: //Outside Edge of Steel
                            support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 + legDepth / 2, "IJUAhsFrameSpan", "SpanValue");
                            break;
                        case 4: //Inside Steel to External Pipe CL
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                            {
                                if (offset1Definition == 4 || offset1Definition == 8)
                                    support.SetPropertyValue(boundingBoxWidth + offset2 - legDepth - member1BeginOverhang - rightPipeDiameter / 2 - insulationThickness, "IJUAhsFrameSpan", "SpanValue");
                                else
                                    support.SetPropertyValue(boundingBoxWidth + offset2 - legDepth / 2 - rightPipeDiameter / 2 - insulationThickness, "IJUAhsFrameSpan", "SpanValue");
                            }
                            else
                            {
                                if (offset1Definition == 4 || offset1Definition == 8)
                                    support.SetPropertyValue(boundingBoxWidth + offset2 - legDepth - member1BeginOverhang - rightPipeDiameter / 2, "IJUAhsFrameSpan", "SpanValue");
                                else
                                    support.SetPropertyValue(boundingBoxWidth + offset2 - legDepth / 2 - rightPipeDiameter / 2, "IJUAhsFrameSpan", "SpanValue");
                            }
                            break;
                        case 5: //Center Steel to External Pipe CL
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                support.SetPropertyValue(boundingBoxWidth + offset2 - rightPipeDiameter / 2 - insulationThickness, "IJUAhsFrameSpan", "SpanValue");
                            else
                                support.SetPropertyValue(boundingBoxWidth + offset2 - rightPipeDiameter / 2, "IJUAhsFrameSpan", "SpanValue");
                            break;
                        case 6: //Outside Steel to External Pipe CL
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                support.SetPropertyValue(boundingBoxWidth + offset2 + legDepth / 2 - rightPipeDiameter / 2 - insulationThickness, "IJUAhsFrameSpan", "SpanValue");
                            else
                                support.SetPropertyValue(boundingBoxWidth + offset2 + legDepth / 2 - rightPipeDiameter / 2, "IJUAhsFrameSpan", "SpanValue");
                            break;
                        case 7: // Steel Length
                            {
                                if (offset1Definition == 4 || offset1Definition == 8)
                                    support.SetPropertyValue(boundingBoxWidth + offset1 + offset2, "IJUAhsFrameSpan", "SpanValue");
                                else
                                    support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 + member1BeginOverhang + member1EndOverhang, "IJUAhsFrameSpan", "SpanValue");
                            }
                            break;
                        default: // Inside Edge of Steel
                            support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 - legDepth / 2, "IJUAhsFrameSpan", "SpanValue");
                            break;
                    }
                }
                //LENGTH1 will be set in the BOM, because they are determined by the DCM Constraint Solver
                //==========================
                // Set the Frame Connections
                //==========================
                //Set Connection between Leg1 and Section
                Boolean connectionSwap = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn1Swap", "Connection1Swap")).PropValue;
                int connection1Type = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn1Type", "Connection1Type")).PropValue;
                Boolean connection1Mirror = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCornerConn1Mirror", "Connection1Mirror")).PropValue;
                if (reflectMember == false)
                {
                    if (offset1Definition == 4 || offset1Definition == 8)
                        if (connectionSwap)
                            FrameAssemblyServices.SetSteelConnection(this, LEG, "BeginCap", ref legData, leg1BeginOverHang, SECTION, "BeginCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)90, member1BeginOverhang + legDepth / 2, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        else
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, LEG, "BeginCap", ref legData, 0, (FrameAssemblyServices.SteelConnectionAngle)270, leg1BeginOverHang + sectionDepth / 2, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                    else
                    {
                        if (connectionSwap)
                            FrameAssemblyServices.SetSteelConnection(this, LEG, "BeginCap", ref legData, leg1BeginOverHang, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        else
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, LEG, "BeginCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                    }
                    componentDictionary[SECTION].SetPropertyValue(member1EndOverhang, "IJUAHgrOccOverLength", "EndOverLength");
                    if (connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                        // Set the Cutback of the other end in case of toggle
                        componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsCutback", "CutbackEndAngle");
                }
                else
                {
                    if (offset1Definition == 4 || offset1Definition == 8)
                    {
                        if (connectionSwap)
                        {
                            if (reflectLeg == false)
                                FrameAssemblyServices.SetSteelConnection(this, LEG, "BeginCap", ref  legData, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)270, -(member1BeginOverhang + legDepth / 2), (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, LEG, "EndCap", ref legData, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)90, member1BeginOverhang + legDepth / 2, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        }
                        else
                        {
                            if (reflectLeg == false)
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG, "BeginCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG, "EndCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        }
                    }
                    else
                    {
                        if (connectionSwap)
                        {
                            if (reflectLeg == false)
                                FrameAssemblyServices.SetSteelConnection(this, LEG, "BeginCap", ref legData, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, LEG, "EndCap", ref legData, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        }
                        else
                        {
                            if (reflectLeg == false)
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG, "BeginCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            else
                                FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG, "EndCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                        }
                    }
                    componentDictionary[SECTION].SetPropertyValue(member1EndOverhang, "IJUAHgrOccOverLength", "BeginOverLength");
                    if (connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                        // Set the Cutback of the other end in case of toggle
                        componentDictionary[SECTION].SetPropertyValue(0.0, "IJOAhsCutback", "CutbackBeginAngle");

                }
                //==========================
                // Set Steel CPs
                //==========================
                // Set the CP for the SECTION Flex Ports (Used to connect the main section to the BBX)
                PropertyValueCodelist cardinalPoint6Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCPFlexPort", "CP7");

                switch (sectionData.Orient)
                {
                    case 0:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)90:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)180:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)270:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                }
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(sectionData.Orient) * Math.PI / 180)), "IJOAhsFlexPort", "FlexPortRotZ");
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(sectionData.Orient) * Math.PI / 180)), "IJOAhsEndFlexPort", "EndFlexPortRotZ");

                //Set the CP for the Member End Port that may attach to the supporting Structure
                PropertyValueCodelist cardinalPoint2Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCP", "CP1");

                if (reflectMember == false)
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(sectionData.CardinalPoint), "IJOAhsSteelCP", "CP1");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(sectionData.Orient) * Math.PI / 180)), "IJOAhsEndCap", "EndCapRotZ");
                }
                else
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(sectionData.CardinalPoint), "IJOAhsSteelCP", "CP2");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(sectionData.Orient) * Math.PI / 180)), "IJOAhsBeginCap", "BeginCapRotZ");
                }

                //Set the CP's for the Leg Ports that Connect to the Supporting Structure
                PropertyValueCodelist cardinalPoint2Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsSteelCP", "CP1");
                if (reflectLeg == false)
                {
                    componentDictionary[LEG].SetPropertyValue(cardinalPoint2Leg.PropValue = (int)Convert.ToDouble(legData.CardinalPoint), "IJOAhsSteelCP", "CP2");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(legData.Orient) * Math.PI / 180)), "IJOAhsEndCap", "EndCapRotZ");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(legData.OffsetX) * Math.PI / 180)), "IJOAhsEndCap", "EndCapXOffset");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(legData.OffsetY) * Math.PI / 180)), "IJOAhsEndCap", "EndCapYOffset");
                }
                else
                {
                    componentDictionary[LEG].SetPropertyValue(cardinalPoint1Leg.PropValue = (int)Convert.ToDouble(legData.CardinalPoint), "IJOAhsSteelCP", "CP1");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(legData.Orient) * Math.PI / 180)), "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(legData.OffsetX) * Math.PI / 180)), "IJOAhsBeginCap", "BeginCapXOffset");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble((Convert.ToDouble(legData.OffsetY) * Math.PI / 180)), "IJOAhsBeginCap", "BeginCapYOffset");
                }
                //==========================
                // Joints To Connect the Main Steel Section to the BBX
                //==========================

                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(boundingBoxWidth + offset1 + offset2), "IJUAHgrOccLength", "Length");

                if (reflectMember == false)
                {
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(offset1), "IJOAhsFlexPort", "FlexPortZOffset");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsEndFlexPort", "EndFlexPortZOffset");

                    if (supportingCount == 2)
                    {
                        // Corner Frame
                        if (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, structurePort[1], PortAxisType.Y) > 0)
                            JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                        else
                        {
                            flipFrame = true;
                            JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, sectionZOffset, boundingBoxWidth, 0);
                        }
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            JointHelper.CreatePrismaticJoint(SECTION, "BeginFlex", "-1", boundingBoxPort, Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, sectionZOffset);
                            JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", STRUCTCONN, "Connection", Plane.ZX);
                        }
                        else
                        {
                            if (SupportHelper.PlacementType != PlacementType.PlaceByReference && supportingCount == 1 && (supportingType == "Steel"))
                            {
                                // Single Steel Parallel to the Route
                                if (structAngle < 45 || structAngle > 135)
                                {
                                    if (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, structurePort[0], PortAxisType.Y) < 0)
                                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                                    else
                                    {
                                        flipFrame = true;
                                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeX, sectionZOffset, boundingBoxWidth, 0);
                                    }
                                }
                                else
                                    JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                            }
                            else
                                JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                        }
                    }
                }
                else
                {
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsFlexPort", "FlexPortZOffset");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(-offset1), "IJOAhsEndFlexPort", "EndFlexPortZOffset");
                    if (supportingCount == 2)
                    {
                        // Corner Frame
                        if (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, structurePort[1], PortAxisType.Y) > 0)
                            JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.ZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                        else
                        {
                            flipFrame = true;
                            JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX, sectionZOffset, boundingBoxWidth, 0);
                        }
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        {
                            JointHelper.CreatePrismaticJoint(SECTION, "EndFlex", "-1", boundingBoxPort, Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, 0, -sectionZOffset);
                            if (reflectLeg == false)
                                JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", STRUCTCONN, "Connection", Plane.ZX);
                            else
                                JointHelper.CreatePointOnPlaneJoint(LEG, "BeginFlex", STRUCTCONN, "Connection", Plane.ZX);
                        }
                        else
                        {
                            if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel"))
                            {
                                // Single Steel Parallel to the Route
                                if (structAngle < 45 || structAngle > 135)
                                {
                                    if (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, structurePort[0], PortAxisType.Y) < 0)
                                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.ZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                                    else
                                    {
                                        flipFrame = true;
                                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX, sectionZOffset, boundingBoxWidth, 0);
                                    }
                                }
                                else
                                    JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.ZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                            }
                            else
                                JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.ZX, Axis.X, Axis.X, sectionZOffset, 0, 0);
                        }
                    }
                }
                //==========================
                // Joints To Connect Leg 1 To Supporting Structure
                //==========================
                //If it is Equipment or a Shape, we will use GetProjectedPointOnSurface
                //For Steel, Slab and Wall, we will use a PointOnJoint

                Vector boundingBox_X = new Vector(0, 0, 0), boundingBox_Y = new Vector(0, 0, 0), boundingBox_Z = new Vector(0, 0, 0), structZ = new Vector(0, 0, 0), structX = new Vector(0, 0, 0), structY = new Vector(0, 0, 0);
                Vector projectXZ = new Vector(0, 0, 0), projectZY = new Vector(0, 0, 0), projectXY = new Vector(0, 0, 0), leg1YOffset = new Vector(0, 0, 0), leg2YOffset = new Vector(0, 0, 0), memberProjectedNormal = new Vector(0, 0, 0), leg1ProjectedNormal = new Vector(0, 0, 0), leg2ProjectedNormal = new Vector(0, 0, 0);
                Position boundingBox_Position = new Position(0, 0, 0), leg2Start = new Position(0, 0, 0), leg1Start = new Position(0, 0, 0), leg1ProjectedPoint = new Position(0, 0, 0), leg2ProjectedPoint = new Position(0, 0, 0);
                Matrix4X4 port = new Matrix4X4();
                double planeAngle = 0, leg1CutbackAngle = 0, leg1Length = 0, memberCutbackAngle = 0;

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
                        leg1YOffset.Length = offset1 - member1BeginOverhang - legDepth / 2;
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
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1Length), "IJUAHgrOccLength", "Length");
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

                    PropertyValueCodelist endCutbackAnchorPointList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsCutback", "EndCutbackAnchorPoint");
                    PropertyValueCodelist beginCutbackAnchorPointList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsCutback", "BeginCutbackAnchorPoint");
                    PropertyValueCodelist cardinalPoint6Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsSteelCP", "CP6");
                    PropertyValueCodelist cardinalPoint7Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsSteelCPFlexPort", "CP7");
                    if (FrameAssemblyServices.AngleBetweenVectors(boundingBox_Z, projectXZ) > FrameAssemblyServices.AngleBetweenVectors(boundingBox_Z, projectZY))
                    {
                        // Cut Across BBX Plane
                        switch (legData.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg == false)
                                        componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg == false)
                                        componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg == false)
                                        componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg == false)
                                        componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                        }
                        if (flipFrame)
                            leg1CutbackAngle = -leg1CutbackAngle;
                        if (reflectLeg == false)
                            componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1CutbackAngle), "IJOAhsCutback", "CutbackEndAngle");
                        else
                            componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1CutbackAngle), "IJOAhsCutback", "CutbackBeginAngle");
                    }
                    else
                    {
                        // Cut in BBX Plane
                        switch (legData.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg == false)
                                        componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg == false)
                                        componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg == false)
                                        componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCPFlexPort", "CP7");
                                    if (reflectLeg == false)
                                        componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    else
                                        componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                        }
                        if (flipFrame)
                            leg1CutbackAngle = -leg1CutbackAngle;
                        if (reflectLeg == false)
                            componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1CutbackAngle), "IJOAhsCutback", "CutbackEndAngle");
                        else
                            componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1CutbackAngle), "IJOAhsCutback", "CutbackBeginAngle");
                    }
                    if (reflectLeg == false)
                        JointHelper.CreatePrismaticJoint(LEG, "BeginFlex", LEG, "EndFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    else
                        JointHelper.CreatePrismaticJoint(LEG, "EndFlex", LEG, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                double basePlate1Th = 0;
                if (isBasePlate1)
                    basePlate1Th = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1], "IJUAhsThickness1", "Thickness1")).PropValue;
                else
                    basePlate1Th = 0;

                if (reflectLeg == false)
                {
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1EndOverHang - basePlate1Th), "IJUAHgrOccOverLength", "EndOverLength");
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[0], PortAxisType.Z, OrientationAlong.Direct);
                    if (isPlacedOnSurface == false)
                    {
                        if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                            JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", "-1", structurePort[0], Plane.NegativeXY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", "-1", structurePort[0], Plane.NegativeZX);
                    }
                }
                else
                {
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1EndOverHang - basePlate1Th), "IJUAHgrOccOverLength", "BeginOverLength");
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, structurePort[0], PortAxisType.Z, OrientationAlong.Direct);
                    if (isPlacedOnSurface == false)
                    {
                        if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                            JointHelper.CreatePointOnPlaneJoint(LEG, "BeginFlex", "-1", structurePort[0], Plane.NegativeXY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG, "BeginFlex", "-1", structurePort[0], Plane.NegativeZX);
                    }
                }
                //==========================
                // Connect Member to Supporting Structure if it is a corner frame
                //==========================
                if (supportingCount == 2)
                {
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Y, structurePort[1], PortAxisType.Z, OrientationAlong.Direct);
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    boundingBox_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    boundingBox_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    boundingBox_Y = boundingBox_Z.Cross(boundingBox_X);

                    port = new Matrix4X4();
                    port = RefPortHelper.PortLCS(structurePort[1]);
                    structX.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    structZ.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    structY = structZ.Cross(structX);

                    if ((planeAngle * 180 / Math.PI) > 45 && (planeAngle * 180 / Math.PI) < 135)
                    {
                        projectXY = FrameAssemblyServices.ProjectVectorIntoPlane(structY, boundingBox_Z);
                        if (FrameAssemblyServices.GetVectorProjection(structY, boundingBox_Y) < 0)
                            projectXY.Set(-projectXY.X, -projectXY.Y, -projectXY.Z);
                        projectZY = FrameAssemblyServices.ProjectVectorIntoPlane(structY, boundingBox_X);
                        if (FrameAssemblyServices.GetVectorProjection(structY, boundingBox_Y) < 0)
                            projectZY.Set(-projectZY.X, -projectZY.Y, -projectZY.Z);
                    }
                    else
                    {
                        projectXY = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, boundingBox_Z);
                        if (FrameAssemblyServices.GetVectorProjection(structZ, boundingBox_Y) < 0)
                            projectXY.Set(-projectXY.X, -projectXY.Y, -projectXY.Z);
                        projectZY = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, boundingBox_X);
                        if (FrameAssemblyServices.GetVectorProjection(structZ, boundingBox_Y) < 0)
                            projectZY.Set(-projectZY.X, -projectZY.Y, -projectZY.Z);
                    }

                    PropertyValueCodelist endCutbackAnchorPointSectionList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsCutback", "EndCutbackAnchorPoint");
                    PropertyValueCodelist beginCutbackAnchorPointSectionList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsCutback", "BeginCutbackAnchorPoint");
                    double prismaticOffset = 0;
                    if (FrameAssemblyServices.AngleBetweenVectors(boundingBox_Y, projectXY) > FrameAssemblyServices.AngleBetweenVectors(boundingBox_Z, projectZY))
                    {
                        // Cut Across BBX Plane
                        switch (sectionData.Orient)
                        {
                            case 0:
                                {
                                    if (reflectMember == false)
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP2");
                                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    }
                                    else
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP1");
                                        componentDictionary[SECTION].SetPropertyValue(beginCutbackAnchorPointSectionList.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    }
                                    memberCutbackAngle = -boundingBox_Y.Angle(projectXY, boundingBox_Z);
                                    prismaticOffset = section.width / 2;
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    if (reflectMember == false)
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP2");
                                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    }
                                    else
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP1");
                                        componentDictionary[SECTION].SetPropertyValue(beginCutbackAnchorPointSectionList.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    }
                                    memberCutbackAngle = boundingBox_Y.Angle(projectXY, boundingBox_Z);
                                    prismaticOffset = section.depth / 2;
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    if (reflectMember == false)
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP2");
                                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    }
                                    else
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP1");
                                        componentDictionary[SECTION].SetPropertyValue(beginCutbackAnchorPointSectionList.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    }
                                    memberCutbackAngle = boundingBox_Y.Angle(projectXY, boundingBox_Z);
                                    prismaticOffset = section.width / 2;
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    if (reflectMember == false)
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP2");
                                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    }
                                    else
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP1");
                                        componentDictionary[SECTION].SetPropertyValue(beginCutbackAnchorPointSectionList.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    }
                                    memberCutbackAngle = -boundingBox_Y.Angle(projectXY, boundingBox_Z);
                                    prismaticOffset = section.depth / 2;
                                    break;
                                }
                        }
                        if (flipFrame)
                            memberCutbackAngle = -memberCutbackAngle;
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(memberCutbackAngle), "IJOAhsCutback", "CutbackEndAngle");
                            JointHelper.CreatePrismaticJoint(SECTION, "BeginCap", SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, prismaticOffset);
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(memberCutbackAngle), "IJOAhsCutback", "CutbackBeginAngle");
                            JointHelper.CreatePrismaticJoint(SECTION, "EndCap", SECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, prismaticOffset);
                        }
                    }
                    else
                    {
                        // Cut in BBX Plane
                        switch (sectionData.Orient)
                        {
                            case 0:
                                {
                                    if (reflectMember == false)
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP2");
                                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    }
                                    else
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP1");
                                        componentDictionary[SECTION].SetPropertyValue(beginCutbackAnchorPointSectionList.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    }
                                    memberCutbackAngle = -boundingBox_Y.Angle(projectZY, boundingBox_X);
                                    prismaticOffset = section.depth / 2;
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    if (reflectMember == false)
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP2");
                                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    }
                                    else
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP1");
                                        componentDictionary[SECTION].SetPropertyValue(beginCutbackAnchorPointSectionList.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    }
                                    memberCutbackAngle = -boundingBox_Y.Angle(projectZY, boundingBox_X);
                                    prismaticOffset = section.width / 2;
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    if (reflectMember == false)
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP2");
                                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    }
                                    else
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP1");
                                        componentDictionary[SECTION].SetPropertyValue(beginCutbackAnchorPointSectionList.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    }
                                    memberCutbackAngle = boundingBox_Y.Angle(projectZY, boundingBox_X);
                                    prismaticOffset = section.depth / 2;
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    if (reflectMember == false)
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP2");
                                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    }
                                    else
                                    {
                                        componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP1");
                                        componentDictionary[SECTION].SetPropertyValue(beginCutbackAnchorPointSectionList.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                    }
                                    memberCutbackAngle = boundingBox_Y.Angle(projectZY, boundingBox_X);
                                    prismaticOffset = section.width / 2;
                                    break;
                                }
                        }
                        if (flipFrame)
                            memberCutbackAngle = -memberCutbackAngle;
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(memberCutbackAngle), "IJOAhsCutback", "CutbackEndAngle");
                            JointHelper.CreatePrismaticJoint(SECTION, "BeginCap", SECTION, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, prismaticOffset, 0);
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(memberCutbackAngle), "IJOAhsCutback", "CutbackBeginAngle");
                            JointHelper.CreatePrismaticJoint(SECTION, "EndCap", SECTION, "BeginCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, prismaticOffset, 0);
                        }
                    }

                    double capPlate3Thickness;
                    if (isCapPlate3)
                        capPlate3Thickness = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3], "IJUAhsThickness1", "Thickness1")).PropValue;
                    else
                        capPlate3Thickness = 0;

                    if (reflectMember == false)
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(member1EndOverhang - capPlate3Thickness), "IJUAHgrOccOverLength", "EndOverLength");
                    else
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(member1EndOverhang - capPlate3Thickness), "IJUAHgrOccOverLength", "BeginOverLength");

                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Y, structurePort[1], PortAxisType.Z, OrientationAlong.Direct);

                    if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                    {
                        if (reflectMember == false)
                            JointHelper.CreatePointOnPlaneJoint(SECTION, "EndCap", "-1", structurePort[1], Plane.NegativeXY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(SECTION, "BeginCap", "-1", structurePort[1], Plane.NegativeXY);
                    }
                    else
                    {
                        if (reflectMember == false)
                            JointHelper.CreatePointOnPlaneJoint(SECTION, "EndCap", "-1", structurePort[1], Plane.NegativeZX);
                        else
                            JointHelper.CreatePointOnPlaneJoint(SECTION, "BeginCap", "-1", structurePort[1], Plane.NegativeZX);
                    }
                }

                // ==========================
                // Joints For Remaining Optional Plates
                // ==========================
                // Joints for End Plates
                double capPlate1Angle = 0, capPlate2Angle = 0, capPlate3Angle = 0, basePlate1Angle = 0;
                if (isCapPlate1)
                    capPlate1Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate1Ang", "CapPlate1Angle")).PropValue;
                if (isCapPlate2)
                    capPlate2Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate2Ang", "CapPlate2Angle")).PropValue;
                if (isCapPlate3)
                    capPlate3Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapPlate3Ang", "CapPlate3Angle")).PropValue;
                if (isBasePlate1)
                    basePlate1Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1Ang", "BasePlate1Angle")).PropValue;
                double capHorizontalOffset = 0, capVerticalOffset = 0, horizontalOffset = 0, verticalOffset = 0;
                string caplate = "";
                if (isCapPlate2)
                    caplate = CAPPLATE2;
                else
                    caplate = CAPPLATE3;
                if (part.SupportsInterface("IJUAhsCapHorOffset"))
                    capHorizontalOffset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapHorOffset", "CapHorOffset")).PropValue;
                if (part.SupportsInterface("IJUAhsCapVerOffset"))
                    capVerticalOffset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsCapVerOffset", "CapVerOffset")).PropValue;
                if (isCapPlate2 || isCapPlate3)
                {
                    horizontalOffset = ((double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[caplate], "IJUAhsWidth1", "Width1")).PropValue) / 2 - (sectionWidth / 2 + capHorizontalOffset);
                    verticalOffset = ((double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[caplate], "IJUAhsLength1", "Length1")).PropValue) / 2 - (sectionDepth / 2 + capVerticalOffset);
                }

                if (isCapPlate1)
                {
                    if (reflectLeg == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE1, "Port2", LEG, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate1Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE1, "Port1", LEG, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate1Angle));
                }
                if (isCapPlate2)
                {
                    if (reflectMember == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE2, "Port2", SECTION, "BeginFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate2Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE2, "Port1", SECTION, "EndFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate2Angle));
                }
                if (isCapPlate3)
                {
                    if (reflectMember == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE3, "Port1", SECTION, "EndFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate3Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE3, "Port2", SECTION, "BeginFace", new Vector(horizontalOffset, 0, 0), new Vector(0, 0, capPlate3Angle));
                }

                //Joints for the Base Plates
                if (isBasePlate1)
                {
                    if (reflectLeg == false)
                        JointHelper.CreateAngularRigidJoint(BASEPLATE1, "Port1", LEG, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(BASEPLATE1, "Port2", LEG, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));
                }

                //==========================
                // Joints For the Remaining Weld Objects
                //==========================
                for (weldCount = 0; weldCount < otherWelds.Count; weldCount++)
                {
                    weld = otherWelds[weldCount];
                    string legWeldPort = string.Empty;
                    if (reflectLeg == false)
                        legWeldPort = "EndFace";
                    else
                        legWeldPort = "BeginFace";
                    string memberWeldPort = string.Empty;
                    if (reflectMember == false)
                        memberWeldPort = "EndFace";
                    else
                        memberWeldPort = "BeginFace";
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
                                            JointHelper.CreateRigidJoint(LEG, legWeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - leg.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                    case 4:
                                        {
                                            JointHelper.CreateRigidJoint(LEG, legWeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - leg.width / 2);
                                            break;
                                        }
                                    case 6:
                                        {
                                            JointHelper.CreateRigidJoint(LEG, legWeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + leg.width / 2);
                                            break;
                                        }
                                    case 8:
                                        {
                                            JointHelper.CreateRigidJoint(LEG, legWeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + leg.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                }
                                break;
                            }
                        case "D":
                            {
                                if (isCapPlate2)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE2], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE2], "IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE2, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE2, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE2, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE2, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
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
                        case "F":
                            {
                                switch (weld.location)
                                {
                                    case 2:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, memberWeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - section.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                    case 4:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, memberWeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - section.width / 2);
                                            break;
                                        }
                                    case 6:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, memberWeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + section.width / 2);
                                            break;
                                        }
                                    case 8:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, memberWeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + section.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                }
                                break;
                            }
                        case "G":
                            {
                                if (isCapPlate3)
                                {
                                    string capPlate3WeldPort = string.Empty;
                                    if (reflectMember == false)
                                        capPlate3WeldPort = "Port2";
                                    else
                                        capPlate3WeldPort = "Port1";

                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3], "IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, capPlate3WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, capPlate3WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, capPlate3WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, capPlate3WeldPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
                                                break;
                                            }
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
                        spanOffset1 = legDepth / 2;
                        spanOffset2 = 0;
                        break;
                    case 2: //Center Line of Steel
                        spanOffset1 = 0;
                        spanOffset2 = 0;
                        break;
                    case 3: //Outside Edge of Steel
                        spanOffset1 = -legDepth / 2;
                        spanOffset2 = 0;
                        break;
                    case 4: //Inside Steel to External Pipe CL
                        spanOffset1 = legDepth / 2;
                        if (offset2Definition == 2)
                            spanOffset2 = offset2;
                        else
                        {
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                spanOffset2 = offset2 + rightPipeDiameter / 2 + insulationThickness;
                            else
                                spanOffset2 = offset2 + rightPipeDiameter / 2;
                        }
                        break;
                    case 5: //Center Steel to External Pipe CL
                        spanOffset1 = 0;
                        if (offset2Definition == 2)
                            spanOffset2 = offset2;
                        else
                        {
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                spanOffset2 = offset2 + rightPipeDiameter / 2 + insulationThickness;
                            else
                                spanOffset2 = offset2 + rightPipeDiameter / 2;
                        }
                        break;
                    case 6: //Outside Steel to External Pipe CL
                        spanOffset1 = -legDepth / 2;
                        if (offset2Definition == 2)
                            spanOffset2 = offset2;
                        else
                        {
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                spanOffset2 = offset2 + rightPipeDiameter / 2 + insulationThickness;
                            else
                                spanOffset2 = offset2 + rightPipeDiameter / 2;
                        }
                        break;
                    case 7: //Steel Length
                        spanOffset1 = -legDepth / 2 - member1BeginOverhang;
                        spanOffset2 = -member1EndOverhang;
                        break;
                    default:
                        spanOffset1 = legDepth / 2;
                        spanOffset2 = 0;
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
                lengthOffset2 = lengthOffset1 + sectionDepth / 2;

                //If lapped offset dim points
                double lappedOffset = 0;
                double length1Offset = 0;
                if ((connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Lapped))
                {
                    lappedOffset = -(legWidth / 2 + sectionWidth / 2);
                    length1Offset = lappedOffset;

                    if (connectionSwap == true)
                    {
                        lappedOffset = 0;
                        length1Offset = sectionWidth / 2;
                    }

                    if (connection1Mirror == true)
                    {
                        lappedOffset = -lappedOffset;
                        length1Offset = -length1Offset;
                    }
                }


                //Set Dimension Points
                ControlPoint controlPoint;
                Note note;
                Boolean excludeNotes = (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                if (reflectMember == false)
                    note = CreateNote("SpanStart", SECTION, "BeginCap", new Position(lappedOffset, lengthOffset1, spanOffset1), " ", true, 2, 1, out controlPoint);
                else
                    note = CreateNote("SpanStart", SECTION, "EndCap", new Position(lappedOffset, -lengthOffset1, -spanOffset1), " ", true, 2, 1, out controlPoint);

                if (reflectMember == false)
                    note = CreateNote("SpanEnd", SECTION, "EndFlex", new Position(0, lengthOffset2, -spanOffset2), " ", true, 2, 1, out controlPoint);
                else
                    note = CreateNote("SpanEnd", SECTION, "BeginFlex", new Position(0, -lengthOffset2, spanOffset2), " ", true, 2, 1, out controlPoint);

                if (reflectLeg == false)
                    note = CreateNote("Length1End", LEG, "EndCap", new Position(length1Offset, spanOffset1, 0), " ", true, 2, 1, out controlPoint);
                else
                    note = CreateNote("Length1End", LEG, "BeginCap", new Position(length1Offset, -spanOffset1, 0), " ", true, 2, 1, out controlPoint);

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
                            note = CreateNote("Elevation", SECTION, "BeginFlex", new Position(0, 0, 0), "Elevation", false, 2, 51, out controlPoint);
                        else
                            note = CreateNote("Elevation", SECTION, "EndFlex", new Position(0, 0, 0), "Elevation", false, 2, 51, out controlPoint);
                    }
                    else
                    {
                        if (reflectMember == false)
                            note = CreateNote("Elevation", SECTION, "BeginFlex", new Position(0, sectionDepth, 0), "Elevation", false, 2, 51, out controlPoint);
                        else
                            note = CreateNote("Elevation", SECTION, "EndFlex", new Position(0, -sectionDepth, 0), "Elevation", false, 2, 51, out controlPoint);
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
                    for (int routeIndex = 1; routeIndex <= SupportHelper.SupportedObjects.Count; routeIndex++)
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
                    structConnections.Add(new ConnectionInfo(LEG, structureIndex[0])); // partindex, routeindex

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
                BusinessObject catalogPart = SupportOrComponent.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                // Get Section Information of each steel part
                double sectionDepth = 0, sectionWidth = 0, legDepth = 0, legWidth = 0;
                IPart part = (IPart)componentDictionary[SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                section = FrameAssemblyServices.GetSectionDataFromPart(part.PartNumber);
                if ((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue == 0 || (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue == 180)
                {
                    sectionDepth = section.depth;
                    sectionWidth = section.width;
                }
                else
                {
                    sectionDepth = section.width;
                    sectionWidth = section.depth;
                }
                part = (IPart)componentDictionary[LEG].GetRelationship("madeFrom", "part").TargetObjects[0];
                leg = FrameAssemblyServices.GetSectionDataFromPart(part.PartNumber);
                if ((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue == 0 || (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue == 180)
                {
                    legDepth = leg.depth;
                    legWidth = leg.width;
                }
                else
                {
                    legDepth = leg.width;
                    legWidth = leg.depth;
                }
                string supportNumber = string.Empty, L1 = string.Empty, span = string.Empty;
                double length1Value = 0, spanValue = 0, shoeHeight = 0;
                // ==========================
                // Shoe Height
                // ==========================
                int shoeHeightDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameShoeHeightDef", "ShoeHeightDefinition")).PropValue;
                if (!string.IsNullOrEmpty((string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue))
                {
                    GenericHelper.GetDataByRule((string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue, null, out shoeHeight);
                    SupportOrComponent.SetPropertyValue(shoeHeight, "IJUAhsFrameShoeHeight", "ShoeHeightValue");
                }
                else
                    shoeHeight = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAhsFrameShoeHeight", "ShoeHeightValue")).PropValue;

                int frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
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

                // Check the Supporting count
                int supportingCount = 0;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    supportingCount = 1;
                else
                    supportingCount = SupportHelper.SupportingObjects.Count;

                // ==========================
                // Span
                // ==========================
                // If it is a Corner Frame, then the Span is determined by the length of the member
                if (supportingCount == 2)
                {
                    spanValue = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJUAHgrOccLength", "Length")).PropValue;
                    // Based on Span Definition - Set dSpan
                    int spanDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameSpanDef", "SpanDefinition")).PropValue;
                    switch (spanDefinition)
                    {
                        case 1:
                            spanValue = spanValue - legDepth / 2;
                            break;
                    }
                    SupportOrComponent.SetPropertyValue(spanValue, "IJUAhsFrameSpan", "SpanValue");
                }
                else
                    spanValue = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAhsFrameSpan", "SpanValue")).PropValue;

                supportNumber = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsSupportNumber", "SupportNumber")).PropValue;
                // ==========================
                // Length 1
                // ==========================
                length1Value = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJUAHgrOccLength", "Length")).PropValue;
                int length1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameLength1Def", "Length1Definition")).PropValue;
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
                            double beginOverLength = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
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
                PropertyValueCodelist steelStandardList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsSteelStandard", "SteelStandard");

                bomDesciption = supportNumber + "," + steelStandardList.PropertyInfo.CodeListInfo.GetCodelistItem(steelStandardList.PropValue).ShortDisplayName + "," + section.sectionName + ", Span =" + span + ", L1=" + L1;

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
