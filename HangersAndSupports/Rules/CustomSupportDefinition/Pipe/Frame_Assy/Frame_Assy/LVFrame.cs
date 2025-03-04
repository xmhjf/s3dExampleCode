//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AK1.cs
//   Frame_Assy,Ingr.SP3D.Content.Support.Rules.LVFrame
//   Author       :  Rajeswari
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 3-Aug-2013  Rajeswari   CR-CP-224474- Convert HS_S3DFrame to C# .Net 
// 30-Mar-2014     PVK     CR-CP-245789	Modify the exsisting .Net Frame_Assy to behave like URS Frame supports
// 27-Apr-2015     PVK     TR-CP-253033	Elevation CP not shown by default for frame supports.
// 06-May-2015     PVK     CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;

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

    public class LVFrame : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string SECTION = "SECTION";
        private const string LEG = "LEG";
        private const string CAPPLATE1 = "CAPPLATE1";
        private const string CAPPLATE2 = "CAPPLATE2";
        private const string CAPPLATE3 = "CAPPLATE3";
        private const string BASEPLATE1 = "BASEPLATE1";

        FrameAssemblyServices.HSSteelMember section = new FrameAssemblyServices.HSSteelMember();
        Collection<FrameAssemblyServices.WeldData> weldCollection = new Collection<FrameAssemblyServices.WeldData>();
        bool isCapPlate1, isCapPlate2, isCapPlate3, isBasePlate1, isbolt1, isbolt2;
        BusinessObject part;
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
                        string family = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support,"IJUAHgrURSCommon", "Family")).PropValue;
                        string type = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support,"IJUAHgrURSCommon", "Type")).PropValue;
                        if (family != "" && type != null)
                            FrameAssemblyServices.CheckSupportWithFamilyAndType(this, family, type);
                    }
                    // Add the Steel Section for the Frame Support
                    string member1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1", "Member1Part")).PropValue;
                    string member1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1Rl", "Member1Rule")).PropValue;
                    FrameAssemblyServices.AddPart(this, SECTION, member1Part, member1Rule, parts);

                    string leg1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1", "Leg1Part")).PropValue;
                    string leg1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1Rl", "Leg1Rule")).PropValue;
                    FrameAssemblyServices.AddPart(this, LEG, leg1Part, leg1Rule, parts);

                    // Add the Plates
                    string capPlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1", "CapPlate1Part")).PropValue;
                    string capPlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1Rl", "CapPlate1Rule")).PropValue;
                    isCapPlate1 = FrameAssemblyServices.AddPart(this, CAPPLATE1, capPlate1Part, capPlate1Rule, parts);
                    string capPlate2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate2", "CapPlate2Part")).PropValue;
                    string capPlate2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate2Rl", "CapPlate2Rule")).PropValue;
                    isCapPlate2 = FrameAssemblyServices.AddPart(this, CAPPLATE2, capPlate2Part, capPlate2Rule, parts);
                    string capPlate3Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate3", "CapPlate3Part")).PropValue;
                    string capPlate3Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate3Rl", "CapPlate3Rule")).PropValue;
                    isCapPlate3 = FrameAssemblyServices.AddPart(this, CAPPLATE3, capPlate3Part, capPlate3Rule, parts);

                    string basePlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate1", "BasePlate1Part")).PropValue;
                    string basePlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate1Rl", "BasePlate1Rule")).PropValue;
                    isBasePlate1 = FrameAssemblyServices.AddPart(this, BASEPLATE1, basePlate1Part, basePlate1Rule, parts);

                    // Add the Weld Objects from the Weld Sheet
                    string weldServiceClassName = string.Empty;
                    if (part.SupportsInterface("IJUAhsWeldServClass"))
                        weldServiceClassName = (String)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsWeldServClass", "WeldServiceClassName")).PropValue;
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
                        string bolt1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1", "Bolt1Part")).PropValue;
                        string bolt1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1Rl", "Bolt1Rule")).PropValue;
                        bolt1Quantity = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1Qty", "Bolt1Quantity")).PropValue;
                        isbolt1 = FrameAssemblyServices.AddImpliedPart(this, bolt1Part, bolt1Rule, impliedParts, null, bolt1Quantity);
                    }
                    if (isCapPlate3)
                    {
                        string bolt2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt2", "Bolt2Part")).PropValue;
                        string bolt2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt2Rl", "Bolt2Rule")).PropValue;
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
                FrameAssemblyServices.HSSteelMember leg = new FrameAssemblyServices.HSSteelMember();
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
                sectionData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue);
                legData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue);

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

                //==========================
                // Handle all the Frame Configurations and Toggles
                // ==========================
                Boolean reflectMember, flipFrame, mirrorBBX;
                if (Configuration == 1 || Configuration == 3)
                    reflectMember = false;
                else
                    reflectMember = true;
                if (Configuration == 3 || Configuration == 4)
                    flipFrame = true;
                else
                    flipFrame = false;

                double routeElevation = 0, structElevation = 0;
                Matrix4X4 port = new Matrix4X4();

                port = RefPortHelper.PortLCS("Route");
                routeElevation = port.Origin.Z;

                port = new Matrix4X4();
                port = RefPortHelper.PortLCS("Structure");
                structElevation = port.Origin.Z;

                if (structElevation > routeElevation)
                    mirrorBBX = false;
                else
                    mirrorBBX = true;

                // ==========================
                // Create the Frame Bounding Box
                // ==========================
                double boundingBoxWidth = 0, boundingBoxDepth = 0;
                string boundingBoxPort = "BBFrame_Low", boundingBoxName = "BBFrame";
                int frameOrientation = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOrientation", "FrameOrientation")).PropValue;
                Boolean includeInsulation = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue;
                Boolean rotateFrame = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsRotateFr", "RotateFrame")).PropValue;
                FrameAssemblyServices.CreateVerticalFrameBoundingBox(this, boundingBoxName, (FrameAssemblyServices.FrameBBXOrientation)frameOrientation, includeInsulation, mirrorBBX, rotateFrame);

                boundingBoxWidth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Width;
                boundingBoxDepth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Height;

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

                // Get the Offset to the Supporting Structure for Lapped Connections
                FrameAssemblyServices.HSSteelMember supportingSection = new FrameAssemblyServices.HSSteelMember();
                int supportingFace = 0;
                double structWidthOffset = 0;
                int structureConnection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsStructureConn", "StructureConnection")).PropValue;

                if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel"))
                {
                    supportingFace = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                    supportingSection = FrameAssemblyServices.GetSupportingSectionData(this, 1);
                }

                switch (structureConnection)
                    {
                        case 1:
                            // Normal Connection
                            structWidthOffset = 0;
                            break;
                        case 2:
                            {
                                // Lapped Connection
                                switch (supportingFace)
                                {
                                    case 513:
                                    case 514:
                                        structWidthOffset = supportingSection.width / 2 + legDepth / 2;
                                        break;
                                    case 257:
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
                                    case 514:
                                        structWidthOffset = -supportingSection.width / 2 - legDepth / 2;
                                        break;
                                    case 257:
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

                // ==========================
                //  Offset 1 (The offset to the Leg)
                // ==========================
                double offset1, structAngle, leftPipeDiameter, rightPipeDiameter, offset2;
                int offset1Definition, offset1Selection, offset2Definition, offset2Selection;
                int routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Left);
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                leftPipeDiameter = pipeInfo.OutsideDiameter;
                double insulationThickness = pipeInfo.InsulationThickness;

                offset1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset1Def", "Offset1Definition")).PropValue;
                offset1Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset1Sel", "Offset1Selection")).PropValue;
                structAngle = (RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, "Structure", PortAxisType.X, OrientationAlong.Direct) * 180 / Math.PI);
                if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel") && (structAngle > 45 && structAngle < 135))
                {
                    if (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Structure", PortAxisType.Z) > 0)
                    {
                        flipFrame = false;
                        offset1 = RefPortHelper.DistanceBetweenPorts("BBFrame_Low", "Structure", PortAxisType.Z) - boundingBoxDepth + structWidthOffset;
                    }
                    else
                    {
                        // Should never go here because the BBX allways has Z Axis towards structure
                        flipFrame = true;
                        offset1 = -RefPortHelper.DistanceBetweenPorts("BBFrame_High", "Structure", PortAxisType.Y) + structWidthOffset;
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
                        case 5:
                            {
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    support.SetPropertyValue(offset1 - legDepth / 2 + leftPipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                                else
                                    support.SetPropertyValue(offset1 - legDepth / 2 + leftPipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            }
                            break;
                        case 6:
                            {
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    support.SetPropertyValue(offset1 + leftPipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                                else
                                    support.SetPropertyValue(offset1 + leftPipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
                            }
                            break;
                        case 7:
                            {
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    support.SetPropertyValue(offset1 + legDepth / 2 + leftPipeDiameter / 2 + insulationThickness, "IJUAhsFrameOffset1", "Offset1Value");
                                else
                                    support.SetPropertyValue(offset1 + legDepth / 2 + leftPipeDiameter / 2, "IJUAhsFrameOffset1", "Offset1Value");
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
                        string offset1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOffset1Rl", "Offset1Rule")).PropValue;
                        GenericHelper.GetDataByRule(offset1Rule, (BusinessObject)support, out offset1);
                        support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                    }
                    else
                        offset1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOffset1", "Offset1Value")).PropValue;
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
                        case 5:
                            {
                                // Inside Steel to CL of Pipe
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    offset1 = offset1 + legDepth / 2 - leftPipeDiameter / 2 - insulationThickness;
                                else
                                    offset1 = offset1 + legDepth / 2 - leftPipeDiameter / 2;
                            }
                            break;
                        case 6:
                            {
                                // Center Steel to CL of Pipe
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    offset1 = offset1 - leftPipeDiameter / 2 - insulationThickness;
                                else
                                    offset1 = offset1 - leftPipeDiameter / 2;
                            }
                            break;
                        case 7:
                            {
                                // Outside Steel to CL of Pipe
                                if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                    offset1 = offset1 - legDepth / 2 - leftPipeDiameter / 2 - insulationThickness;
                                else
                                    offset1 = offset1 - legDepth / 2 - leftPipeDiameter / 2;
                            }
                            break;
                        default:
                            // Inside Steel to Edge of Pipe
                            offset1 = offset1 + leftPipeDiameter / 2;
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

                offset2Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset2Def", "Offset2Definition")).PropValue;
                //lOffset1Selection = HH.GetAttr("Offset1Selection", , , False)
                offset2Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset2Sel", "Offset2Selection")).PropValue;

                //Get the Offset From the Input Attributes
                string offset2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset2Rl", "Offset2Rule")).PropValue;
                if (offset2Selection == 1)
                {
                    GenericHelper.GetDataByRule(offset2Rule, (BusinessObject)support, out offset2);
                    support.SetPropertyValue(offset2, "IJUAhsFrameOffset2", "Offset2Value");
                }
                else
                    offset2 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOffset2", "Offset2Value")).PropValue;

                switch (offset2Definition)
                {
                    case 1:
                        // Edge of Pipe
                        break;
                    case 2:
                        {
                            // Pipe Center Line
                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                offset2 = offset2 - rightPipeDiameter / 2 - insulationThickness;
                            else
                                offset2 = offset2 - rightPipeDiameter / 2;
                        }
                        break;
                    default:
                        // Edge of Pipe
                        break;
                }

                //==========================
                // Shoe Height
                //==========================
                double shoeHeight;
                int shoeHeightDefinition, shoeHeightSelection;
                shoeHeightDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameShoeHeightDef", "ShoeHeightDefinition")).PropValue;
                shoeHeightSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameShoeHeightSel", "ShoeHeightSelection")).PropValue;
                string shoeHeightRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue;

                //Get the Shoe Height From the Input Attributes
                if (shoeHeightSelection == 1)
                {
                    GenericHelper.GetDataByRule(shoeHeightRule, (BusinessObject)support, out shoeHeight);
                    support.SetPropertyValue(shoeHeight, "IJUAhsFrameShoeHeight", "ShoeHeightValue");
                }
                else
                    shoeHeight = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameShoeHeight", "ShoeHeightValue")).PropValue;

                switch (shoeHeightDefinition)
                {
                    case 1: //Edge of Bounding Box
                        break;
                    case 2: //Centerline of Primary Pipe
                        {
                            if (Configuration == 1 || Configuration == 3)
                                shoeHeight = shoeHeight - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Y));

                            shoeHeight = shoeHeight - (boundingBoxDepth - RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z));
                        }
                        break;
                }

                // Section Offset
                double sectionYOffset = 0;
                if (Configuration == 1 || Configuration == 3)
                    sectionYOffset = shoeHeight + boundingBoxWidth + sectionWidth / 2;
                else
                    sectionYOffset = -shoeHeight - sectionWidth / 2;

                //==========================
                //Leg 1 Begin Overhang
                //========================== 
                double leg1BeginOverHang;
                int leg1BeginOverHangDefinition, leg1BeginOverHangSelection;
                leg1BeginOverHangDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1BeginOHDef", "Leg1BeginOverhangDefinition")).PropValue;
                leg1BeginOverHangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1BeginOHSel", "Leg1BeginOverhangSelection")).PropValue;
                string legBeginOverHangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1BeginOHRl", "Leg1BeginOverhangRule")).PropValue;
                if (leg1BeginOverHangSelection == 1)
                {
                    GenericHelper.GetDataByRule(legBeginOverHangRule, (BusinessObject)support, out leg1BeginOverHang);
                    support.SetPropertyValue(leg1BeginOverHang, "IJUAhsLeg1BeginOH", "Leg1BeginOverhangValue");
                }
                else
                    leg1BeginOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsLeg1BeginOH", "Leg1BeginOverhangValue")).PropValue;
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
                }

                //==========================
                // Leg 1 End Overhang
                //==========================
                double leg1EndOverHang;
                int leg1EndOverHangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1EndOHSel", "Leg1EndOverhangSelection")).PropValue;
                string leg1EndOverHangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1EndOHRl", "Leg1EndOverhangRule")).PropValue;
                if (structureConnection != 1)
                {
                    if (leg1EndOverHangSelection == 1)
                    {
                        GenericHelper.GetDataByRule(leg1EndOverHangRule, (BusinessObject)support, out leg1EndOverHang);
                        support.SetPropertyValue(leg1EndOverHang, "IJUAhsLeg1EndOH", "Leg1EndOverhangValue");
                    }
                    else
                        leg1EndOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsLeg1EndOH", "Leg1EndOverhangValue")).PropValue;

                    switch (supportingFace)
                    {
                        case 513:
                        case 514:
                            leg1EndOverHang = supportingSection.depth + leg1EndOverHang;
                            break;
                        case 257:
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
                        leg1EndOverHang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsLeg1EndOH", "Leg1EndOverhangValue")).PropValue;
                    }
                }
                //==========================
                // Member 1 Begin Overhang
                //==========================
                double member1BeginOverhang = 0;
                int member1BeginOverhangDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1BeginOHDef", "Member1BeginOverhangDefinition")).PropValue;
                int member1BeginOverhangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1BeginOHSel", "Member1BeginOverhangSelection")).PropValue;
                string member1BeginOverhangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1BeginOHRl", "Member1BeginOverhangRule")).PropValue;
                if (member1BeginOverhangSelection == 1)
                {
                    GenericHelper.GetDataByRule(member1BeginOverhangRule, (BusinessObject)support, out member1BeginOverhang);
                    support.SetPropertyValue(member1BeginOverhang, "IJUAhsMember1BeginOH", "Member1BeginOverhangValue");
                }
                else
                    member1BeginOverhang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsMember1BeginOH", "Member1BeginOverhangValue")).PropValue;

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
                int member1EndOverhangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1EndOHSel", "Member1EndOverhangSelection")).PropValue;
                string member1EndOverhangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1EndOHRl", "Member1EndOverhangRule")).PropValue;
                if (member1EndOverhangSelection == 1)
                {
                    GenericHelper.GetDataByRule(member1EndOverhangRule, (BusinessObject)support, out member1EndOverhang);
                    support.SetPropertyValue(member1EndOverhang, "IJUAhsMember1EndOH", "Member1EndOverhangValue");
                }
                else
                    member1EndOverhang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsMember1EndOH", "Member1EndOverhangValue")).PropValue;

                //==========================
                // Set the Frame Outputs as per their Definitions
                //==========================
                // SPAN
                int spanDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameSpanDef", "SpanDefinition")).PropValue;
                switch (spanDefinition)
                {
                    case 1: //Inside Edge of Steel
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 - legDepth / 2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    case 2: //Center Line of Steel
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    case 3: //Outside Edge of Steel
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 + legDepth / 2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    case 4: // Steel Length
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 + member1BeginOverhang + member1EndOverhang, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    default: // Inside Edge of Steel
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2 - legDepth / 2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                }
                // LENGTH1 will be set in the BOM, because they are determined by the DCM Constraint Solver
                //==========================
                // Set the Frame Connections
                //==========================
                //Set Connection between Leg1 and Section
                PropertyValueCodelist connection1TypeList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCornerConn1Type", "Connection1Type");
                Boolean flipLeg, flipSection;
                Boolean connectionSwap = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCornerConn1Swap", "Connection1Swap")).PropValue;
                int connection1Type = connection1TypeList.PropValue;
                Boolean connection1Mirror = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCornerConn1Mirror", "Connection1Mirror")).PropValue;
                if (reflectMember == false)
                {
                    if (connectionSwap)
                    {
                        if (flipFrame == false)
                        {
                            FrameAssemblyServices.SetSteelConnection(this, LEG, "BeginCap", ref legData, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            flipLeg = false;
                            flipSection = false;
                        }
                        else
                        {
                            FrameAssemblyServices.SetSteelConnection(this, LEG, "EndCap", ref legData, leg1BeginOverHang, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            flipLeg = true;
                            flipSection = true;
                        }
                    }
                    else
                    {
                        if (flipFrame == false)
                        {
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG, "BeginCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            flipLeg = false;
                            flipSection = false;
                        }
                        else
                        {
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, LEG, "EndCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            flipLeg = true;
                            flipSection = true;
                        }
                    }
                    if (flipFrame == false)
                    {
                        componentDictionary[SECTION].SetPropertyValue(member1EndOverhang, "IJUAHgrOccOverLength", "BeginOverLength");
                        if (connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                            // Set the Cutback of the other end in case of toggle
                            componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsCutback", "CutbackBeginAngle");
                    }
                    else
                    {
                        componentDictionary[SECTION].SetPropertyValue(member1EndOverhang, "IJUAHgrOccOverLength", "EndOverLength");
                        if (connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                            // Set the Cutback of the other end in case of toggle
                            componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsCutback", "CutbackEndAngle");
                    }
                }
                else
                {
                    if (connectionSwap)
                    {
                        if (flipFrame == false)
                        {
                            FrameAssemblyServices.SetSteelConnection(this, LEG, "EndCap", ref legData, leg1BeginOverHang, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            flipLeg = true;
                            flipSection = true;
                        }
                        else
                        {
                            FrameAssemblyServices.SetSteelConnection(this, LEG, "BeginCap", ref legData, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, member1BeginOverhang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            flipLeg = false;
                            flipSection = false;
                        }
                    }
                    else
                    {
                        if (flipFrame == false)
                        {
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "BeginCap", ref sectionData, member1BeginOverhang, LEG, "EndCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)90, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            flipLeg = true;
                            flipSection = true;
                        }
                        else
                        {
                            FrameAssemblyServices.SetSteelConnection(this, SECTION, "EndCap", ref sectionData, member1BeginOverhang, LEG, "BeginCap", ref legData, leg1BeginOverHang, (FrameAssemblyServices.SteelConnectionAngle)270, 0, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                            flipLeg = false;
                            flipSection = false;
                        }
                    }
                    if (flipFrame == false)
                    {
                        componentDictionary[SECTION].SetPropertyValue(member1EndOverhang, "IJUAHgrOccOverLength", "EndOverLength");
                        if (connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                            // Set the Cutback of the other end in case of toggle
                            componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsCutback", "CutbackEndAngle");
                    }
                    else
                    {
                        componentDictionary[SECTION].SetPropertyValue(member1EndOverhang, "IJUAHgrOccOverLength", "BeginOverLength");
                        if (connection1TypeList.PropValue == (int)FrameAssemblyServices.SteelConnection.SteelConnection_Mitered)
                            // Set the Cutback of the other end in case of toggle
                            componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsCutback", "CutbackBeginAngle");
                    }
                }
                //==========================
                // Set Steel CPs
                //==========================
                // Set the CP for the SECTION Flex Ports (Used to connect the main section to the BBX)
                PropertyValueCodelist cardinalPoint6Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsSteelCPFlexPort", "CP7");

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
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsFlexPort", "FlexPortRotZ");
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsEndFlexPort", "EndFlexPortRotZ");

                //Set the CP for the Member End Port that may attach to the supporting Structure
                PropertyValueCodelist cardinalPoint2Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsSteelCP", "CP1");

                PropertyValueCodelist cardinalPoint2Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG],"IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG],"IJOAhsSteelCP", "CP1");

                if (flipSection == true)
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = (int)Convert.ToDouble(sectionData.CardinalPoint), "IJOAhsSteelCP", "CP2");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsEndCap", "EndCapRotZ");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(sectionData.OffsetX), "IJOAhsEndCap", "EndCapXOffset");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(sectionData.OffsetY), "IJOAhsEndCap", "EndCapYOffset");
                }
                else
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = (int)Convert.ToDouble(sectionData.CardinalPoint), "IJOAhsSteelCP", "CP1");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(sectionData.OffsetX), "IJOAhsBeginCap", "BeginCapXOffset");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(sectionData.OffsetY), "IJOAhsBeginCap", "BeginCapYOffset");
                }

                // Set the CP's for the Leg Port that Connects to the Supporting Structure
                if (flipLeg == false)
                {
                    componentDictionary[LEG].SetPropertyValue(cardinalPoint2Leg.PropValue = (int)Convert.ToDouble(legData.CardinalPoint), "IJOAhsSteelCP", "CP2");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(legData.Orient) * Math.PI / 180), "IJOAhsEndCap", "EndCapRotZ");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(legData.OffsetX) * Math.PI / 180), "IJOAhsEndCap", "EndCapXOffset");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(legData.OffsetY) * Math.PI / 180), "IJOAhsEndCap", "EndCapYOffset");
                }
                else
                {
                    componentDictionary[LEG].SetPropertyValue(cardinalPoint1Leg.PropValue = (int)Convert.ToDouble(legData.CardinalPoint), "IJOAhsSteelCP", "CP1");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(legData.Orient) * Math.PI / 180), "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(legData.OffsetX) * Math.PI / 180), "IJOAhsBeginCap", "BeginCapXOffset");
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(legData.OffsetY) * Math.PI / 180), "IJOAhsBeginCap", "BeginCapYOffset");
                }

                   //==========================
                // Joints To Connect the Main Steel Section to the BBX
                //==========================

                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(boundingBoxWidth + offset1 + offset2), "IJUAHgrOccLength", "Length");

                if (reflectMember == false)
                {
                    if (flipFrame == false)
                    {
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(offset2), "IJOAhsFlexPort", "FlexPortZOffset");
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(-offset1), "IJOAhsEndFlexPort", "EndFlexPortZOffset");
                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, sectionYOffset, 0);
                    }
                    else
                    {
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(offset1), "IJOAhsFlexPort", "FlexPortZOffset");
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(-offset2), "IJOAhsEndFlexPort", "EndFlexPortZOffset");
                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, boundingBoxDepth, sectionYOffset, 0);
                    }
                }
                else
                {
                    if (flipFrame == false)
                    {
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(offset1), "IJOAhsFlexPort", "FlexPortZOffset");
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(-offset2), "IJOAhsEndFlexPort", "EndFlexPortZOffset");
                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "EndFlex", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, sectionYOffset, -sectionDepth);
                    }
                    else
                    {
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(offset2), "IJOAhsFlexPort", "FlexPortZOffset");
                        componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(-offset1), "IJOAhsEndFlexPort", "EndFlexPortZOffset");
                        JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, boundingBoxDepth, sectionYOffset, -sectionDepth);
                    }
                }

                //==========================
                // Joints To Connect Leg 1 To Supporting Structure
                //==========================
                //If it is Equipment or a Shape, we will use GetProjectedPointOnSurface
                //For Steel, Slab and Wall, we will use a PointOnJoint

                Vector boundingBox_X = new Vector(0, 0, 0), boundingBox_Y = new Vector(0, 0, 0), boundingBox_Z = new Vector(0, 0, 0), structZ = new Vector(0, 0, 0), structX = new Vector(0, 0, 0), structY = new Vector(0, 0, 0);
                Vector projectXZ = new Vector(0, 0, 0), projectZY = new Vector(0, 0, 0), projectXY = new Vector(0, 0, 0), leg1YOffset = new Vector(0, 0, 0), memberProjectedNormal = new Vector(0, 0, 0), leg1ProjectedNormal = new Vector(0, 0, 0);
                Position boundingBox_Position = new Position(0, 0, 0), leg1Start = new Position(0, 0, 0), leg1ProjectedPoint = new Position(0, 0, 0);
                double leg1CutbackAngle = 0;
                double planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.X, "Structure", PortAxisType.Z, OrientationAlong.Direct);

                port = new Matrix4X4();
                port = RefPortHelper.PortLCS(boundingBoxPort);
                boundingBox_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                boundingBox_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                boundingBox_Y = boundingBox_Z.Cross(boundingBox_X);

                port = new Matrix4X4();
                port = RefPortHelper.PortLCS("Structure");
                structZ.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                structX.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                structY = structZ.Cross(structX);

                if ((planeAngle * 180 / Math.PI) > 45 && (planeAngle * 180 / Math.PI) < 135)
                {
                    projectXZ = FrameAssemblyServices.ProjectVectorIntoPlane(structY, boundingBox_Y);
                    if (FrameAssemblyServices.GetVectorProjection(structY, boundingBox_X) < 0)
                        projectXZ.Set(-projectXZ.X, -projectXZ.Y, -projectXZ.Z);
                    projectXY = FrameAssemblyServices.ProjectVectorIntoPlane(structY, boundingBox_Z);
                    if (FrameAssemblyServices.GetVectorProjection(structY, boundingBox_X) < 0)
                        projectXY.Set(-projectXY.X, -projectXY.Y, -projectXY.Z);
                }
                else
                {
                    projectXZ = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, boundingBox_Y);
                    if (FrameAssemblyServices.GetVectorProjection(structZ, boundingBox_X) < 0)
                        projectXZ.Set(-projectXZ.X, -projectXZ.Y, -projectXZ.Z);
                    projectXY = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, boundingBox_Z);
                    if (FrameAssemblyServices.GetVectorProjection(structZ, boundingBox_X) < 0)
                        projectXY.Set(-projectXY.X, -projectXY.Y, -projectXY.Z);
                }

                PropertyValueCodelist cardinalPoint6Section6Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG],"IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint6Section7Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG],"IJOAhsSteelCPFlexPort", "CP7");

                PropertyValueCodelist endCutbackAnchorPointList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG],"IJOAhsCutback", "EndCutbackAnchorPoint");
                PropertyValueCodelist beginCutbackAnchorPointList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG],"IJOAhsCutback", "BeginCutbackAnchorPoint");
                if ((FrameAssemblyServices.AngleBetweenVectors(boundingBox_X, projectXZ)) > (FrameAssemblyServices.AngleBetweenVectors(boundingBox_X, projectXY)))
                {
                    // Trim Steel in XZ Plane of BBX
                    switch (legData.Orient)
                    {
                        case 0:
                            {
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section6Leg.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP6");
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section7Leg.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCPFlexPort", "CP7");
                                if (flipLeg == false)
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                else
                                    componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                leg1CutbackAngle = -boundingBox_X.Angle(projectXZ, boundingBox_Y);
                                break;
                            }
                        case (FrameAssemblyServices.SteelOrientationAngle)90:
                            {
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section6Leg.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP6");
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section7Leg.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCPFlexPort", "CP7");
                                if (flipLeg == false)
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                else
                                    componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                leg1CutbackAngle = boundingBox_X.Angle(projectXZ, boundingBox_Y);
                                break;
                            }
                        case (FrameAssemblyServices.SteelOrientationAngle)180:
                            {
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section6Leg.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP6");
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section7Leg.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCPFlexPort", "CP7");
                                if (flipLeg == false)
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                else
                                    componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                leg1CutbackAngle = boundingBox_X.Angle(projectXZ, boundingBox_Y);
                                break;
                            }
                        case (FrameAssemblyServices.SteelOrientationAngle)270:
                            {
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section6Leg.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP6");
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section7Leg.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCPFlexPort", "CP7");
                                if (flipLeg == false)
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                else
                                    componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                leg1CutbackAngle = -boundingBox_X.Angle(projectXZ, boundingBox_Y);
                                break;
                            }
                    }
                    if (flipFrame == false)
                    {
                        if (flipLeg)
                            leg1CutbackAngle = -leg1CutbackAngle;
                    }
                    else
                    {
                        if (flipLeg == false)
                            leg1CutbackAngle = -leg1CutbackAngle;
                    }
                    if (flipLeg == false)
                        componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1CutbackAngle), "IJOAhsCutback", "CutbackEndAngle");
                    else
                        componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1CutbackAngle), "IJOAhsCutback", "CutbackBeginAngle");
                }
                else
                {
                    // Trim Steel in XY Plane of BBX
                    switch (legData.Orient)
                    {
                        case 0:
                            {
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section6Leg.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP6");
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section7Leg.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCPFlexPort", "CP7");
                                if (flipLeg == false)
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                else
                                    componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                leg1CutbackAngle = -boundingBox_X.Angle(projectXY, boundingBox_Z);
                                break;
                            }
                        case (FrameAssemblyServices.SteelOrientationAngle)90:
                            {
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section6Leg.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP6");
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section7Leg.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCPFlexPort", "CP7");
                                if (flipLeg == false)
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                else
                                    componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                leg1CutbackAngle = -boundingBox_X.Angle(projectZY, boundingBox_Z);
                                break;
                            }
                        case (FrameAssemblyServices.SteelOrientationAngle)180:
                            {
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section6Leg.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP6");
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section7Leg.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCPFlexPort", "CP7");
                                if (flipLeg == false)
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                else
                                    componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                leg1CutbackAngle = boundingBox_X.Angle(projectZY, boundingBox_Z);
                                break;
                            }
                        case (FrameAssemblyServices.SteelOrientationAngle)270:
                            {
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section6Leg.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP6");
                                componentDictionary[LEG].SetPropertyValue(cardinalPoint6Section7Leg.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCPFlexPort", "CP7");
                                if (flipLeg == false)
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPointList.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                else
                                    componentDictionary[LEG].SetPropertyValue(beginCutbackAnchorPointList.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                leg1CutbackAngle = boundingBox_X.Angle(projectZY, boundingBox_Z);
                                break;
                            }
                    }
                    if (flipLeg == false)
                        componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1CutbackAngle), "IJOAhsCutback", "CutbackEndAngle");
                    else
                        componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1CutbackAngle), "IJOAhsCutback", "CutbackBeginAngle");
                }

                if (flipLeg == false)
                    JointHelper.CreatePrismaticJoint(LEG, "BeginFlex", LEG, "EndFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                else
                    JointHelper.CreatePrismaticJoint(LEG, "EndFlex", LEG, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                double basePlate1Th = 0;
                if (isBasePlate1)
                    basePlate1Th = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1],"IJUAhsThickness1", "Thickness1")).PropValue;
                else
                    basePlate1Th = 0;

                if (flipLeg == false)
                {
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1EndOverHang - basePlate1Th), "IJUAHgrOccOverLength", "EndOverLength");
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.X, "Structure", PortAxisType.Z, OrientationAlong.Direct);
                    if (isPlacedOnSurface == false)
                    {
                        if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                            JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", "-1", "Structure", Plane.NegativeXY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", "-1", "Structure", Plane.NegativeZX);
                    }
                }
                else
                {
                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(leg1EndOverHang - basePlate1Th), "IJUAHgrOccOverLength", "BeginOverLength");
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.X, "Structure", PortAxisType.Z, OrientationAlong.Direct);
                    if (isPlacedOnSurface == false)
                    {
                        if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                            JointHelper.CreatePointOnPlaneJoint(LEG, "BeginFlex", "-1", "Structure", Plane.NegativeXY);
                        else
                            JointHelper.CreatePointOnPlaneJoint(LEG, "BeginFlex", "-1", "Structure", Plane.NegativeZX);
                    }
                }

                // ==========================
                // Joints For Remaining Optional Plates
                // ==========================
                // Joints for End Plates
                double capPlate1Angle = 0, capPlate2Angle = 0, capPlate3Angle = 0, basePlate1Angle = 0;
                if (isCapPlate1)
                    capPlate1Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1Ang", "CapPlate1Angle")).PropValue;
                if (isCapPlate1)
                    capPlate2Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate2Ang", "CapPlate2Angle")).PropValue;
                if (isCapPlate1)
                    capPlate3Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate3Ang", "CapPlate3Angle")).PropValue;
                if (isBasePlate1)
                    basePlate1Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1Ang", "BasePlate1Angle")).PropValue;

                if (isCapPlate1)
                {
                    if (flipLeg == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE1, "Port2", LEG, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate1Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE1, "Port1", LEG, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate1Angle));
                }
                if (isCapPlate2)
                {
                    if (reflectMember == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE2, "Port2", SECTION, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate2Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE2, "Port1", SECTION, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate2Angle));
                }
                if (isCapPlate3)
                {
                    if (reflectMember == false)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE3, "Port1", SECTION, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate3Angle));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE3, "Port2", SECTION, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate3Angle));
                }

                //Joints for the Base Plates
                if (isBasePlate1)
                {
                    if (flipLeg == false)
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
                    if (flipLeg == false)
                        legWeldPort = "EndFace";
                    else
                        legWeldPort = "BeginFace";
                    double length1, width1;
                    switch (weld.connection)
                    {
                        case "A":
                            {
                                if (isBasePlate1)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1],"IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1],"IJUAhsWidth1", "Width1")).PropValue;
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
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE2],"IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE2],"IJUAhsWidth1", "Width1")).PropValue;
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
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE1],"IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE1],"IJUAhsWidth1", "Width1")).PropValue;
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
                                if (isCapPlate3)
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3],"IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3],"IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(CAPPLATE3, "Port1", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
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
                double startSpanOffset = 0, endSpanOffset = 0, farPipeDia, farPipeInsul;
                int farPipeIndex = 0;
                Boolean excludeNotes = (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;

                switch (spanDefinition)
                {
                    case 1:
                        {
                            // Inside Edge of Steel
                            startSpanOffset = -legDepth / 2;
                            endSpanOffset = 0;
                        }
                        break;
                    case 2:
                        {
                            // CL Steel
                            startSpanOffset = 0;
                            endSpanOffset = 0;
                        }
                        break;
                    case 3:
                        {
                            // Outside Edge of Steel
                            startSpanOffset = legDepth / 2;
                            endSpanOffset = 0;
                        }
                        break;
                    case 4:
                        {
                            // Inside Edge Steel to CL Pipe
                            startSpanOffset = -legDepth / 2;
                            if (flipFrame == false)
                                farPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Bottom);
                            else
                                farPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Top);
                            pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(farPipeIndex);
                            farPipeDia = pipeInfo.OutsideDiameter;
                            farPipeInsul = pipeInfo.InsulationThickness;

                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                endSpanOffset = offset2 + farPipeInsul + farPipeDia / 2;
                            else
                                endSpanOffset = offset2 + farPipeDia / 2;
                        }
                        break;
                    case 5:
                        {
                            // CL Steel to CL Pipe
                            startSpanOffset = 0;
                            if (flipFrame == false)
                                farPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Bottom);
                            else
                                farPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Top);
                            pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(farPipeIndex);
                            farPipeDia = pipeInfo.OutsideDiameter;
                            farPipeInsul = pipeInfo.InsulationThickness;

                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                endSpanOffset = offset2 + farPipeInsul + farPipeDia / 2;
                            else
                                endSpanOffset = offset2 + farPipeDia / 2;
                        }
                        break;
                    case 6:
                        {
                            // Outside Edge of Steel to CL Pipe
                            startSpanOffset = legDepth / 2;
                            if (flipFrame == false)
                                farPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Bottom);
                            else
                                farPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Top);
                            pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(farPipeIndex);
                            farPipeDia = pipeInfo.OutsideDiameter;
                            farPipeInsul = pipeInfo.InsulationThickness;

                            if ((bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue)
                                endSpanOffset = offset2 + farPipeInsul + farPipeDia / 2;
                            else
                                endSpanOffset = offset2 + farPipeDia / 2;
                        }
                        break;
                    case 7:
                        {
                            // Section Length
                            startSpanOffset = legDepth / 2 + member1BeginOverhang;
                            endSpanOffset = 0;
                        }
                        break;
                }
                double spanYOffset, lengthOffsetV;
                if (reflectMember)
                    spanYOffset = -sectionDepth;
                else
                    spanYOffset = 0;

                switch ((int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameLength1Def", "Length1Definition")).PropValue)
                {
                    case 1:
                        // Internal Edge of Steel
                        lengthOffsetV = 0;
                        break;
                    case 2:
                        // External Edge of Steel
                        lengthOffsetV = sectionDepth;
                        break;
                    case 3:
                        // End of Steel
                        lengthOffsetV = sectionDepth + leg1BeginOverHang;
                        break;
                    default:
                        lengthOffsetV = 0;
                        break;
                }

                // Span & Elevation
                ControlPoint controlPoint;
                Note note;
                if (flipSection == false)
                {
                    note = CreateNote("SpanStart", SECTION, "EndFlex", new Position(-sectionWidth / 2, spanYOffset + lengthOffsetV, offset1 + startSpanOffset), " ", true, 2, 1, out controlPoint);
                    note = CreateNote("SpanEnd", SECTION, "BeginFlex", new Position(-sectionWidth / 2, spanYOffset + lengthOffsetV, -offset2 + endSpanOffset), " ", true, 2, 1, out controlPoint);
                    if (excludeNotes == false)
                        note = CreateNote("Elevation", SECTION, "EndFlex", new Position(0, spanYOffset, 0), "Elevation", false, 2, 51, out controlPoint);
                    else
                        DeleteNoteIfExists("Elevation");
                }
                else
                {
                    note = CreateNote("SpanStart", SECTION, "BeginFlex", new Position(-sectionWidth / 2, spanYOffset + lengthOffsetV, -offset1 - startSpanOffset), " ", true, 2, 1, out controlPoint);
                    note = CreateNote("SpanEnd", SECTION, "EndFlex", new Position(-sectionWidth / 2, spanYOffset + lengthOffsetV, offset2 - endSpanOffset), " ", true, 2, 1, out controlPoint);
                    if (excludeNotes == false)
                        note = CreateNote("Elevation", SECTION, "BeginFlex", new Position(0, spanYOffset, 0), "Elevation", false, 2, 51, out controlPoint);
                    else
                        DeleteNoteIfExists("Elevation");
                }

                // Length
                if (flipLeg == false)
                    note = CreateNote("Length", LEG, "EndFlex", new Position(0, startSpanOffset, 0), " ", true, 2, 1, out controlPoint);
                else
                    note = CreateNote("Length", LEG, "BeginFlex", new Position(0, startSpanOffset, 0), " ", true, 2, 1, out controlPoint);

                // Pipe Centerlines
                string botPipePort = string.Empty, topPipePort = string.Empty;
                int botPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Bottom);
                if (botPipeIndex == 1)
                    botPipePort = "Route";
                else
                    botPipePort = "Route_" + botPipeIndex.ToString();

                int topPipeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Top);
                if (topPipeIndex == 1)
                    topPipePort = "Route";
                else
                    topPipePort = "Route_" + topPipeIndex.ToString();

                note = CreateNote("Pipe CL", "-1", topPipePort, new Position(0, 0, 0), " ", true, 2, 1, out controlPoint);
                if (botPipeIndex != topPipeIndex)
                    note = CreateNote("Pipe CL", "-1", botPipePort, new Position(0, 0, 0), " ", true, 2, 1, out controlPoint);

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
                    structConnections.Add(new ConnectionInfo(LEG, 1)); // partindex, routeindex

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
        public override void OrientLocalCoordinateSystem()
        {
            Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

            Boolean rotateFrame = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsRotateFr", "RotateFrame")).PropValue;
            Plane planeA=new Plane(); Plane planeB=new Plane();
            Axis axisA=new Axis();Axis axisB=new Axis();
            if (rotateFrame == true)
            {
                planeA = Plane.YZ;
                planeB = Plane.XY;
                axisA = Axis.Z;
                axisB = Axis.NegativeY;
            }
            else
            {
                planeA = Plane.YZ;
                planeB = Plane.XY;
                axisA = Axis.Z;
                axisB = Axis.X;
            }

            JointHelper.CreateRigidJoint("-1", "Route", "-1", "LocalCS", planeA, planeB, axisA, axisB, 0, 0, 0);

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
                double sectionDepth = 0, sectionWidth = 0;
                IPart part = (IPart)componentDictionary[SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                section = FrameAssemblyServices.GetSectionDataFromPart(part.PartNumber);
                if ((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart,"IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue == 0 || (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart,"IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue == 180)
                {
                    sectionDepth = section.depth;
                    sectionWidth = section.width;
                }
                else
                {
                    sectionDepth = section.width;
                    sectionWidth = section.depth;
                }
             
                string supportNumber = string.Empty, L1 = string.Empty, span = string.Empty;
                double length1Value = 0, spanValue = 0;
             
                // Check the Supporting count
                int supportingCount = 0;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    supportingCount = 1;
                else
                    supportingCount = SupportHelper.SupportingObjects.Count;

                // ==========================
                // Span
                // ==========================
                spanValue = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAhsFrameSpan", "SpanValue")).PropValue;

                supportNumber = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogPart,"IJUAhsSupportNumber", "SupportNumber")).PropValue;
                // ==========================
                // Length 1
                // ==========================
                length1Value = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG],"IJUAHgrOccLength", "Length")).PropValue;
                int length1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart,"IJUAhsFrameLength1Def", "Length1Definition")).PropValue;
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
                            double beginOverLength = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG],"IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                            length1Value = length1Value + Convert.ToDouble(beginOverLength);
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
                PropertyValueCodelist steelStandardList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart,"IJUAhsSteelStandard", "SteelStandard");

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
