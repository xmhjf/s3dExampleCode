//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   TFrame.cs
//   Frame_Assy,Ingr.SP3D.Content.Support.Rules.TFrame
//   Author       :  Hema
//   Creation Date:  26-Jul-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   26-Jul-2013     Hema    CR-CP-224474 Convert HS_S3DFrame to C# .Net 
//   30-Mar-2014     PVK     CR-CP-245789	Modify the exsisting .Net Frame_Assy to behave like URS Frame supports
//   27-Apr-2015     PVK     TR-CP-253033	Elevation CP not shown by default for frame supports.
//   06-May-2015     PVK     CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   25/05/2016     PVK      TR-CP-292743	Corrected the VerticalOffset attribute
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Linq;
using Ingr.SP3D.ReferenceData.Middle;
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

    public class TFrame : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Steel Sections
        private const string SECTION = "SECTION"; //Main Member
        private const string LEG = "LEG"; //Leg of T-Frame

        //Other Optional Parts
        private const string CAPPLATE1 = "CAPPLATE1"; //Plate on the end of Leg
        private const string CAPPLATE2 = "CAPPLATE2"; //Plate on the end of the Main Member
        private const string CAPPLATE3 = "CAPPLATE3"; //Plate on the end of the Main Member

        private const string BASEPLATE1 = "BASEPLATE1"; //Base Plate for Leg
        private const string STRUCTCONN = "STRUCTCONN";
        string[] PIPEATT1;
        //Collections for the Weld Data and Weld Part Index's
        private Collection<FrameAssemblyServices.WeldData> weldCollection = new Collection<FrameAssemblyServices.WeldData>();  //Collection of Welds (hsWeldData Type)

        //Steel Members - Stores Data related to steel member (Part Number, Depth, Width etc...)
        private FrameAssemblyServices.HSSteelMember section;
        private FrameAssemblyServices.HSSteelMember leg;

        //Steel Configuration - Stores Data related to steel configuration
        private FrameAssemblyServices.HSSteelConfig sectionData;
        private FrameAssemblyServices.HSSteelConfig legData;
        Boolean isMember, isLeg1, isCapPlate1, isCapPlate2, isCapPlate3, isBasePlate1, isBolt1Part, includePipeAtt1;
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
                    isLeg1 = FrameAssemblyServices.AddPart(this, LEG, leg1Part, leg1Rule, parts);

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

                    string basePlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1", "BasePlate1Part")).PropValue;
                    string basePlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1Rl", "BasePlate1Rule")).PropValue;
                    isBasePlate1 = FrameAssemblyServices.AddPart(this, BASEPLATE1, basePlate1Part, basePlate1Rule, parts);

                    // Add the Structure Connection Object if it is Place-By-Structure
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        parts.Add(new PartInfo(STRUCTCONN, "Log_Conn_Part_1"));

                    //Add PipeAttachemnts
                    string pipeAtt1Part = "";
                    string pipeAtt1Rule = "";
                    if (part.SupportsInterface("IJUAhsFrPipeAtt1"))
                        pipeAtt1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrPipeAtt1", "PipeAtt1Part")).PropValue;
                    if (part.SupportsInterface("IJUAhsFrPipeAtt1Rl"))
                        pipeAtt1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrPipeAtt1Rl", "PipeAtt1Rule")).PropValue;

                    PIPEATT1 = new string[1];
                    if (SupportHelper.SupportedObjects.Count > 0 && (!string.IsNullOrEmpty(pipeAtt1Part) || !string.IsNullOrEmpty(pipeAtt1Rule)))
                    {
                        includePipeAtt1 = true;
                        Array.Resize(ref PIPEATT1, SupportHelper.SupportedObjects.Count);
                        for (int index = 0; index < SupportHelper.SupportedObjects.Count; index++)
                        {
                            PIPEATT1[index] = "PIPEATT1_" + index;
                            FrameAssemblyServices.AddPart(this, PIPEATT1[index], pipeAtt1Part, pipeAtt1Rule, parts);
                        }
                    }

                    //Add the Weld Objects from the Weld Sheet
                    string weldServiceClassName = string.Empty;
                    if (part.SupportsInterface("IJUAhsWeldServClass"))
                        weldServiceClassName = (String)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsWeldServClass", "WeldServiceClassName")).PropValue;
                    if (weldServiceClassName == null)
                        weldServiceClassName = string.Empty;
                    weldCollection = FrameAssemblyServices.AddWeldsFromCatalog(this, parts, "IJUAhsTFrameWelds", ((IPart)part).PartNumber, weldServiceClassName);

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
                    int bolt1Quantity = 0;
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
                        isBolt1Part = FrameAssemblyServices.AddImpliedPart(this, bolt1Part, bolt1Rule, impliedParts, null, bolt1Quantity);
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

                BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                double sectionDepth, sectionWidth, legDepth, legWidth;

                //==========================
                // Get Required Information about the Steel Parts
                //==========================

                //Get the Steel Cross Section Data
                section = FrameAssemblyServices.GetSectionDataFromPartIndex(this, SECTION);
                leg = FrameAssemblyServices.GetSectionDataFromPartIndex(this, LEG);

                sectionData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue;
                legData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue;

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

                //Get the Width and Depth of each steel part according to the orientation angle
                if (sectionData.Orient == 0 || sectionData.Orient == (FrameAssemblyServices.SteelOrientationAngle)180)
                {
                    sectionDepth = section.depth;
                    sectionWidth = section.width;
                }
                else
                {
                    sectionDepth = section.width;
                    sectionWidth = section.depth;
                }
                if (legData.Orient == 0 || legData.Orient == (FrameAssemblyServices.SteelOrientationAngle)180)
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
                // Create the Frame Bounding Box
                //==========================

                string boundingBoxName = "BBFrame", boundingBoxPort = "BBFrame_Low";
                Boolean includeInsulation = false;
                includeInsulation = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue;
                int frameOrientation = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOrientation", "FrameOrientation")).PropValue;

                FrameAssemblyServices.CreateFrameBoundingBox(this, boundingBoxName, (FrameAssemblyServices.FrameBBXOrientation)frameOrientation, includeInsulation, (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsMirrorFrame", "MirrorFrame")).PropValue, SupportedHelper.IsSupportedObjectVertical(1, 45));

                double boundingBoxWidth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Width;
                double boundingBoxDepth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Height;
                Boolean isPlacedOnSurface = false;

                Collection<FrameAssemblyServices.WeldData> frameConnection1Welds = new Collection<FrameAssemblyServices.WeldData>();
                Collection<FrameAssemblyServices.WeldData> pipeConnection1Welds = new Collection<FrameAssemblyServices.WeldData>();
                Collection<FrameAssemblyServices.WeldData> otherWelds = new Collection<FrameAssemblyServices.WeldData>();
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByReference && SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.GenericSurface)
                        isPlacedOnSurface = true;
                }
                else
                    isPlacedOnSurface = false;
                // ==========================
                // Determine the Connections to the Structural Steel
                // ==========================
                double basePlate1Th = 0;
                if (isBasePlate1)
                {
                    basePlate1Th = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1], "IJUAhsThickness1", "Thickness1")).PropValue;
                }

                FrameAssemblyServices.HSSteelMember supportingSection = new FrameAssemblyServices.HSSteelMember();
                int supportingFace = 0, structWidthOffsetType = 0, structureConnection = 1;
                double structWidthOffset = 0;
                if (part.SupportsInterface("IJUAhsStructureConn"))
                    structureConnection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsStructureConn", "StructureConnection")).PropValue;

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
                            // Lapped Connection
                            structWidthOffsetType = 1;
                            break;
                        case 3:
                            // Reverse Lapped Connection
                            structWidthOffsetType = 2;
                            break;
                    }
                    switch (structWidthOffsetType)
                    {
                        case 1:
                            switch (supportingFace)
                            {
                                case 513:
                                case 514:
                                    structWidthOffset = -supportingSection.width / 2 - sectionWidth / 2 - basePlate1Th;
                                    break;
                                case 257:
                                case 258:
                                    structWidthOffset = -supportingSection.depth / 2 - sectionWidth / 2 - basePlate1Th;
                                    break;
                                default:
                                    structWidthOffset = 0;
                                    break;
                            }
                            break;
                        case 2:
                            switch (supportingFace)
                            {
                                case 513:
                                case 514:
                                    structWidthOffset = supportingSection.width / 2 + sectionWidth / 2 + basePlate1Th;
                                    break;
                                case 257:
                                case 258:
                                    structWidthOffset = supportingSection.depth / 2 + sectionWidth / 2 + basePlate1Th;
                                    break;
                                default:
                                    structWidthOffset = 0;
                                    break;
                            }
                            break;
                    }
                    // Attach the Structure Conn to the Structure
                    if (RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Y, boundingBoxPort, PortAxisType.X, OrientationAlong.Direct) < Math.Atan(1) * 4.0 / 2)
                        JointHelper.CreateRigidJoint("-1", "Structure", STRUCTCONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, structWidthOffset, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "Structure", STRUCTCONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -structWidthOffset, 0);
                }




                //==========================
                // Organize the Welds into two Collections
                //==========================

                FrameAssemblyServices.WeldData weld = new FrameAssemblyServices.WeldData();
                int weldCount;
                for (weldCount = 0; weldCount < weldCollection.Count; weldCount++)
                {
                    weld = weldCollection[weldCount];
                    if ((weld.connection).ToUpper().Equals("C"))
                        frameConnection1Welds.Add(weld);
                    if ((weld.connection).ToUpper().Equals("G"))
                        pipeConnection1Welds.Add(weld);
                    else
                        otherWelds.Add(weld);
                }

                //==========================
                // Offset 1 (The offset to the Leg)
                //==========================
                double offset1, structAngle, leftPipeDiameter, rightPipeDiameter, offset2, offset3 = 0;
                int offset1Definition, offset1Selection, offset2Definition, offset3Definition, offset2Selection, offset3Selection;
                int routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Left);
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                leftPipeDiameter = pipeInfo.OutsideDiameter;

                offset1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset1Def", "Offset1Definition")).PropValue;
                offset1Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset1Sel", "Offset1Selection")).PropValue;

                //Get the Offset From the Input Attributes
                Boolean flipFrame = false;
                structAngle = (RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct)) * 180 / Math.PI;
                string offset1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset1Rl", "Offset1Rule")).PropValue;
                if (offset1Selection == 1)
                {
                    GenericHelper.GetDataByRule(offset1Rule, (BusinessObject)support, out offset1);
                    support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                }
                else
                    offset1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOffset1", "Offset1Value")).PropValue;

                switch (offset1Definition)
                {
                    case 1: //Edge of Pipe
                        break;
                    case 2: //Pipe Center Line
                        if (includeInsulation)
                            offset1 = offset1 - leftPipeDiameter / 2 - pipeInfo.InsulationThickness;
                        else
                            offset1 = offset1 - leftPipeDiameter / 2;
                        break;
                    case 3: //Leg Center to Section Edge
                        offset1 = offset1 - boundingBoxWidth / 2;
                        break;
                    case 4:  //Leg Left Edge to Section Edge
                        offset1 = offset1 - boundingBoxWidth / 2 + legDepth / 2;
                        break;
                    case 5: //Leg Right Edge to Section Edge
                        offset1 = offset1 - boundingBoxWidth / 2 - legDepth / 2;
                        break;
                    default: //Edge of Pipe
                        break;
                }

                //==========================
                // Offset 2
                //==========================
                routeIndex = BoundingBoxHelper.GetBoundaryRouteIndex(boundingBoxName, BoundingBoxEdge.Right);
                pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(routeIndex);
                rightPipeDiameter = pipeInfo.OutsideDiameter;

                offset2Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Def", "Offset2Definition")).PropValue;
                offset2Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Sel", "Offset2Selection")).PropValue;

                //Get the Offset From the Input Attributes
                string offset2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset2Rl", "Offset2Rule")).PropValue;
                if (offset1Selection == 1)
                {
                    GenericHelper.GetDataByRule(offset2Rule, (BusinessObject)support, out offset2);
                    support.SetPropertyValue(offset2, "IJUAhsFrameOffset2", "Offset2Value");
                }
                else
                    offset2 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOffset2", "Offset2Value")).PropValue;

                switch (offset2Definition)
                {
                    case 1: //Edge of Pipe
                        break;
                    case 2: //Pipe Center Line
                        if (includeInsulation)
                            offset2 = offset2 - rightPipeDiameter / 2 - pipeInfo.InsulationThickness;
                        else
                            offset2 = offset2 - rightPipeDiameter / 2;
                        break;
                    case 3: //Leg Center to Section Edge
                        offset2 = offset2 - boundingBoxWidth / 2;
                        break;
                    case 4: //Leg Left Edge to Section Edge
                        offset2 = offset2 - boundingBoxWidth / 2 - legDepth / 2;
                        break;
                    case 5: //Leg Right Edge to Section Edge
                        offset2 = offset2 - boundingBoxWidth / 2 - legDepth / 2;
                        break;
                    default: //Edge of Pipe
                        break;
                }

                //==========================
                // Offset 3
                //==========================
                offset3Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset3Def", "Offset3Definition")).PropValue;
                //lOffset1Selection = HH.GetAttr("Offset1Selection", , , False)
                offset3Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset3Sel", "Offset3Selection")).PropValue;
                //Get the Offset From the Input Attributes
                string offset3Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameOffset3Rl", "Offset3Rule")).PropValue;
                if (offset3Selection == 1)
                {
                    GenericHelper.GetDataByRule(offset3Rule, (BusinessObject)support, out offset3);
                    support.SetPropertyValue(offset3, "IJUAhsFrameOffset3", "Offset3Value");
                }
                else
                    offset3 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameOffset3", "Offset3Value")).PropValue;

                //Ignore the Offset3 value when the Offset1 and Offset2 is calculated from Leg
                if (offset3 != 0)
                {
                    if (offset3Definition == 3 || offset1Definition == 4 || offset1Definition == 5 || offset2Definition == 3 || offset2Definition == 4 || offset2Definition == 5)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Offset3 value will be ignored.", "", "TFrame", 346);
                        offset3 = 0;
                    }
                }
                else
                {
                    switch (offset3Definition)
                    {
                        case 1: //Left Edge of the Steel
                            offset3 = offset3 + legDepth / 2;
                            break;
                        case 2: //Center of Steel
                            break;
                        case 3: //Right edge of Steel
                            offset3 = offset3 - legDepth / 2;
                            break;
                        default: //Left Edge of the Steel
                            offset3 = offset3 + legDepth / 2;
                            break;
                    }
                }

                //==========================
                // Shoe Height
                //==========================
                double shoeHeight = 0;
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
                        frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            shoeHeight = shoeHeight - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z));
                        else
                            shoeHeight = shoeHeight + (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) - boundingBoxDepth;
                        break;
                    default: //Edge of Bounding Box
                        break;
                }
                //==========================
                // Shoe Width
                //==========================
                double shoeWidth = 0;
                int shoeWidthDefinition = 0, shoewidthSelection = 0;
                string shoeWidthRule = "";
                if (part.SupportsInterface("IJUAhsFrameShoeWidthDef"))
                    shoeWidthDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameShoeWidthDef", "ShoeWidthDefinition")).PropValue;
                if (part.SupportsInterface("IJUAhsFrameShoeWidthSel"))
                    shoewidthSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameShoeWidthSel", "ShoeWidthSelection")).PropValue;
                if (part.SupportsInterface("IJUAhsFrameShoeWidthRl"))
                    shoeWidthRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameShoeWidthRl", "ShoeWidthRule")).PropValue;

                //Get the Shoe Height From the Input Attributes
                if (shoewidthSelection == 1)
                {
                    GenericHelper.GetDataByRule(shoeHeightRule, (BusinessObject)support, out shoeHeight);
                    support.SetPropertyValue(shoeWidth, "IJUAhsFrameShoeWidth", "ShoeWidthValue");
                }
                else if (shoewidthSelection == 2)
                    shoeWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameShoeWidth", "ShoeWidthValue")).PropValue;

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
                        frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                        if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                            leg1BeginOverHang = leg1BeginOverHang - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) - shoeHeight - sectionDepth;
                        else
                            leg1BeginOverHang = leg1BeginOverHang - (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z)) + boundingBoxDepth + shoeHeight;
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
                //Handle all the Frame Configurations and Toggles
                //==========================
                double sectionZOffset;
                frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                    sectionZOffset = 0 - shoeHeight;
                else
                    sectionZOffset = boundingBoxDepth + sectionDepth + shoeHeight;
                Boolean reflectMember;
                if ((frameConfiguration == 1 && Configuration == 1) || (frameConfiguration == 2 && Configuration == 2))
                    reflectMember = false;
                else
                    reflectMember = true;

                //==========================
                // Set the Frame Outputs as per their Definitions
                //==========================
                int spanDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrameSpanDef", "SpanDefinition")).PropValue;
                switch (spanDefinition)
                {
                    case 1: //Pipe Center Lines
                        support.SetPropertyValue(boundingBoxWidth - leftPipeDiameter / 2 - rightPipeDiameter / 2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    case 2: //Steel Length
                        support.SetPropertyValue(boundingBoxWidth + offset1 + offset2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                    default: //Pipe Center Lines
                        support.SetPropertyValue(boundingBoxWidth - leftPipeDiameter / 2 - rightPipeDiameter / 2, "IJUAhsFrameSpan", "SpanValue");
                        break;
                }

                //LENGTH1 will be set in the BOM, because they are determined by the DCM Constraint Solver
                //==========================
                // Set the Frame Connections
                //==========================
                //Set Connection between Leg1 and Section
                double axialOffset;
                if (offset1Definition == 3 || offset1Definition == 4 || offset1Definition == 5)
                    axialOffset = -(boundingBoxWidth / 2 + offset1);
                else
                    axialOffset = -(boundingBoxWidth + offset1 + offset2) / 2 - offset3;

                int connection1Type = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsTeeConn1Type", "Connection1Type")).PropValue;
                Boolean connection1Mirror = (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsTeeConn1Mirror", "Connection1Mirror")).PropValue;

                if (reflectMember == false)
                    FrameAssemblyServices.SetSteelConnection(this, LEG, "BeginCap", ref legData, leg1BeginOverHang, SECTION, "BeginCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)90, -axialOffset, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);
                else
                    FrameAssemblyServices.SetSteelConnection(this, LEG, "BeginCap", ref legData, leg1BeginOverHang, SECTION, "EndCap", ref sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)270, axialOffset, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID, frameConnection1Welds);

                //==========================
                // Set Steel CPs
                //==========================
                PropertyValueCodelist cardinalPoint6Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist endCutbackAnchorPoint = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsCutback", "EndCutbackAnchorPoint");
                PropertyValueCodelist cardinalPoint2Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Section = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION], "IJOAhsSteelCP", "CP1");

                PropertyValueCodelist cardinalPoint6Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsSteelCP", "CP6");
                PropertyValueCodelist cardinalPoint7Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist cardinalPoint2Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint1Leg = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJOAhsSteelCP", "CP1");

                switch (sectionData.Orient)
                {
                    case 0:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = 2, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = 8, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)90:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = 6, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = 4, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)180:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = 8, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = 2, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                    case (FrameAssemblyServices.SteelOrientationAngle)270:
                        if (reflectMember == false)
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = 4, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        else
                        {
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = 6, "IJOAhsSteelCP", "CP6");
                            componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                        }
                        break;
                }
                componentDictionary[SECTION].SetPropertyValue((Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180)), "IJOAhsFlexPort", "FlexPortRotZ");
                componentDictionary[SECTION].SetPropertyValue((Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180)), "IJOAhsEndFlexPort", "EndFlexPortRotZ");

                //Set the CP for the Member End Port that may attach to the supporting Structure


                if (reflectMember == false)
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint2Section.PropValue = sectionData.CardinalPoint, "IJOAhsSteelCP", "CP2");
                    componentDictionary[SECTION].SetPropertyValue((Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180)), "IJOAhsEndCap", "EndCapRotZ");
                }
                else
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint1Section.PropValue = sectionData.CardinalPoint, "IJOAhsSteelCP", "CP1");
                    componentDictionary[SECTION].SetPropertyValue((Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180)), "IJOAhsBeginCap", "BeginCapRotZ");
                }

                //Set the CP's for the Leg Ports that Connect to the Supporting Structure
                componentDictionary[LEG].SetPropertyValue(cardinalPoint2Leg.PropValue = legData.CardinalPoint, "IJOAhsSteelCP", "CP2");
                componentDictionary[LEG].SetPropertyValue((Convert.ToDouble(Convert.ToDouble(legData.Orient) * Math.PI / 180)), "IJOAhsEndCap", "EndCapRotZ");
                componentDictionary[LEG].SetPropertyValue((Convert.ToDouble(Convert.ToDouble(legData.OffsetX) * Math.PI / 180)), "IJOAhsEndCap", "EndCapXOffset");
                componentDictionary[LEG].SetPropertyValue((Convert.ToDouble(Convert.ToDouble(legData.OffsetY) * Math.PI / 180)), "IJOAhsEndCap", "EndCapYOffset");

                //==========================
                // Joints To Connect the Main Steel Section to the BBX
                //==========================

                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(boundingBoxWidth + offset1 + offset2), "IJUAHgrOccLength", "Length");

                if (reflectMember == false)
                {
                    componentDictionary[SECTION].SetPropertyValue(offset1, "IJOAhsFlexPort", "FlexPortZOffset");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsEndFlexPort", "EndFlexPortZOffset");
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreatePrismaticJoint(SECTION, "BeginFlex", "-1", boundingBoxPort, Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, sectionZOffset);
                        JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", STRUCTCONN, "Connection", Plane.ZX);
                    }
                    else
                    {
                        if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel"))
                        {
                            //Single Steel Parallel to the Route
                            if (structAngle < 45 || structAngle > 135)
                            {
                                if (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Structure", PortAxisType.Y) < 0)
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
                else
                {
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsFlexPort", "FlexPortZOffset");
                    componentDictionary[SECTION].SetPropertyValue(-offset1, "IJOAhsEndFlexPort", "EndFlexPortZOffset");
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        JointHelper.CreatePrismaticJoint(SECTION, "EndFlex", "-1", boundingBoxPort, Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, 0, -sectionZOffset);
                        JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", STRUCTCONN, "Connection", Plane.ZX);
                    }
                    else
                    {
                        if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel"))
                        {
                            //Single Steel Parallel to the Route
                            if (structAngle < 45 || structAngle > 135)
                            {
                                if (RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Structure", PortAxisType.Y) < 0)
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
                //==========================
                // Joints To Connect Leg 1 To Supporting Structure
                //==========================
                //If it is Equipment or a Shape, we will use GetProjectedPointOnSurface
                //For Steel, Slab and Wall, we will use a PointOnJoint

                Vector boundingBox_X = new Vector(0, 0, 0), boundingBox_Y = new Vector(0, 0, 0), boundingBox_Z = new Vector(0, 0, 0), structZ = new Vector(0, 0, 0), structX = new Vector(0, 0, 0), structY = new Vector(0, 0, 0);
                Vector projectXZ = new Vector(0, 0, 0), projectZY = new Vector(0, 0, 0), projectXY = new Vector(0, 0, 0), leg1YOffset = new Vector(0, 0, 0), leg2YOffset = new Vector(0, 0, 0), memberProjectedNormal = new Vector(0, 0, 0), leg1ProjectedNormal = new Vector(0, 0, 0), leg2ProjectedNormal = new Vector(0, 0, 0);
                Position boundingBox_Position = new Position(0, 0, 0), leg2Start = new Position(0, 0, 0), leg1Start = new Position(0, 0, 0), leg1ProjectedPoint = new Position(0, 0, 0), leg2ProjectedPoint = new Position(0, 0, 0);
                Matrix4X4 port = new Matrix4X4();
                double planeAngle = 0, leg1CutbackAngle = 0, leg1Length = 0;

                if (isPlacedOnSurface == true)
                {
                    // Get Projection from calculated point
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    boundingBox_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    boundingBox_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    boundingBox_Y = boundingBox_Z.Cross(boundingBox_X);
                    boundingBox_Position.Set(port.Origin.X, port.Origin.Y, port.Origin.Z);

                    leg1YOffset.Set(boundingBox_Y.X, boundingBox_Y.Y, boundingBox_Y.Z);
                    leg1YOffset.Length = boundingBoxWidth / 2 + offset3;
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
                    componentDictionary[LEG].SetPropertyValue(leg1Length, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Direct);
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    boundingBox_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    boundingBox_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    boundingBox_Y = boundingBox_Z.Cross(boundingBox_X);

                    port = new Matrix4X4();
                    port = RefPortHelper.PortLCS("Structure");
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
                        switch (legData.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = 6, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = 8, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = 4, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = 2, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectXZ, boundingBox_Y);
                                    break;
                                }
                        }
                        if (flipFrame)
                            leg1CutbackAngle = -leg1CutbackAngle;
                        componentDictionary[LEG].SetPropertyValue(leg1CutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                    }
                    else
                    {
                        // Cut in BBX Plane
                        switch (legData.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = 8, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = 8, "IJOAhsSteelCPFlexPort", "CP7");
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = 4, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = 4, "IJOAhsSteelCPFlexPort", "CP7");
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    leg1CutbackAngle = -boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = 2, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = 2, "IJOAhsSteelCPFlexPort", "CP7");
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint6Leg.PropValue = 6, "IJOAhsSteelCP", "CP6");
                                    componentDictionary[LEG].SetPropertyValue(cardinalPoint7Leg.PropValue = 6, "IJOAhsSteelCPFlexPort", "CP7");
                                    componentDictionary[LEG].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    leg1CutbackAngle = boundingBox_Z.Angle(projectZY, boundingBox_X);
                                    break;
                                }
                        }
                        if (flipFrame)
                            leg1CutbackAngle = -leg1CutbackAngle;
                        componentDictionary[LEG].SetPropertyValue(leg1CutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                    }
                    JointHelper.CreatePrismaticJoint(LEG, "BeginFlex", LEG, "EndFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                double basePlateWidth = 0;
                double basePlateLength = 0;
                if (isBasePlate1)
                {
                    basePlate1Th = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1], "IJUAhsThickness1", "Thickness1")).PropValue;
                    basePlateWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1], "IJUAhsWidth1", "Width1")).PropValue;
                    basePlateLength = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE1], "IJUAhsLength1", "Length1")).PropValue;
                }
                componentDictionary[LEG].SetPropertyValue(leg1EndOverHang - basePlate1Th, "IJUAHgrOccOverLength", "EndOverLength");

                planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Direct);
                if (isPlacedOnSurface == false)
                {
                    if ((planeAngle * 180 / Math.PI) <= 45 || (planeAngle * 180 / Math.PI) >= 135)
                        JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", "-1", "Structure", Plane.XY);
                    else
                        JointHelper.CreatePointOnPlaneJoint(LEG, "EndFlex", "-1", "Structure", Plane.ZX);
                }
                // ==========================
                // Joints For Remaining Plates and Grout
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
                    JointHelper.CreateAngularRigidJoint(CAPPLATE1, "Port2", LEG, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, capPlate1Angle));
                if (isCapPlate2)
                    JointHelper.CreateAngularRigidJoint(CAPPLATE2, "Port2", SECTION, "BeginFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate2Angle));
                if (isCapPlate3)
                    JointHelper.CreateAngularRigidJoint(CAPPLATE3, "Port1", SECTION, "EndFace", new Vector(horizontalOffset, verticalOffset, 0), new Vector(0, 0, capPlate3Angle));

                //Joints for the Base Plates
                double basePlateOffset = 0;
                if (part.SupportsInterface("IJUAhsBasePlate1Off"))
                    basePlateOffset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBasePlate1Off", "BasePlate1Offset")).PropValue;
                if (isBasePlate1)
                {
                    if (structureConnection == 1)
                        JointHelper.CreateAngularRigidJoint(BASEPLATE1, "Port1", LEG, "EndFace", new Vector(0, 0, basePlateOffset), new Vector(0, 0, basePlate1Angle));
                    else if (structureConnection == 2)
                        JointHelper.CreateAngularRigidJoint(BASEPLATE1, "Port1", LEG, "EndFace", new Vector(0, legDepth / 2 + basePlate1Th, basePlateOffset), new Vector(Math.PI / 2, 0, basePlate1Angle));
                    else if (structureConnection == 3)
                        JointHelper.CreateAngularRigidJoint(BASEPLATE1, "Port1", LEG, "EndFace", new Vector(0, -legDepth / 2, basePlateOffset), new Vector(Math.PI / 2, 0, basePlate1Angle));
                }
                // ==========================
                // Joints For the Pipe Attachments
                // ==========================
                string partProgId = string.Empty;
                int index = 0;


                if (includePipeAtt1 == true)
                {
                    Part pipeAtt1Part = (Part)componentDictionary[PIPEATT1[0]].GetRelationship("madeFrom", "part").TargetObjects[0];
                    partProgId = (string)((PropertyValueString)(pipeAtt1Part).GetPropertyValue("IJDSymbolDefHelper", "ProgId")).PropValue;
                    double[] pipedia = new double[SupportHelper.SupportedObjects.Count];
                    for (index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                    {
                        PipeObjectInfo pipeInfo1 = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(index);
                        pipedia[index - 1] = pipeInfo1.OutsideDiameter;
                    }

                    switch (partProgId)
                    {
                        case "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Guide":
                            double guidegap = 0;
                            for (index = 0; index < SupportHelper.SupportedObjects.Count; index++)
                            {
                                guidegap = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT1[index]], "IJUAhsGap1", "Gap1")).PropValue;
                                componentDictionary[PIPEATT1[index]].SetPropertyValue(shoeHeight + pipedia[index] / 2, "IJUAhsShoeHeight", "ShoeHeight");
                                componentDictionary[PIPEATT1[index]].SetPropertyValue(pipedia[index], "IJOAhsPipeOD", "PipeOD");
                                componentDictionary[PIPEATT1[index]].SetPropertyValue(shoeWidth + guidegap, "IJUAhsGap1", "Gap1");
                            }
                            break;
                        case "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.UBolt":
                            for (index = 0; index < SupportHelper.SupportedObjects.Count; index++)
                            {
                                componentDictionary[PIPEATT1[index]].SetPropertyValue(shoeHeight + pipedia[index] / 2, "IJOAhsSteelThickness", "SteelThickness");
                            }
                            break;
                        case "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strap":
                            double strapWidthInside = 0;
                            for (index = 0; index < SupportHelper.SupportedObjects.Count; index++)
                            {
                                strapWidthInside = (double)((PropertyValueDouble)pipeAtt1Part.GetPropertyValue("IJUAhsStrap", "StrapWidthInside")).PropValue;
                                if (pipeConnection1Welds.Count > 0)
                                {
                                    weld = pipeConnection1Welds[index];
                                    JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", PIPEATT1[index], "Steel", new Vector(strapWidthInside / 2, 0, 0), new Vector(0, 0, 0));
                                }
                            }
                            break;
                    }
                    string routePort = "";
                    for (index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                    {
                        if (index == 1)
                            routePort = "Route";
                        else
                            routePort = "Route_" + index;
                        JointHelper.CreateTranslationalJoint(SECTION, "BeginFlex", PIPEATT1[index - 1], "Route", Plane.YZ, Plane.NegativeYZ, Axis.Z, Axis.Y, 0);
                        JointHelper.CreatePointOnAxisJoint(PIPEATT1[index - 1], "Route", "-1", routePort, Axis.X);
                    }

                }

                //==========================
                // Joints For the Remaining Weld Objects
                //==========================
                for (weldCount = 0; weldCount < otherWelds.Count; weldCount++)
                {
                    weld = otherWelds[weldCount];
                    string legWeldPort = "EndFace";
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
                                if (!string.IsNullOrEmpty(CAPPLATE2))
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
                                if (!string.IsNullOrEmpty(CAPPLATE3))
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
                                if (!string.IsNullOrEmpty(CAPPLATE3))
                                {
                                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3], "IJUAhsLength1", "Length1")).PropValue;
                                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE3], "IJUAhsWidth1", "Width1")).PropValue;
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
                double spanOffset1, spanOffset2, lengthOffset1, lengthOffset2;
                Boolean excludeNotes = (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;

                //SPAN
                switch (spanDefinition)
                {
                    case 1: //Pipe Center Lines
                        spanOffset1 = offset1 + leftPipeDiameter / 2;
                        spanOffset2 = offset2 + rightPipeDiameter / 2;
                        break;
                    case 2: //steel length
                        spanOffset1 = 0;
                        spanOffset2 = 0;
                        break;
                    default: //Pipe Center Lines
                        spanOffset1 = offset1 + leftPipeDiameter / 2;
                        spanOffset2 = offset2 + rightPipeDiameter / 2;
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
                ControlPoint controlPoint;
                Note note;

                //If lapped offset dim points
                double length1Offset = 0;
                if ((FrameAssemblyServices.SteelConnection)connection1Type == FrameAssemblyServices.SteelConnection.SteelConnection_Lapped)
                {
                    length1Offset = sectionWidth / 2;

                    if (connection1Mirror == true)
                        length1Offset = -length1Offset;
                }
                else if ((FrameAssemblyServices.SteelConnection)connection1Type == FrameAssemblyServices.SteelConnection.SteelConnection_Nested)
                {
                    length1Offset = sectionWidth / 2 - section.flangeThickness;

                    if (connection1Mirror == true)
                        length1Offset = -length1Offset;
                }

                //Set Dimension Points
                if (reflectMember == false)
                    note = CreateNote("SpanStart", SECTION, "BeginCap", new Position(0, lengthOffset1, spanOffset1), " ", true, 2, 1, out controlPoint);
                else
                    note = CreateNote("SpanStart", SECTION, "EndCap", new Position(0, -lengthOffset1, -spanOffset1), " ", true, 2, 1, out controlPoint);

                if (reflectMember == false)
                    note = CreateNote("SpanEnd", SECTION, "EndFlex", new Position(0, lengthOffset2, -spanOffset2), " ", true, 2, 1, out controlPoint);
                else
                    note = CreateNote("SpanEnd", SECTION, "BeginFlex", new Position(0, -lengthOffset2, spanOffset2), " ", true, 2, 1, out controlPoint);

                note = CreateNote("Length1End", LEG, "EndCap", new Position(0, legDepth / 2, 0), " ", true, 2, 1, out controlPoint);

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
                double sectionDepth = 0, sectionWidth = 0, legDepth = 0, legWidth = 0;
                part = (IPart)componentDictionary[SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                section = FrameAssemblyServices.GetSectionDataFromPart(part.PartNumber);
                if ((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue == 0 || (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue == 180)
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
                if ((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue == 0 || (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue == 180)
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
                int shoeHeightDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsFrameShoeHeightDef", "ShoeHeightDefinition")).PropValue;
                if (!string.IsNullOrEmpty(((string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue)))
                {
                    GenericHelper.GetDataByRule((string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue, null, out shoeHeight);
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
                length1Value = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[LEG], "IJUAHgrOccLength", "Length")).PropValue;
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
                PropertyValueCodelist steelStandardList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogpart, "IJUAhsSteelStandard", "SteelStandard");

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
