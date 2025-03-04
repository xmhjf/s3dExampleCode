//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AK1.cs
//   PipeSupport_Assy,Ingr.SP3D.Content.Support.Rules.Cantilever
//   Author       :  Patrick Weckworth
//   Creation Date:  11/24/2014
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//24-November-2014  Pweckw  CR-CP-245791- Convert HS_S3DCantilever to C# .Net 
//30-March-2014     PVK     CR-CP-245789	Modify the exsisting .Net Frame_Assy to behave like URS Frame supports
//28-Dec-2015       PVK	    Resolve Coverity issues found in April
//27-Apr-2015       PVK     TR-CP-253033	Elevation CP not shown by default for frame supports.
//06-May-2015       PVK     CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
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
using System.Linq;
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

    public class Cantilever : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string SECTION = "SECTION";
        private const string BRACE = "BRACE";
        private const string CAPPLATE = "CAPPLATE";
        private const string BASEPLATE = "BASEPLATE";
        private const string BRACEPLATE = "BRACEPLATE";
        private const string STRUCT_CONN = "STRUCT_CONN";

        string[] PIPEATT1, PIPEATT2, ANCHORS;
        private Boolean includeBolts, includePipeAtt1, includePipeAtt2;
        private int boltQty, pipeAtt1Qty, pipeAtt2Qty;
        private string boltPart, boltRule;
        FrameAssemblyServices.HSSteelMember section = new FrameAssemblyServices.HSSteelMember();
        FrameAssemblyServices.HSSteelMember brace = new FrameAssemblyServices.HSSteelMember();
        Collection<FrameAssemblyServices.WeldData> weldCollection = new Collection<FrameAssemblyServices.WeldData>();
        bool isBrace, isCapPlate, isBasePlate, isBracePlate, isPipeAtt1, isPipeAtt2, isbolt1;
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
                        string family = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support,"IJUAHgrURSCommon", "Family")).PropValue;
                        string type = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support,"IJUAHgrURSCommon", "Type")).PropValue;
                        if (family != "" && type != null)
                            FrameAssemblyServices.CheckSupportWithFamilyAndType(this, family, type);
                    }

                    // Add the Steel Section for the Frame Support
                    string member1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1", "Member1Part")).PropValue;
                    string member1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1Rl", "Member1Rule")).PropValue;
                    FrameAssemblyServices.AddPart(this, SECTION, member1Part, member1Rule, parts);

                    try
                    {
                        string bracePart = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1", "Leg1Part")).PropValue;
                        string braceRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1Rl", "Leg1Rule")).PropValue;
                        isBrace = FrameAssemblyServices.AddPart(this, BRACE, bracePart, braceRule, parts);
                    }
                    catch { isBrace = false; }

                    // Add the Plates
                    string capPlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1", "CapPlate1Part")).PropValue;
                    string capPlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1Rl", "CapPlate1Rule")).PropValue;
                    isCapPlate = FrameAssemblyServices.AddPart(this, CAPPLATE, capPlate1Part, capPlate1Rule, parts);
                    string basePlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate1", "BasePlate1Part")).PropValue;
                    string basePlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate1Rl", "BasePlate1Rule")).PropValue;
                    isBasePlate = FrameAssemblyServices.AddPart(this, BASEPLATE, basePlate1Part, basePlate1Rule, parts);
                    if (isBrace)
                    {
                        string bracePlatePart = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate2", "BasePlate2Part")).PropValue;
                        string bracePlateRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate2Rl", "BasePlate2Rule")).PropValue;
                        isBracePlate = FrameAssemblyServices.AddPart(this, BRACEPLATE, bracePlatePart, bracePlateRule, parts);
                    }
                    else
                        isBracePlate = false;

                    includePipeAtt1 = false;
                    string pipeAtt1Part = "";
                    string pipeAtt1Rule = "";
                    try
                    {
                        pipeAtt1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt1", "PipeAtt1Part")).PropValue;
                        pipeAtt1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt1Rl", "PipeAtt1Rule")).PropValue;
                        pipeAtt1Qty = SupportHelper.SupportedObjects.Count;
                    }
                    catch 
                    {
                        includePipeAtt1 = false;
                        isPipeAtt1 = false;
                        includePipeAtt2 = false;
                        isPipeAtt2 = false;
                        pipeAtt1Qty = 0;
                    }

                    PIPEATT1 = new string[pipeAtt1Qty];
                    if (pipeAtt1Qty > 0 && (!string.IsNullOrEmpty(pipeAtt1Part) || !string.IsNullOrEmpty(pipeAtt1Rule)))
                    {
                        includePipeAtt1 = true;
                        Array.Resize(ref PIPEATT1, pipeAtt1Qty);
                        for (int index = 0; index < pipeAtt1Qty; index++)
                        {
                            PIPEATT1[index] = "PIPEATT1_" + index;
                            isPipeAtt1 = FrameAssemblyServices.AddPart(this, PIPEATT1[index], pipeAtt1Part, pipeAtt1Rule, parts);
                        }
                    }

                    includePipeAtt2 = false;
                    if (includePipeAtt1)
                    {
                        string pipeAtt2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt2", "PipeAtt2Part")).PropValue;
                        string pipeAtt2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt2Rl", "PipeAtt2Rule")).PropValue;
                        pipeAtt2Qty = 2 * pipeAtt1Qty;

                        PIPEATT2 = new string[pipeAtt2Qty];
                        if (pipeAtt2Qty > 0 && (!string.IsNullOrEmpty(pipeAtt2Part) || !string.IsNullOrEmpty(pipeAtt2Rule)))
                        {
                            includePipeAtt2 = true;
                            Array.Resize(ref PIPEATT2, pipeAtt2Qty);
                            for (int index = 0; index < pipeAtt2Qty; index++)
                            {
                                PIPEATT2[index] = "PIPEATT2_" + index;
                                isPipeAtt2 = FrameAssemblyServices.AddPart(this, PIPEATT2[index], pipeAtt2Part, pipeAtt2Rule, parts);
                            }
                        }
                    }


                    // Add the Anchor Bolts (Requires Base Plate)
                    includeBolts = false;
                    if (isBasePlate)
                    {
                        boltPart = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1", "Bolt1Part")).PropValue;
                        boltRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1Rl", "Bolt1Rule")).PropValue;
                        boltQty = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1Qty", "Bolt1Quantity")).PropValue;

                        ANCHORS = new string[boltQty];
                        if (boltQty > 0 && (!string.IsNullOrEmpty(boltPart) || !string.IsNullOrEmpty(boltRule)))
                        {
                            includeBolts = true;
                            Array.Resize(ref ANCHORS, boltQty);
                            for (int index = 0; index < boltQty; index++)
                            {
                                ANCHORS[index] = "ANCHORS_" + index;
                                isbolt1 = FrameAssemblyServices.AddPart(this, ANCHORS[index], boltPart, boltRule, parts);
                            }
                        }

                    }

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        parts.Add(new PartInfo(STRUCT_CONN, "Log_Conn_Part_1"));
                    }

                    // Add the Weld Objects from the Weld Sheet
                    string weldServiceClassName = string.Empty;
                    if (part.SupportsInterface("IJUAhsWeldServClass"))
                        weldServiceClassName = (String)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsWeldServClass", "WeldServiceClassName")).PropValue;
                    if (weldServiceClassName == null)
                        weldServiceClassName = string.Empty;
                    weldCollection = FrameAssemblyServices.AddWeldsFromCatalog(this, parts, "IJUAhsBCWelds", ((IPart)part).PartNumber, weldServiceClassName);
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
                    string ImpliedPartServiceClass = null;

                    // Add the Optional Bolts as Implied Parts
                    if (part.SupportsInterface("IJUAhsImpServClass"))
                        ImpliedPartServiceClass = (String)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsImpServClass", "ImpServiceClassName")).PropValue;
                    if (ImpliedPartServiceClass == string.Empty)
                        ImpliedPartServiceClass = null;

                    if (ImpliedPartServiceClass != null)
                        FrameAssemblyServices.AddImpliedPartFromCatalog(this, impliedParts, ImpliedPartServiceClass);
                    
                     if (isBasePlate)
                        isbolt1 = FrameAssemblyServices.AddImpliedPart(this, boltPart, boltRule, impliedParts, null, boltQty);

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
                FrameAssemblyServices.HSSteelConfig sectionData = new FrameAssemblyServices.HSSteelConfig();
                FrameAssemblyServices.HSSteelConfig braceData = new FrameAssemblyServices.HSSteelConfig();

                // ==========================
                // Get Required Information about the Steel Parts
                // ==========================
                // Get the Steel Cross Section Data
                section = FrameAssemblyServices.GetSectionDataFromPartIndex(this, SECTION);
                if (isBrace)
                    brace = FrameAssemblyServices.GetSectionDataFromPartIndex(this, BRACE);

                // Get the Steel Configuration Data
                sectionData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue);

                if (isBrace)
                    braceData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue);

                double sectionDepth = 0, sectionWidth = 0;
                double braceDepth = 0, braceWidth = 0;
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

                if (isBrace)
                {
                    if (braceData.Orient == (FrameAssemblyServices.SteelOrientationAngle)0 || braceData.Orient == (FrameAssemblyServices.SteelOrientationAngle)180)
                    {
                        braceDepth = brace.depth;
                        braceWidth = brace.width;
                    }
                    else
                    {
                        braceDepth = brace.width;
                        braceWidth = brace.depth;
                    }
                }

                // ==========================
                // Create the Frame Bounding Box
                // ==========================
                double boundingBoxWidth = 0, boundingBoxDepth = 0;
                string boundingBoxPort = "BBFrame_Low", boundingBoxName = "BBFrame";
                
                int frameOrientation = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOrientation", "FrameOrientation" )).PropValue;
                Boolean includeInsulation = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue;
                Boolean mirrorFrame = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsMirrorFrame", "MirrorFrame")).PropValue;
                FrameAssemblyServices.CreateFrameBoundingBox(this, boundingBoxName, (FrameAssemblyServices.FrameBBXOrientation)frameOrientation, includeInsulation, mirrorFrame, SupportedHelper.IsSupportedObjectVertical(1, 45));

                boundingBoxWidth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Width;
                boundingBoxDepth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Height;
                // ==========================
                // Determine the number of Supporting Objects
                // ==========================
                int supportingCopunt = 0;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    // For Place-By-Reference Command No Supporting Objects are Selected
                    // However, we want to treat it like 1 was selected
                    supportingCopunt = 1;
                else
                    supportingCopunt = SupportHelper.SupportingObjects.Count;

                // Check if the Support is placed on a Surface, such as Equipment
                Boolean isPlacedOnSurface = false;
                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByReference && SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.GenericSurface)
                        isPlacedOnSurface = true;
                }
                else
                    isPlacedOnSurface = false;
                //==========================
                //Handle all the Frame Configurations and Toggles
                //==========================
                double sectionZOffset;
                bool reflectMember;

                double routeToZAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, OrientationAlong.Global_Z) * 180 / Math.PI;

                if (Configuration == 1 && (routeToZAngle < 45 || routeToZAngle > 315) ||
                        Configuration == 2 && (routeToZAngle >= 45 && routeToZAngle <= 315))
                {
                    reflectMember = false;
                }
                else
                {
                    reflectMember = true;
                }

                if (!mirrorFrame)
                    reflectMember = !reflectMember;

                // ==========================
                // Determine the Connections to the Structural Steel
                // ==========================
                double basePlate1Th = 0;
                if (isBasePlate)
                    basePlate1Th = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsThickness1", "Thickness1")).PropValue;
                else
                    basePlate1Th = 0;

                FrameAssemblyServices.HSSteelMember supportingSection = new FrameAssemblyServices.HSSteelMember();
                int supportingFace = 0;
                double structWidthOffset = 0;
                long structWidthOffsetType = 0;
                int structureConnection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsStructureConn", "StructureConnection")).PropValue;

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

                if (SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel"))
                {
                    supportingFace = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                    supportingSection = FrameAssemblyServices.GetSupportingSectionData(this, 1);
                }

                bool structPortWrong = false;
                double routeZBeamYAngle = 0;
                Vector beamXAxis = new Vector(0,0,0);
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    double structXRouteY = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, "Route", PortAxisType.Y, OrientationAlong.Direct);
                    structPortWrong = structXRouteY > 3 * Math.PI / 4 && structXRouteY < 5 * Math.PI / 4;

                    ILocalCoordinateSystem localCoord = support.SupportingObjects[0] as ILocalCoordinateSystem;

                    Matrix4X4 routeMatrix = RefPortHelper.PortLCS("Route");
                    Vector routeZAxis = routeMatrix.ZAxis;
                    Vector beamYAxis = localCoord.YAxis;
                    beamXAxis = localCoord.XAxis;
                    routeZBeamYAngle = routeZAxis.Angle(beamYAxis, new Vector(0,0,1));

                    make90DegAngle(ref routeZBeamYAngle);

                    if (Configuration == 2)
                        flip90DegAngle(ref routeZBeamYAngle);

                    if (beamXAxis.Z < 0 && (HgrCompareDoubleService.cmpdbl(routeZBeamYAngle, 0) || HgrCompareDoubleService.cmpdbl(routeZBeamYAngle, Math.PI)))
                        flip90DegAngle(ref routeZBeamYAngle);
                    
                    switch (structureConnection)
                    {
                        case 1: //normal
                            structWidthOffset = 0;
                            break;
                        case 2: //lapped
                            {
                                if (reflectMember)
                                    structWidthOffsetType = 1;
                                else
                                    structWidthOffsetType = 2;
                                break;
                            }
                        case 3: //reverse lapped
                            {
                                if (reflectMember)
                                    structWidthOffsetType = 2;
                                else
                                    structWidthOffsetType = 1;
                                break;
                            }
                    }
                    if (structWidthOffsetType == 1)
                    {
                        if (HgrCompareDoubleService.cmpdbl(Math.PI, routeZBeamYAngle))
                        {
                            switch (supportingFace)
                            {
                                case 513:
                                    structWidthOffset = -sectionWidth / 2;
                                    break;
                                case 514:
                                    structWidthOffset = -supportingSection.depth - sectionWidth / 2;
                                    break;
                                case 257:
                                case 258:
                                    structWidthOffset = -supportingSection.depth / 2 - sectionWidth / 2;
                                    break;
                            }
                        }
                        else if (HgrCompareDoubleService.cmpdbl(0, routeZBeamYAngle))
                        {
                            switch (supportingFace)
                            {
                                case 513:
                                    structWidthOffset = -supportingSection.depth - sectionWidth / 2;
                                    break;
                                case 514:
                                    structWidthOffset = -sectionWidth / 2;
                                    break;
                                case 257:
                                case 258:
                                    structWidthOffset = -supportingSection.depth / 2 - sectionWidth / 2;
                                    break;
                            }
                        }
                        else if (HgrCompareDoubleService.cmpdbl(3 * Math.PI / 2, routeZBeamYAngle))
                        {
                            structWidthOffset = -supportingSection.width / 2 - sectionWidth / 2;
                            switch (supportingFace)
                            {
                                case 257:
                                    if (supportingSection.sectionType == "W")
                                        //structWidthOffset += supportingSection.webThickness / 2;
                                        structWidthOffset -= supportingSection.webThickness / 2;
                                    else
                                        structWidthOffset = -supportingSection.width - sectionWidth / 2;
                                    break;
                                case 258:
                                    if (supportingSection.sectionType == "W") 
                                        //structWidthOffset -= supportingSection.webThickness / 2;
                                        structWidthOffset += supportingSection.webThickness / 2;
                                    else
                                        structWidthOffset = -sectionWidth / 2;
                                    break;
                            }
                        }
                        else if (HgrCompareDoubleService.cmpdbl(Math.PI / 2, routeZBeamYAngle))
                        {
                            structWidthOffset = -supportingSection.width / 2 - sectionWidth / 2;
                            switch (supportingFace)
                            {
                                case 257:
                                    if (supportingSection.sectionType == "W")
                                        //structWidthOffset -= supportingSection.webThickness / 2;
                                        structWidthOffset += supportingSection.webThickness / 2;
                                    else
                                        structWidthOffset = -sectionWidth / 2;
                                    break;
                                case 258:
                                    if (supportingSection.sectionType == "W")
                                        //structWidthOffset += supportingSection.webThickness / 2;
                                        structWidthOffset -= supportingSection.webThickness / 2;
                                    else
                                        structWidthOffset = -supportingSection.width - sectionWidth / 2;
                                    break;
                            }
                        }
                        if (isBasePlate && SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            structWidthOffset -= basePlate1Th;
                    }
                    else if (structWidthOffsetType == 2)
                    {
                        if (HgrCompareDoubleService.cmpdbl(Math.PI, routeZBeamYAngle))
                        {
                            switch (supportingFace)
                            {
                                case 513:
                                    structWidthOffset = supportingSection.depth + sectionWidth / 2;
                                    break;
                                case 514:
                                    structWidthOffset = sectionWidth / 2;
                                    break;
                                case 257:
                                case 258:
                                    structWidthOffset = supportingSection.depth / 2 + sectionWidth / 2;
                                    break;
                            }
                        }
                        else if (HgrCompareDoubleService.cmpdbl(0, routeZBeamYAngle))
                        {
                            switch (supportingFace)
                            {
                                case 513:
                                    structWidthOffset = sectionWidth / 2;
                                    break;
                                case 514:
                                    structWidthOffset = supportingSection.depth + sectionWidth / 2;
                                    break;
                                case 257:
                                case 258:
                                    structWidthOffset = supportingSection.depth / 2 + sectionWidth / 2;
                                    break;
                            }
                        }
                        else if (HgrCompareDoubleService.cmpdbl(3 * Math.PI / 2, routeZBeamYAngle))
                        {
                            structWidthOffset = supportingSection.width / 2 + sectionWidth / 2;
                            switch (supportingFace)
                            {
                                case 257:
                                    if (supportingSection.sectionType == "W")
                                        structWidthOffset -= supportingSection.webThickness / 2;
                                    else
                                        structWidthOffset = sectionWidth / 2;
                                    break;
                                case 258:
                                    if (supportingSection.sectionType == "W")
                                        structWidthOffset += supportingSection.webThickness / 2;
                                    else
                                        structWidthOffset = supportingSection.width + sectionWidth / 2;
                                    break;
                            }
                        }
                        else if (HgrCompareDoubleService.cmpdbl(Math.PI / 2, routeZBeamYAngle))
                        {
                            structWidthOffset = supportingSection.width / 2 + sectionWidth / 2;
                            switch (supportingFace)
                            {
                                case 257:
                                    if (supportingSection.sectionType == "W")
                                        structWidthOffset += supportingSection.webThickness / 2;
                                    else
                                        structWidthOffset = supportingSection.width + sectionWidth / 2;
                                    break;
                                case 258:
                                    if (supportingSection.sectionType == "W")
                                        structWidthOffset -= supportingSection.webThickness / 2;
                                    else
                                        structWidthOffset = sectionWidth / 2;
                                    break;
                            }
                        }
                        if (isBasePlate && SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            structWidthOffset += basePlate1Th;
                    }

                    if (structPortWrong && structWidthOffsetType > 0)
                        reflectMember = !reflectMember;

                    if (structureConnection > 1 && !reflectMember)
                        structWidthOffset = -structWidthOffset;                    

                    // Attach the Structure Conn to the Structure
                    double structYBoxXAngle = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Y, boundingBoxPort, PortAxisType.X, OrientationAlong.Direct);
                    if (structYBoxXAngle < Math.PI / 2 || structYBoxXAngle > 7 * Math.PI / 4)
                        JointHelper.CreateRigidJoint("-1", "Structure", STRUCT_CONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, structWidthOffset, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "Structure", STRUCT_CONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -structWidthOffset, 0);
                }

                // ==========================
                // Organize the Welds into several Collections
                // ==========================
                Collection<FrameAssemblyServices.WeldData> frameConnection1Welds = new Collection<FrameAssemblyServices.WeldData>();
                Collection<FrameAssemblyServices.WeldData> pipeConnection1Welds = new Collection<FrameAssemblyServices.WeldData>();
                Collection<FrameAssemblyServices.WeldData> pipeConnection2Welds = new Collection<FrameAssemblyServices.WeldData>();
                Collection<FrameAssemblyServices.WeldData> pipeConnection3Welds = new Collection<FrameAssemblyServices.WeldData>();

                FrameAssemblyServices.WeldData weld = new FrameAssemblyServices.WeldData();
                for (int weldCount = 0; weldCount < weldCollection.Count; weldCount++)
                {
                    weld = weldCollection[weldCount];
                    if ((weld.connection).ToUpper() == "D")
                        pipeConnection1Welds.Add(weld);
                    if ((weld.connection).ToUpper() == "E")
                        pipeConnection2Welds.Add(weld);
                    if ((weld.connection).ToUpper() == "F")
                        pipeConnection3Welds.Add(weld);
                    else
                        frameConnection1Welds.Add(weld);
                }

                // ==========================
                // Get ExtremPipeIdx and the Max Distance between Structure and the Bottom Pipe.
                // ==========================
                int botPipeIdx = 1;
                int primParallelPipeIdx = 0, primPerpendicularPipeIdx = 0;
                double primParPipeDia, primPerPipeDia = 0;
                double strBotPipeDist;
                double beginOverLength; 
                double botPipeDia, insulTh;
                bool bMixedPipes = false;
                PipeObjectInfo pipeInfo;
                
                if (SupportHelper.SupportedObjects.Count > 1)
                {
                    strBotPipeDist = GetStructBotPipeDist(ref botPipeIdx);
                    //Get Primary Parallel and Perpendicular Pipe Diameter
                    GetPrimaryPipeIdxes(ref primParallelPipeIdx, ref primPerpendicularPipeIdx); //gets index of first parallel and first perp. pipe
        
                    bool bParallelPipes = false, bPerpentclrPipes = false;
        
                    if (primParallelPipeIdx != 0)
                    {
                        pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(primParallelPipeIdx);
                        primParPipeDia = pipeInfo.OutsideDiameter;
                        insulTh = pipeInfo.InsulationThickness;
                        bParallelPipes = true;
                    }
                    if (primPerpendicularPipeIdx != 0)
                    {
                        pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(primPerpendicularPipeIdx);
                        primPerPipeDia = pipeInfo.OutsideDiameter;
                        insulTh = pipeInfo.InsulationThickness;
                        bPerpentclrPipes = true;
                    }

                    if (bPerpentclrPipes && bParallelPipes)
                        bMixedPipes = true;

                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct && bMixedPipes)
                    {
                        if (reflectMember)
                            structWidthOffset = primPerPipeDia / 2 + sectionWidth / 2;
                        else
                            structWidthOffset = -(primPerPipeDia / 2 + sectionWidth / 2);
                    }
                }
                else
                {
                    botPipeIdx = 1;
                    strBotPipeDist = RefPortHelper.DistanceBetweenPorts("Structure", "Route", PortAxisType.Z);
                }

                //Get Bottom Pipe Diameter
                pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(botPipeIdx);
                botPipeDia = pipeInfo.OutsideDiameter;
                insulTh = pipeInfo.InsulationThickness;
    
                //Compare the Bottom pipe distance with BBX distance for the over Length
                double tempDist;
                tempDist = Math.Abs(RefPortHelper.DistanceBetweenPorts("Structure", boundingBoxPort, PortAxisType.Z));

                if (tempDist < (strBotPipeDist + botPipeDia / 2))
                    beginOverLength = (strBotPipeDist + botPipeDia / 2) - tempDist;
                else
                    beginOverLength = 0;

                //===========================

                // ==========================
                // Offset 1 (The offset to the Leg)
                // ==========================
                //PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                //double pipeDiameter = routeInfo.OutsideDiameter;
                //double insulationThickness = routeInfo.InsulationThickness;
                double offset1 = 0;
                int offset1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset1Def", "Offset1Definition")).PropValue;
                int offset1Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset1Sel", "Offset1Selection")).PropValue;
                if (offset1Selection == 1)
                {
                    string offset1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset1Rl", "Offset1Rule")).PropValue;
                    GenericHelper.GetDataByRule(offset1Rule, null, out offset1);
                    support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                }
                else
                    offset1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOffset1", "Offset1Value")).PropValue;

                if (offset1Definition == 2) //pipe center line
                {
                    if (includeInsulation)
                        offset1 = offset1 - botPipeDia / 2 - insulTh;
                    else
                        offset1 = offset1 - botPipeDia / 2;
                }
                else //edge of pipe
                { }

                // ==========================
                // Shoe Height
                // ==========================
                double shoeHeight = 0;
                int shoeHeightDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameShoeHeightDef", "ShoeHeightDefinition")).PropValue;
                int shoeHeightSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameShoeHeightSel", "ShoeHeightSelection")).PropValue;
                // Get the Shoe Height From the Input Attributes
                if (shoeHeightSelection == 1)
                {
                    string shoeHeightRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameShoeHeightRl", "ShoeHeightRule")).PropValue;
                    GenericHelper.GetDataByRule(shoeHeightRule, null, out shoeHeight);
                    support.SetPropertyValue(shoeHeight, "IJUAhsFrameShoeHeight", "ShoeHeightValue");
                }
                else
                    shoeHeight = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameShoeHeight", "ShoeHeightValue")).PropValue;

                switch (shoeHeightDefinition)
                {
                    case 1:
                        // Edge of Bounding Box
                        break;
                    case 2:
                        // Centerline of Primary Pipe
                        shoeHeight = shoeHeight + RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Y);
                        break;
                    default:
                        //Edge of Bounding Box
                        break;
                }

                // ==========================
                // Member 1 End Overhang
                // ==========================
                double member1EndOverhang = 0;
                int member1EndOverhangDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1EndOHDef", "Member1EndOverhangDefinition")).PropValue;
                int member1EndOverhangSelection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1EndOHSel", "Member1EndOverhangSelection")).PropValue;
                if (member1EndOverhangSelection == 1)
                {
                    string member1EndOverhangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsMember1EndOHRl", "Member1EndOverhangRule")).PropValue;
                    GenericHelper.GetDataByRule(member1EndOverhangRule, null, out member1EndOverhang);
                    support.SetPropertyValue(member1EndOverhang, "IJUAhsMember1EndOH", "Member1EndOverhangValue");
                }
                else
                    member1EndOverhang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsMember1EndOH", "Member1EndOverhangValue")).PropValue;

                if (structureConnection != 1)
                {
                    if (Configuration == 1)
                        flip90DegAngle(ref routeZBeamYAngle);

                    if (beamXAxis.Z < 0)
                        flip90DegAngle(ref routeZBeamYAngle);

                    double distToFarEdge = 0, distToCenter = 0, distToCloseEdge = 0;
                    if (HgrCompareDoubleService.cmpdbl(routeZBeamYAngle, 0))
                    {
                        if (supportingFace == 257)
                        {
                            if (supportingSection.sectionType == "W")
                            {
                                distToCloseEdge = supportingSection.webThickness / 2 - supportingSection.width / 2;
                                distToCenter = supportingSection.webThickness / 2;
                                distToFarEdge = supportingSection.webThickness / 2 + supportingSection.width / 2;
                            }
                            else
                            {
                                distToCloseEdge = 0;
                                distToCenter = supportingSection.width / 2;
                                distToFarEdge = supportingSection.width;
                            }
                        }
                        else if (supportingFace == 258)
                        {
                            if (supportingSection.sectionType == "W")
                            {
                                distToCloseEdge = -supportingSection.webThickness / 2 - supportingSection.width / 2;
                                distToCenter = -supportingSection.webThickness / 2;
                                distToFarEdge = -supportingSection.webThickness / 2 + supportingSection.width / 2;
                            }
                            else
                            {
                                distToCloseEdge = -supportingSection.width;
                                distToCenter = -supportingSection.width / 2;
                                distToFarEdge = 0;
                            }
                        }
                        else 
                        {
                            distToCloseEdge = -supportingSection.width / 2;
                            distToCenter = 0;
                            distToFarEdge = supportingSection.width / 2;
                        }
                    }
                    else if (HgrCompareDoubleService.cmpdbl(routeZBeamYAngle, Math.PI / 2))
                    {
                        if (supportingFace == 513)
                        {
                            distToCloseEdge = 0;
                            distToCenter = supportingSection.depth / 2;
                            distToFarEdge = supportingSection.depth; 
                        }
                        else if (supportingFace == 514)
                        {
                            distToCloseEdge = -supportingSection.depth;
                            distToCenter = -supportingSection.depth / 2;
                            distToFarEdge = 0;
                        }
                        else
                        {
                            distToCloseEdge = -supportingSection.depth / 2;
                            distToCenter = 0;
                            distToFarEdge = supportingSection.depth / 2;
                        }
                    }
                    else if (HgrCompareDoubleService.cmpdbl(routeZBeamYAngle, Math.PI))
                    {
                        if (supportingFace == 257)
                        {
                            if (supportingSection.sectionType == "W")
                            {
                                distToCloseEdge = -supportingSection.webThickness / 2 - supportingSection.width / 2;
                                distToCenter = -supportingSection.webThickness / 2;
                                distToFarEdge = -supportingSection.webThickness / 2 + supportingSection.width / 2;
                            }
                            else
                            {
                                distToCloseEdge = -supportingSection.width;
                                distToCenter = -supportingSection.width / 2;
                                distToFarEdge = 0;
                            }
                        }
                        else if (supportingFace == 258)
                        {
                            if (supportingSection.sectionType == "W")
                            {
                                distToCloseEdge = supportingSection.webThickness / 2 - supportingSection.width / 2;
                                distToCenter = supportingSection.webThickness / 2;
                                distToFarEdge = supportingSection.webThickness / 2 + supportingSection.width / 2;
                            }
                            else
                            {
                                distToCloseEdge = 0;
                                distToCenter = supportingSection.width / 2;
                                distToFarEdge = supportingSection.width;
                            }
                        }
                        else
                        {
                            distToCloseEdge = -supportingSection.width / 2;
                            distToCenter = 0;
                            distToFarEdge = supportingSection.width / 2;
                        }
                    }
                    else if (HgrCompareDoubleService.cmpdbl(routeZBeamYAngle, 3 * Math.PI / 2))
                    {
                        if (supportingFace == 513)
                        {
                            distToCloseEdge = -supportingSection.depth;
                            distToCenter = -supportingSection.depth / 2;
                            distToFarEdge = 0; 
                        }
                        else if (supportingFace == 514)
                        {
                            distToCloseEdge = 0;
                            distToCenter = supportingSection.depth / 2;
                            distToFarEdge = supportingSection.depth;
                        }
                        else
                        {
                            distToCloseEdge = -supportingSection.depth / 2;
                            distToCenter = 0;
                            distToFarEdge = supportingSection.depth / 2;
                        }
                    }

                    if (member1EndOverhangDefinition == 1) //Far edge of the steel
                        member1EndOverhang += distToFarEdge;
                    else if (member1EndOverhangDefinition == 3) //Near Edge of the steel
                        member1EndOverhang += distToCloseEdge;
                    else //Center line of the steel
                        member1EndOverhang += distToCenter;
                }
                //==========================
                //Handle all the Frame Configurations and Toggles
                //==========================
                if (!reflectMember)
                    sectionZOffset = shoeHeight;
                else
                    sectionZOffset = -(boundingBoxWidth + sectionDepth + shoeHeight);
                // ==========================
                // Set Steel CPs
                // ==========================
                // Set the CP for the SECTION Flex Ports (Used to connect the main section to the BBX)
                PropertyValueCodelist CP6 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsSteelCP", "CP6");
                PropertyValueCodelist CP7 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist CP2 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsSteelCP", "CP2");
                PropertyValueCodelist CP1 = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsSteelCP", "CP1");

                componentDictionary[SECTION].SetPropertyValue(CP6.PropValue = (int)Convert.ToDouble(5), "IJOAhsSteelCP", "CP6");
                componentDictionary[SECTION].SetPropertyValue(CP7.PropValue = (int)Convert.ToDouble(5), "IJOAhsSteelCPFlexPort", "CP7");
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsFlexPort", "FlexPortRotZ");
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsEndFlexPort", "EndFlexPortRotZ");

                // Set the CP's for the Section Ports that Connect to the Supporting Structure
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsEndCap", "EndCapRotZ");

                // ==========================
                // Joints To Connect the Main Steel Section to the BBX
                // ==========================
                double capPlateTh;
                if (isCapPlate)
                    capPlateTh = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE],"IJUAhsThickness1","Thickness1")).PropValue;
                else
                    capPlateTh = 0;
                if (reflectMember)
                    JointHelper.CreateRigidJoint(SECTION, "EndFlex", "-1", boundingBoxName + "_High", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -(boundingBoxDepth + offset1), sectionDepth / 2 + sectionZOffset, structWidthOffset);
                else
                    JointHelper.CreateRigidJoint(SECTION, "BeginFlex", "-1", boundingBoxName + "_High", Plane.XY, Plane.XY, Axis.X, Axis.X, boundingBoxDepth + offset1, -(sectionDepth / 2 + sectionZOffset), structWidthOffset);

                // ==========================
                // Joints To Connect Section To Supporting Structure
                // ==========================
                // If it is Equipment or a Shape, we will use GetProjectedPointOnSurface
                // For Steel, Slab and Wall, we will use a PointOnJoint
                Vector BB_X = new Vector(0, 0, 0), BB_Y = new Vector(0, 0, 0), BB_Z = new Vector(0, 0, 0), structZ = new Vector(0, 0, 0), structX = new Vector(0, 0, 0), structY = new Vector(0, 0, 0);
                Vector projectXZ = new Vector(0, 0, 0), projectZY = new Vector(0, 0, 0), memberYOffset = new Vector(0, 0, 0), memberZOffset = new Vector(0, 0, 0), memberProjectedNormal = new Vector(0, 0, 0);
                Position BB_Pos = new Position(0, 0, 0), memberProjectedPoint = new Position(0, 0, 0), memberStart = new Position(0, 0, 0);
                Matrix4X4 port = new Matrix4X4();
                double memberCutbackAngle = 0, memberLength = 0, planeAngle = 0;

                if (isPlacedOnSurface)
                {
                    // Get Projection from calculated point
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    BB_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    BB_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    BB_Y = BB_Z.Cross(BB_X);
                    BB_Pos.Set(port.Origin.X, port.Origin.Y, port.Origin.Z);

                    memberYOffset.Set(BB_Y.X, BB_Y.Y, BB_Y.Z);
                    if (!reflectMember)
                        sectionZOffset = -shoeHeight - sectionDepth / 2;
                    else
                        sectionZOffset = +boundingBoxWidth + shoeHeight + sectionDepth / 2;
                    memberYOffset.Length = sectionZOffset;
                    memberZOffset.Set(BB_Z.X, BB_Z.Y, BB_Z.Z);
                    memberZOffset.Length = -offset1;

                    memberStart = BB_Pos.Offset(memberYOffset);
                    memberStart = memberStart.Offset(memberZOffset);

                    try
                    {
                        BusinessObject SupportingFace = (BusinessObject)support.SupportingFaces.First();
                        SupportingHelper.GetProjectedPointOnSurface(memberStart, BB_Z, SupportingFace, out memberProjectedPoint, out memberProjectedNormal);
                        // Get Projection from calculated point
                        memberLength = memberStart.DistanceToPoint(memberProjectedPoint);
                    }
                    catch { }
                    componentDictionary[SECTION].SetPropertyValue(memberLength, "IJUAHgrOccLength", "Length");
                }
                else
                {
                    planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Direct);
                    port = new Matrix4X4();
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    BB_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    BB_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    BB_Y = BB_Z.Cross(BB_X);

                    port = new Matrix4X4();
                    port = RefPortHelper.PortLCS("Structure");
                    structX.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    structZ.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    structY = structZ.Cross(structX);

                    if ((planeAngle * 180 / Math.PI) > 45 && (planeAngle * 180 / Math.PI) < 135)
                    {
                        projectXZ = FrameAssemblyServices.ProjectVectorIntoPlane(structY, BB_Y);
                        if (FrameAssemblyServices.GetVectorProjection(structY, BB_Z) < 0)
                            projectXZ.Set(-projectXZ.X, -projectXZ.Y, -projectXZ.Z);
                        projectZY = FrameAssemblyServices.ProjectVectorIntoPlane(structY, BB_X);
                        if (FrameAssemblyServices.GetVectorProjection(structY, BB_Z) < 0)
                            projectZY.Set(-projectZY.X, -projectZY.Y, -projectZY.Z);
                    }
                    else
                    {
                        projectXZ = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, BB_Y);
                        if (FrameAssemblyServices.GetVectorProjection(structZ, BB_Z) < 0)
                            projectXZ.Set(-projectXZ.X, -projectXZ.Y, -projectXZ.Z);
                        projectZY = FrameAssemblyServices.ProjectVectorIntoPlane(structZ, BB_X);
                        if (FrameAssemblyServices.GetVectorProjection(structZ, BB_Z) < 0)
                            projectZY.Set(-projectZY.X, -projectZY.Y, -projectZY.Z);
                    }

                    PropertyValueCodelist endCutbackAnchorPointSectionList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJOAhsCutback", "EndCutbackAnchorPoint");
                    int CP = 2;
                    if (FrameAssemblyServices.AngleBetweenVectors(BB_Z, projectXZ) > FrameAssemblyServices.AngleBetweenVectors(BB_Z, projectZY))
                    {
                        // Cut Across BBX Plane
                        switch (sectionData.Orient)
                        {
                            case 0:
                                    CP = 6;
                                    memberCutbackAngle = -BB_Z.Angle(projectXZ, BB_Y);
                                    break;
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                    CP = 8;
                                    memberCutbackAngle = BB_Z.Angle(projectXZ, BB_Y);
                                    break;
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                    CP = 4;
                                    memberCutbackAngle = BB_Z.Angle(projectXZ, BB_Y);
                                    break;
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                    CP = 2;
                                    memberCutbackAngle = -BB_Z.Angle(projectXZ, BB_Y);
                                    break;
                        }
                    }
                    else
                    {
                        // Cut in BBX Plane
                        switch (sectionData.Orient)
                        {
                            case 0:
                                    CP = 8;
                                    memberCutbackAngle = -BB_Z.Angle(projectZY, BB_X);
                                    break;
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                    CP = 4;
                                    memberCutbackAngle = -BB_Z.Angle(projectZY, BB_X);
                                    break;
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                    CP = 2;
                                    memberCutbackAngle = BB_Z.Angle(projectZY, BB_X);
                                    break;
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                    CP = 6;
                                    memberCutbackAngle = BB_Z.Angle(projectZY, BB_X);
                                    break;
                        }
                    }

                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(CP), "IJOAhsSteelCP", "CP1");
                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(CP), "IJOAhsSteelCP", "CP2");
                    if (!reflectMember)
                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(CP), "IJOAhsCutback", "EndCutbackAnchorPoint");
                    else
                        componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(CP), "IJOAhsCutback", "BeginCutbackAnchorPoint");
                    
                    if (!reflectMember)
                        componentDictionary[SECTION].SetPropertyValue(memberCutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                    else
                        componentDictionary[SECTION].SetPropertyValue(memberCutbackAngle, "IJOAhsCutback", "CutbackBeginAngle");

                    if (reflectMember)
                        JointHelper.CreatePrismaticJoint(SECTION, "BeginFlex", SECTION, "EndFlex", Plane.ZX, Plane.ZX, Axis.NegativeZ, Axis.NegativeZ, 0, 0);
                    else
                        JointHelper.CreatePrismaticJoint(SECTION, "EndFlex", SECTION, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }

                //Set the OverLengths
                double tempOverLength = 0.0;
                if (structureConnection == 1)   //Normal
                    tempOverLength = member1EndOverhang - basePlate1Th;
                //else //For Lapped
                //{
                //    if (supportingFace == 257)
                //        tempOverLength = member1EndOverhang + supportingSection.webThickness / 2;
                //    else if (supportingFace == 258)
                //        tempOverLength = member1EndOverhang - supportingSection.webThickness / 2;
                //    else
                //        tempOverLength = member1EndOverhang;
                //}
                else
                    tempOverLength = member1EndOverhang;

                if (reflectMember)
                {
                    componentDictionary[SECTION].SetPropertyValue(tempOverLength, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[SECTION].SetPropertyValue(beginOverLength, "IJUAHgrOccOverLength", "EndOverLength");
                }
                else
                {
                    componentDictionary[SECTION].SetPropertyValue(tempOverLength, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[SECTION].SetPropertyValue(beginOverLength, "IJUAHgrOccOverLength", "BeginOverLength");
                }

                string tempSectPort;
                Plane tempPlane = new Plane();
                planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Direct);
                if (!isPlacedOnSurface)
                {
                    if (reflectMember)
                        tempSectPort = "BeginCap";
                    else
                        tempSectPort = "EndCap";

                    if ((planeAngle * 180 / Math.PI) < 45 || (planeAngle * 180 / Math.PI) >= 135)
                        tempPlane = Plane.NegativeXY;
                    else
                        tempPlane = Plane.NegativeZX;

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePointOnPlaneJoint(SECTION, tempSectPort, STRUCT_CONN, "Connection", tempPlane);
                    else
                        JointHelper.CreatePointOnPlaneJoint(SECTION, tempSectPort, "-1", "Structure", tempPlane);
                }
                //==========================
                // Joints For Brace
                //==========================
                double braceAngle = 0;
                if (isBrace)
                {
                    braceAngle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsBraceAngle","BraceAngle")).PropValue;
    
                    double braceHorOffset;
                    braceHorOffset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsBraceHorOff", "BraceHorOffset")).PropValue;
                    int braceOffsetDefinition;
                    braceOffsetDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsBraceOffDef", "BraceOffsetDef")).PropValue;
   
                    switch (braceOffsetDefinition)
                    {
                        case 1:  //From Edge of the Horizontal Member to Brace End
                            braceHorOffset += (braceDepth / Math.Sin(braceAngle));
                            break;
                        case 2:  //From Center of Pipe to Brace End
                            braceHorOffset = braceHorOffset + offset1 + botPipeDia / 2 + (braceDepth / Math.Sin(braceAngle));
                            break;
                        case 3:  //From Edge of Pipe to Brace End
                            braceHorOffset = braceHorOffset + offset1 + (braceDepth / Math.Sin(braceAngle));
                            break;
                        case 4:  //Edge of Horizontal Member to intersection of centerlines of members
                            braceHorOffset = braceHorOffset + (sectionDepth / 2) / Math.Tan(braceAngle) + (braceDepth / Math.Sin(braceAngle)) / 2;
                            break;
                        default:
                            braceHorOffset += (braceDepth / Math.Sin(braceAngle));
                            break; //From Edge of the Horizontal Member to Brace End
                    }

    
                    double dYoffset, dZoffset, dXOffset;

                    //if (componentDictionary[BRACE].SupportsInterface)
                    //componentDictionary[BRACE].SetPropertyValue(2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                    componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsCutback", "EndCutbackAnchorPoint", 2));
                    componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsCutback", "BeginCutbackAnchorPoint", 2));
                    if (!reflectMember)
                    {
                        
                        if (section.sectionType == "WT")
                        {

                            componentDictionary[BRACE].SetPropertyValue(braceAngle - memberCutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                            componentDictionary[BRACE].SetPropertyValue(Math.PI, "IJOAhsBeginCap", "BeginCapRotZ");
                            componentDictionary[BRACE].SetPropertyValue(Math.PI, "IJOAhsEndCap", "EndCapRotZ");
                            dXOffset = brace.width / 2;
                        }
                        else if (section.sectionType == "L")
                        {

                            componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsSteelCP", "CP1", 2));
                            componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsSteelCP", "CP2", 2));
                            componentDictionary[BRACE].SetPropertyValue(-braceAngle + memberCutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                            dXOffset = section.webThickness;
                        }
                        else
                        {
                            componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsSteelCP", "CP1", 2));
                            componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsSteelCP", "CP2", 2));
                            componentDictionary[BRACE].SetPropertyValue(-braceAngle + memberCutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                            dXOffset = 0;
                        }
                    
                        if (brace.sectionType != "L")
                        {
                            dYoffset = sectionDepth / 2;
                            if (brace.sectionType == "WT")
                            {
                                componentDictionary[BRACE].SetPropertyValue(-(Math.PI / 2 - braceAngle), "IJOAhsCutback", "CutbackBeginAngle");
                                dZoffset = braceHorOffset - brace.depth / Math.Sin(braceAngle);
                            }
                            else
                            {
                                componentDictionary[BRACE].SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                                dZoffset = braceHorOffset;
                            }
                        }   
                        else
                        {
                            componentDictionary[BRACE].SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            dZoffset = braceHorOffset - sectionDepth / Math.Tan(braceAngle);
                            dYoffset = -sectionDepth / 2;
                        }
       
                        JointHelper.CreateAngularRigidJoint(BRACE, "BeginCap", SECTION, "BeginFlex", new Vector(dXOffset, dYoffset, dZoffset), new Vector(-braceAngle, 0, 0));
                        //JointFactory.MakePrismaticJoint(BRACE, "BeginCap", BRACE, "EndCap", 11757)
                    }
                    else
                    {
                        if (brace.sectionType == "WT")
                        {
                            componentDictionary[BRACE].SetPropertyValue(-braceAngle - memberCutbackAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            componentDictionary[BRACE].SetPropertyValue(Math.PI, "IJOAhsBeginCap", "BeginCapRotZ");
                            componentDictionary[BRACE].SetPropertyValue(Math.PI, "IJOAhsEndCap", "EndCapRotZ");
                            dXOffset = brace.width / 2;
                        }
                        else if (section.sectionType == "L")
                        {

                            componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsSteelCP", "CP1", 2));
                            componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsSteelCP", "CP2", 2));
                            componentDictionary[BRACE].SetPropertyValue(braceAngle + memberCutbackAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            dXOffset = section.webThickness;
                        }
                        else
                        {
                            componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsSteelCP", "CP1", 2));
                            componentDictionary[BRACE].SetPropertyValue(new PropertyValueCodelist("IJOAhsSteelCP", "CP2", 2));
                            componentDictionary[BRACE].SetPropertyValue(braceAngle + memberCutbackAngle, "IJOAhsCutback", "CutbackBeginAngle");
                            dXOffset = 0;
                        }
                    
                        if (brace.sectionType != "L")
                        {
                            dYoffset = sectionDepth / 2;
                            if (brace.sectionType == "WT")
                            {
                                componentDictionary[BRACE].SetPropertyValue(Math.PI / 2 - braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                                dZoffset = -braceHorOffset + brace.depth / Math.Sin(braceAngle);
                            }
                            else
                            {
                                componentDictionary[BRACE].SetPropertyValue(-Math.PI / 2 + braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                                dZoffset = -braceHorOffset;
                            }
                        } 
                        else
                        {
                            componentDictionary[BRACE].SetPropertyValue(-Math.PI / 2 + braceAngle, "IJOAhsCutback", "CutbackEndAngle");
                            dZoffset = -braceHorOffset + sectionDepth / Math.Tan(braceAngle);
                            dYoffset = -sectionDepth / 2;
                        }
                    
        
                        JointHelper.CreateAngularRigidJoint(BRACE, "EndCap", SECTION, "EndFlex", new Vector(dXOffset, dYoffset, dZoffset), new Vector(braceAngle, 0, 0));
                    }
                
    
                    if (!isPlacedOnSurface)
                    {
                        if (reflectMember)
                            JointHelper.CreatePrismaticJoint(BRACE, "EndCap", BRACE, "BeginCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        else
                            JointHelper.CreatePrismaticJoint(BRACE, "BeginCap", BRACE, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            if (reflectMember)      //For By Structure
                            {
                                if (planeAngle * 180 / Math.PI <= 45 || planeAngle * 180 / Math.PI >= 135)
                                    JointHelper.CreatePointOnPlaneJoint(BRACE, "BeginCap", STRUCT_CONN, "Connection", Plane.NegativeXY);
                                else
                                    JointHelper.CreatePointOnPlaneJoint(BRACE, "BeginCap", STRUCT_CONN, "Connection", Plane.NegativeZX);
                            }
                            else
                            {
                                if (planeAngle * 180 / Math.PI <= 45 || planeAngle * 180 / Math.PI >= 135)
                                    JointHelper.CreatePointOnPlaneJoint(BRACE, "EndCap", STRUCT_CONN, "Connection", Plane.NegativeXY);
                                else
                                    JointHelper.CreatePointOnPlaneJoint(BRACE, "EndCap", STRUCT_CONN, "Connection", Plane.NegativeZX);
                            }

                        else            //For By point
                        {
                            if (reflectMember)
                            {
                                if (planeAngle * 180 / Math.PI <= 45 || planeAngle * 180 / Math.PI >= 135)
                                    JointHelper.CreatePointOnPlaneJoint(BRACE, "BeginCap", "-1", "Structure", Plane.NegativeXY);
                                else
                                    JointHelper.CreatePointOnPlaneJoint(BRACE, "BeginCap", "-1", "Structure", Plane.NegativeZX);
                            }
                            else
                            {
                                if (planeAngle * 180 / Math.PI <= 45 || planeAngle * 180 / Math.PI >= 135)
                                    JointHelper.CreatePointOnPlaneJoint(BRACE, "EndCap", "-1", "Structure", Plane.NegativeXY);
                                else
                                    JointHelper.CreatePointOnPlaneJoint(BRACE, "EndCap", "-1", "Structure", Plane.NegativeZX);
                            }
                        }                             
                    }  
                    else
                        componentDictionary[BRACE].SetPropertyValue((memberLength - braceHorOffset) / Math.Cos(braceAngle), "IJUAHgrOccLength", "Length");
                
    
                    //Set the overlength for Brace
                    double bracePlateTh;
                    if (isBracePlate)
                        bracePlateTh = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BRACEPLATE],"IJUAhsThickness1", "Thickness1")).PropValue;
                    else
                        bracePlateTh = 0;
                

                    //Set the OverLengths
                    if (structureConnection == 1)    //Normal
                    {
                        if (reflectMember)
                            componentDictionary[BRACE].SetPropertyValue((member1EndOverhang - bracePlateTh) / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "BeginOverLength");
                        else
                            componentDictionary[BRACE].SetPropertyValue((member1EndOverhang - bracePlateTh) / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "EndOverLength");
                    }
                    else        //For Lapped
                    {
                        if (supportingFace == 257 )
                        {
                            if (reflectMember)
                                componentDictionary[BRACE].SetPropertyValue((member1EndOverhang) / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "BeginOverLength"); 
                            else
                                componentDictionary[BRACE].SetPropertyValue((member1EndOverhang) / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "EndOverLength"); 
                        }
                        else if (supportingFace == 258)
                        {
                            if (reflectMember)
                                componentDictionary[BRACE].SetPropertyValue((member1EndOverhang) / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "BeginOverLength"); 
                            else
                                componentDictionary[BRACE].SetPropertyValue((member1EndOverhang) / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "EndOverLength");
                        }
                        else
                        {
                            if (reflectMember)
                                componentDictionary[BRACE].SetPropertyValue(member1EndOverhang / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "BeginOverLength"); 
                            else
                                componentDictionary[BRACE].SetPropertyValue(member1EndOverhang / Math.Cos(braceAngle), "IJUAHgrOccOverLength", "EndOverLength");  
                        }
                    }

                    if (brace.sectionType == "L")
                    {
                        if (reflectMember)
                            componentDictionary[BRACE].SetPropertyValue(-section.flangeThickness / Math.Sin(braceAngle), "IJUAHgrOccOverLength", "EndOverLength");
                        else
                            componentDictionary[BRACE].SetPropertyValue(-section.flangeThickness / Math.Sin(braceAngle), "IJUAHgrOccOverLength", "BeginOverLength");
                    }
                }   
                
                // ==========================
                // Joints For Remaining Plates
                // ==========================
                // Joints for End Plates
                double capHorOffset = 0.0, capVertOffset = 0.0, horOffset = 0.0, vertOffset = 0.0;
                if (isCapPlate)
                {
                    capHorOffset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsCapHorOffset", "CapHorOffset")).PropValue;
                    capVertOffset= (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsCapVerOffset", "CapVerOffset")).PropValue;
                    double capPlateWidth1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE],"IJUAhsWidth1", "Width1")).PropValue;
                    double capPlateLength1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE],"IJUAhsLength1", "Length1")).PropValue;
                    double capPlateAngle1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1Ang", "CapPlate1Angle")).PropValue;

                    if (HgrCompareDoubleService.cmpdbl(capHorOffset , 0) == false || HgrCompareDoubleService.cmpdbl(capVertOffset ,0) == false)
                    {
                        horOffset = capPlateWidth1 / 2 - sectionWidth / 2 - capHorOffset;
                        vertOffset = capPlateLength1 / 2 - sectionDepth / 2 - capVertOffset;
                    }

                    if (reflectMember)
                        JointHelper.CreateAngularRigidJoint(CAPPLATE, "Port1", SECTION, "EndFace", new Vector(horOffset, vertOffset, 0), new Vector(0, 0, capPlateAngle1));
                    else
                        JointHelper.CreateAngularRigidJoint(CAPPLATE, "Port2", SECTION, "BeginFace", new Vector(horOffset, vertOffset, 0), new Vector(0, 0, capPlateAngle1));
                }

                // Joints for the Base Plate
                if (isBasePlate)
                {
                    double basePlateWidth1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsWidth1", "Width1")).PropValue;
                    double basePlate1Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate1Ang", "BasePlate1Angle")).PropValue;
                    double basePlate1Offset;
                    try
                    {
                        basePlate1Offset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsBasePlate1Off", "BasePlate1Offset")).PropValue;
                    }
                    catch { basePlate1Offset = 0; }

                    if (structureConnection == 1) //Normal
                    {
                        if (reflectMember)
                            JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port2", SECTION, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));
                        else
                            JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port1", SECTION, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));
                    }
                    else if (structureConnection == 2) //Lapped
                    {
                        if (structPortWrong)
                        {
                            if (reflectMember)
                                JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port1", SECTION, "BeginFace", new Vector(sectionWidth / 2, 0, basePlateWidth1 / 2 - basePlate1Offset), new Vector(0, Math.PI / 2, basePlate1Angle));
                            else
                                JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port1", SECTION, "EndFace", new Vector(sectionWidth / 2, 0, -basePlateWidth1 / 2 + basePlate1Offset), new Vector(0, Math.PI / 2, basePlate1Angle));
                        }
                        else
                        {
                            if (reflectMember)
                                JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port2", SECTION, "BeginFace", new Vector(-sectionWidth / 2, 0, basePlateWidth1 / 2 - basePlate1Offset), new Vector(0, Math.PI / 2, basePlate1Angle));
                            else
                                JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port2", SECTION, "EndFace", new Vector(-sectionWidth / 2, 0, -basePlateWidth1 / 2 + basePlate1Offset), new Vector(0, Math.PI / 2, basePlate1Angle));
                        }
                    }
                    else
                    {
                        if (structPortWrong)
                        {
                            if (reflectMember)
                                JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port2", SECTION, "BeginFace", new Vector(-sectionWidth / 2, 0, basePlateWidth1 / 2 - basePlate1Offset), new Vector(0, Math.PI / 2, basePlate1Angle));
                            else
                                JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port1", SECTION, "EndFace", new Vector(-sectionWidth / 2, 0, -basePlateWidth1 / 2 + basePlate1Offset), new Vector(0, Math.PI / 2, basePlate1Angle));
                        }
                        else
                        {
                            if (reflectMember)
                                JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port1", SECTION, "BeginFace", new Vector(sectionWidth / 2, 0, basePlateWidth1 / 2 - basePlate1Offset), new Vector(0, Math.PI / 2, basePlate1Angle));
                            else
                                JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port2", SECTION, "EndFace", new Vector(sectionWidth / 2, 0, -basePlateWidth1 / 2 + basePlate1Offset), new Vector(0, Math.PI / 2, basePlate1Angle));
                        }
                    }
                }

                //Joints for Brace Plates
                if (isBracePlate)
                {
                    double basePlate2Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsBasePlate2Ang", "BasePlate2Angle")).PropValue;
                    double bracePlateWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BRACEPLATE], "IJUAhsWidth1", "Width1")).PropValue;
                    double basePlate2Offset;
                    try
                    {
                        basePlate2Offset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsBasePlate2Off", "BasePlate2Offset")).PropValue;
                    }
                    catch { basePlate2Offset = 0; }

                    if (structureConnection == 1) //Normal
                    {
                        if (reflectMember)
                            JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port2", BRACE, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate2Angle));
                        else
                            JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port1", BRACE, "EndFace", new Vector(0, 0 ,0), new Vector(0, 0, basePlate2Angle));
                    }
                    else if (structureConnection == 2) //Lapped
                    {
                        if (structPortWrong)
                        {
                            if (reflectMember)
                                JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port1", BRACE, "BeginFace", new Vector(brace.width / 2, 0, bracePlateWidth / 2 - basePlate2Offset), new Vector(0, Math.PI / 2, basePlate2Angle));
                            else
                                JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port1", BRACE, "EndFace", new Vector(brace.width / 2, 0, -bracePlateWidth / 2 + basePlate2Offset), new Vector(0, Math.PI / 2, basePlate2Angle));
                        }
                        else
                        {
                            if (reflectMember)
                                JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port2", BRACE, "BeginFace", new Vector(-brace.width / 2, 0, bracePlateWidth / 2 - basePlate2Offset), new Vector(0, Math.PI / 2, basePlate2Angle));
                            else
                                JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port2", BRACE, "EndFace", new Vector(-brace.width / 2, 0, -bracePlateWidth / 2 + basePlate2Offset), new Vector(0, Math.PI / 2, basePlate2Angle));
                        }
                    }
                    else
                    {
                        if (structPortWrong)
                        {
                            if (reflectMember)
                                JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port2", BRACE, "BeginFace", new Vector(-brace.width / 2, 0, bracePlateWidth / 2 - basePlate2Offset), new Vector(0, Math.PI / 2, basePlate2Angle));
                            else
                                JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port1", BRACE, "EndFace", new Vector(-brace.width / 2, 0, -bracePlateWidth / 2 + basePlate2Offset), new Vector(0, Math.PI / 2, basePlate2Angle));
                        }
                        else
                        {
                            if (reflectMember)
                                JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port1", BRACE, "BeginFace", new Vector(brace.width / 2, 0, bracePlateWidth / 2 - basePlate2Offset), new Vector(0, Math.PI / 2, basePlate2Angle));
                            else
                                JointHelper.CreateAngularRigidJoint(BRACEPLATE, "Port2", BRACE, "EndFace", new Vector(brace.width / 2, 0, -bracePlateWidth / 2 + basePlate2Offset), new Vector(0, Math.PI / 2, basePlate2Angle));
                        }
                    }
                }
                // ==========================
                // Joints For the Pipe Attachments
                // ==========================
                string partProgId = string.Empty;
                double lugOffset = 0.0;
                double pipeAttOffset = 0.0;
                double[] pipeDiameters = new double[SupportHelper.SupportedObjects.Count];

                if (includePipeAtt1)
                {
                    Part pipeAtt1Part = (Part)componentDictionary[PIPEATT1[0]].GetRelationship("madeFrom", "part").TargetObjects[0];
                    partProgId = (string)((PropertyValueString)(pipeAtt1Part).GetPropertyValue("IJDSymbolDefHelper", "ProgId")).PropValue;

                    try
                    {
                        pipeAttOffset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrPipeAtt1Offset", "PipeAtt1Offset")).PropValue;
                    }
                    catch { pipeAttOffset = 0; }

                    for (int i = 0; i < SupportHelper.SupportedObjects.Count; i++)
                    {
                        pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(i+1);
                        pipeDiameters[i] = pipeInfo.OutsideDiameter;
                        insulTh = pipeInfo.InsulationThickness;
                    }
                    if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Guide")
                    {
                        for (int i = 0; i < SupportHelper.SupportedObjects.Count; i++)
                        {
                            componentDictionary[PIPEATT1[i]].SetPropertyValue(shoeHeight + pipeDiameters[i] / 2, "IJUAhsShoeHeight", "ShoeHeight");
                            componentDictionary[PIPEATT1[i]].SetPropertyValue(0, "IJOAhsPipeOD", "PipeOD");
                        }
                    }
                    else if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.UBolt")
                    {
                        double steelThickness;
                        if (section.HSS_DesignWallThickness > 0)
                            steelThickness = section.HSS_DesignWallThickness;
                        else
                            steelThickness = section.flangeThickness;
                        for (int i = 0; i < SupportHelper.SupportedObjects.Count; i++)
                        {
                            if (componentDictionary[PIPEATT1[i]].SupportsInterface("IJOAhsSteelThickness"))
                                componentDictionary[PIPEATT1[i]].SetPropertyValue(steelThickness, "IJOAhsSteelThickness", "SteelThickness");
                        }
                    }
                    else if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Shoe")
                    {
                        double endPlateThickness = 0.0, slidePlateThickness = 0.0, shoePartWidth = 0.0;
                        string guideShape, strapShape, slidePlateShape;
                        BusinessObject guideShapePart, strapShapePart, slidePlateShapePart;
                        Catalog catalog = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog;
                        
                        for (int i = 0; i < SupportHelper.SupportedObjects.Count; i++)
                        {
                            guideShape = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT1[i]],"IJUAhsGuideShape", "GuideShape")).PropValue;
                            strapShape = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT1[i]],"IJUAhsStrapShape", "StrapShape")).PropValue;
                            slidePlateShape = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT1[i]],"IJUAhPlateShape", "SlidePlateShape")).PropValue;

                            shoePartWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT1[i]],"IJUAhsShoeWidth", "ShoeWidth")).PropValue;
                            if (!String.IsNullOrEmpty(guideShape))
                            {
                                guideShapePart = catalog.GetNamedObject(guideShape);
                                shoePartWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(guideShapePart,"IJUAhsWidth1", "Width1")).PropValue;
                                endPlateThickness = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(guideShapePart,"IJUAhsThickness1", "Thickness1")).PropValue;
                            }
                            if (!String.IsNullOrEmpty(strapShape))
                            {
                                strapShapePart = catalog.GetNamedObject(strapShape);
                                shoePartWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(strapShapePart,"IJUAhsStrap", "StrapWidthInside")).PropValue;
                            }
                            if (!String.IsNullOrEmpty(slidePlateShape))
                            {
                                slidePlateShapePart = catalog.GetNamedObject(slidePlateShape);
                                slidePlateThickness = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(slidePlateShapePart,"IJUAhsThickness1", "Thickness1")).PropValue;
                            }

                            lugOffset = endPlateThickness + slidePlateThickness;

                            componentDictionary[PIPEATT1[i]].SetPropertyValue(shoeHeight, "IJOAhsShoeHeight", "ShoeHeight");
                            weld = pipeConnection1Welds[i];
                            JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", PIPEATT1[i], "Steel", new Vector(0, -pipeDiameters[i] / 2, 0), new Vector(0, 0, 0));

                            if (pipeConnection3Welds.Count > 0)
                            {
                                weld = pipeConnection3Welds[i];
                                JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", "-1", "Route", new Vector(0, -pipeDiameters[i] / 2, 0), new Vector(0, 0, 0));
                            }
                        }
                    }
                    else if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.BlockClamp")
                    {
                        double blockClampWidth;
                        for (int i = 0; i < SupportHelper.SupportedObjects.Count; i++)
                        {
                            blockClampWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT1[i]],"IJUAhsWidth3", "Width3")).PropValue;
                            weld = pipeConnection1Welds[i];
                            JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", PIPEATT1[i], "Structure", new Vector(0, blockClampWidth / 2, 0), new Vector(0, 0, 0));
                        
                            if (pipeConnection2Welds.Count > 0)
                            {
                                weld = pipeConnection2Welds[i];
                                JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", PIPEATT1[i], "Structure", new Vector(0, 0, 0), new Vector(0, 0, 0));
                            }
                        }
                    }
                    else if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strap")
                    {
                        double strapWidthInside;
                        for (int i = 0; i < SupportHelper.SupportedObjects.Count; i++)
                        {
                            strapWidthInside = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT1[i]],"IJUAhsStrap", "StrapWidthInside")).PropValue;
                            weld = pipeConnection1Welds[i];
                            JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", PIPEATT1[i], "Steel", new Vector(strapWidthInside / 2, 0, 0), new Vector(0, 0, 0));
                        }
                    }

                    string routePort;
                    for (int i = 0; i < SupportHelper.SupportedObjects.Count; i++)
                    {
                        if (i == 0)
                            routePort = "Route";
                        else
                            routePort = "Route_" + (i + 1);

                        JointHelper.CreateTranslationalJoint(SECTION, "BeginFlex", PIPEATT1[i], "Route", Plane.YZ, Plane.NegativeYZ, Axis.Z, Axis.Y, -pipeAttOffset);
                        JointHelper.CreatePointOnAxisJoint(PIPEATT1[i], "Route", "-1", routePort, Axis.NegativeX);
                    }
                }

                if (includePipeAtt1 && includePipeAtt2)
                {
                    Part pipeAtt2Part = (Part)componentDictionary[PIPEATT2[0]].GetRelationship("madeFrom", "part").TargetObjects[0];
                    string part2ProgId = (string)((PropertyValueString)(pipeAtt2Part).GetPropertyValue("IJDSymbolDefHelper", "ProgId")).PropValue;
                    double pipeAtt2Offset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt2Offset", "PipeAtt2Offset")).PropValue;

                    if (part2ProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.TwoStandardPort")
                    {
                        double lugWidth, lugThick;

                        if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strap")
                        {
                            for (int i = 0; i < pipeAtt2Qty; i++)
                            {
                                int pipeAtt1Idx = i / 2;

                                Vector translation = new Vector();
                                lugWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT2[i]],"IJUAhsWidth1", "Width1")).PropValue;
                                lugThick = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT2[i]],"IJUAhsThickness1", "Thickness1")).PropValue;
                                translation.X = -lugThick / 2;
                                translation.Y = sectionWidth / 2 + lugWidth / 2 + pipeAtt2Offset;
                                translation.Z = -lugWidth / 2;
                                if (i % 2 != 0)
                                    translation.Y = -translation.Y;
                                JointHelper.CreateAngularRigidJoint(PIPEATT2[i], "Port1", PIPEATT1[pipeAtt1Idx], "Steel", translation, new Vector(Math.PI / 2, Math.PI / 2, 0));

                                if (i % 2 != 0)
                                {
                                    weld = pipeConnection2Welds[pipeAtt1Idx];
                                    JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", PIPEATT2[i], "Port1", new Vector(0, lugWidth / 2, 0), new Vector(0, 0, 0));
                                }
                            }
                        }
                        else if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Shoe")
                        {
                            for (int i = 0; i < pipeAtt2Qty; i++)
                            {
                                int pipeAtt1Idx = i / 2;

                                Vector translation = new Vector();
                                lugWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT2[i]],"IJUAhsWidth1", "Width1")).PropValue;
                                lugThick = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[PIPEATT2[i]],"IJUAhsThickness1", "Thickness1")).PropValue;
                                translation.X = sectionWidth / 2 + lugWidth / 2 + pipeAtt2Offset;
                                translation.Y = lugThick / 2;
                                translation.Z = lugOffset - lugWidth / 2;
                                if (i % 2 != 0)
                                    translation.X = -translation.X;
                                JointHelper.CreateAngularRigidJoint(PIPEATT2[i], "Port1", PIPEATT1[pipeAtt1Idx], "Steel", translation, new Vector(Math.PI / 2, 0, 0));

                                if (i % 2 != 0)
                                {
                                    weld = pipeConnection2Welds[pipeAtt1Idx];
                                    JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", PIPEATT2[i], "Port1", new Vector(0, lugWidth / 2, 0), new Vector(0, 0, 0));
                                }
                            }
                        }
                    }
                }

                

                // ==========================
                // Joints For the Anchor Bolts
                // ==========================
                double boltOffset1, boltOffset2;
                if (includeBolts == true)
                {
                    double basePlateWidth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsWidth1", "Width1")).PropValue;
                    double basePlateDepth = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsLength1", "Length1")).PropValue;
                    try
                    {
                        boltOffset1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1Offset1", "Bolt1Offset1")).PropValue;
                    }
                    catch
                    {
                        boltOffset1 = 0;
                    }
                    try
                    {
                        boltOffset2 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1Offset2", "Bolt1Offset2")).PropValue;
                    }
                    catch
                    {
                        boltOffset2 = 0;
                    }

                    double dBoltOffsetX = 0, dBoltOffsetY = 0;
                    if (boltOffset1 <= 0)
                        dBoltOffsetX = basePlateWidth / 2;
                    else
                        dBoltOffsetX = basePlateWidth / 2 - boltOffset1;

                    if (boltOffset2 <= 0)
                        dBoltOffsetY = basePlateDepth / 2;
                    else
                        dBoltOffsetY = basePlateDepth / 2 - boltOffset2;

                    switch (boltQty)
                    {
                        case 0:
                            // No Bolts Included
                            break;
                        case 2:
                            {
                                JointHelper.CreateRigidJoint(BASEPLATE, "Port1", ANCHORS[0], "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, dBoltOffsetY, dBoltOffsetX);
                                JointHelper.CreateRigidJoint(BASEPLATE, "Port1", ANCHORS[1], "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -dBoltOffsetY, -dBoltOffsetX);
                                break;
                            }
                        case 4:
                            {
                                JointHelper.CreateRigidJoint(BASEPLATE, "Port1", ANCHORS[0], "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, dBoltOffsetY, dBoltOffsetX);
                                JointHelper.CreateRigidJoint(BASEPLATE, "Port1", ANCHORS[1], "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -dBoltOffsetY, -dBoltOffsetX);
                                JointHelper.CreateRigidJoint(BASEPLATE, "Port1", ANCHORS[2], "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -dBoltOffsetY, dBoltOffsetX);
                                JointHelper.CreateRigidJoint(BASEPLATE, "Port1", ANCHORS[3], "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, dBoltOffsetY, -dBoltOffsetX);
                                break;
                            }
                    }
                }

                // ==========================
                // Joints For the Weld Objects
                // ==========================
                //FrameAssemblyServices.WeldData weld = new FrameAssemblyServices.WeldData();
                string plateWeldPort, bWeldPort, cWeldPort;
                if (reflectMember)
                {
                    plateWeldPort = "Port1";
                    bWeldPort = "BeginFace";
                    cWeldPort = "EndFace";
                }
                else
                {
                     plateWeldPort =  "Port2";
                     bWeldPort = "EndFace";
                     cWeldPort = "BeginFace";
                }
                double length1 = 0;
                double width1 = 0;
                if (isBasePlate)
                {
                    length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsLength1", "Length1")).PropValue;
                    width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsWidth1", "Width1")).PropValue;
                }
                for (int weldCount = 0; weldCount < weldCollection.Count; weldCount++)
                {
                    weld = weldCollection[weldCount];
                    switch (weld.connection)
                    {
                        case "A":
                            {
                                // Weld for the Base Plate
                                if (isBasePlate)
                                {
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BASEPLATE, plateWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue - length1 / 2, weld.offsetZValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BASEPLATE, plateWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue - width1 / 2, weld.offsetYValue, weld.offsetZValue);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BASEPLATE, plateWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue + width1 / 2, weld.offsetYValue, weld.offsetZValue);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BASEPLATE, plateWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue + length1 / 2, weld.offsetZValue);
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
                                            JointHelper.CreateRigidJoint(weld.partKey, "Other", SECTION, bWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue - sectionDepth / 2, weld.offsetZValue);
                                            break;
                                        }
                                    case 4:
                                        {
                                            JointHelper.CreateRigidJoint(weld.partKey, "Other", SECTION, bWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue - sectionWidth / 2, weld.offsetYValue, weld.offsetZValue);
                                            break;
                                        }
                                    case 6:
                                        {
                                            JointHelper.CreateRigidJoint(weld.partKey, "Other", SECTION, bWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue + sectionWidth / 2, weld.offsetYValue, weld.offsetZValue);
                                            break;
                                        }
                                    case 8:
                                        {
                                            JointHelper.CreateRigidJoint(weld.partKey, "Other", SECTION, bWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue + sectionDepth / 2, weld.offsetZValue);
                                            break;
                                        }
                                }
                                break;
                            }
                        case "C":
                            {
                                switch (weld.location)
                                {
                                    case 2:
                                        {
                                            JointHelper.CreateRigidJoint(weld.partKey, "Other", SECTION, cWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue - sectionDepth / 2, weld.offsetZValue);
                                            break;
                                        }
                                    case 4:
                                        {
                                            JointHelper.CreateRigidJoint(weld.partKey, "Other", SECTION, cWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue - sectionWidth / 2, weld.offsetYValue, weld.offsetZValue);
                                            break;
                                        }
                                    case 6:
                                        {
                                            JointHelper.CreateRigidJoint(weld.partKey, "Other", SECTION, cWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue + sectionWidth / 2, weld.offsetYValue, weld.offsetZValue);
                                            break;
                                        }
                                    case 8:
                                        {
                                            JointHelper.CreateRigidJoint(weld.partKey, "Other", SECTION, cWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue + sectionDepth / 2, weld.offsetZValue);
                                            break;
                                        }
                                }
                                break;
                            }
                        case "D":
                            {
                                if (isBrace)
                                {
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACE, "BeginFace", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue - (brace.depth / Math.Sin(braceAngle)) / 2, weld.offsetZValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACE, "BeginFace", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue - sectionWidth / 2, weld.offsetYValue, weld.offsetZValue);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACE, "BeginFace", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue + sectionWidth / 2, weld.offsetYValue, weld.offsetZValue);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACE, "BeginFace", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue + (brace.depth / Math.Sin(braceAngle)) / 2, weld.offsetZValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                        case "E":
                            {
                                if (isBrace)
                                {
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACE, "EndFace", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue - (brace.depth / Math.Cos(braceAngle)) / 2, weld.offsetZValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACE, "EndFace", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue - sectionWidth / 2, weld.offsetYValue, weld.offsetZValue);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACE, "EndFace", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue + sectionWidth / 2, weld.offsetYValue, weld.offsetZValue);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACE, "EndFace", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue + (brace.depth / Math.Cos(braceAngle)) / 2, weld.offsetZValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                        case "F":
                            {
                                if (isBrace && isBracePlate)
                                {
                                    double bracePlateLength1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BRACEPLATE],"IJUAhsLength1","Length1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACEPLATE, plateWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue - bracePlateLength1 / 2, weld.offsetZValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACEPLATE, plateWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue - bracePlateLength1 / 2, weld.offsetYValue, weld.offsetZValue);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACEPLATE, plateWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue + bracePlateLength1 / 2, weld.offsetYValue, weld.offsetZValue);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(weld.partKey, "Other", BRACEPLATE, plateWeldPort, Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetXValue, weld.offsetYValue + bracePlateLength1 / 2, weld.offsetZValue);
                                                break;
                                            }
                                    }
                                }
                                break;
                            }
                    }
                }

                // Create the Dimensions and Labels
                Boolean excludeNotes = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                ControlPoint controlPoint;
                Note note;
                if (!excludeNotes)
                {
                    double offsetX = 0, offsetY = 0;

                    if (sectionData.Orient == FrameAssemblyServices.SteelOrientationAngle.SteelOrientationAngle_90)
                        offsetX = sectionDepth / 2;
                    else if (sectionData.Orient == FrameAssemblyServices.SteelOrientationAngle.SteelOrientationAngle_180)
                        offsetY = sectionDepth / 2;
                    else if (sectionData.Orient == FrameAssemblyServices.SteelOrientationAngle.SteelOrientationAngle_90)
                        offsetX = -sectionDepth / 2;
                    else
                        offsetY = -sectionDepth / 2;

                    if (reflectMember)
                        note = CreateNote("Elevation", SECTION, "EndFlex", new Position(offsetX, offsetY, 0.0), "Elevation", false, 2, 51, out controlPoint);
                    else
                        note = CreateNote("Elevation", SECTION, "BeginFlex", new Position(offsetX, offsetY, 0.0), "Elevation", false, 2, 51, out controlPoint);
                }
                else
                    DeleteNoteIfExists("Elevation");

                if (reflectMember)
                {
                    note = CreateNote("Dim1", SECTION, "EndFlex", new Position(0.0, 0.0, -capPlateTh), " ", true, 2, 1, out controlPoint);
                    note = CreateNote("Dim2", SECTION, "BeginFlex", new Position(0.0, 0.0, 0.0), " ", true, 2, 1, out controlPoint);
                }
                else
                {
                    note = CreateNote("Dim1", SECTION, "BeginFlex", new Position(0.0, 0.0, -capPlateTh), " ", true, 2, 1, out controlPoint);
                    note = CreateNote("Dim2", SECTION, "EndFlex", new Position(0.0, 0.0, 0.0), " ", true, 2, 1, out controlPoint);
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
                    for (int i = 0; i < SupportHelper.SupportedObjects.Count; i++)
                    {
                        routeConnections.Add(new ConnectionInfo(SECTION, i+1)); // partindex, routeindex
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
                    structConnections.Add(new ConnectionInfo(SECTION, 1)); // partindex, routeindex

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
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject catalogPart = SupportOrComponent.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                // Get Section Information of each steel part
                IPart part = (IPart)componentDictionary[SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                section = FrameAssemblyServices.GetSectionDataFromPart(part.PartNumber);
                double capPlateTh = 0;
                if (isCapPlate)
                    capPlateTh = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE],"IJUAhsThickness1", "Thickness1")).PropValue;

                // ==========================
                // Shoe Height
                // ==========================
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = pipeInfo.OutsideDiameter;
                double insulationThickness = pipeInfo.InsulationThickness;
                double pipeCLOffset = 0;

                Boolean includeInsulation = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(SupportOrComponent, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue;
                if (includeInsulation)
                    pipeCLOffset = pipeDiameter / 2 + insulationThickness;
                else
                    pipeCLOffset = pipeDiameter / 2;

                int shoeHeightDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart,"IJUAhsFrameShoeHeightDef", "ShoeHeightDefinition")).PropValue;
                // Get the Shoe Height From the Input Attributes
                double shoeHeight = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameShoeHeight", "ShoeHeightValue")).PropValue;
                switch (shoeHeightDefinition)
                {
                    case 1:
                        // Edge of Bounding Box
                        break;
                    case 2:
                        // Centerline of Primary Pipe
                        shoeHeight = shoeHeight - RefPortHelper.DistanceBetweenPorts("BBFrame_High", "Route", PortAxisType.Y);
                        break;
                }

                // ==========================
                // Length 1
                // ==========================
                double L1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJUAHgrOccLength", "Length")).PropValue;
                int length1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameLength1Def", "Length1Definition")).PropValue;
                double offset1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOffset1", "Offset1Value")).PropValue;
                switch (length1Definition)
                {
                    case 1:
                        break;
                    case 2:
                        L1 = L1 + capPlateTh;
                        break;
                    case 3:
                        L1 = L1 + capPlateTh - pipeCLOffset - offset1;
                        break;
                    default:
                        break;
                }
                SupportOrComponent.SetPropertyValue(L1, "IJUAhsFrameLength1", "Length1Value");

                isBrace = componentDictionary.ContainsKey(BRACE);

                double L2 = 0;
                if (isBrace)
                {
                    L2 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BRACE],"IJUAhsVarLength", "VarLength")).PropValue;
                    int length2Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameLength2Def", "Length2Definition")).PropValue;
                    switch (length2Definition)
                    {
                        case 1: //Cut Length
                            break;
                        case 2: //Length
                            L2 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BRACE],"IJUAHgrOccLength", "Length")).PropValue;
                            break;
                        default:
                            break;
                    }
                    SupportOrComponent.SetPropertyValue(L2, "IJUAhsFrameLength2", "Length2Value");
                }

                

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

                string length1 = FrameAssemblyServices.FormatValueWithUnits(SupportOrComponent, "IJUAhsFrameLength1", "Length1Value", L1, primaryUnits, secondaryUnits, precisionType, precision);
                string length2 = FrameAssemblyServices.FormatValueWithUnits(SupportOrComponent, "IJUAhsFrameLength1", "Length1Value", L2, primaryUnits, secondaryUnits, precisionType, precision);
                string supportNumber = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsSupportNumber", "SupportNumber")).PropValue;
                PropertyValueCodelist steelStandardList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart,"IJUAhsSteelStandard", "SteelStandard");

                if (isBrace)
                    bomDesciption = supportNumber + "," + steelStandardList.PropertyInfo.CodeListInfo.GetCodelistItem(steelStandardList.PropValue).ShortDisplayName + "," + section.sectionName + ", L1=" + length1 + ", L2=" + length2;
                else
                    bomDesciption = supportNumber + "," + steelStandardList.PropertyInfo.CodeListInfo.GetCodelistItem(steelStandardList.PropValue).ShortDisplayName + "," + section.sectionName + ", L1=" + length1;

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

        /// <summary>
        ///  Returns the distance between the structure and bottom pipe, and gives the index of the bottom pipe.
        /// </summary>
        /// <param name="botPipeIdx">Gets set to the index of bottom pipe.</param>
        private double GetStructBotPipeDist(ref int botPipeIdx)
        {
            int numRoutes = SupportHelper.SupportedObjects.Count;
            int maxI = 0, maxJ = 0;
            double maxDist = 0.0;
            double pipeToStructDist1, pipeToStructDist2;
            string routePort1, routePort2;
            string[] pipePortNames = new string[numRoutes];
            double[,] pipesPortDist = new double[numRoutes, numRoutes];

            
            for (int i = 0; i < numRoutes; i++)
            {
                if (i == 0)
                    pipePortNames[i] = "Route";
                else
                    pipePortNames[i] = "Route_" + (i + 1);
            }

            for (int i = 0; i < numRoutes; i++)
            {
                for (int j = i + 1; j < numRoutes; j++)
                {
                    if (i == j)
                        pipesPortDist[i, j] = 0.0;
                    else
                    {
                        pipesPortDist[i, j] = RefPortHelper.DistanceBetweenPorts(pipePortNames[i], pipePortNames[j], PortDistanceType.Direct);
                        if (pipesPortDist[i, j] > maxDist)
                        {
                            maxDist = pipesPortDist[i, j];
                            maxI = i;
                            maxJ = j;
                        }
                    }
                }
            }

            if (maxI == 0)
                routePort1 = "Route";
            else
                routePort1 = "Route_" + (maxI + 1);

            if (maxJ == 0)
                routePort2 = "Route";
            else
                routePort2 = "Route_" + (maxJ + 1);

            pipeToStructDist1 = Math.Abs(RefPortHelper.DistanceBetweenPorts("Structure", routePort1, PortAxisType.Z));
            pipeToStructDist2 = Math.Abs(RefPortHelper.DistanceBetweenPorts("Structure", routePort2, PortAxisType.Z));

            if (pipeToStructDist1 > pipeToStructDist2)
            {
                botPipeIdx = maxI + 1;
                return pipeToStructDist1;
            }
            else
            {
                botPipeIdx = maxJ + 1;
                return pipeToStructDist2;
            }
        }

        private void GetPrimaryPipeIdxes(ref int primParallelPipeIdx, ref int primPerpentclrPipeIdx)
        {
            Collection<int> perpendicularPipeCollection = new Collection<int>();
            Collection<int> parallelPipeCollection = new Collection<int>();
            int numRoutes = SupportHelper.SupportedObjects.Count;
            int numParallelPipes, numPerpendicularPipes;
            bool[] routeStructParallel = new bool[numRoutes];

            for (int i = 0; i < numRoutes; i++)
            {
                routeStructParallel[i] = IsRouteStructParallel(i+1);
                if (routeStructParallel[i])
                    parallelPipeCollection.Add(i+1);
                else
                    perpendicularPipeCollection.Add(i+1);
            }

            numParallelPipes = parallelPipeCollection.Count;
            numPerpendicularPipes = perpendicularPipeCollection.Count;

            if (numParallelPipes > 0)
                primParallelPipeIdx = parallelPipeCollection[0];
            else
                primParallelPipeIdx = 0;

            if (numPerpendicularPipes > 0)
                primPerpentclrPipeIdx = perpendicularPipeCollection[0];
            else
                primPerpentclrPipeIdx = 0;

        }

        private bool IsRouteStructParallel(int routeIdx)
        {
            double routeStructAngle;
            string routePort;

            if (routeIdx == 1)
                routePort = "Route";
            else
                routePort = "Route_" + routeIdx;

            routeStructAngle = RefPortHelper.AngleBetweenPorts(routePort, PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

            if ((routeStructAngle > -0.0001 && routeStructAngle < 0.0001) || (routeStructAngle > Math.PI - 0.0001 && routeStructAngle < Math.PI + 0.001))
                return true;
            else if (routeStructAngle < Math.PI / 4 || routeStructAngle > 3 * Math.PI / 4)
                return true;
            else
                return false;

        }

        private void make90DegAngle(ref double angle)
        {
            if (angle <= Math.PI / 4 || angle > 7 * Math.PI / 4)
                angle = 0;
            else if (angle > Math.PI / 4 && angle <= 3 * Math.PI / 4)
                angle = Math.PI / 2;
            else if (angle > 3 * Math.PI / 4 && angle <= 5 * Math.PI / 4)
                angle = Math.PI;
            else
                angle = 3 * Math.PI / 2;
        }

        private void flip90DegAngle(ref double angle) 
        {
            if (angle < 3 * Math.PI / 4)
                angle = angle + Math.PI;
            else
                angle = angle - Math.PI;
        }

    }
}
