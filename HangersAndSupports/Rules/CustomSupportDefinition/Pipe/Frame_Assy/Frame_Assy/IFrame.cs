//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AK1.cs
//   Frame_Assy,Ingr.SP3D.Content.Support.Rules.IFrame
//   Author       :  Rajeswari
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 26-Jul-2013  Rajeswari   CR-CP-224474- Convert HS_S3DFrame to C# .Net 
// 30-Mar-2014     PVK      CR-CP-245789	Modify the exsisting .Net Frame_Assy to behave like URS Frame supports
// 27-Apr-2015     PVK      TR-CP-253033	Elevation CP not shown by default for frame supports.
// 06-May-2015     PVK      CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
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

    public class IFrame : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string SECTION = "SECTION";
        private const string CAPPLATE = "CAPPLATE";
        private const string BASEPLATE = "BASEPLATE";
        private const string GROUT = "GROUT";

        string[] PIPEATT1, PIPEATT2, ANCHORS;
        private Boolean includeBolts, includePipeAtt1, includePipeAtt2;
        private int boltQty, pipeAtt1Qty, pipeAtt2Qty;
        FrameAssemblyServices.HSSteelMember section = new FrameAssemblyServices.HSSteelMember();
        Collection<FrameAssemblyServices.WeldData> weldCollection = new Collection<FrameAssemblyServices.WeldData>();
        bool isGrout, isCapPlate, isBasePlate, isPipeAtt1, isPipeAtt2, isbolt1, isBoltPart;
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
                        string type = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support,"IJUAHgrURSCommon", "Type")).PropValue;
                        if (family !="" && type !=null)
                            FrameAssemblyServices.CheckSupportWithFamilyAndType(this, family, type);
                    }

                    // Add the Steel Section for the Frame Support
                    string member1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1", "Member1Part")).PropValue;
                    string member1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1Rl", "Member1Rule")).PropValue;
                    FrameAssemblyServices.AddPart(this, SECTION, member1Part, member1Rule, parts);

                    // Add the Plates
                    string capPlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1", "CapPlate1Part")).PropValue;
                    string capPlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1Rl", "CapPlate1Rule")).PropValue;
                    isCapPlate = FrameAssemblyServices.AddPart(this, CAPPLATE, capPlate1Part, capPlate1Rule, parts);
                    string basePlate1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate1", "BasePlate1Part")).PropValue;
                    string basePlate1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate1Rl", "BasePlate1Rule")).PropValue;
                    isBasePlate = FrameAssemblyServices.AddPart(this, BASEPLATE, basePlate1Part, basePlate1Rule, parts);


                    includePipeAtt1 = false;
                    string pipeAtt1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt1", "PipeAtt1Part")).PropValue;
                    string pipeAtt1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt1Rl", "PipeAtt1Rule")).PropValue;
                    pipeAtt1Qty = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrPipeAtt1Qty", "PipeAtt1Quantity")).PropValue;

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
                    string pipeAtt2Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt2", "PipeAtt2Part")).PropValue;
                    string pipeAtt2Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt2Rl", "PipeAtt2Rule")).PropValue;
                    pipeAtt2Qty = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsFrPipeAtt2Qty", "PipeAtt2Quantity")).PropValue;

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

                    // Add the Grout (Requires Base Plate)
                    if (isBasePlate)
                    {
                        string groutPart = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrGrout", "GroutPart")).PropValue;
                        string groutRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrGroutRl", "GroutRule")).PropValue;
                        isGrout = FrameAssemblyServices.AddPart(this, GROUT, groutPart, groutRule, parts);
                    }

                    // Add the Anchor Bolts (Requires Base Plate)
                    includeBolts = false;
                    if (isBasePlate)
                    {
                        string boltPart = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1", "Bolt1Part")).PropValue;
                        string boltRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1Rl", "Bolt1Rule")).PropValue;
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

                    // Add the Weld Objects from the Weld Sheet
                    string weldServiceClassName = string.Empty;
                    if (part.SupportsInterface("IJUAhsWeldServClass"))
                        weldServiceClassName = (String)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsWeldServClass", "WeldServiceClassName")).PropValue;
                    if (weldServiceClassName == null)
                        weldServiceClassName = string.Empty;
                    weldCollection = FrameAssemblyServices.AddWeldsFromCatalog(this, parts, "IJUAhsIFrameWelds",((IPart)part).PartNumber, weldServiceClassName);
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
                    
                    if (isBasePlate)
                    {
                        string bolt1Part, bolt1Rule;
                        bolt1Part = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1", "Bolt1Part")).PropValue;
                        bolt1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBolt1Rl", "Bolt1Rule")).PropValue;
                        bolt1Quantity = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(part, "IJUAhsBolt1Qty", "Bolt1Quantity")).PropValue;
                        isBoltPart = FrameAssemblyServices.AddImpliedPart(this, bolt1Part, bolt1Rule, impliedParts, null, bolt1Quantity);
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
                FrameAssemblyServices.HSSteelConfig sectionData = new FrameAssemblyServices.HSSteelConfig();

                Double ndFrom = 0, ndTo = 0;
                Boolean checkPipeSize = false;

                if (support.SupportsInterface("IJUAHgrCheckPipeSize"))
                    checkPipeSize = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(part,"IJUAHgrCheckPipeSize", "CheckPipeSize")).PropValue;
                else
                    checkPipeSize = false;
                //Check the Pipe size within limit and throw warning
                if (includePipeAtt1 == true || includePipeAtt2 == true)
                {
                    if (checkPipeSize == true)
                    {
                        // Get the NDFrom and NDTo values
                        CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                        PartClass iframePartClass = (PartClass)catalogBaseHelper.GetPartClass("hsS3D_IFrame");
                        ReadOnlyCollection<BusinessObject> iframeParts = iframePartClass.Parts;
                        foreach (BusinessObject iframePart in iframeParts)
                        {
                            if ((string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(iframePart,"IJDPart", "PartNumber")).PropValue == support.SupportDefinition.PartNumber)
                            {
                                ndFrom = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(iframePart,"IJHgrSupportDefinition", "NDFrom")).PropValue;
                                ndTo = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(iframePart,"IJHgrSupportDefinition", "NDTo")).PropValue;
                                break;
                            }
                        }
                        PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                        double nominalPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                        // Check the pipe size on which axial stops are placing
                        if (nominalPipeDiameter < ndFrom || nominalPipeDiameter > ndTo)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "" + "can only be placed from" + ndFrom + "in to" + ndTo + "in pipe sizes.", "", "IFrame.cs", 195);
                    }

                }
                // ==========================
                // Get Required Information about the Steel Parts
                // ==========================
                // Get the Steel Cross Section Data
                section = FrameAssemblyServices.GetSectionDataFromPartIndex(this, SECTION);
                // Get the Steel Configuration Data
                sectionData.Orient = (FrameAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue);
                double sectionDepth = 0, sectionWidth = 0;
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
                // ==========================
                // Create the Frame Bounding Box
                // ==========================
                double boundingBoxWidth = 0, boundingBoxDepth = 0;
                Boolean includeInsulation = false;
                string boundingBoxPort = "BBFrame_Low", boundingBoxName = "BBFrame";
                int frameOrientation = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOrientation", "FrameOrientation")).PropValue;
                includeInsulation = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue;
                Boolean mirrorFrame = (bool)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsMirrorFrame", "MirrorFrame")).PropValue;

                if(frameOrientation==3)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "" + "Tangent orientation is not applicable for I Frame.", "", "IFrame.cs", 201);
                else
                    FrameAssemblyServices.CreateFrameBoundingBox(this, boundingBoxName, (FrameAssemblyServices.FrameBBXOrientation)frameOrientation, includeInsulation, mirrorFrame, SupportedHelper.IsSupportedObjectVertical(1, 45));

                boundingBoxWidth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Width;
                boundingBoxDepth = BoundingBoxHelper.GetBoundingBox(boundingBoxName).Height;

                // ==========================
                // Organize the Welds into two Collections
                // ==========================
                Collection<FrameAssemblyServices.WeldData> frameConnection1Welds = new Collection<FrameAssemblyServices.WeldData>(); // Welds at the connection between Leg and Main Member
                Collection<FrameAssemblyServices.WeldData> pipeConnection1Welds = new Collection<FrameAssemblyServices.WeldData>();  // Welds at the connection for Pipe Attachment
                FrameAssemblyServices.WeldData weld = new FrameAssemblyServices.WeldData();
                for (int weldCount = 0; weldCount < weldCollection.Count; weldCount++)
                {
                    weld = weldCollection[weldCount];
                    if ((weld.connection).ToUpper().Equals("D"))
                        pipeConnection1Welds.Add(weld);
                    else
                        frameConnection1Welds.Add(weld);
                }
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
                // ==========================
                // Offset 1 (The offset to the Leg)
                // ==========================
                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter;
                double insulationThickness = routeInfo.InsulationThickness;
                double offset1 = 0;
                int offset1Selection = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset1Sel", "Offset1Selection")).PropValue;
                if (offset1Selection == 1)
                {
                    string offset1Rule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrameOffset1Rl", "Offset1Rule")).PropValue;
                    GenericHelper.GetDataByRule(offset1Rule, null, out offset1);
                    support.SetPropertyValue(offset1, "IJUAhsFrameOffset1", "Offset1Value");
                }
                else
                    offset1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrameOffset1", "Offset1Value")).PropValue;
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
                    shoeHeight = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support, "IJUAhsFrameShoeHeight", "ShoeHeightValue")).PropValue; 

                switch (shoeHeightDefinition)
                {
                    case 1:
                        // Edge of Bounding Box
                        break;
                    case 2:
                        // Centerline of Primary Pipe
                        shoeHeight = shoeHeight + RefPortHelper.DistanceBetweenPorts(boundingBoxPort, "Route", PortAxisType.Z) - boundingBoxDepth;
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
                    string member1EndOverhangRule = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsMember1EndOHRl", "Member1EndOverhangRule")).PropValue;
                    GenericHelper.GetDataByRule(member1EndOverhangRule, null, out member1EndOverhang);
                    support.SetPropertyValue(member1EndOverhang, "IJUAhsMember1EndOH", "Member1EndOverhangValue");
                }
                else
                    member1EndOverhang = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsMember1EndOH", "Member1EndOverhangValue")).PropValue;
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
                //controlPointsList = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCP", "CP2");
                sectionData.CardinalPoint = CP2.PropValue;
                componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(sectionData.CardinalPoint), "IJOAhsSteelCP", "CP2");
                componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(sectionData.Orient) * Math.PI / 180), "IJOAhsEndCap", "EndCapRotZ");

                // ==========================
                // Joints To Connect the Main Steel Section to the BBX
                // ==========================
                double capPlateTh;
                if (isCapPlate)
                    capPlateTh = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE],"IJUAhsThickness1", "Thickness1")).PropValue;
                else
                    capPlateTh = 0;
                JointHelper.CreateRigidJoint("-1", boundingBoxPort, SECTION, "BeginFlex", Plane.XY, Plane.XY, Axis.X, Axis.X, boundingBoxDepth + shoeHeight + capPlateTh, boundingBoxWidth / 2 + offset1, 0);

                // ==========================
                // Joints To Connect Section To Supporting Structure
                // ==========================
                // If it is Equipment or a Shape, we will use GetProjectedPointOnSurface
                // For Steel, Slab and Wall, we will use a PointOnJoint
                Vector BB_X = new Vector(0, 0, 0), BB_Y = new Vector(0, 0, 0), BB_Z = new Vector(0, 0, 0), structZ = new Vector(0, 0, 0), structX = new Vector(0, 0, 0), structY = new Vector(0, 0, 0);
                Vector projectXZ = new Vector(0, 0, 0), projectZY = new Vector(0, 0, 0), projectXY = new Vector(0, 0, 0), memberYOffset = new Vector(0, 0, 0), memberZOffset = new Vector(0, 0, 0), memberProjectedNormal = new Vector(0, 0, 0);
                Position BB_Pos = new Position(0, 0, 0), memberProjectedPoint = new Position(0, 0, 0), memberStart = new Position(0, 0, 0);
                Matrix4X4 port = new Matrix4X4();
                double memberCutbackAngle = 0, memberLength = 0, planeAngle = 0;

                if (isPlacedOnSurface == true)
                {
                    // Get Projection from calculated point
                    port = RefPortHelper.PortLCS(boundingBoxPort);
                    BB_Z.Set(port.ZAxis.X, port.ZAxis.Y, port.ZAxis.Z);
                    BB_X.Set(port.XAxis.X, port.XAxis.Y, port.XAxis.Z);
                    BB_Y = BB_Z.Cross(BB_X);
                    BB_Pos.Set(port.Origin.X, port.Origin.Y, port.Origin.Z);

                    memberYOffset.Set(BB_Y.X, BB_Y.Y, BB_Y.Z);
                    memberYOffset.Length = boundingBoxWidth / 2 + offset1;
                    memberZOffset.Set(BB_Z.X, BB_Z.Y, BB_Z.Z);
                    memberZOffset.Length = boundingBoxDepth + shoeHeight + capPlateTh;

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
                    if (FrameAssemblyServices.AngleBetweenVectors(BB_Z, projectXZ) > FrameAssemblyServices.AngleBetweenVectors(BB_Z, projectZY))
                    {
                        // Cut Across BBX Plane
                        switch (sectionData.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP1");
                                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP2");
                                    componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(6), "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    memberCutbackAngle = -BB_Z.Angle(projectXZ, BB_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP1");
                                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP2");
                                    componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(8), "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    memberCutbackAngle = BB_Z.Angle(projectXZ, BB_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP1");
                                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP2");
                                    componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(4), "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    memberCutbackAngle = BB_Z.Angle(projectXZ, BB_Y);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP1");
                                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP2");
                                    componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(2), "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    memberCutbackAngle = -BB_Z.Angle(projectXZ, BB_Y);
                                    break;
                                }
                        }
                        componentDictionary[SECTION].SetPropertyValue(memberCutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                    }
                    else
                    {
                        // Cut in BBX Plane
                        switch (sectionData.Orient)
                        {
                            case 0:
                                {
                                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP1");
                                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(8), "IJOAhsSteelCP", "CP2");
                                    componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(8), "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    memberCutbackAngle = -BB_Z.Angle(projectZY, BB_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)90:
                                {
                                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP1");
                                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(4), "IJOAhsSteelCP", "CP2");
                                    componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(4), "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    memberCutbackAngle = -BB_Z.Angle(projectZY, BB_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)180:
                                {
                                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP1");
                                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(2), "IJOAhsSteelCP", "CP2");
                                    componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(2), "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    memberCutbackAngle = BB_Z.Angle(projectZY, BB_X);
                                    break;
                                }
                            case (FrameAssemblyServices.SteelOrientationAngle)270:
                                {
                                    componentDictionary[SECTION].SetPropertyValue(CP1.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP1");
                                    componentDictionary[SECTION].SetPropertyValue(CP2.PropValue = (int)Convert.ToDouble(6), "IJOAhsSteelCP", "CP2");
                                    componentDictionary[SECTION].SetPropertyValue(endCutbackAnchorPointSectionList.PropValue = (int)Convert.ToDouble(6), "IJOAhsCutback", "EndCutbackAnchorPoint");
                                    memberCutbackAngle = BB_Z.Angle(projectZY, BB_X);
                                    break;
                                }
                        }
                        componentDictionary[SECTION].SetPropertyValue(memberCutbackAngle, "IJOAhsCutback", "CutbackEndAngle");
                    }
                    JointHelper.CreatePrismaticJoint(SECTION, "EndFlex", SECTION, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }

                double basePlate1Th = 0, groutTh = 0;
                if (isBasePlate)
                    basePlate1Th = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsThickness1", "Thickness1")).PropValue;
                else
                    basePlate1Th = 0;

                if (isGrout)
                    groutTh = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrGroutTh", "GroutThickness")).PropValue; 
                else
                    groutTh = 0;

                componentDictionary[SECTION].SetPropertyValue(member1EndOverhang - basePlate1Th - groutTh, "IJUAHgrOccOverLength", "EndOverLength");
                planeAngle = RefPortHelper.AngleBetweenPorts(boundingBoxPort, PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Direct);
                if (isPlacedOnSurface == false)
                {
                    if ((planeAngle * 180 / Math.PI) < 45 || (planeAngle * 180 / Math.PI) >= 135)
                        JointHelper.CreatePointOnPlaneJoint(SECTION, "EndCap", "-1", "Structure", Plane.NegativeXY);
                    else
                        JointHelper.CreatePointOnPlaneJoint(SECTION, "EndCap", "-1", "Structure", Plane.NegativeZX);
                }

                // ==========================
                // Joints For the Pipe Attachments
                // ==========================
                string partProgId = string.Empty;
                Boolean oddQty;
                int index = 0;

                if (includePipeAtt1 == true)
                {
                    Part pipeAtt1Part = (Part)componentDictionary[PIPEATT1[0]].GetRelationship("madeFrom", "part").TargetObjects[0];
                    partProgId = (string)((PropertyValueString)(pipeAtt1Part).GetPropertyValue("IJDSymbolDefHelper", "ProgId")).PropValue;

                    if (pipeAtt1Qty % 2 > 0)
                        oddQty = true;
                    else
                        oddQty = false;
                    double pipeAtt1Offset;
                    try
                    {
                        pipeAtt1Offset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsFrPipeAtt1Offset", "PipeAtt1Offset")).PropValue;
                    }
                    catch
                    {
                        pipeAtt1Offset = 0;
                    }

                    for (index = 0; index < pipeAtt1Qty; index++)
                    {
                        if (oddQty == true)
                            JointHelper.CreateRigidJoint("-1", boundingBoxPort, PIPEATT1[index], "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, boundingBoxDepth / 2, boundingBoxWidth / 2, pipeAtt1Offset * (pipeAtt1Qty - (pipeAtt1Qty - 1) / 2 - index));
                        else
                            JointHelper.CreateRigidJoint("-1", boundingBoxPort, PIPEATT1[index], "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, boundingBoxDepth / 2, boundingBoxWidth / 2, pipeAtt1Offset * ((pipeAtt1Qty - 1) / 2 + 1 - index));

                        // If Pipe Attachment is a U-Bolt then set the Steel Thickness
                        if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.UBolt")
                            componentDictionary[PIPEATT1[index]].SetPropertyValue(capPlateTh, "IJOAhsSteelThickness", "SteelThickness");
                        double strapWidthInside = 0;
                        if (partProgId == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strap")
                        {
                            strapWidthInside = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(pipeAtt1Part,"IJUAhsStrap", "StrapWidthInside")).PropValue;
                            if( pipeConnection1Welds.Count>0)
                            {
                                weld = pipeConnection1Welds[index];
                                JointHelper.CreateAngularRigidJoint(weld.partKey, "Other", PIPEATT1[index], "Steel", new Vector(strapWidthInside / 2, 0, 0), new Vector(0, 0, 0));
                            }
                        }
                            

                    }
                }

                if (includePipeAtt2 == true)
                {
                    Part pipeAtt2Part = (Part)componentDictionary[PIPEATT2[0]].GetRelationship("madeFrom", "part").TargetObjects[0];
                    partProgId = (string)((PropertyValueString)(pipeAtt2Part).GetPropertyValue("IJDSymbolDefHelper", "ProgId")).PropValue;

                    if (pipeAtt2Qty % 2 > 0)
                        oddQty = true;
                    else
                        oddQty = false;
                    double pipeAtt2Offset;
                    try
                    {
                        pipeAtt2Offset = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsFrPipeAtt2Offset", "PipeAtt2Offset")).PropValue;
                    }
                    catch
                    {
                        pipeAtt2Offset = 0;
                    }
                    for (index = 0; index < pipeAtt2Qty; index++)
                    {
                        if (oddQty == true)
                            JointHelper.CreateRigidJoint("-1", boundingBoxPort, PIPEATT2[index], "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, boundingBoxDepth / 2, boundingBoxWidth / 2, pipeAtt2Offset * (pipeAtt2Qty - (pipeAtt2Qty - 1) / 2 - index));
                        else
                            JointHelper.CreateRigidJoint("-1", boundingBoxPort, PIPEATT2[index], "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, boundingBoxDepth / 2, boundingBoxWidth / 2, pipeAtt2Offset * ((pipeAtt2Qty - 1) / 2 + 1 - index));

                        // If Pipe Attachment is a U-Bolt then set the Steel Thickness
                        if (partProgId.ToUpper() == "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.UBolt")
                            componentDictionary[PIPEATT2[index]].SetPropertyValue(capPlateTh, "IJOAhsSteelThickness", "SteelThickness");
                    }
                }

                // ==========================
                // Joints For Remaining Plates and Grout
                // ==========================
                // Joints for End Plates
                double capPlateAngle1 = 0, basePlate1Angle = 0;
                if (isCapPlate)
                    capPlateAngle1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsCapPlate1Ang", "CapPlate1Angle")).PropValue;
                if (isBasePlate)
                    basePlate1Angle = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(part,"IJUAhsBasePlate1Ang", "BasePlate1Angle")).PropValue;

                if (isCapPlate)
                    JointHelper.CreateAngularRigidJoint(CAPPLATE, "Port2", SECTION, "BeginFace", new Vector(0, 0, 0), new Vector(0, 0, capPlateAngle1));

                // Joints for the Base Plates and Grout
                if (isBasePlate)
                    JointHelper.CreateAngularRigidJoint(BASEPLATE, "Port1", SECTION, "EndFace", new Vector(0, 0, 0), new Vector(0, 0, basePlate1Angle));

                if (isGrout)
                {
                    // Set the Grout Width and Length to match the base plate.
                    componentDictionary[GROUT].SetPropertyValue((double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsWidth1", "Width1")).PropValue, "IJUAhsWidth1", "Width1");
                    componentDictionary[GROUT].SetPropertyValue((double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsLength1", "Length1")).PropValue, "IJUAhsLength1", "Length1");
                    // Set the Grout Thickness
                    componentDictionary[GROUT].SetPropertyValue(groutTh, "IJUAhsThickness1", "Thickness1");

                    // Joint for the Grout
                    JointHelper.CreateRigidJoint(BASEPLATE, "Port2", GROUT, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
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
                                    double length1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsLength1", "Length1")).PropValue;
                                    double width1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[BASEPLATE],"IJUAhsWidth1", "Width1")).PropValue;
                                    switch (weld.location)
                                    {
                                        case 2:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - length1 / 2, weld.offsetXValue);
                                                break;
                                            }
                                        case 4:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - width1 / 2);
                                                break;
                                            }
                                        case 6:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + width1 / 2);
                                                break;
                                            }
                                        case 8:
                                            {
                                                JointHelper.CreateRigidJoint(BASEPLATE, "Port2", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + length1 / 2, weld.offsetXValue);
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
                                            JointHelper.CreateRigidJoint(SECTION, "EndFace", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - section.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                    case 4:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, "EndFace", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - section.width / 2);
                                            break;
                                        }
                                    case 6:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, "EndFace", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + section.width / 2);
                                            break;
                                        }
                                    case 8:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, "EndFace", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + section.depth / 2, weld.offsetXValue);
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
                                            JointHelper.CreateRigidJoint(SECTION, "BeginFace", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue - section.depth / 2, weld.offsetXValue);
                                            break;
                                        }
                                    case 4:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, "BeginFace", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue - section.width / 2);
                                            break;
                                        }
                                    case 6:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, "BeginFace", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, weld.offsetXValue + section.width / 2);
                                            break;
                                        }
                                    case 8:
                                        {
                                            JointHelper.CreateRigidJoint(SECTION, "BeginFace", weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue + section.depth / 2, weld.offsetXValue);
                                            break;
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
                if (excludeNotes == false)
                    note = CreateNote("Elevation", SECTION, "BeginFlex", new Position(0.0, 0.0, -capPlateTh), "Elevation", false, 2, 51, out controlPoint);
                else
                    DeleteNoteIfExists("Elevation");

                note = CreateNote("PipeCL", "-1", boundingBoxPort, new Position(0.0, boundingBoxWidth / 2, boundingBoxDepth / 2), " ", true, 2, 1, out controlPoint);
                note = CreateNote("Dim1", SECTION, "BeginFlex", new Position(0.0, 0.0, -capPlateTh), " ", true, 2, 1, out controlPoint);
                note = CreateNote("Dim2", SECTION, "EndFlex", new Position(0.0, 0.0, 0.0), " ", true, 2, 1, out controlPoint);
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
                    routeConnections.Add(new ConnectionInfo(SECTION, 1)); // partindex, routeindex

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
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject catalogPart = SupportOrComponent.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                // Get Section Information of each steel part
                IPart part = (IPart)componentDictionary[SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                section = FrameAssemblyServices.GetSectionDataFromPart(part.PartNumber);
                double capPlateTh = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[CAPPLATE],"IJUAhsThickness1", "Thickness1")).PropValue;

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

                int shoeHeightDefinition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameShoeHeightDef", "ShoeHeightDefinition")).PropValue;
                // Get the Shoe Height From the Input Attributes
                double shoeHeight = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameShoeHeight", "ShoeHeightValue")).PropValue;
                switch (shoeHeightDefinition)
                {
                    case 1:
                        // Edge of Bounding Box
                        break;
                    case 2:
                        // Centerline of Primary Pipe
                        shoeHeight = shoeHeight - pipeCLOffset;
                        break;
                }

                // ==========================
                // Length 1
                // ==========================
                double L1 = (double)((PropertyValueDouble)FrameAssemblyServices.GetPropertyValue(componentDictionary[SECTION],"IJUAHgrOccLength", "Length")).PropValue;
                int length1Definition = (int)((PropertyValueInt)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsFrameLength1Def", "Length1Definition")).PropValue;
                switch (length1Definition)
                {
                    case 1:
                        L1 = L1 + 0;
                        break;
                    case 2:
                        L1 = L1 + capPlateTh;
                        break;
                    case 3:
                        L1 = L1 + capPlateTh + shoeHeight + pipeCLOffset;
                        break;
                    default:
                        L1 = L1 + 0;
                        break;
                }
                SupportOrComponent.SetPropertyValue(L1, "IJUAhsFrameLength1", "Length1Value");

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
                string supportNumber = (string)((PropertyValueString)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsSupportNumber", "SupportNumber")).PropValue;
                PropertyValueCodelist steelStandardList = (PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(catalogPart, "IJUAhsSteelStandard", "SteelStandard");

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
    }
}
