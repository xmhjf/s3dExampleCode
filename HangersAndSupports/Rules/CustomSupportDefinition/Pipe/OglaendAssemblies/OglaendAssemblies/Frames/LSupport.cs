//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   LSupport.cs
//   OglaendAssemblies,Ingr.SP3D.Content.Support.Rules.LSupport
//   Author       :  PVK
//   Creation Date:  07-11-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET  
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
    public class LSupport : CustomSupportDefinition, ICustomHgrBOMDescription
    {

        private const string SECTION = "SECTION";
        private const string LEG = "LEG";

        private const string WELDSTAT_1 = "WELDSTAT_1";


        //Steel Members - Stores Data related to steel member (Part Number, Depth, Width etc...)
        private OglaendAssemblyServices.HSSteelMember section;
        private OglaendAssemblyServices.HSSteelMember leg1;

        //Steel Configuration - Stores Data related to steel configuration
        private OglaendAssemblyServices.HSSteelConfig sectionData;
        private OglaendAssemblyServices.HSSteelConfig leg1Data;


        int[] structureIndex = new int[2];
        Boolean isMember, isLeg1, isweldstat1;
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
                    isMember = OglaendAssemblyServices.AddPart(this, SECTION, member1Part, strRule, parts, ruleBO);

                    string leg1Part = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLeg1", "Leg1Part")).PropValue;
                    isLeg1 = OglaendAssemblyServices.AddPart(this, LEG, leg1Part, null, parts);

                    // Add the Plates
                    string Weldstat1 = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsWeldStarter1", "WeldStarter1Part")).PropValue;
                    isweldstat1 = OglaendAssemblyServices.AddPart(this, WELDSTAT_1, Weldstat1, null, parts);

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
                leg1 = OglaendAssemblyServices.GetSectionDataFromPartIndex(this, LEG);

                // Get the Steel Configuration Data
                sectionData.Orient = (OglaendAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsMember1Ang", "Member1OrientationAngle")).PropValue);
                leg1Data.Orient = (OglaendAssemblyServices.SteelOrientationAngle)((int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsLeg1Ang", "Leg1OrientationAngle")).PropValue);

                // Get the Width and Depth of each steel part according to the orientation angle
                double leg1Depth = 0, leg1Width = 0, sectionDepth = 0, sectionWidth = 0;
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


                //'==========================
                // 'SPAN
                // '==========================

                Double SpanWidth, MaxSpanWidth = 0, MinSpanWidth = 0, BBXWidth = 0, BBXWidth2 = 0, MaxHeight = 0;
                String sSupportNumber;
                
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
                    if (HgrCompareDoubleService.cmpdbl(Math.Round(BBXWidth2 * 1000, 0), Math.Round(BBXWidth * 1000, 0) )==false)
                    {
                        if ((boundingBoxWidth + leg1Depth / 2 ) > MaxSpanWidth)
                        {
                            SpanWidth = MaxSpanWidth;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMaxWidth,"Width Exceeds Maximum Allowable Width. Resetting it to Maximum Width."), "", "USupport_1", 346);
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                        else
                        {
                            SpanWidth = boundingBoxWidth + leg1Depth ;
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
                        else if (HgrCompareDoubleService.cmpdbl(SpanWidth, 0)==true)
                        {
                            SpanWidth = boundingBoxWidth + leg1Depth ;
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }
                        else if (SpanWidth < MinSpanWidth)
                        {
                            SpanWidth = MinSpanWidth;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMaxWidth,"Width is less than Minimumn Allowable Width. Resetting it to Maximum Width."), "", "USupport_1", 346);
                            support.SetPropertyValue(SpanWidth, "IJUAhsOGSpanWidth", "SpanWidth");
                        }

                    }

                

                //Height//
                double ShoeHeight, RSDistance, InputHeight, SupportHeight, RSdistance2;
                string Refport;
    
                    Refport = "StructAlt";

                ShoeHeight = 0;
                try
                {
                    InputHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGLegHeight", "LegHeight")).PropValue;
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

                if ((HgrCompareDoubleService.cmpdbl(Math.Round(RSdistance2 * 1000, 0), Math.Round(RSDistance * 1000, 0))==false))
                {
                    if (((Boolean)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsLengthOverride", "LengthOverride")).PropValue) == true)
                    {
                        if (RSDistance > MaxHeight)
                        {
                            ShoeHeight = ShoeHeight - (SupportHeight - MaxHeight);
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE",OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxHeight, "Height Exceeds Maximum Allowable Height. Resetting it to Maximum Height."), "", "USupport_1", 346);
                            support.SetPropertyValue(MaxHeight, "IJUAhsOGLegHeight", "LegHeight");                                                         
                        }                                            
                
                        else
                        {
                            ShoeHeight = 0;
                            support.SetPropertyValue(SupportHeight, "IJUAhsOGLegHeight", "LegHeight");
                        }
                    }
                    else
                    {
                        if (InputHeight > MaxHeight)
                        {
                            ShoeHeight = ShoeHeight - (SupportHeight - MaxHeight);
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxHeight,"Height Exceeds Maximum Allowable Height. Resetting it to Maximum Height."), "", "USupport_1", 346);
                            support.SetPropertyValue(MaxHeight, "IJUAhsOGLegHeight", "LegHeight");
                        }
                        else
                        {
                            ShoeHeight = ShoeHeight - (SupportHeight - InputHeight);
                            support.SetPropertyValue(InputHeight, "IJUAhsOGLegHeight", "LegHeight");
                        }
                    }
                }
                else
                {
                    if (InputHeight > MaxHeight)
                    {
                        ShoeHeight = ShoeHeight - (SupportHeight - MaxHeight);
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxHeight,"Height Exceeds Maximum Allowable Height. Resetting it to Maximum Height."), "", "USupport_1", 346);
                        support.SetPropertyValue(MaxHeight, "IJUAhsOGLegHeight", "LegHeight");
                    }
                    else if (HgrCompareDoubleService.cmpdbl(InputHeight,0)==true)
                    {
                        ShoeHeight = 0;
                        support.SetPropertyValue(SupportHeight, "IJUAhsOGLegHeight", "LegHeight");
                    }
                    else
                    {
                        ShoeHeight = ShoeHeight - (SupportHeight - InputHeight);
                        support.SetPropertyValue(InputHeight, "IJUAhsOGLegHeight", "LegHeight");
                    }

                }
                int connection1Type = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsCornerConn1Type", "Connection1Type")).PropValue;
                double FrameOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsOGFrameOffset", "FrameOffset")).PropValue;


                PropertyValueCodelist cardinalPoint7VerSection = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint2Leg = (PropertyValueCodelist)componentDictionary[LEG].GetPropertyValue("IJOAhsSteelCP", "CP2");
                PropertyValueCodelist cardinalPoint6HorSection = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCP", "CP1");
                PropertyValueCodelist cardinalPoint1Leg = (PropertyValueCodelist)componentDictionary[LEG].GetPropertyValue("IJOAhsSteelCP", "CP1");

                if (connection1Type == 0)
                {
                    componentDictionary[LEG].SetPropertyValue(cardinalPoint2Leg.PropValue =(int)5, "IJOAhsSteelCP", "CP2");
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue =(int)5, "IJOAhsSteelCP", "CP1");

                    componentDictionary[SECTION].SetPropertyValue(Math.PI, "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[SECTION].SetPropertyValue(FrameOffset, "IJOAhsFlexPort", "FlexPortZOffset");

                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(sectionDepth/2), "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble( - sectionDepth/2), "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(SpanWidth - sectionDepth / 2), "IJUAHgrOccLength", "Length");

                    JointHelper.CreateAngularRigidJoint(LEG,"EndCap",SECTION ,"BeginCap", new Vector(0,leg1Depth /2,0),new Vector(0,Math.PI /2,Math.PI ));

                    JointHelper.CreateAngularRigidJoint(SECTION, "BeginFlex", "-1", boundingBoxPort, new Vector(0, -leg1Depth / 2, -ShoeHeight), new Vector((3 * Math.PI) / 2, 0, (3 * Math.PI) / 2));

                    JointHelper.CreatePrismaticJoint(LEG, "EndFlex", LEG, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    JointHelper.CreatePointOnPlaneJoint(LEG, "BeginFlex", "-1", "Structure", Plane.NegativeXY);


                }
                else
                {

                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint6HorSection.PropValue = (int)1, "IJOAhsSteelCP", "CP1");
                    componentDictionary[LEG].SetPropertyValue(cardinalPoint1Leg.PropValue = (int)1, "IJOAhsSteelCP", "CP1");

                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(0), "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[SECTION].SetPropertyValue(FrameOffset, "IJOAhsFlexPort", "FlexPortZOffset");

                    componentDictionary[LEG].SetPropertyValue(Convert.ToDouble(sectionDepth), "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(-leg1Depth / 2), "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(leg1Depth), "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[SECTION].SetPropertyValue(Convert.ToDouble(SpanWidth - leg1Depth / 2), "IJUAHgrOccLength", "Length");

                    JointHelper.CreateAngularRigidJoint(LEG, "EndCap", SECTION, "BeginCap", new Vector(0, 0, 0), new Vector(-Math.PI / 2, 0, Math.PI / 2));

                    JointHelper.CreateAngularRigidJoint(SECTION, "BeginFlex", "-1", boundingBoxPort, new Vector(leg1Depth / 2, 0, -leg1Depth / 2 - ShoeHeight), new Vector((3 * Math.PI) / 2, 0, 0));

                    JointHelper.CreatePrismaticJoint(LEG, "EndFlex", LEG, "BeginFlex", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    JointHelper.CreatePointOnPlaneJoint(LEG, "BeginFlex", "-1", "Structure", Plane.NegativeXY);
            }
                //'========================
                //'Joints for Wels Starters
                //'========================
                double Wlength1, WeldFlangeWidth, Xoffset,Yoffset, Zoffset,StructureIntPosZ,WeldThickness;

                PropertyValue oWlength1 = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDSTAT_1], "IJUAHgrOGLength", "Length", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);
                Wlength1 = (double)(((PropertyValueDouble)oWlength1).PropValue);

                PropertyValue oWeldFlangeWidth = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDSTAT_1], "IJUAHgrOGFlangeWidth", "FlangeWidth", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);
                WeldFlangeWidth = (double)(((PropertyValueDouble)oWeldFlangeWidth).PropValue);                

                PropertyValue oStructureIntPosZ = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDSTAT_1], "IJUAHgrOGStructureIntPort", "StuctureIntPosZ", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);
                StructureIntPosZ = (double)(((PropertyValueDouble)oStructureIntPosZ).PropValue);

                PropertyValue oThickness = OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELDSTAT_1], "IJUAHgrOGThickness", "Thickness", OglaendAssemblyServices.FromCatalogOrOccurrence.FromCatalog, false);
                WeldThickness = (double)(((PropertyValueDouble)oThickness).PropValue);

                Xoffset = leg1Depth / 2 - WeldFlangeWidth / 2;
                Zoffset = Wlength1 / 2 + StructureIntPosZ;
                Yoffset = leg1Depth / 2;

                if (connection1Type == 0) // for butted
                {
                    if (Xoffset < 0)//for weld starters outside the lug
                        Xoffset = WeldFlangeWidth / 2 - WeldThickness;
                    else
                    {
                        Xoffset = WeldFlangeWidth / 2 + leg1.webThickness;
                    }
                    JointHelper.CreateAngularRigidJoint(WELDSTAT_1, "Structure", LEG, "BeginFlex", new Vector(Xoffset, 0, 0), new Vector(0, Math.PI, Math.PI));
                }
                else// for lapped
                {
                    JointHelper.CreateAngularRigidJoint(WELDSTAT_1, "StructureInt", LEG, "BeginFlex", new Vector(-Xoffset, -Yoffset, Zoffset), new Vector(0, 0, 0));
                }
                // setend overlength
                    
                
                if(connection1Type==0)
                    componentDictionary[LEG].SetPropertyValue(-Wlength1 / 2, "IJUAHgrOccOverLength", "BeginOverLength");
                else
                    componentDictionary[LEG].SetPropertyValue(-Wlength1 / 2, "IJUAHgrOccOverLength", "BeginOverLength");


                //'Query For Loads
                Double CapacityPy = 0, CapacityPx = 0, CapacityDWPy = 0, PipingPy = 0, PipingPx = 0, BlastPy = 0, BlastPx = 0;
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
                        CapacityDWPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "CapacityDWPy")).PropValue);
                        PipingPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "PipingPy")).PropValue);
                        PipingPx = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "PipingPx")).PropValue);
                        BlastPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "BlastPy")).PropValue);
                        BlastPx = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "BlastPx")).PropValue);
                        break;

                    }

                }
                //set load properties
                if (support.SupportsInterface("IJUAhsOGCapacityPy"))
                    support.SetPropertyValue(CapacityPy, "IJUAhsOGCapacityPy", "CapacityPy");
                if (support.SupportsInterface("IJUAhsOGCapacityPx"))
                    support.SetPropertyValue(CapacityPx, "IJUAhsOGCapacityPx", "CapacityPx");
                if (support.SupportsInterface("IJUAhsOGCapacityDWPy"))
                    support.SetPropertyValue(CapacityDWPy, "IJUAhsOGCapacityDWPy", "CapacityDWPy");
                if (support.SupportsInterface("IJUAhsOGPipingPy"))
                    support.SetPropertyValue(PipingPy, "IJUAhsOGPipingPy", "PipingPy");
                if (support.SupportsInterface("IJUAhsOGPipingPx"))
                    support.SetPropertyValue(PipingPx, "IJUAhsOGPipingPx", "PipingPx");
                if (support.SupportsInterface("IJUAhsOGBlastPy"))
                    support.SetPropertyValue(BlastPy, "IJUAhsOGBlastPy", "BlastPy");
                if (support.SupportsInterface("IJUAhsOGBlastPx"))
                    support.SetPropertyValue(BlastPx, "IJUAhsOGBlastPx", "BlastPx");


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
                    if (isweldstat1)
                        structConnections.Add(new ConnectionInfo(WELDSTAT_1, 1)); // partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(LEG, 1));

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
