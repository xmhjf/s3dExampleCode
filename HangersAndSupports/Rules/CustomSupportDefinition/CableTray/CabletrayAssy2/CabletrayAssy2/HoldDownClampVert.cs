//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HoldDownClampVert.cs
//   CabletrayAssy2,Ingr.SP3D.Content.Support.Rules.ClipHoldClampHoldDownClampVert
//   Author       : Vinay
//   Creation Date:  20/11/2015
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Support.Middle.Hidden;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class HoldDownClampVert : CustomSupportDefinition
    {
        int numOfCTs, numOfPart, clampBegin, clampEnd, hgrBeam, hgrBmBegin, hgrBmEnd;
        private const string CTHOLDDOWNCLAMP = "CTClipHoldClamp";
        private const string BEAM_ATT_1 = "WELDEDBEAMATTACHMENT1";
        private const string EYE_NUT_1 = "EYENUT1";
        private const string ROD_1 = "ROD1";
        private const string BEAM_ATT_2 = "WELDEDBEAMATTACHMENT2";
        private const string EYE_NUT_2 = "EYENUT2";
        private const string ROD_2 = "ROD2";
        private const string HGRBEAM = "HGRBEAM";
        string[] part = new string[20];
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    SupportComponentUtils supportComponentUtils = new SupportComponentUtils();
                    numOfCTs = SupportHelper.SupportedObjects.Count;
                    clampBegin = 4 + 1;
                    clampEnd = clampBegin + 2 * numOfCTs - 1;
                    hgrBmBegin = clampEnd + 1;
                    hgrBmEnd = hgrBmBegin + numOfCTs - 1;
                    numOfPart = hgrBmEnd;
                    string[] partClass = new string[numOfPart + 1];
                    for (int i = clampBegin; i <= clampEnd; i++)
                    {
                        partClass[i] = "CTHoldDownClamp";
                    }
                    for (int i = hgrBmBegin; i <= hgrBmEnd; i++)
                    {
                        partClass[i] = "RichHgrAISC31_L";

                    }

                    PropertyValueCodelist rodDia = ((PropertyValueCodelist)support.GetPropertyValue("IJUAHSA_RodDia", "RodDiameter"));
                    string rodDiametr1 = rodDia.PropertyInfo.CodeListInfo.GetCodelistItem(rodDia.PropValue).DisplayName;

                    // WBAAttachment 1
                    parts.Add(new PartInfo(BEAM_ATT_1, "S3Dhs_WBABolt-" + rodDiametr1));

                    //Flexible Rod 1
                    parts.Add(new PartInfo(ROD_1, "S3Dhs_RodCT-" + rodDiametr1));

                    //TopEyenut 1
                    parts.Add(new PartInfo(EYE_NUT_1, "S3Dhs_EyeNut-" + rodDiametr1));

                    // WBAAttachment 2
                    parts.Add(new PartInfo(BEAM_ATT_2, "S3Dhs_WBABolt-" + rodDiametr1));

                    //Flexible Rod 2
                    parts.Add(new PartInfo(ROD_2, "S3Dhs_RodCT-" + rodDiametr1));

                    //TopEyenut 2
                    parts.Add(new PartInfo(EYE_NUT_2, "S3Dhs_EyeNut-" + rodDiametr1));
                    for (int i = clampBegin; i <= clampEnd; i++)
                    {
                        part[i] = "part" + i;
                        Part FlatPlate = supportComponentUtils.GetPartFromPartClass(partClass[i], "", support);
                        parts.Add(new PartInfo(part[i], FlatPlate.ToString()));
                    }
                    for (int i = hgrBmBegin; i <= hgrBmEnd; i++)
                    {
                        part[i] = "part" + i;
                        parts.Add(new PartInfo(part[i], partClass[i], "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    }
                    return parts;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
                }
            }
        }
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========
                double depth;
                double width;
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                int iNumRoutes = SupportHelper.SupportedObjects.Count;
                int idxClamp = clampBegin;
                int idxCT = 1;
                double[] ctdepth = new double[iNumRoutes + 1];
                double[] ctwidth = new double[iNumRoutes + 1];
                double[] ctradius = new double[iNumRoutes + 1];
                for (int i = 1; i <= iNumRoutes; i++)
                {
                    CableTrayObjectInfo cableInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(i);
                    ctdepth[i] = cableInfo.Depth;
                    ctwidth[i] = cableInfo.Width;
                    ctradius[i] = cableInfo.BendRadius;
                    if (ctwidth[i] <= 0 || ctdepth[i] <= 0)
                    {
                        ctwidth[i] = ctradius[i] * 2;
                        ctdepth[i] = ctradius[i] * 2;
                    }
                }
                depth = ctdepth[1];
                width = ctwidth[1];
                for (int i = 1; i <= iNumRoutes; i++)
                {
                    if (width < ctwidth[i])
                    {
                        width = ctwidth[i];
                        depth = ctdepth[i];
                    }
                }

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;


                (componentDictionary[part[clampBegin]]).SetPropertyValue(ctwidth[1], "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[part[clampBegin]]).SetPropertyValue(ctdepth[1], "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[part[clampEnd]]).SetPropertyValue(ctwidth[1], "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[part[clampEnd]]).SetPropertyValue(ctdepth[1], "IJUAHgrCTOffset", "TrayDepth");
                double bOverLength;
                double eOverLength;
                for (int i = hgrBmBegin; i <= hgrBmEnd; i++)
                {
                    try
                    {
                        bOverLength = (double)((PropertyValueDouble)(componentDictionary[part[i]]).GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                        eOverLength = (double)((PropertyValueDouble)(componentDictionary[part[i]]).GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                    }
                    catch (InvalidOperationException )
                    {

                        bOverLength = 0.0;
                        eOverLength = 0.0;
                    }
                    
                }
                              
                Collection<object> collection = new Collection<object>();
                
                bool value = GenericHelper.GetDataByRule("HgrSupStructOffset", (componentDictionary[part[hgrBmBegin]]), out collection);

                double Offset = 0;
                if (collection != null)
                    Offset = (double)(collection[0]);
                double lugOffset = 0;
                if (Offset > lugOffset)
                    lugOffset = Offset;

                double dOffset = lugOffset;
                lugOffset = 2.0 * dOffset;
                bOverLength = lugOffset;
                eOverLength = lugOffset;

                for (int i = hgrBmBegin; i <= hgrBmEnd; i++)
                {
                    
                    (componentDictionary[part[i]]).SetPropertyValue(lugOffset, "IJUAHgrOccOverLength", "BeginOverLength");
                    (componentDictionary[part[i]]).SetPropertyValue(lugOffset, "IJUAHgrOccOverLength", "EndOverLength");
                }
                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================
                string strBBLow, strBBHigh;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    strBBLow = "BBSR_Low";
                    strBBHigh = "BBSR_High";
                }
                else
                {
                    strBBLow = "BBR_Low";
                    strBBHigh = "BBR_High";
                }
                BoundingBoxHelper.CreateStandardBoundingBoxes(true);
                BoundingBox boundingBox;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double dWidth = boundingBox.Width;
                double dHeight = boundingBox.Height;
                double endOffset = dWidth / 2 + dOffset;
            
                BusinessObject horizontalSectionPart = componentDictionary[part[hgrBmBegin]].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
            
                hgrBeam = hgrBmBegin + GetBottomRouteIndex() - 1;
                string[] strRoute = new string[numOfCTs + 1];
                double distance;
                ControlPoint controlPoint;
                int idxnote = hgrBmEnd + 3;
                double endNotePosition = 0;
                try
                {
                    endNotePosition = (double)((PropertyValueDouble)(componentDictionary[part[hgrBeam]].GetPropertyValue("IJUAHgrOccLength", "Length"))).PropValue;
                }
                catch
                {
                    endNotePosition = 0;
                }
                //Create Notes
                    Note note3 = CreateNote("L_Start", part[hgrBeam], "BeginCap", new Position(0, 0, -bOverLength), " ", true, 2, 1, out controlPoint);
                    note3.SetPropertyValue("", "IJGeneralNote", "Text");
                    PropertyValueCodelist note3PropertyValueCL = (PropertyValueCodelist)note3.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList3 = note3PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                    note3.SetPropertyValue(codeList3, "IJGeneralNote", "Purpose");
                    note3.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                    Note note4 = CreateNote("L_End", part[hgrBeam], "BeginCap", new Position(0, 0, endNotePosition+lugOffset), " ", true, 2, 1, out controlPoint);
                    note4.SetPropertyValue("", "IJGeneralNote", "Text");
                    PropertyValueCodelist note4PropertyValueCL = (PropertyValueCodelist)note3.GetPropertyValue("IJGeneralNote", "Purpose");
                    CodelistItem codeList4 = note4PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                    note4.SetPropertyValue(codeList4, "IJGeneralNote", "Purpose");
                    note4.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                    if (numOfCTs == 1)
                    {
                        distance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                        //JointHelper.CreateRigidJoint(part[hgrBmEnd + 3], "Dimension", part[hgrBmEnd + 1], "Dimension", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, distance);
                        Note note5 = CreateNote("L_1", part[hgrBeam], "BeginCap", new Position(-distance, 0, -bOverLength), " ", true, 2, 1, out controlPoint);
                        note5.SetPropertyValue("", "IJGeneralNote", "Text");
                        PropertyValueCodelist notePropertyValueCL = (PropertyValueCodelist)note5.GetPropertyValue("IJGeneralNote", "Purpose");
                        CodelistItem codeList5 = notePropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                        note5.SetPropertyValue(codeList5, "IJGeneralNote", "Purpose");
                        note5.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
                    }
                    else
                    {
                        int[] intRouteOrder = new int[numOfCTs + 1];
                        double[] dDistanceroutestructure = new double[numOfCTs + 1];

                        string strRouteTemp;
                        int intSmallest = 0;
                        for (int i = 1; i <= numOfCTs; i++)
                        {
                            if (i == 1)
                                strRouteTemp = "Route";
                            else
                                strRouteTemp = "Route_" + i;
                            distance = RefPortHelper.DistanceBetweenPorts(strRouteTemp, "Structure", PortDistanceType.Vertical);
                            dDistanceroutestructure[i] = distance;
                        }
                        for (int i = 1; i <= numOfCTs; i++)
                        {
                            for (int j = 1; j <= numOfCTs; j++)
                            {
                                if (dDistanceroutestructure[i] > dDistanceroutestructure[j])
                                    intRouteOrder[i] = i;
                                else
                                {
                                    intSmallest = i;
                                    intRouteOrder[i] = j;
                                }
                            }
                        }
                        intRouteOrder[numOfCTs] = intSmallest;

                        for (int i = 1; i <= numOfCTs; i++)
                        {
                            if (intRouteOrder[1] == 1)
                                strRouteTemp = "Route";
                            else
                                strRouteTemp = "Route_" + intRouteOrder[1];

                            if (i == 1)
                            {
                                if (intRouteOrder[1] == 1)
                                    distance = RefPortHelper.DistanceBetweenPorts(strRouteTemp, "Route_" + numOfCTs, PortDistanceType.Vertical);
                                else
                                    distance = RefPortHelper.DistanceBetweenPorts(strRouteTemp, "Route", PortDistanceType.Vertical);
                            }
                            else
                            {
                                if (i == numOfCTs)
                                {
                                    distance = RefPortHelper.DistanceBetweenPorts(strRouteTemp, "Structure", PortDistanceType.Vertical) + ctdepth[i] / 2;
                                    bOverLength = dOffset;
                                }
                                else
                                    distance = RefPortHelper.DistanceBetweenPorts(strRouteTemp, "Route_" + i, PortDistanceType.Vertical);
                            }
                            int noteIndex = i + 2;
                            Note note5 = CreateNote("L_" + noteIndex, part[hgrBeam], "BeginCap", new Position(-distance, 0, -bOverLength), " ", true, 2, 1, out controlPoint);
                            note5.SetPropertyValue("", "IJGeneralNote", "Text");
                            PropertyValueCodelist note5PropertyValueCL = (PropertyValueCodelist)note5.GetPropertyValue("IJGeneralNote", "Purpose");
                            CodelistItem codeList5 = note5PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                            note5.SetPropertyValue(codeList5, "IJGeneralNote", "Purpose");
                            note5.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                            idxnote = idxnote + 1;
                        }
                    }
                for (int i = hgrBmBegin; i <= hgrBmEnd; i++)
                {
                    
                    if (idxCT == 1)
                        strRoute[idxCT] = "Route";
                    else
                        strRoute[idxCT] = "Route_" + idxCT;
                    //Create the Joint between the RteLow Reference Port and the HgrBeam BeginCap
                    JointHelper.CreatePlanarJoint("-1", strBBLow, part[i], "BeginCap", Plane.ZX, Plane.XY, -dOffset);

                    //Create the Joint between the RteHigh Reference Port and the HgrBeam EndCap
                    JointHelper.CreatePlanarJoint("-1", strBBHigh, part[i], "EndCap", Plane.ZX, Plane.XY, dOffset);

                    //Create the Joint between the igh Reference Port
                    JointHelper.CreatePlanarJoint("-1", strRoute[idxCT], part[i], "BeginCap", Plane.XY, Plane.NegativeYZ, -0.5 * ctdepth[idxCT]);
                    
                    //if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        JointHelper.CreatePointOnPlaneJoint(part[i], "Neutral", "-1", strRoute[idxCT], Plane.YZ);

                    //Add a Joint between cable tray Clamp and Route
                    JointHelper.CreatePointOnAxisJoint(part[idxClamp], "Route", "-1", strRoute[idxCT], Axis.X);
                    JointHelper.CreatePointOnAxisJoint(part[idxClamp + 1], "Route", "-1", strRoute[idxCT], Axis.X);

                    //Add a Joint between cable tray Clamp and support beam
                    JointHelper.CreateTranslationalJoint(part[idxClamp], "Structure", part[i], "Neutral", Plane.YZ, Plane.ZX, Axis.Z, Axis.NegativeX, 0);
                    JointHelper.CreateTranslationalJoint(part[idxClamp + 1], "Structure", part[i], "Neutral", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.NegativeX, 0);
                    JointHelper.CreatePrismaticJoint(part[i], "BeginCap", part[i], "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    idxCT = idxCT + 1;
                    idxClamp = idxClamp + 2;
                }
                hgrBeam = hgrBmBegin + GetBottomRouteIndex() - 1;


                //Add a Spherical Joint between Support beam and Bottom of Rod
                JointHelper.CreateRigidJoint(ROD_1, "RodEnd1", part[hgrBeam], "Neutral", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, -0.5 * width - lugOffset, 0);
                JointHelper.CreateRigidJoint(ROD_2, "RodEnd1", part[hgrBeam], "Neutral", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, 0.5 * width + lugOffset, 0);
                for (int i = hgrBmBegin; i <= hgrBmEnd; i++)
                {
                    if (i != hgrBeam)
                        JointHelper.CreatePointOnPlaneJoint(part[i], "Neutral", part[hgrBeam], "Neutral", Plane.ZX);
                }


                // Add a Planar Joint between the Beam Attachment and the structure
                JointHelper.CreatePlanarJoint(BEAM_ATT_1, "Structure", "-1", "Structure", Plane.XY, Plane.XY, 0);

                // Add a revolute Joint between the Eye nut and Beam Attachment
                JointHelper.CreateRevoluteJoint(EYE_NUT_1, "Eye", BEAM_ATT_1, "Pin", Axis.Y, Axis.Y);

                //Add a Rigid Joint between Rod and Eyenut
                JointHelper.CreateRigidJoint(ROD_1, "RodEnd2", EYE_NUT_1, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                // Add a Vertical Joint to the Rod Z axis
                JointHelper.CreateGlobalAxesAlignedJoint(ROD_1, "RodEnd1", Axis.Z, Axis.Z);

                // Add the Flexible (Prismatic) Joint between the ports of the bottom rod
                JointHelper.CreatePrismaticJoint(ROD_1, "RodEnd1", ROD_1, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                // Add a Planar Joint between the Beam Attachment and the structure
                JointHelper.CreatePlanarJoint(BEAM_ATT_2, "Structure", "-1", "Structure", Plane.XY, Plane.XY, 0);

                // Add a revolute Joint between the Eye nut and Beam Attachment
                JointHelper.CreateRevoluteJoint(EYE_NUT_2, "Eye", BEAM_ATT_2, "Pin", Axis.Y, Axis.Y);

                //Add a Rigid Joint between Rod and Eyenut
                JointHelper.CreateRigidJoint(ROD_2, "RodEnd2", EYE_NUT_2, "RodEnd", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                // Add a Vertical Joint to the Rod Z axis
                JointHelper.CreateGlobalAxesAlignedJoint(ROD_2, "RodEnd1", Axis.Z, Axis.Z);

                // Add the Flexible (Prismatic) Joint between the ports of the bottom rod
                JointHelper.CreatePrismaticJoint(ROD_2, "RodEnd1", ROD_2, "RodEnd2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

            }
            catch (Exception exception)
            {
                Type myType = this.GetType();
                CmnException exception1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                throw exception1;
            }
        }
        public int GetBottomRouteIndex()
        {
            int index = 0;
            double dMin = 10000000;

            for (int i = 1; i <= numOfCTs; i++)
            {
                CableTrayObjectInfo cableInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(i);
                Position pos = cableInfo.StartLocation;
                if (pos.Z < dMin)
                {
                    dMin = pos.Z;
                    index = i;
                }
            }
            return index;
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
                    for (int i = clampBegin; i <= SupportHelper.SupportedObjects.Count; i++)
                    {
                        routeConnections.Add(new ConnectionInfo(part[i], 1)); // partindex, routeindex
                        routeConnections.Add(new ConnectionInfo(part[i + 1], 1)); // partindex, routeindex
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

                    structConnections.Add(new ConnectionInfo(BEAM_ATT_1, 1)); // partindex, structureindex
                    structConnections.Add(new ConnectionInfo(BEAM_ATT_2, 1)); // partindex, structureindex


                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
                }
            }
        }
    }
}