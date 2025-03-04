//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HoldDownClampHori.cs
//   CabletrayAssy2,Ingr.SP3D.Content.Support.Rules.ClipHoldClampHoldDownClampHori
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

namespace Ingr.SP3D.Content.Support.Rules
{
    public class HoldDownClampHori : CustomSupportDefinition
    {
        int cableTrays, numOfPart, clampBegin, clampEnd, hgrBeam;
        private const string CTHOLDDOWNCLAMP = "CTHOLDDOWNCLAMP";
        private const string BEAM_ATT_1 = "WELDEDBEAMATTACHMENT1";
        private const string EYE_NUT_1 = "EYENUT1";
        private const string ROD_1 = "ROD1";
        private const string BEAM_ATT_2 = "WELDEDBEAMATTACHMENT2";
        private const string EYE_NUT_2 = "EYENUT2";
        private const string ROD_2 = "ROD2"; 
        private const string HGRBEAM = "HGRBEAM";
        string[] part = new string[10];
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
                    cableTrays = SupportHelper.SupportedObjects.Count;
                    clampBegin = 4 + 1;
                    clampEnd = clampBegin + 2 * cableTrays - 1;
                    hgrBeam = clampEnd + 1;
                    numOfPart = hgrBeam;
                    string[] partClass = new string[numOfPart + 1];
                    for (int i = clampBegin; i <= clampEnd; i++)
                    {
                        partClass[i] = "CTHoldDownClamp";
                    }
                    partClass[hgrBeam] = "RichHgrAISC31_L";


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
                    for (int i = clampBegin; i <= numOfPart-1; i++)
                    {
                        part[i] = "part" + i;
                        Part FlatPlate = supportComponentUtils.GetPartFromPartClass(partClass[i], "", support);
                        parts.Add(new PartInfo(part[i], FlatPlate.ToString()));
                    }
                    parts.Add(new PartInfo(HGRBEAM, partClass[hgrBeam], "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========
              
                double width;
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                int iNumRoutes = SupportHelper.SupportedObjects.Count;
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
                //Set dWidth and dDepth to the largest CT
               
                width = ctwidth[1];
                for (int i = 1; i <= iNumRoutes; i++)
                {
                    if (width < ctwidth[i])
                    {
                        width = ctwidth[i];
                        
                    }
                }

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                for (int i = 1; i <= iNumRoutes; i++)
                {
                    (componentDictionary[part[clampBegin]]).SetPropertyValue(ctwidth[i], "IJUAHgrCTOffset", "TrayWidth");
                    (componentDictionary[part[clampBegin]]).SetPropertyValue(ctdepth[i], "IJUAHgrCTOffset", "TrayDepth");
                    (componentDictionary[part[clampBegin + 1]]).SetPropertyValue(ctwidth[i], "IJUAHgrCTOffset", "TrayWidth");
                    (componentDictionary[part[clampBegin + 1]]).SetPropertyValue(ctdepth[i], "IJUAHgrCTOffset", "TrayDepth");
                }
                double eOverLength, bOverLength;                
                Collection<object> collection = new Collection<object>();
                bool value = GenericHelper.GetDataByRule("HgrSupStructOffset", (componentDictionary[HGRBEAM]), out collection);
                double offset = 0;
                if (collection != null)
                    offset = (double)(collection[0]);
                double lugOffset = 0;
                lugOffset = 2 * offset;
                bOverLength = eOverLength = lugOffset;
                (componentDictionary[HGRBEAM]).SetPropertyValue((bOverLength), "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HGRBEAM]).SetPropertyValue((eOverLength), "IJUAHgrOccOverLength", "BeginOverLength");


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
                
                double endOffset = 0;
                if (collection != null)
                    endOffset = dWidth / 2 + (double)(collection[0]);

                //Create the Joint between the RteLow Reference Port and the HgrBeam BeginCap
                JointHelper.CreatePlanarJoint("-1", strBBLow, HGRBEAM, "BeginCap", Plane.ZX, Plane.XY, -offset);

                //Create the Joint between the RteHigh Reference Port and the HgrBeam EndCap
                JointHelper.CreatePlanarJoint("-1", strBBHigh, HGRBEAM, "EndCap", Plane.ZX, Plane.XY, offset);

                //Create the Joint between the igh Reference Port
                JointHelper.CreatePlanarJoint("-1", strBBLow, HGRBEAM, "BeginCap", Plane.XY, Plane.NegativeYZ, 0);
                //if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    JointHelper.CreatePointOnPlaneJoint(HGRBEAM, "Neutral", "-1", strBBLow, Plane.YZ);

                //Add a flexable Joint for HgrBeam
                JointHelper.CreatePrismaticJoint(HGRBEAM, "BeginCap", HGRBEAM, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                int clamp = clampBegin;
                string[] strRoute = new string[iNumRoutes + 1];
                for (int i = 1; i <= iNumRoutes; i++)
                {
                    if (i == 1)
                        strRoute[i] = "Route";
                    else
                        strRoute[i] = "Route_" + i;

                    //Add a Joint between cable tray Clamp and Route
                    JointHelper.CreatePrismaticJoint(part[clamp], "Route", "-1", strRoute[i], Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    JointHelper.CreatePrismaticJoint(part[clamp + 1], "Route", "-1", strRoute[i], Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);

                    //Add a Joint between cable tray Clamp and support beam
                    JointHelper.CreatePointOnPlaneJoint(part[clamp], "Structure", HGRBEAM, "Neutral", Plane.ZX);
                    JointHelper.CreatePointOnPlaneJoint(part[clamp + 1], "Structure", HGRBEAM, "Neutral", Plane.ZX);
                    clamp = clamp + 2;
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

                //Add a Rigid Joint between cable tray Clamp and Bottom of Rod
                JointHelper.CreateRigidJoint(ROD_1, "RodEnd1", HGRBEAM, "Neutral", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, -0.5 * dWidth - lugOffset, 0);

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

                //Add a Spherical Joint between Support beam and Bottom of Rod

                JointHelper.CreateRigidJoint(ROD_2, "RodEnd1", HGRBEAM, "Neutral", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, 0.5 * dWidth + lugOffset, 0);



            }
            catch (Exception exception)
            {
                Type myType = this.GetType();
                CmnException exception1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                throw exception1;
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
                    for (int i = clampBegin; i <= SupportHelper.SupportedObjects.Count; i++)
                    {
                        routeConnections.Add(new ConnectionInfo(part[i], 1)); // partindex, routeindex
                        routeConnections.Add(new ConnectionInfo(part[i + 1], 1)); // partindex, routeindex
                    }
                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception exception)
                {
                    Type myType = this.GetType();
                    CmnException exception1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + exception.Message, exception);
                    throw exception1;
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