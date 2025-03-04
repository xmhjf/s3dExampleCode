//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HoldDownClamp.cs
//   CabletrayAssy2,Ingr.SP3D.Content.Support.Rules.ClipHoldClampHoldDownClamp
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
    public class HoldDownClamp : CustomSupportDefinition
    {
        private const string CTHOLDDOWNCLAMP1 = "CTHOLDDOWNCLAMP1";
        private const string CTHOLDDOWNCLAMP2 = "CTHOLDDOWNCLAMP2";
        private const string BEAM_ATT_1 = "WELDEDBEAMATTACHMENT1";
        private const string EYE_NUT_1 = "EYENUT1";
        private const string ROD_1 = "ROD1";
        private const string BEAM_ATT_2 = "WELDEDBEAMATTACHMENT2";
        private const string EYE_NUT_2 = "EYENUT2";
        private const string ROD_2 = "ROD2"; 
        private const string HGRBEAM = "HGRBEAM";   

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
                    Collection<object> colllection = new Collection<object>();
                    PartClass ctClipHoldClamp = (PartClass)catalogBaseHelper.GetPartClass("CTHoldDownClamp");
                    string strPartSelection = ctClipHoldClamp.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();

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

                    parts.Add(new PartInfo(CTHOLDDOWNCLAMP1, "CTHoldDownClamp", strPartSelection));
                    parts.Add(new PartInfo(CTHOLDDOWNCLAMP2, "CTHoldDownClamp", strPartSelection));
                    parts.Add(new PartInfo(HGRBEAM, "RichHgrAISC31_L", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
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
                CableTrayObjectInfo cableInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                double depth = cableInfo.Depth;
                double width = cableInfo.Width;
                double radius = cableInfo.BendRadius;

                if (width <= 0 || depth <= 0)
                {
                    width = radius * 2;
                    depth = radius * 2;
                }

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                (componentDictionary[CTHOLDDOWNCLAMP1]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTHOLDDOWNCLAMP1]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTHOLDDOWNCLAMP2]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTHOLDDOWNCLAMP2]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");

                double eOverLength, bOverLength;
                Collection<object> collection = new Collection<object>();
                bool value = GenericHelper.GetDataByRule("HgrSupStructOffset", (componentDictionary[HGRBEAM]), out collection);
                double lugOffset = 0;
                if (collection != null)
                    lugOffset = 2 * (double)(collection[0]);
                bOverLength = eOverLength = lugOffset;
                (componentDictionary[HGRBEAM]).SetPropertyValue((bOverLength), "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[HGRBEAM]).SetPropertyValue((eOverLength), "IJUAHgrOccOverLength", "BeginOverLength");


                string strBBLow = string.Empty, strBBHigh = string.Empty;
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

                JointHelper.CreatePlanarJoint("-1", strBBLow, HGRBEAM, "BeginCap", Plane.ZX, Plane.XY, -lugOffset / 2);
                JointHelper.CreateRigidJoint("-1", strBBHigh, HGRBEAM, "EndCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -depth, lugOffset / 2, -lugOffset / 2);

                JointHelper.CreatePrismaticJoint(HGRBEAM, "BeginCap", HGRBEAM, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                JointHelper.CreatePrismaticJoint(CTHOLDDOWNCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                JointHelper.CreatePrismaticJoint(CTHOLDDOWNCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);
                JointHelper.CreatePlanarJoint(HGRBEAM, "BeginCap", CTHOLDDOWNCLAMP1, "Structure", Plane.YZ, Plane.YZ, lugOffset/2);
                JointHelper.CreatePlanarJoint(HGRBEAM, "EndCap", CTHOLDDOWNCLAMP2, "Structure", Plane.YZ, Plane.NegativeYZ, lugOffset/2);

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

                JointHelper.CreateRigidJoint(ROD_1, "RodEnd1", HGRBEAM, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X,lugOffset/2, lugOffset/2, -lugOffset/2);
                JointHelper.CreateRigidJoint(ROD_2, "RodEnd1", HGRBEAM, "EndCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, lugOffset/2, -lugOffset/2, -lugOffset/2);


            }
            // return the collection of ocpmmittedmtd joints,
            catch (Exception excepion)
            {
                Type myType = this.GetType();
                CmnException exception1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + excepion.Message, excepion);
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

                    routeConnections.Add(new ConnectionInfo(CTHOLDDOWNCLAMP1, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(CTHOLDDOWNCLAMP2, 1)); // partindex, routeindex

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