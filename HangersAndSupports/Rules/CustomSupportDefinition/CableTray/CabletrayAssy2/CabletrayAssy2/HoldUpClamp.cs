//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HoldUpClamp.cs
//   CabletrayAssy2,Ingr.SP3D.Content.Support.Rules.ClipHoldClampHoldUpClamp
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
    public class HoldUpClamp : CustomSupportDefinition
    {
        private const string CTHOLDUPCLAMP1 = "CTHOLDUPCLAMP1";
        private const string CTHOLDUPCLAMP2 = "CTHOLDUPCLAMP2";
        private const string BEAM_ATT_1 = "WELDEDBEAMATTACHMENT1";
        private const string EYE_NUT_1 = "EYENUT1";
        private const string ROD_1 = "ROD1";
        private const string BEAM_ATT_2 = "WELDEDBEAMATTACHMENT2";
        private const string EYE_NUT_2 = "EYENUT2";
        private const string ROD_2 = "ROD2"; 

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
                    PartClass ctHoldSideClamp = (PartClass)catalogBaseHelper.GetPartClass("CTHoldUpClamp");
                    string partselection = ctHoldSideClamp.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    Part g4gPart1 = supportComponentUtils.GetPartFromPartClass("G4G_1461_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize", support);
                    Part ctHoldSideClampPart1 = supportComponentUtils.GetPartFromPartClass("CTHoldUpClamp", partselection, support);
                    Part ctHoldSideClampPart2 = supportComponentUtils.GetPartFromPartClass("CTHoldUpClamp", partselection, support);
                    parts.Add(new PartInfo(CTHOLDUPCLAMP1, ctHoldSideClampPart1.ToString()));
                    parts.Add(new PartInfo(CTHOLDUPCLAMP2, ctHoldSideClampPart2.ToString()));

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
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========    
                CableTrayObjectInfo cableInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                double depth = cableInfo.Depth;
                double width = cableInfo.Width;
                double radius = cableInfo.BendRadius;
                if (width <= 0 || depth <= 0)
                {
                    width = radius * 2;
                    depth = radius * 2;
                }
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                (componentDictionary[CTHOLDUPCLAMP1]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTHOLDUPCLAMP1]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTHOLDUPCLAMP1]).SetPropertyValue(depth / 4.0, "IJUAHgrCTOffset", "TrayBeamWidth");

                (componentDictionary[CTHOLDUPCLAMP2]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTHOLDUPCLAMP2]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTHOLDUPCLAMP2]).SetPropertyValue(depth / 4.0, "IJUAHgrCTOffset", "TrayBeamWidth");

                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================


                //======================================================
                //Create Joints
                //======================================================

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
                JointHelper.CreateRigidJoint(ROD_1, "RodEnd1", CTHOLDUPCLAMP1, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

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

                //Add a Rigid Joint between cable tray Clamp and Bottom of Rod
                JointHelper.CreateRigidJoint(ROD_2, "RodEnd1", CTHOLDUPCLAMP2, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                //Add a Rigid Joint between clamp and Cable Tray
                JointHelper.CreateRigidJoint(CTHOLDUPCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(CTHOLDUPCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);

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

                    routeConnections.Add(new ConnectionInfo(CTHOLDUPCLAMP1, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(CTHOLDUPCLAMP2, 1)); // partindex, routeindex

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