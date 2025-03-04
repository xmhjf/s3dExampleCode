//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ClipHoldClamp.cs
//   CabletrayAssy2,Ingr.SP3D.Content.Support.Rules.ClipHoldClampClipHoldClamp
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
    public class ClipHoldClamp : CustomSupportDefinition
    {
        private const string CTCLIPHOLDCLAMP1 = "CTCLIPHOLDCLAMP1";
        private const string CTCLIPHOLDCLAMP2 = "CTCLIPHOLDCLAMP2";
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
                    PartClass ctClipHoldClamp = (PartClass)catalogBaseHelper.GetPartClass("CTClipHoldClamp");
                    string partSelection = ctClipHoldClamp.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();

                    //Create the list of part classes required by the type
                    Part g4GPart1 = supportComponentUtils.GetPartFromPartClass("G4G_1461_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize", support);
                    Part clipHoldClampPart1 = supportComponentUtils.GetPartFromPartClass("CTClipHoldClamp", partSelection, support);
                    Part clipHoldClampPart2 = supportComponentUtils.GetPartFromPartClass("CTClipHoldClamp", partSelection, support);

                    parts.Add(new PartInfo(CTCLIPHOLDCLAMP1, clipHoldClampPart1.ToString()));
                    parts.Add(new PartInfo(CTCLIPHOLDCLAMP2, clipHoldClampPart2.ToString()));
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
                CableTrayObjectInfo cableInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                double depth = cableInfo.Depth;
                double width = cableInfo.Width;
                double radius = cableInfo.BendRadius;
                if (width <= 0 || depth <= 0)
                {
                    width = radius * 2;
                    depth = radius * 2;
                }

                //Set occurance cable tray width
                double thickness = depth / 20;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                (componentDictionary[CTCLIPHOLDCLAMP1]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTCLIPHOLDCLAMP1]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTCLIPHOLDCLAMP1]).SetPropertyValue(thickness, "IJUAHgrCTOffset", "TrayThickness");

                (componentDictionary[CTCLIPHOLDCLAMP2]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTCLIPHOLDCLAMP2]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTCLIPHOLDCLAMP2]).SetPropertyValue(thickness, "IJUAHgrCTOffset", "TrayThickness");


                JointHelper.CreateRigidJoint(CTCLIPHOLDCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(CTCLIPHOLDCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                JointHelper.CreateRigidJoint(CTCLIPHOLDCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(CTCLIPHOLDCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);

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
                JointHelper.CreateRigidJoint(ROD_1, "RodEnd1", CTCLIPHOLDCLAMP1, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

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
                JointHelper.CreateRigidJoint(ROD_2, "RodEnd1", CTCLIPHOLDCLAMP2, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

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

                    routeConnections.Add(new ConnectionInfo(CTCLIPHOLDCLAMP1, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(CTCLIPHOLDCLAMP2, 1)); // partindex, routeindex

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
