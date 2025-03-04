//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ClipHoldClamp.cs
//   CabletrayAssy,Ingr.SP3D.Content.Support.Rules.ClipHoldClamp
//   Author       :  MK
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013       MK      CR-CP-224477 - Converted CabletrayAssemblies to C# .Net
//  30/04/2015       Chethan   TR-CP-271643  Rec exc. minidumps created while placing support by ref. on cabletray 
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
        private const string G4G1461011 = "G4G1461011";   //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1461012 = "G4G1461012";   //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1460011 = "G4G1460011";   //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)
        private const string G4G1460012 = "G4G1460012";   //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)

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
                    PartClass g4G146101 = (PartClass)catalogBaseHelper.GetPartClass("G4G_1460_01");

                    Part g4GPart2 = null;
                    string g4gPart2Part = string.Empty;
                    foreach (BusinessObject businessObject in g4G146101.Parts)
                    {
                        if (((double)((PropertyValueDouble)businessObject.GetPropertyValue("IJUAHgrDia", "RodSizeDia")).PropValue) >= (((double)((PropertyValueDouble)clipHoldClampPart1.GetPropertyValue("IJUAHgrCTClipHdClamp", "RodDiameter")).PropValue)))
                        {
                            g4GPart2 = (Part)businessObject;
                            g4gPart2Part = g4GPart2.ToString();
                            break;
                        }
                    }
                    parts.Add(new PartInfo(CTCLIPHOLDCLAMP1, clipHoldClampPart1.ToString()));
                    parts.Add(new PartInfo(CTCLIPHOLDCLAMP2, clipHoldClampPart2.ToString()));
                    parts.Add(new PartInfo(G4G1461011, g4GPart1.ToString()));
                    parts.Add(new PartInfo(G4G1461012, g4GPart1.ToString()));
                    parts.Add(new PartInfo(G4G1460011, g4gPart2Part));
                    parts.Add(new PartInfo(G4G1460012, g4gPart2Part));

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
                double beamWidth;
                if (width <= 0 || depth <= 0)
                {
                    width = radius * 2;
                    depth = radius * 2;
                }

                if (SupportHelper.SupportingObjects.Count>0)
                {
                    beamWidth = SupportingHelper.SupportingObjectInfo(1).Width;
                }
                else
                {
                    beamWidth = 0;
                }


                // If unable to retrieve BeamWidth, Structure could be a slab.
                // Use the cable tray width to set the clamp dimensions.
                if (beamWidth <= 0)
                    beamWidth = width;
                //Set occurance cable tray width
                double thickness = depth / 20;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                (componentDictionary[CTCLIPHOLDCLAMP1]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTCLIPHOLDCLAMP1]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTCLIPHOLDCLAMP1]).SetPropertyValue(thickness, "IJUAHgrCTOffset", "TrayThickness");

                (componentDictionary[CTCLIPHOLDCLAMP2]).SetPropertyValue(width, "IJUAHgrCTOffset", "TrayWidth");
                (componentDictionary[CTCLIPHOLDCLAMP2]).SetPropertyValue(depth, "IJUAHgrCTOffset", "TrayDepth");
                (componentDictionary[CTCLIPHOLDCLAMP2]).SetPropertyValue(thickness, "IJUAHgrCTOffset", "TrayThickness");

                (componentDictionary[G4G1461011]).SetPropertyValue((0.75 * beamWidth), "IJUAHgrOccGeometry", "Width");
                (componentDictionary[G4G1461012]).SetPropertyValue((0.75 * beamWidth), "IJUAHgrOccGeometry", "Width");


                JointHelper.CreateCylindricalJoint(G4G1460011, "RodTop", G4G1460011, "RodBottom", Axis.Z, Axis.Z, 0);
                JointHelper.CreateCylindricalJoint(G4G1460012, "RodTop", G4G1460012, "RodBottom", Axis.Z, Axis.Z, 0);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreatePrismaticJoint(CTCLIPHOLDCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    JointHelper.CreatePrismaticJoint(CTCLIPHOLDCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);
                }
                else
                {
                    JointHelper.CreateRigidJoint(CTCLIPHOLDCLAMP1, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(CTCLIPHOLDCLAMP2, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                }
                JointHelper.CreateSphericalJoint(CTCLIPHOLDCLAMP1, "Structure", G4G1460011, "RodBottom");
                JointHelper.CreateSphericalJoint(CTCLIPHOLDCLAMP2, "Structure", G4G1460012, "RodBottom");

                JointHelper.CreateSphericalJoint(G4G1460011, "RodTop", G4G1461011, "Rod");
                JointHelper.CreateSphericalJoint(G4G1460012, "RodTop", G4G1461012, "Rod");

                JointHelper.CreateGlobalAxesAlignedJoint(G4G1460011, "RodBottom", Axis.Z, Axis.Z);
                JointHelper.CreateGlobalAxesAlignedJoint(G4G1460012, "RodBottom", Axis.Z, Axis.Z);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreatePrismaticJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0);
                    JointHelper.CreatePrismaticJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeX, 0, 0);
                }
                else
                {
                    JointHelper.CreateTranslationalJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                    JointHelper.CreateTranslationalJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                }

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

                    structConnections.Add(new ConnectionInfo(G4G1461011, 1)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(G4G1461012, 1)); // partindex, routeindex


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
