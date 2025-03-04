//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   HoldDownClampHori.cs
//   CabletrayAssy,Ingr.SP3D.Content.Support.Rules.HoldDownClampHori
//   Author       :  MK
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013       MK        CR-CP-224477 - Converted CabletrayAssemblies to C# .Net
//  22-Jan-2015     PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report
//  30/04/2015       Chethan   TR-CP-271643  Rec exc. minidumps created while placing support by ref. on cabletray 
//  16-Jul-2015      PVK       Resolve coverity issues found in July 2015 report
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
        private const string G4G1461011 = "G4G1461011"; //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1461012 = "G4G1461012"; //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
        private const string G4G1460011 = "G4G1460011"; //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)
        private const string G4G1460012 = "G4G1460012"; //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)
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
                    partClass[hgrBeam] = "HgrBeam";

                  
                    parts.Add(new PartInfo(G4G1461011, "G4G_1461_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize"));                   
                    parts.Add(new PartInfo(G4G1461012, "G4G_1461_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize"));                    
                    parts.Add(new PartInfo(G4G1460011, "G4G_1460_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize"));                    
                    parts.Add(new PartInfo(G4G1460012, "G4G_1460_01", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByCTSize"));
                    for (int i = clampBegin; i <= numOfPart; i++)
                    {
                        part[i] = "part" + i;
                        Part FlatPlate = supportComponentUtils.GetPartFromPartClass(partClass[i], "", support);
                        parts.Add(new PartInfo(part[i], FlatPlate.ToString()));
                    }
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

                double beamWidth;

                if (SupportHelper.SupportingObjects.Count>0)
                {
                    beamWidth = 0.5*SupportingHelper.SupportingObjectInfo(1).Width;
                }
                else
                {
                    beamWidth = 0;
                }
                // If unable to retrieve BeamWidth, Structure could be a slab.
                // Use the cable tray width to set the clamp dimensions.
                if (beamWidth <= 0)
                    beamWidth = 0.5 * width;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                (componentDictionary[G4G1461011]).SetPropertyValue(beamWidth, "IJUAHgrOccGeometry", "Width");
                (componentDictionary[G4G1461012]).SetPropertyValue(beamWidth, "IJUAHgrOccGeometry", "Width");

                for (int i = 1; i <= iNumRoutes; i++)
                {
                    (componentDictionary[part[clampBegin]]).SetPropertyValue(ctwidth[i], "IJUAHgrCTOffset", "TrayWidth");
                    (componentDictionary[part[clampBegin]]).SetPropertyValue(ctdepth[i], "IJUAHgrCTOffset", "TrayDepth");
                    (componentDictionary[part[clampBegin + 1]]).SetPropertyValue(ctwidth[i], "IJUAHgrCTOffset", "TrayWidth");
                    (componentDictionary[part[clampBegin + 1]]).SetPropertyValue(ctdepth[i], "IJUAHgrCTOffset", "TrayDepth");
                }
                double eOverLength, bOverLength;                
                Collection<object> colllection = new Collection<object>();
                bool value = GenericHelper.GetDataByRule("HgrSupStructOffset", (componentDictionary[part[hgrBeam]]), out colllection);
                double offset = 0;
                if(colllection!=null)
                    offset= (double)(colllection[0]);
                double lugOffset = 0;
                lugOffset = 2 * offset;
                bOverLength = eOverLength = lugOffset;
                (componentDictionary[part[hgrBeam]]).SetPropertyValue((bOverLength), "IJUAHgrOccOverLength", "EndOverLength");
                (componentDictionary[part[hgrBeam]]).SetPropertyValue((eOverLength), "IJUAHgrOccOverLength", "BeginOverLength");


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
                if (colllection != null)
                    endOffset = dWidth / 2 + (double)(colllection[0]);
                //Create the Joint between the RteLow Reference Port and the HgrBeam BeginCap
                JointHelper.CreatePlanarJoint("-1", strBBLow, part[hgrBeam], "BeginCap", Plane.ZX, Plane.XY, -offset);

                //Create the Joint between the RteHigh Reference Port and the HgrBeam EndCap
                JointHelper.CreatePlanarJoint("-1", strBBHigh, part[hgrBeam], "EndCap", Plane.ZX, Plane.XY, offset);

                //Create the Joint between the igh Reference Port
                JointHelper.CreatePlanarJoint("-1", strBBLow, part[hgrBeam], "BeginCap", Plane.XY, Plane.NegativeYZ, 0);
                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    JointHelper.CreatePointOnPlaneJoint(part[hgrBeam], "Neutral", "-1", strBBLow, Plane.YZ);

                //Add a flexable Joint for HgrBeam
                JointHelper.CreatePrismaticJoint(part[hgrBeam], "BeginCap", part[hgrBeam], "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

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
                    JointHelper.CreatePointOnPlaneJoint(part[clamp], "Structure", part[hgrBeam], "Neutral", Plane.ZX);
                    JointHelper.CreatePointOnPlaneJoint(part[clamp + 1], "Structure", part[hgrBeam], "Neutral", Plane.ZX);
                    clamp = clamp + 2;
                }


                //Add a Spherical Joint between Support beam and Bottom of Rod
                JointHelper.CreateRigidJoint(G4G1460011, "RodBottom", part[hgrBeam], "Neutral", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, -0.5 * dWidth - lugOffset, 0);
                JointHelper.CreateRigidJoint(G4G1460012, "RodBottom", part[hgrBeam], "Neutral", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, 0.5 * dWidth + lugOffset, 0);

                //Add a Spherical Joint between Beam Clamp and Top of Rod
                JointHelper.CreateSphericalJoint(G4G1460011, "RodTop", G4G1461011, "Rod");
                JointHelper.CreateSphericalJoint(G4G1460012, "RodTop", G4G1461012, "Rod");

                //Create the Flexible (Cylindrical) Joint between the ports of the rod_1
                JointHelper.CreateCylindricalJoint(G4G1460011, "RodTop", G4G1460011, "RodBottom", Axis.Z, Axis.Z, 0);
                //Create the Flexible (Cylindrical) Joint between the ports of the rod_2
                JointHelper.CreateCylindricalJoint(G4G1460012, "RodTop", G4G1460012, "RodBottom", Axis.Z, Axis.Z, 0);

                //Add a Joint between the Supporting Object (Beam or Plate) and the Beam Clamp
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    //Beam Structure ... Add a Prismatic Joint between Beam and Beam Clamp
                    JointHelper.CreatePrismaticJoint(G4G1461011, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0);
                    JointHelper.CreateTranslationalJoint(G4G1461012, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                }
                else
                {
                    //Plate Structure ... Add a Translation Joint Between Plate and Beam Clamp
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