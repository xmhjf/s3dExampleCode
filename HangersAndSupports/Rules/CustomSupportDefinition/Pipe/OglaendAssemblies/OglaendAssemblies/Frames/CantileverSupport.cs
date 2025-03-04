//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   CantileverSupport.cs
//   OglaendAssemblies,Ingr.SP3D.Content.Support.Rules.CantileverSupport
//   Author       :  Durga Prasad
//   Creation Date:  18-12-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET  
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
    public class CantileverSupport : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        #region Global Constants, Fields and properties
        //Part Keys
        private
        const string SECTION = "SECTION";
        const string WELD_STARTER = "WELD_STARTER";
        const string END_PROTECTION = "END_PROTECTION";
        const string STRUCT_CONN = "STRUCT_CONN";

        //Steel Members - Stores Data related to steel member (Part Number, Depth, Width etc...)
        OglaendAssemblyServices.HSSteelMember section;

        //Fields
        string sectionName, weldStarter, endProtection;
        #endregion

        #region Get Assembly Catalog Parts
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support supportOcc = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //Get Support Attributes
                    sectionName = (string)((PropertyValueString)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGHorSection", "HorSection")).PropValue;
                    weldStarter = (string)((PropertyValueString)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGWeldedStarter", "WeldedStarter")).PropValue;
                    endProtection = (string)((PropertyValueString)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGEndProtection", "EndProtection")).PropValue;

                    //Add Parts
                    OglaendAssemblyServices.AddPart(this, SECTION, sectionName, parts);
                    OglaendAssemblyServices.AddPart(this, WELD_STARTER, weldStarter, parts);
                    OglaendAssemblyServices.AddPart(this, END_PROTECTION, endProtection, parts);
                    OglaendAssemblyServices.AddPart(this, STRUCT_CONN, "Log_Conn_Part_1", parts);

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
        #endregion

        #region Get Assembly Implied Parts
        //-----------------------------------------------------------------------------------
        //Get Assembly Implied Parts
        //-----------------------------------------------------------------------------------
        public override ReadOnlyCollection<PartInfo> ImpliedParts
        {
            get
            {
                try
                {
                    SP3D.Support.Middle.Support supportOcc = SupportHelper.Support;
                    Collection<PartInfo> impliedParts = new Collection<PartInfo>();

                    string supportpartnumber = supportOcc.SupportDefinition.PartNumber;
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
        #endregion

        #region Get Assembly Joints
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                SP3D.Support.Middle.Support supportOcc = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                BusinessObject supportPart = supportOcc.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                // Get the Steel Cross Section Data
                section = OglaendAssemblyServices.GetSectionDataFromPartIndex(this, SECTION);

                //Load standard bounding box definition
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                //=========================
                //1. Get bounding box boundary objects dimension information
                //=========================

                boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                //====== ======
                //2. Retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry
                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | boundingBoxHeight
                // |____________________|
                //    boundingBoxWidth

                double boundingBoxWidth = boundingBox.Width;
                double boundingBoxHeight = boundingBox.Height;

                //Get and Set Support/Part Attributes
                double offset = 0, totalOffset = 0, cantileverLength = 0;
                double weldStarterLength = 0, weldedStarterFlangeWidth = 0, weldedStarterThickness = 0;
                double cantileverRotationAngle = 0;
                double structureIntPosZ = 0;
                double endProtectionThickness = 0;
                double offsetX = 0, offsetY = 0;

                offset = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJOAhsOGOffset", "Offset")).PropValue;
                cantileverLength = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGHSLength", "HSLength")).PropValue;

                PropertyValueCodelist cardinalPoint7Section = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCPFlexPort", "CP7");
                PropertyValueCodelist cardinalPoint6Section = (PropertyValueCodelist)componentDictionary[SECTION].GetPropertyValue("IJOAhsSteelCP", "CP6");

                if (section.sectionType == "Oglaend_Tri_StrSec" || section.sectionType == "Oglaend_L_StrSec")
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)5, "IJOAhsSteelCPFlexPort", "CP7");
                    componentDictionary[SECTION].SetPropertyValue((double)0.5, "IJUAHgrOccLength", "Length");
                    componentDictionary[SECTION].SetPropertyValue((double)Math.PI / 2, "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)1, "IJOAhsSteelCP", "CP6");
                    componentDictionary[WELD_STARTER].SetPropertyValue((double)0, "IJUAHgrOGStructureIntPort", "StuctureIntPosZ");

                    weldStarterLength = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELD_STARTER], "IJUAHgrOGLength", "Length")).PropValue;
                    structureIntPosZ = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELD_STARTER], "IJUAHgrOGStructureIntPort", "StuctureIntPosZ")).PropValue;
                    weldedStarterFlangeWidth = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELD_STARTER], "IJUAHgrOGFlangeWidth", "FlangeWidth")).PropValue;
                    weldedStarterThickness = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELD_STARTER], "IJUAHgrOGThickness", "Thickness")).PropValue;

                    offsetX = (weldStarterLength/2) - structureIntPosZ;
                    offsetY = 0;
                }
                else if (section.sectionType == "Oglaend_C_StrSec")
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)5, "IJOAhsSteelCPFlexPort", "CP7");
                    componentDictionary[SECTION].SetPropertyValue((double)0.5, "IJUAHgrOccLength", "Length");
                    componentDictionary[SECTION].SetPropertyValue((double)0, "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)5, "IJOAhsSteelCP", "CP6");

                    weldStarterLength = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELD_STARTER], "IJUAHgrOGLength", "Length")).PropValue;
                    weldedStarterThickness = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELD_STARTER], "IJUAHgrOGThickness", "Thickness")).PropValue;
                    endProtectionThickness = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[END_PROTECTION], "IJUAhsThickness1", "Thickness1")).PropValue;

                    offsetX = weldStarterLength/2;
                    offsetY = 0.00175;
                }
                else
                {
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint7Section.PropValue = (int)5, "IJOAhsSteelCPFlexPort", "CP7");
                    componentDictionary[SECTION].SetPropertyValue((double)0.5, "IJUAHgrOccLength", "Length");
                    componentDictionary[SECTION].SetPropertyValue((double)0, "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[SECTION].SetPropertyValue(cardinalPoint6Section.PropValue = (int)5, "IJOAhsSteelCP", "CP6");

                    weldStarterLength = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELD_STARTER], "IJUAHgrOGLength", "Length")).PropValue;
                    weldedStarterThickness = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[WELD_STARTER], "IJUAHgrOGThickness", "Thickness")).PropValue;
                    endProtectionThickness = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(componentDictionary[END_PROTECTION], "IJUAhsThickness1", "Thickness1")).PropValue;

                    offsetX = weldStarterLength/2;
                    offsetY = weldedStarterThickness/2;
                }

                //Calculate Ports Orientation Angle
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);

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

                if (supportingType == "Steel")
                {
                    if ((routeStructAngle > (Math.PI / 4) && routeStructAngle < (3 * (Math.PI / 4))) || (routeStructAngle > (5 * (Math.PI / 4)) && routeStructAngle < (7 * (Math.PI / 4))))
                    {
                        if (HgrCompareDoubleService.cmpdbl(offset, 0) == true)
                            supportOcc.SetPropertyValue(boundingBoxWidth / 2, "IJOAhsOGOffset", "Offset");
                        else
                            offset = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJOAhsOGOffset", "Offset")).PropValue;

                        if (HgrCompareDoubleService.cmpdbl(cantileverLength, 0) == true)
                        {
                            cantileverLength = RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Structure", PortAxisType.Z);
                            supportOcc.SetPropertyValue(cantileverLength, "IJUAhsOGHSLength", "HSLength");
                        }
                        else
                            cantileverLength = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGHSLength", "HSLength")).PropValue;

                        //Change Support Configurations
                        if (section.sectionType == "Oglaend_Tri_StrSec" || section.sectionType == "Oglaend_L_StrSec")
                        {
                            switch (Configuration)
                            {
                                case 1:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset + ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) - weldedStarterThickness;
                                    break;
                                case 2:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset + ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) - weldedStarterThickness;
                                    break;
                                case 3:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = -offset - ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) + weldedStarterThickness;
                                    break;
                                case 4:
                                    cantileverRotationAngle = 0;
                                    totalOffset = -offset - ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) + weldedStarterThickness;
                                    break;
                                case 5:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset + ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) - weldedStarterThickness;
                                    break;
                                case 6:
                                    cantileverRotationAngle =  Math.PI;
                                    totalOffset = offset + ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) - weldedStarterThickness;
                                    break;
                                case 7:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = -offset - ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) + weldedStarterThickness;
                                    break;
                                case 8:
                                    cantileverRotationAngle = 0;
                                    totalOffset = -offset - ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) + weldedStarterThickness;
                                    break;
                            }
                        }
                        else if (section.sectionType == "Oglaend_C_StrSec")
                        {
                            switch (Configuration)
                            {
                                case 1:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset + (section.depth/2);
                                    break;
                                case 2:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset + (section.width / 2) - offsetY;
                                    break;
                                case 3:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = offset + (section.depth / 2);
                                    break;
                                case 4:
                                    cantileverRotationAngle = 0;
                                    totalOffset = offset + (section.width / 2) + offsetY;
                                    break;
                                case 5:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = -offset - (section.depth / 2);
                                    break;
                                case 6:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = -offset - (section.width / 2) - offsetY;
                                    break;
                                case 7:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = -offset - (section.depth / 2);
                                    break;
                                case 8:
                                    cantileverRotationAngle = 0;
                                    totalOffset = -offset - (section.width / 2) + offsetY;
                                    break;
                            }
                        }
                        else
                        {
                            switch (Configuration)
                            {
                                case 1:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset + (section.depth / 2);
                                    break;
                                case 2:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset + (section.width / 2) - (weldedStarterThickness/2);
                                    break;
                                case 3:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = -offset - (section.depth / 2);
                                    break;
                                case 4:
                                    cantileverRotationAngle = 0;
                                    totalOffset = -offset - (section.width / 2) + (weldedStarterThickness / 2);
                                    break;
                                case 5:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset + (section.depth / 2);
                                    break;
                                case 6:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset + (section.width / 2) - (weldedStarterThickness / 2);
                                    break;
                                case 7:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = -offset - (section.depth / 2);
                                    break;
                                case 8:
                                    cantileverRotationAngle = 0;
                                    totalOffset = -offset - (section.width / 2) + (weldedStarterThickness / 2);
                                    break;
                            }
                        }

                        //Add Joints
                        if (section.sectionType == "Oglaend_Tri_StrSec" || section.sectionType == "Oglaend_L_StrSec")
                        {
                            JointHelper.CreateRigidJoint(SECTION, "BeginFlex", WELD_STARTER, "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint(SECTION, "BeginFlex", WELD_STARTER, "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -offsetX, 0, -offsetY);
                        }

                        JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, totalOffset);
                        JointHelper.CreateAngularRigidJoint(WELD_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, cantileverRotationAngle));
                    }
                    else
                    {
                        if (HgrCompareDoubleService.cmpdbl(offset, 0) == true)
                            supportOcc.SetPropertyValue((double)0, "IJOAhsOGOffset", "Offset");
                        else
                            offset = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJOAhsOGOffset", "Offset")).PropValue;

                        if (HgrCompareDoubleService.cmpdbl(cantileverLength, 0) == true)
                        {
                            cantileverLength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z);
                            supportOcc.SetPropertyValue(cantileverLength, "IJUAhsOGHSLength", "HSLength");
                        }
                        else
                            cantileverLength = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGHSLength", "HSLength")).PropValue;

                        //Get Distance Between Ports
                        double projectionY = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.Y);
                        double projectionZ = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortAxisType.Z);

                        //Get Normal Vector along the Surface
                        Matrix4X4 hangerPort = RefPortHelper.PortLCS("Structure");
                        Vector structureZVector = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);

                        //Get the Vector Normal to the Structure Face (May Not be the Z Axis in certain Cases)
                        IPlane iplane = (IPlane)supportOcc.SupportingFaces[0];
                        Vector structureFaceNormal = new Vector(iplane.Normal.X, iplane.Normal.Y, iplane.Normal.Z);

                        //Change Support Configurations
                        if (section.sectionType == "Oglaend_Tri_StrSec" || section.sectionType == "Oglaend_L_StrSec")
                        {
                            switch (Configuration)
                            {
                                case 1:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 2:
                                    cantileverRotationAngle = 0;
                                    totalOffset = offset;
                                    break;
                                case 3:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 4:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset;
                                    break;
                                case 5:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 6:
                                    cantileverRotationAngle = 0;
                                    totalOffset = offset;
                                    break;
                                case 7:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 8:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset;
                                    break;
                            }
                        }
                        else if (section.sectionType == "Oglaend_C_StrSec")
                        {
                            switch (Configuration)
                            {
                                case 1:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 2:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset;
                                    break;
                                case 3:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 4:
                                    cantileverRotationAngle = 0;
                                    totalOffset = offset;
                                    break;
                                case 5:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 6:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset;
                                    break;
                                case 7:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 8:
                                    cantileverRotationAngle = 0;
                                    totalOffset = offset;
                                    break;
                            }
                        }
                        else
                        {
                            switch (Configuration)
                            {
                                case 1:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 2:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset;
                                    break;
                                case 3:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 4:
                                    cantileverRotationAngle = 0;
                                    totalOffset = offset;
                                    break;
                                case 5:
                                    cantileverRotationAngle = Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 6:
                                    cantileverRotationAngle = Math.PI;
                                    totalOffset = offset;
                                    break;
                                case 7:
                                    cantileverRotationAngle = 3 * Math.PI / 2;
                                    totalOffset = offset;
                                    break;
                                case 8:
                                    cantileverRotationAngle = 0;
                                    totalOffset = offset;
                                    break;
                            }
                        }

                        if (Math.Abs(projectionZ) > Math.Abs(projectionY))
                        {
                            if (OglaendAssemblyServices.AngleBetweenVectors(structureFaceNormal, structureZVector) < Math.PI / 2)
                            {
                                if (projectionY > 0)
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, totalOffset, 0);
                                }
                                else
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -totalOffset, 0);
                                }
                            }
                            else
                            {
                                if (projectionY > 0)
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, totalOffset, 0);
                                }
                                else
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -totalOffset, 0);
                                }
                            }
                        }
                        else
                        {
                            if (OglaendAssemblyServices.AngleBetweenVectors(structureFaceNormal, structureZVector) < Math.PI / 2)
                            {
                                if (projectionY > 0)
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, totalOffset, 0);
                                }
                                else
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, totalOffset, 0);
                                }
                            }
                            else
                            {
                                if (projectionY > 0)
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -totalOffset, 0);
                                }
                                else
                                {
                                    //Add Joint between Structure and Connection Object
                                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, -totalOffset, 0);
                                }
                            }
                        }

                        //Add Remaining joints
                        if (section.sectionType == "Oglaend_Tri_StrSec" || section.sectionType == "Oglaend_L_StrSec")
                        {
                            JointHelper.CreateRigidJoint(SECTION, "BeginFlex", WELD_STARTER, "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint(SECTION, "BeginFlex", WELD_STARTER, "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -offsetX, 0, -offsetY);
                        }

                        JointHelper.CreateAngularRigidJoint(WELD_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, (-Math.PI/2) + cantileverRotationAngle));
                    }
                }
                else
                {
                    if (HgrCompareDoubleService.cmpdbl(offset, 0) == true)
                        supportOcc.SetPropertyValue(boundingBoxWidth / 2, "IJOAhsOGOffset", "Offset");
                    else
                        offset = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJOAhsOGOffset", "Offset")).PropValue;

                    if (HgrCompareDoubleService.cmpdbl(cantileverLength, 0) == true)
                    {
                        cantileverLength = RefPortHelper.DistanceBetweenPorts("BBR_Low", "Structure", PortAxisType.Z);
                        supportOcc.SetPropertyValue(cantileverLength, "IJUAhsOGHSLength", "HSLength");
                    }
                    else
                        cantileverLength = (double)((PropertyValueDouble)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsOGHSLength", "HSLength")).PropValue;

                    //Change Support Configurations
                    if (section.sectionType == "Oglaend_Tri_StrSec" || section.sectionType == "Oglaend_L_StrSec")
                    {
                        switch (Configuration)
                        {
                            case 1:
                                cantileverRotationAngle = Math.PI / 2;
                                totalOffset = offset + ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) - weldedStarterThickness;
                                break;
                            case 2:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = offset + ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) - weldedStarterThickness;
                                break;
                            case 3:
                                cantileverRotationAngle = 3 * Math.PI / 2;
                                totalOffset = -offset - ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) + weldedStarterThickness;
                                break;
                            case 4:
                                cantileverRotationAngle = 0;
                                totalOffset = -offset - ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) + weldedStarterThickness;
                                break;
                            case 5:
                                cantileverRotationAngle = Math.PI / 2;
                                totalOffset = offset + ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) - weldedStarterThickness;
                                break;
                            case 6:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = offset + ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) - weldedStarterThickness;
                                break;
                            case 7:
                                cantileverRotationAngle = 3 * Math.PI / 2;
                                totalOffset = -offset - ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) + weldedStarterThickness;
                                break;
                            case 8:
                                cantileverRotationAngle = 0;
                                totalOffset = -offset - ((weldedStarterFlangeWidth - weldedStarterThickness) / 2) + weldedStarterThickness;
                                break;
                        }
                    }
                    else if (section.sectionType == "Oglaend_C_StrSec")
                    {
                        switch (Configuration)
                        {
                            case 1:
                                cantileverRotationAngle = Math.PI / 2;
                                totalOffset = offset + (section.depth / 2);
                                break;
                            case 2:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = offset + (section.width / 2) - offsetY;
                                break;
                            case 3:
                                cantileverRotationAngle = 3 * Math.PI / 2;
                                totalOffset = offset + (section.depth / 2);
                                break;
                            case 4:
                                cantileverRotationAngle = 0;
                                totalOffset = offset + (section.width / 2) + offsetY;
                                break;
                            case 5:
                                cantileverRotationAngle = Math.PI / 2;
                                totalOffset = -offset - (section.depth / 2);
                                break;
                            case 6:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = -offset - (section.width / 2) - offsetY;
                                break;
                            case 7:
                                cantileverRotationAngle = 3 * Math.PI / 2;
                                totalOffset = -offset - (section.depth / 2);
                                break;
                            case 8:
                                cantileverRotationAngle = 0;
                                totalOffset = -offset - (section.width / 2) + offsetY;
                                break;
                        }
                    }
                    else
                    {
                        switch (Configuration)
                        {
                            case 1:
                                cantileverRotationAngle = Math.PI / 2;
                                totalOffset = offset + (section.depth / 2);
                                break;
                            case 2:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = offset + (section.width / 2) - (weldedStarterThickness / 2);
                                break;
                            case 3:
                                cantileverRotationAngle = 3 * Math.PI / 2;
                                totalOffset = -offset - (section.depth / 2);
                                break;
                            case 4:
                                cantileverRotationAngle = 0;
                                totalOffset = -offset - (section.width / 2) + (weldedStarterThickness / 2);
                                break;
                            case 5:
                                cantileverRotationAngle = Math.PI / 2;
                                totalOffset = offset + (section.depth / 2);
                                break;
                            case 6:
                                cantileverRotationAngle = Math.PI;
                                totalOffset = offset + (section.width / 2) - (weldedStarterThickness / 2);
                                break;
                            case 7:
                                cantileverRotationAngle = 3 * Math.PI / 2;
                                totalOffset = -offset - (section.depth / 2);
                                break;
                            case 8:
                                cantileverRotationAngle = 0;
                                totalOffset = -offset - (section.width / 2) + (weldedStarterThickness / 2);
                                break;
                        }
                    }

                    //Add Joints
                    if (section.sectionType == "Oglaend_Tri_StrSec" || section.sectionType == "Oglaend_L_StrSec")
                    {
                        JointHelper.CreateRigidJoint(SECTION, "BeginFlex", WELD_STARTER, "StructureInt", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateAngularRigidJoint(WELD_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, cantileverRotationAngle - (Math.PI / 2)));
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(SECTION, "BeginFlex", WELD_STARTER, "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -offsetX, 0, -offsetY);
                        JointHelper.CreateAngularRigidJoint(WELD_STARTER, "Structure", STRUCT_CONN, "Connection", new Vector(0, 0, 0), new Vector(Math.PI, 0, cantileverRotationAngle));
                    }

                    JointHelper.CreateRigidJoint(STRUCT_CONN, "Connection", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, totalOffset, 0);
                    
                }
                if(componentDictionary.ContainsKey(END_PROTECTION))
                {
                    JointHelper.CreateRigidJoint(SECTION, "EndFlex", END_PROTECTION, "Port1", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                }

                //Check for Maximum and Minimum Permissible Values and Capacity Values
                double capacityDWPy = 0, pipingPy = 0, pipingPx = 0, blastPy = 0, blastPx = 0;
                double maximumLength = 0, minimumLength = 0;
                double rCantileverLength = 0;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass maxValuesServiceClass = (PartClass)catalogBaseHelper.GetPartClass("HS_OGAssy_MaxValues");
                ReadOnlyCollection<BusinessObject> maxValuesClassItems = maxValuesServiceClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                string supportNumber = (string)((PropertyValueString)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsSupportNumber", "SupportNumber")).PropValue;

                foreach (BusinessObject classItem in maxValuesClassItems)
                {
                    bool isEqual = String.Equals(supportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGMaxValue", "SupportNumber")).PropValue), StringComparison.Ordinal);
                    if (isEqual == true)
                    {
                        maximumLength = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGMaxValue", "MaximumLength")).PropValue);
                        break;
                    }
                }

                PartClass minValuesServiceClass = (PartClass)catalogBaseHelper.GetPartClass("HS_OGAssy_MinValues");
                ReadOnlyCollection<BusinessObject> minValuesClassItems = minValuesServiceClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in minValuesClassItems)
                {
                    bool isEqual = String.Equals(supportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGMinValue", "SupportNumber")).PropValue), StringComparison.Ordinal);
                    if (isEqual == true)
                    {
                        minimumLength = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGMinValue", "MinimumLength")).PropValue);
                        break;
                    }
                }

                if (cantileverLength > maximumLength)
                {
                    cantileverLength = maximumLength;

                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxLength,"Total Calculated Length of Cantilever is greater than the Maximum allowable length.Hence,resetting the length to maximum length"), "", "CantileverSupport", 696);
                    supportOcc.SetPropertyValue(cantileverLength, "IJUAhsOGHSLength", "HSLength");
                }

                if (cantileverLength < minimumLength)
                {
                    cantileverLength = minimumLength;

                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", OglaendAssembliesLocalizer.GetString(OglaendAssembliesResourceIDs.ErrInvalidMinMaxLength,"Total Calculated Length of Cantilever is less than the Minimum allowable length.Hence,resetting the length to minimum length."), "", "CantileverSupport", 696);
                    supportOcc.SetPropertyValue(cantileverLength, "IJUAhsOGHSLength", "HSLength");
                }

                rCantileverLength = Math.Round(cantileverLength * 1000, 0);

                PartClass loadTableServiceClass = (PartClass)catalogBaseHelper.GetPartClass("HS_OGAssy_LoadTable");
                ReadOnlyCollection<BusinessObject> loadTableClassItems = loadTableServiceClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                foreach (BusinessObject classItem in loadTableClassItems)
                {
                    double minLength = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "MinLength")).PropValue);
                    double maxLength = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "MaxLength")).PropValue);

                    bool isEqual = String.Equals(supportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsOGLoadTable", "SupportNumber")).PropValue), StringComparison.Ordinal);
                    if ((isEqual == true) && (rCantileverLength > minLength && rCantileverLength < maxLength + 0.00001))
                    {
                        capacityDWPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "CapacityDWPy")).PropValue);
                        pipingPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "PipingPy")).PropValue);
                        pipingPx = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "PipingPx")).PropValue);
                        blastPy = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "BlastPy")).PropValue);
                        blastPx = ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAhsOGLoadTable", "BlastPx")).PropValue);
                        break;
                    }

                }
                //set support properties
                componentDictionary[SECTION].SetPropertyValue(cantileverLength-offsetX-endProtectionThickness, "IJUAHgrOccLength", "Length");

                if (supportOcc.SupportsInterface("IJUAhsOGMaxLength"))
                    supportOcc.SetPropertyValue(maximumLength, "IJUAhsOGMaxLength", "MaxLength");
                if (supportOcc.SupportsInterface("IJUAhsOGMinLength"))
                    supportOcc.SetPropertyValue(minimumLength, "IJUAhsOGMinLength", "MinLength");
                if (supportOcc.SupportsInterface("IJUAhsOGCapacityDWPy"))
                    supportOcc.SetPropertyValue(capacityDWPy, "IJUAhsOGCapacityDWPy", "CapacityDWPy");
                if (supportOcc.SupportsInterface("IJUAhsOGPipingPy"))
                    supportOcc.SetPropertyValue(pipingPy, "IJUAhsOGPipingPy", "PipingPy");
                if (supportOcc.SupportsInterface("IJUAhsOGPipingPx"))
                    supportOcc.SetPropertyValue(pipingPx, "IJUAhsOGPipingPx", "PipingPx");
                if (supportOcc.SupportsInterface("IJUAhsOGBlastPy"))
                    supportOcc.SetPropertyValue(blastPy, "IJUAhsOGBlastPy", "BlastPy");
                if (supportOcc.SupportsInterface("IJUAhsOGBlastPx"))
                    supportOcc.SetPropertyValue(blastPx, "IJUAhsOGBlastPx", "BlastPx");

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion

        #region ConfigurationCount
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 8;
            }
        }
        #endregion

        #region Get Route Connections
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
        #endregion

        #region Get Struct Connections
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

                    structConnections.Add(new ConnectionInfo(WELD_STARTER, 1)); // partindex, routeindex

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
        #endregion

        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject supportOcc)
        {
            string bomDesciption = "";
            try
            {
                string supportNumber = (string)((PropertyValueString)OglaendAssemblyServices.GetAttributeValue(supportOcc, "IJUAhsSupportNumber", "SupportNumber")).PropValue;
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
