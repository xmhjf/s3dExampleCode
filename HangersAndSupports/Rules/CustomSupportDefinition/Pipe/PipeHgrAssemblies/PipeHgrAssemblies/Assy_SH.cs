//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_SH.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_SH
//   Author       : Rajeswari
//   Creation Date: 17-April-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 17-April-2013  Rajeswari CR-CP-224484 C#.Net HS_Assembly Project Creation 
// 22-Jan-2015       PVK    TR-CP-264951  Resolve coverity issues found in November 2014 report 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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

    public class Assy_SH : CustomSupportDefinition
    {
        private const string T_SECTION = "T_SECTION";//1
        private const string END_PLATE1 = "END_PLATE1";//2
        private const string END_PLATE2 = "END_PLATE2";//3

        private Double shoeLength, endplateInset, endplateClearance;
        private String sectionSize, plateSize;
        private int plates;
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

                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWTSize", "WTSize")).PropValue;
                    shoeLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssySH", "SHOE_L")).PropValue;
                    endplateInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssySH", "ENDPLATE_INSET")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlateSize", "PLATE_SIZE")).PropValue;
                    endplateClearance = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssySH", "ENDPLATE_CLEARANCE")).PropValue;
                    plates = (int)((PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssySH", "PLATES")).PropValue;

                    // Create the list of part classes required by the type
                    if (plates == 1)
                    {
                        parts.Add(new PartInfo(T_SECTION, sectionSize));
                        parts.Add(new PartInfo(END_PLATE1, "Utility_END_PLATE"));
                        parts.Add(new PartInfo(END_PLATE2, "Utility_END_PLATE"));
                    }
                    else
                        parts.Add(new PartInfo(T_SECTION, sectionSize));
                   
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                BusinessObject tSectionPart = componentDictionary[T_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];

                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter;

                Double steelWidth, steelDepth, flangeThickness, endplateT=0.0;

                CrossSection crosssection = (CrossSection)tSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth 
                steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                flangeThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                if (plates == 1)
                {
                    PropertyValueCodelist utilityCodelist = (PropertyValueCodelist)componentDictionary[END_PLATE1].GetPropertyValue("IJOAHgrUtility_END_PLATE", "THICKNESS");
                    CodelistItem codeList = utilityCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(Convert.ToInt32(plateSize));

                    endplateT =Convert.ToDouble( codeList.ShortDisplayName) * 25.4 / 1000.0;

                    componentDictionary[END_PLATE1].SetPropertyValue(utilityCodelist.PropValue, "IJOAHgrUtility_END_PLATE", "THICKNESS");
                    componentDictionary[END_PLATE1].SetPropertyValue(steelWidth, "IJOAHgrUtility_END_PLATE", "W");
                    componentDictionary[END_PLATE1].SetPropertyValue(steelDepth - flangeThickness - endplateClearance, "IJOAHgrUtility_END_PLATE", "H");

                    componentDictionary[END_PLATE2].SetPropertyValue(utilityCodelist.PropValue, "IJOAHgrUtility_END_PLATE", "THICKNESS");
                    componentDictionary[END_PLATE2].SetPropertyValue(steelWidth, "IJOAHgrUtility_END_PLATE", "W");
                    componentDictionary[END_PLATE2].SetPropertyValue(steelDepth - flangeThickness - endplateClearance, "IJOAHgrUtility_END_PLATE", "H");
                }
                componentDictionary[T_SECTION].SetPropertyValue(shoeLength, "IJUAHgrOccLength", "Length");

                // ====== ======
                // Set Values of Part Occurance Attributes
                // ====== ======
                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[T_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[T_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;

                componentDictionary[T_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[T_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[T_SECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[T_SECTION].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[T_SECTION].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                // ====== ======
                // Create Joints
                // ====== ======
                // Add a Joint between Pipe and T Section
                JointHelper.CreateRigidJoint(T_SECTION, "BeginCap", "-1", "Route", Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeX, -pipeDiameter / 2.0, steelWidth / 2.0, shoeLength / 2.0);

                if (plates == 1)
                {
                    // Create the Rigid Joint between the End Plate and the T Section

                    JointHelper.CreateRigidJoint(T_SECTION, "BeginCap", END_PLATE1, "Route", Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX, -pipeDiameter / 2.0 + endplateClearance, steelWidth / 2.0, endplateT + endplateInset);

                    // Create the Rigid Joint between the End Plate and the T Section

                    JointHelper.CreateRigidJoint(T_SECTION, "EndCap", END_PLATE2, "Route", Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX, -pipeDiameter / 2.0 + endplateClearance, steelWidth / 2.0, -endplateInset);
                }
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

                    //Find the number of route that we are connecting to
                    for (int index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                        routeConnections.Add(new ConnectionInfo(T_SECTION, 1)); // partindex, routeindex

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

                    int numberOfStructures;
                    if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        numberOfStructures = 1;
                    else
                        numberOfStructures = SupportHelper.SupportingObjects.Count;

                    //Get the number of structure connections
                    for (int index = 1; index <= numberOfStructures; index++)
                        structConnections.Add(new ConnectionInfo(T_SECTION, index)); // partindex, routeindex

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
    }
}
