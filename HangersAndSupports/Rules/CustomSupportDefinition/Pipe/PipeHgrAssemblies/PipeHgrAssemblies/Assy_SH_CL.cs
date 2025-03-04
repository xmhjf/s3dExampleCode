//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_SH_CL.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_SH_CL
//   Author       : Rajeswari
//   Creation Date: 17-April-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 17-April-2013  Rajeswari CR-CP-224484 C#.Net HS_Assembly Project Creation
// 22-Jan-2015       PVK    TR-CP-264951  Resolve coverity issues found in November 2014 report  
// 27-April-2015     PVK    TR-CP-253033  Elevation CP not shown by default
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

    public class Assy_SH_CL : CustomSupportDefinition
    {
        private const string T_SECTION = "T_SECTION";//1
        private const string CONNECTION1 = "CONNECTION1";//2
        private const string CONNECTION2 = "CONNECTION2";//3
        private const string CONNECTION3 = "CONNECTION3";//4
        private const string CONNECTION4 = "CONNECTION4";//5
        private const string PIPE_CLAMP1 = "PIPE_CLAMP1";//6
        private const string PIPE_CLAMP2 = "PIPE_CLAMP2";//7
        private const string END_PLATE1 = "END_PLATE1";//8
        private const string END_PLATE2 = "END_PLATE2";//9

        private Double shoeLength, clampInset, clampAngle;
        private String sectionSize, clampType, clamps, plateSize, plates;
        private Boolean userClampAng;
        //private int PLATES;
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
                    shoeLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssySH_CL", "SHOE_L")).PropValue;
                    clampInset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssySH_CL", "CLAMP_INSET")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlateSize", "PLATE_SIZE")).PropValue;
                    clamps = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssySH_CL", "CLAMPS")).PropValue;
                    plates = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlates", "Plates")).PropValue;
                   
                    // Initialise
                    userClampAng = false;
                    try
                    {
                        // Get the User ClampAngle
                        userClampAng = true;
                        clampAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrShoeClampAng", "CLAMP_ANGLE")).PropValue;
                    }
                   catch
                    {
                        userClampAng = false;
                    }

                    // Create the list of part classes required by the type
                    if(clamps=="Medium")
                        clampType="Anvil_FIG212";
                    else
                        clampType = "Anvil_FIG216";
                     if (plates == "With")
                    {
                        parts.Add(new PartInfo(T_SECTION, sectionSize));
                        parts.Add(new PartInfo(CONNECTION1, "Log_Conn_Part_1"));  // to rotate clamp
                        parts.Add(new PartInfo(CONNECTION2, "Log_Conn_Part_1"));  // to rotate clamp
                        parts.Add(new PartInfo(CONNECTION3, "Log_Conn_Part_1"));  // to rotate clamp
                        parts.Add(new PartInfo(CONNECTION4, "Log_Conn_Part_1"));  // to rotate clamp
                        parts.Add(new PartInfo(PIPE_CLAMP1, clampType));
                        parts.Add(new PartInfo(PIPE_CLAMP2, clampType));
                        parts.Add(new PartInfo(END_PLATE1, "Utility_END_PLATE"));
                        parts.Add(new PartInfo(END_PLATE2, "Utility_END_PLATE"));
                    }
                    else
                    {
                        parts.Add(new PartInfo(T_SECTION, sectionSize));
                        parts.Add(new PartInfo(CONNECTION1, "Log_Conn_Part_1"));  // to rotate clamp
                        parts.Add(new PartInfo(CONNECTION2, "Log_Conn_Part_1"));  // to rotate clamp
                        parts.Add(new PartInfo(CONNECTION3, "Log_Conn_Part_1"));  // to rotate clamp
                        parts.Add(new PartInfo(CONNECTION4, "Log_Conn_Part_1"));  // to rotate clamp
                        parts.Add(new PartInfo(PIPE_CLAMP1, clampType));
                        parts.Add(new PartInfo(PIPE_CLAMP2, clampType));
                    }

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

                String[] collectionOfInterfaces = new String[] { "IJUAHgrAnvil_FIG212", "IJUAHgrAnvil_FIG216" };
                CommonAssembly commonAssembly = new CommonAssembly();

                BusinessObject tSectionPart = componentDictionary[T_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                BusinessObject pipeClamp2Part = componentDictionary[PIPE_CLAMP2].GetRelationship("madeFrom", "part").TargetObjects[0];

                PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = routeInfo.OutsideDiameter;

                Double clampTO, calc1, steelWidth, steelDepth, flangeThickness, endplateT = 0.0;

                CrossSection crosssection = (CrossSection)tSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth 
                steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                flangeThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;

                clampTO = (double)((PropertyValueDouble)pipeClamp2Part.GetPropertyValue("IJUAHgrTake_Out", "TAKE_OUT")).PropValue;
                calc1 = clampTO / Math.Sqrt(2.0);

                if (plates == "With")
                {
                    PropertyValueCodelist utilityCodelist = (PropertyValueCodelist)componentDictionary[END_PLATE1].GetPropertyValue("IJOAHgrUtility_END_PLATE", "THICKNESS");
                    CodelistItem codeList = utilityCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(Convert.ToInt32(plateSize));

                    endplateT = Convert.ToDouble(codeList.ShortDisplayName) * 25.4 / 1000.0;

                    componentDictionary[END_PLATE1].SetPropertyValue(utilityCodelist.PropValue, "IJOAHgrUtility_END_PLATE", "THICKNESS");
                    componentDictionary[END_PLATE1].SetPropertyValue(steelWidth, "IJOAHgrUtility_END_PLATE", "W");
                    componentDictionary[END_PLATE1].SetPropertyValue(steelDepth - flangeThickness, "IJOAHgrUtility_END_PLATE", "H");

                    componentDictionary[END_PLATE2].SetPropertyValue(utilityCodelist.PropValue, "IJOAHgrUtility_END_PLATE", "THICKNESS");
                    componentDictionary[END_PLATE2].SetPropertyValue(steelWidth, "IJOAHgrUtility_END_PLATE", "W");
                    componentDictionary[END_PLATE2].SetPropertyValue(steelDepth - flangeThickness, "IJOAHgrUtility_END_PLATE", "H");
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
                Boolean excludeNotes;
                // Check for "ExcludeNotes" attribute (for migrated DB)

                if (support.SupportsInterface("IJUAhsExcludeNotes"))
                    excludeNotes = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsExcludeNotes", "ExcludeNotes")).PropValue;
                else
                    excludeNotes = false;

                ControlPoint controlPoint;
                Note note;

                if (excludeNotes == false)
                {
                    note = CreateNote("CAD Detail Vertical Section", T_SECTION, "BeginCap", new Position(steelWidth / 2.0, 0.0, shoeLength), @"\HangersAndSupports\CAD Details\Shoe Vertical Part.sym", false, 2, 53, out controlPoint);
                    controlPoint.SetUserDefinedName("CAD Detail");
                }
                else
                    DeleteNoteIfExists("CAD Detail Vertical Section");
                if (excludeNotes == false)
                {
                    note = CreateNote("CAD Detail Bottom", T_SECTION, "BeginCap", new Position(steelWidth / 2.0, steelDepth, shoeLength), @"\HangersAndSupports\CAD Details\Shoe Base.sym", false, 2, 53, out controlPoint);
                    controlPoint.SetUserDefinedName("CAD Detail");
                }
                else
                    DeleteNoteIfExists("CAD Detail Bottom");
                // Add a Joint between Pipe and T Section

                 JointHelper.CreateRigidJoint(T_SECTION, "BeginCap", "-1", "Route", Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeX, -pipeDiameter / 2.0, steelWidth / 2.0, shoeLength / 2.0);

                 if (plates == "With")
                 {
                     if (excludeNotes == false)
                     {
                         note = CreateNote("CAD Detail Plate", END_PLATE1, "Route", new Position(0.0, steelWidth / 2.0, (-pipeDiameter / 2.0) - steelDepth / 2.0), @"\HangersAndSupports\CAD Details\End Plate.sym", false, 2, 53, out controlPoint);
                         controlPoint.SetUserDefinedName("CAD Detail");
                     }
                     else
                         DeleteNoteIfExists("CAD Detail Plate");

                     // Create the Rigid Joint between the End Plate and the T Section
                     JointHelper.CreateRigidJoint(T_SECTION, "BeginCap", END_PLATE1, "Route", Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX, -pipeDiameter / 2.0, steelWidth / 2.0, endplateT / 2.0 + clampInset);

                     // Create the Rigid Joint between the End Plate and the T Section
                     JointHelper.CreateRigidJoint(T_SECTION, "EndCap", END_PLATE2, "Route", Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX, -pipeDiameter / 2.0, steelWidth / 2.0, -clampInset + endplateT / 2.0);
                 }

                if (userClampAng == false)
                {
                    double pi = Math.Atan(1) * 4;
                    clampAngle = 7 * pi / 4;
                }

                // User Angle
                // Create the Angular Rigid Joint between the Clamp1 and Route
                JointHelper.CreateAngularRigidJoint("-1", "Route", PIPE_CLAMP1, "Route", new Vector(-shoeLength / 2 + clampInset, 0, 0), new Vector(clampAngle, 0, 0));

                // Create the Angular Rigid Joint between the Clamp2 and Route
                JointHelper.CreateRigidJoint(PIPE_CLAMP1, "Route", PIPE_CLAMP2, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0,-(shoeLength - 2 * clampInset));
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
                    structConnections.Add(new ConnectionInfo(T_SECTION, 1)); // partindex, routeindex

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
