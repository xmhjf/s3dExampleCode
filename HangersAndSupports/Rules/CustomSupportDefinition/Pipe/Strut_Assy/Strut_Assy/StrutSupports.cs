//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   StrutSupports.cs
//   Strut_Assy,Ingr.SP3D.Content.Support.Rules.StrutSupports
//   Author       :  Vijay
//   Creation Date:  14/07/2013   
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   14/07/2013     Vijay    CR-CP-224475 Convert HS_S3DStrut_Assy to C# .Net 
//   04-11-2014     PVK      CR-CP-245790 Modify the exsisting .Net Strut_Assy to new URS Strut supports
//   22/01/2015     PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report 
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

    public class StrutSupports : CustomSupportDefinition
    {
        //Part Index's
        private const string STRUT_1A = "STRUT_1A";
        private const string STRUT_1B = "STRUT_1B";
        private const string STRUT_2A = "STRUT_2A";
        private const string STRUT_2B = "STRUT_2B";
        private const string PIPEATT_1 = "PIPEATT_1";
        private const string PIPEATT_2 = "PIPEATT_2";
        private const string STRUCTATT_1 = "STRUCTATT_1";
        private const string STRUCTATT_2 = "STRUCTATT_2";

        private const string STRUT1CONN = "STRUT1CONN";
        private const string STRUT2CONN = "STRUT2CONN";
        private const string PIPECONN = "PIPECONN";

        private const string SUPCONN_1 = "SUPCONN_1";
        private const string SUPCONN_2 = "SUPCONN_2";
        private const string SUPCONN_3 = "SUPCONN_3";
        private const string SUPCONN_4 = "SUPCONN_4";
        private const string STRUCTCONN_1 = "STRUCTCONN_1";
        private const string STRUCTCONN_2 = "STRUCTCONN_2";

        private const string sWeld = "WELD";

        //Part Classes
        private string strut1A;
        private string strut1B;
        private string strut2A;
        private string strut2B;
        private string pipeAtt1;
        private string pipeAtt2;

        private string structAtt1;
        private string structAtt2;

        //Strut 1 Attributes
        private double strut1Angle1;
        private double strut1Angle2;
        private double strut1Offset1;
        private double strut1MinP_P;
        private int strut1Config;
        private int strut1EndOrientation;

        //Strut 2 Attributes
        private double strut2Angle1;
        private double strut2Angle2;
        private double strut2Offset1;
        private double strut2MinP_P;
        private int strut2Config;
        private int strut2EndOrientation;

        //Weld
        private string weld;

        //Maximum Angles
        private double maxStructAngle;
        private double maxRouteAngle;

        //Booleans
        private bool withSecondStrut;
        private bool twoStructures;

        //BOM Length and Finish Attributes
        private int bomLengthUnitsValue;
        private string bomLengthUnits;
        private int finish;
        private string struct1APartClass, struct1BPartClass, pipeAtt1PartClass, structAtt1PartClass, struct2APartClass, struct2BPartClass, pipeAtt2PartClass, structAtt2PartClass;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                    if (part.SupportsInterface("IJUAHgrURSCommon"))
                    {
                        string family = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrURSCommon", "Family")).PropValue;
                        string type = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrURSCommon", "Type")).PropValue;
                        StrutAssemblyServices.CheckSupportWithFamilyAndType(this, family, type);
                    }
                   
                    //Get the Part Classes
                    if (part.SupportsInterface("IJUAhsStrut1_ClassA"))
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut1_ClassA", "Strut1ClassA", ref strut1A);
                    else
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut1", "Strut1ClassA", ref strut1A);
                    if (part.SupportsInterface("IJUAhsStrut1_ClassB"))
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut1_ClassB", "Strut1ClassB", ref strut1B);
                    else
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut1", "Strut1ClassB", ref strut1B);
                    if (part.SupportsInterface("IJUAhsStrut2_ClassA"))
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut2_ClassA", "Strut2ClassA", ref strut2A);
                    else
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut2", "Strut2ClassA", ref strut2A);
                    if (part.SupportsInterface("IJUAhsStrut2_ClassB"))
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut2_ClassB", "Strut2ClassB", ref strut2B);
                    else
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut2", "Strut2ClassB", ref strut2B);

                    GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_PipeAtt", "PipeAttachment", ref pipeAtt1);
                    GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_PipeAtt2", "PipeAttachment2", ref pipeAtt2);
                    GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_StructAtt", "StructAttachment", ref structAtt1);
                    GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_StructAtt2", "StructAttachment2", ref structAtt2);

                    //Weld Part Class
                    if (part.SupportsInterface("IJUAHgrURS_WeldAtt1"))
                        GetStringPropertyValueFromOccurrenceOrPart(support, "IJUAHgrURS_WeldAtt1", "Weld1", ref weld);

                    //Get the Attributes for Strut 1
                    if (part.SupportsInterface("IJUAhsStrut1_Angle1"))
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut1_Angle1", "Strut1Angle1", ref strut1Angle1);
                    else
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut1", "Strut1Angle1", ref strut1Angle1);
                    if (part.SupportsInterface("IJUAhsStrut1_Angle2"))
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut1_Angle2", "Strut1Angle2", ref strut1Angle2);
                    else
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut1", "Strut1Angle2", ref strut1Angle2);
                    if (part.SupportsInterface("IJUAhsStrut1_Offset1"))
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut1_Offset1", "Strut1Offset1", ref strut1Offset1);
                    else
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut1", "Strut1Offset1", ref strut1Offset1);
                    if (part.SupportsInterface("IJUAhsStrut1_MinP_P"))
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut1_MinP_P", "Strut1MinP_P", ref strut1MinP_P);
                    else
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut1", "Strut1MinP_P", ref strut1MinP_P);
                    if (part.SupportsInterface("IJUAhsStrut1_Config"))
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut1_Config", "Strut1Config", ref strut1Config);
                    else
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut1", "Strut1Config", ref strut1Config);
                    if (part.SupportsInterface("IJUAhsStrut1_EndOrient"))
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut1_EndOrient", "Strut1EndOrientation", ref strut1EndOrientation);
                    else
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut1", "Strut1EndOrientation", ref strut1EndOrientation);



                    //Get the Attributes for Strut 2
                    if (part.SupportsInterface("IJUAhsStrut2_Angle1"))
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut2_Angle1", "Strut2Angle1", ref strut2Angle1);
                    else
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut2", "Strut2Angle1", ref strut2Angle1);
                    if (part.SupportsInterface("IJUAhsStrut2_Angle2"))
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut2_Angle2", "Strut2Angle2", ref strut2Angle2);
                    else
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut2", "Strut2Angle2", ref strut2Angle2);
                    if (part.SupportsInterface("IJUAhsStrut2_Offset1"))
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut2_Offset1", "Strut2Offset1", ref strut2Offset1);
                    else
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut2", "Strut2Offset1", ref strut2Offset1);
                    if (part.SupportsInterface("IJUAhsStrut2_MinP_P"))
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut2_MinP_P", "Strut2MinP_P", ref strut2MinP_P);
                    else
                        GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut2", "Strut2MinP_P", ref strut2MinP_P);
                    if (part.SupportsInterface("IJUAhsStrut2_Config"))
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut2_Config", "Strut2Config", ref strut2Config);
                    else
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut2", "Strut2Config", ref strut2Config);
                    if (part.SupportsInterface("IJUAhsStrut2_EndOrient"))
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut2_EndOrient", "Strut2EndOrientation", ref strut2EndOrientation);
                    else
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_Strut2", "Strut2EndOrientation", ref strut2EndOrientation);

                    //Get the Max Angle Attributes                 
                    GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_MaxStructAngle", "MaxStructAngle", ref maxStructAngle);
                    GetDoublePropertyValueFromOccurrenceOrPart(support, "IJUAhsStrut_MaxRouteAngle", "MaxRouteAngle", ref maxRouteAngle);

                    //Get the BOM Length Units and Finish
                    if (part.SupportsInterface("IJOAhsBOMLenUnits"))
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJOAhsBOMLenUnits", "BOMLenUnits", ref bomLengthUnitsValue);
                    else
                        bomLengthUnitsValue = 1;

                    if (part.SupportsInterface("IJOAhsPTPFinish"))
                        GetCodeListPropertyValueFromOccurrenceOrPart(support, "IJOAhsPTPFinish", "Finish", ref finish);
                    else
                        finish = 1;

                    if (HgrCompareDoubleService.cmpdbl(maxRouteAngle , 0)==false)
                    {
                        if (maxRouteAngle < strut1Angle1)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Strut 1 pipe clamp angle exceeds the maximum allowable.", "", "StrutSupports.cs", 153);

                        if (maxRouteAngle < strut2Angle1)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Strut 2 pipe clamp angle exceeds the maximum allowable.", "", "StrutSupports.cs", 156);
                    }

                    if (HgrCompareDoubleService.cmpdbl(maxStructAngle, 0) == false)
                    {
                        if (maxStructAngle < strut1Angle2)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Strut 1 structure angle exceeds the maximum allowable.", "", "StrutSupports.cs", 162);

                        if (maxStructAngle < strut2Angle2)
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "Parts" + ": " + "WARNING: " + "Strut 2 structure angle exceeds the maximum allowable.", "", "StrutSupports.cs", 165);
                    }

                    //Get PartClass
                    GetPartClassValue(strut1A, ref struct1APartClass);
                    GetPartClassValue(strut1B, ref struct1BPartClass);
                    GetPartClassValue(pipeAtt1, ref pipeAtt1PartClass);
                    GetPartClassValue(structAtt1, ref structAtt1PartClass);

                    //Create a new collection to hold the caltalog parts
                    parts.Add(new PartInfo(STRUT_1A, strut1A, struct1APartClass,1));
                    parts.Add(new PartInfo(STRUT_1B, strut1B, struct1BPartClass, 1));
                    parts.Add(new PartInfo(PIPEATT_1, pipeAtt1, pipeAtt1PartClass, 1));
                    parts.Add(new PartInfo(STRUCTATT_1, structAtt1, structAtt1PartClass, 1));
                    parts.Add(new PartInfo(SUPCONN_1, "Log_Conn_Part_1", "", 1));
                    parts.Add(new PartInfo(SUPCONN_2, "Log_Conn_Part_1", "", 1));
                    parts.Add(new PartInfo(STRUCTCONN_1, "Log_Conn_Part_1", "", 1));
                    parts.Add(new PartInfo(STRUT1CONN, "Utility_Connection_1", "", 1));
                    parts.Add(new PartInfo(PIPECONN, "Log_Conn_Part_1", "", 1));

                    //Get the Parts for the Second Strut
                    //Set the Referenct Input to 2 so the rule knows to look up the size for the second strut

                    if (strut2A != "No Value" && strut2B != "No Value")
                    {
                        withSecondStrut = true;

                        GetPartClassValue(strut2A, ref struct2APartClass);
                        GetPartClassValue(strut2B, ref struct2BPartClass);
                        GetPartClassValue(pipeAtt2, ref pipeAtt2PartClass);
                        GetPartClassValue(structAtt2, ref structAtt2PartClass);

                        parts.Add(new PartInfo(STRUT_2A, strut2A, struct2APartClass, 2));
                        parts.Add(new PartInfo(STRUT_2B, strut2B, struct2BPartClass, 2));
                        parts.Add(new PartInfo(PIPEATT_2, pipeAtt2, pipeAtt2PartClass, 2));
                        parts.Add(new PartInfo(STRUCTATT_2, structAtt2, structAtt2PartClass, 2));
                        parts.Add(new PartInfo(SUPCONN_3, "Log_Conn_Part_1", "", 2));
                        parts.Add(new PartInfo(SUPCONN_4, "Log_Conn_Part_1", "", 2));
                        parts.Add(new PartInfo(STRUCTCONN_2, "Log_Conn_Part_1", "", 2));
                        parts.Add(new PartInfo(STRUT2CONN, "Utility_Connection_1", "", 2));
                    }

                    //Weld
                    if (weld == null)
                        weld = "No Value";
                    if (weld != "No Value")
                        parts.Add(new PartInfo(sWeld, weld, "", 3));

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
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------

        public override int ConfigurationCount
        {
            get
            {
                return 4;
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
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //#################################################
                // Get the Structure Colection
                //#################################################
                if (SupportHelper.SupportingObjects.Count != 0)
                {
                    if (SupportHelper.SupportingObjects.Count == 2)
                        twoStructures = true;
                    else
                        twoStructures = false;
                }

                //#################################################
                // Parameters for Vector Algebra to determine the Angle to Vertical Plane for PipeAtt and for strut
                //#################################################

                Vector locationRoute = new Vector(), locationStruct = new Vector(), locationStruct2 = new Vector(), vector = new Vector(), routeZ = new Vector(), routeX = new Vector(), routeY = new Vector(), structX = new Vector(), structY = new Vector(), structZ = new Vector(), plane = new Vector(), globalZ = new Vector(), struct1FaceNormal = new Vector(), pz = new Vector(), py = new Vector(), projection = new Vector(), negSz = new Vector();
                Vector planeNormal = new Vector(), perpComponent = new Vector(), struct2FaceNormal = new Vector();
                Matrix4X4 matrix = new Matrix4X4();
                double pipeAttOffset1, pipeAttOffset2, strut1CorrectionAngle, strut2CorrectionAngle;
                Plane[] structAtt1Plane = new Plane[2]; Plane[] structAtt2Plane = new Plane[2];
                if (withSecondStrut)
                {
                    BusinessObject partPipeAtt1 = componentDictionary[PIPEATT_1].GetRelationship("madeFrom", "part").TargetObjects[0];
                    BusinessObject partPipeAtt2 = componentDictionary[PIPEATT_2].GetRelationship("madeFrom", "part").TargetObjects[0];
                    pipeAttOffset1 = (double)((PropertyValueDouble)partPipeAtt1.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue / 2;
                    pipeAttOffset2 = (double)((PropertyValueDouble)partPipeAtt2.GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue / 2;
                }
                else
                {
                    pipeAttOffset1 = 0;
                    pipeAttOffset2 = 0;
                }

                Matrix4X4 hangerPort = new Matrix4X4();
                Position portPosition1 = new Position(), portPosition2 = new Position();
                Vector portX = new Vector(), portY = new Vector(), portZ = new Vector();

                IConnectable connectable1A = (IConnectable)componentDictionary[STRUT_1A];
                ReadOnlyCollection<IPort> ports1A = connectable1A.GetPorts(PortType.All);

                foreach (IPort iport in ports1A)
                {
                    Port port = iport as Port;
                    if (port != null)
                    {
                        if (port.Name.Contains("Port1"))
                        {
                            port.GetOrientation(out portX, out portZ);
                            portPosition1 = new Position(port.Origin.X, port.Origin.Y, port.Origin.Z);
                            portY = portX.Cross(portZ);
                            break;
                        }
                    }
                }

                IConnectable connectable1B = (IConnectable)componentDictionary[STRUT_1B];
                ReadOnlyCollection<IPort> ports1B = connectable1B.GetPorts(PortType.All);

                foreach (IPort iport in ports1B)
                {
                    Port port = iport as Port;
                    if (port != null)
                    {
                        if (port.Name.Contains("Port1"))
                        {
                            port.GetOrientation(out portX, out portZ);
                            portPosition2 = new Position(port.Origin.X, port.Origin.Y, port.Origin.Z);
                            portY = portX.Cross(portZ);
                            break;
                        }
                    }
                }

                double strutLengthValue = portPosition1.DistanceToPoint(portPosition2);
                if (support.SupportsInterface("IJOAhsStrut_CCLength"))
                    support.SetPropertyValue(strutLengthValue, "IJOAhsStrut_CCLength", "CCLength");

                if (bomLengthUnitsValue == -1)
                    bomLengthUnitsValue = 6;

                if (componentDictionary[STRUT_1A].SupportsInterface("IJOAhsPTPStrutAssyInfo"))
                {
                    bomLengthUnits = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, bomLengthUnitsValue, UnitName.DISTANCE_INCH);
                    string strutLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, strutLengthValue, UnitName.DISTANCE_INCH);
                    componentDictionary[STRUT_1A].SetPropertyValue("CCLength:" + strutLength + "," + bomLengthUnits + ",Finish:" + finish, "IJOAhsPTPStrutAssyInfo", "StrutAssyInfo");
                }

                switch (strut1Config)
                {
                    case 1:
                        //Configuration 1: Angles are set with respect to the line from the Route Reference Port to the Structure
                        //                   Reference Port (First Structure Selected)

                        //Get the Angle From the Route Z to the Vector from Route to Structure to set the 0 angle of the pipeclamp

                        hangerPort = new Matrix4X4();
                        hangerPort = RefPortHelper.PortLCS("Route");
                        routeX = new Vector(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                        routeZ = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);
                        locationRoute = new Vector(hangerPort.Origin.X, hangerPort.Origin.Y, hangerPort.Origin.Z);

                        hangerPort = new Matrix4X4();
                        hangerPort = RefPortHelper.PortLCS("Structure");
                        structX = new Vector(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                        structZ = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);
                        structY = structZ.Cross(structX);
                        locationStruct = new Vector(hangerPort.Origin.X, hangerPort.Origin.Y, hangerPort.Origin.Z);

                        //Vector From Route to Structure
                        vector = new Vector(locationStruct.X - locationRoute.X, locationStruct.Y - locationRoute.Y, locationStruct.Z - locationRoute.Z);

                        //If Route and Structure are Parralel, then Add the WBA Offset to the Vector from route to struct
                        //This will ensure that the StructAngle is calculated such that the StructAttachment is centered on the
                        //structure

                        double routeStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);
                        if (routeStructAngle < (Math.Atan(1) * 4.0) / 4.0 || routeStructAngle > 3 * Math.Atan(1) * 4.0 / 4.0)
                        {
                            BusinessObject part = componentDictionary[STRUCTATT_1].GetRelationship("madeFrom", "part").TargetObjects[0];
                            double structAttOffset = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsHeight2", "Height2")).PropValue;

                            if (SupportHelper.SupportingObjects.Count != 0)
                            {
                                //Get the Vector Normal to the Structure Face (May Not be the Z Axis in certain Cases)
                                IPlane iplane = (IPlane)support.SupportingFaces[0];

                                struct1FaceNormal = new Vector(iplane.Normal.X, iplane.Normal.Y, iplane.Normal.Z);
                                struct1FaceNormal.Length = structAttOffset;

                                double struct1NormalAngle = Math.Acos(struct1FaceNormal.Dot(structZ) / (struct1FaceNormal.Length * structZ.Length));
                                if (struct1NormalAngle <= Math.Atan(1) * 4.0 / 4.0)
                                {
                                    //Structure Z Axis is Pointing along the Face Normal
                                    structAtt1Plane[0] = Plane.XY;
                                    structAtt1Plane[1] = Plane.XY;
                                }
                                else if (struct1NormalAngle > Math.Atan(1) * 4.0 / 4.0 && struct1NormalAngle < 3 * Math.Atan(1) * 4.0 / 4.0)
                                {
                                    //Structure Z Axis is Perpendicular to Face Normal (Therefore, Use Y Instead of Z when defining the joints)
                                    struct1NormalAngle = Math.Acos(struct1FaceNormal.Dot(structY) / (struct1FaceNormal.Length * structY.Length));
                                    if (struct1NormalAngle >= Math.Atan(1) * 4.0 / 2.0)
                                    {
                                        structAtt1Plane[0] = Plane.XY;
                                        structAtt1Plane[1] = Plane.NegativeZX;
                                    }
                                    else
                                    {
                                        structAtt1Plane[0] = Plane.XY;
                                        structAtt1Plane[1] = Plane.ZX;
                                    }
                                }
                                else
                                {
                                    //Structure Z Axis is Pointing opposite to the Face Normal (Flip the Config Index)
                                    structAtt1Plane[0] = Plane.XY;
                                    structAtt1Plane[1] = Plane.NegativeXY;
                                }
                            }
                            else            //Place By Reference
                            {
                                //Structure Z Axis is Pointing opposite to the Face Normal (Flip the Config Index)
                                structAtt1Plane[0] = Plane.XY;
                                structAtt1Plane[1] = Plane.NegativeXY;
                            }

                            vector = vector.Add(struct1FaceNormal);

                            //If Vector and FaceNormal are Perpendicular, then Add an Extra Constraint
                            if (HgrCompareDoubleService.cmpdbl(Math.Round(vector.Dot(struct1FaceNormal), 5) , 0)==true)
                                JointHelper.CreatePointOnAxisJoint(STRUCTATT_1, "Structure", "-1", "Structure", Axis.X);
                        }

                        plane = vector.Cross(routeX);
                        routeY = routeZ.Cross(routeX);

                        double structAngle = routeY.Angle(plane, routeX);

                        //Determine the Angle to Rotate the Strut Around the Pipe-Clamp Pin

                        matrix.SetIdentity();

                        routeZ.Length = 1.0;
                        matrix.SetIndexValue(0, routeZ.X);
                        matrix.SetIndexValue(4, routeZ.Y);
                        matrix.SetIndexValue(8, routeZ.Z);
                        matrix.SetIndexValue(12, 1);

                        matrix.Rotate(strut1Angle1 + structAngle, routeX);

                        pz = new Vector(matrix.GetIndexValue(0), matrix.GetIndexValue(4), matrix.GetIndexValue(8));

                        py = pz.Cross(routeX);
                        py.Length = 1.0;

                        //Project Pz into a plane defined by -Sz
                        negSz = new Vector(-structZ.X, -structZ.Y, -structZ.Z);

                        planeNormal = negSz.Cross(negSz.Cross(routeX));

                        planeNormal.Length = 1.0;

                        perpComponent = new Vector(planeNormal.X * pz.Dot(planeNormal), planeNormal.Y * pz.Dot(planeNormal), planeNormal.Z * pz.Dot(planeNormal));

                        projection = pz.Subtract(perpComponent);

                        strut1CorrectionAngle = pz.Angle(projection, py);

                        //Attach Pipe Clamp to Pipe
                        if (componentDictionary[PIPEATT_1].SupportsInterface("IJUAhsAngle3"))
                        componentDictionary[PIPEATT_1].SetPropertyValue(strut1Angle1 + structAngle,"IJUAhsAngle3","Angle3");

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct && withSecondStrut != true)
                        {
                            JointHelper.CreatePrismaticJoint(PIPECONN, "Connection", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                            JointHelper.CreatePointOnAxisJoint(STRUCTATT_1, "Structure", "-1", "Structure", Axis.X);
                        }
                        else
                            JointHelper.CreateRigidJoint("-1", "Route", PIPECONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreateRigidJoint(PIPECONN, "Connection", PIPEATT_1, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, strut1Offset1 + pipeAttOffset1);

                        componentDictionary[STRUT1CONN].SetPropertyValue(strut1Angle2 + strut1CorrectionAngle, "IJUAhsConnectionEx", "FlexPortRotY");

                        //Previously AngularRigidJoint was not applicable.So instead of the below two RigidJoint,AngularRigidJoint is implemented.
                        //JointHelper.CreateRigidJoint(nPipeAtt1, "Wing", nStrut1Conn, "FlexPort", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //JointHelper.CreateRigidJoint(nStrut1Conn, "Connection", nStrut1B, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateAngularRigidJoint(STRUT_1B, "Port1", PIPEATT_1, "Wing", new Vector(0, 0, 0), new Vector(0, strut1Angle2 + strut1CorrectionAngle, 0));

                        if (strut1EndOrientation == 1)
                            JointHelper.CreateRevoluteJoint(STRUT_1B, "Port2", STRUT_1A, "Port2", Axis.Z, Axis.Z);
                        else if (strut1EndOrientation == 2)
                            JointHelper.CreateRigidJoint(STRUT_1A, "Port2", STRUT_1B, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(STRUT_1A, "Port2", STRUT_1B, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                        JointHelper.CreatePrismaticJoint(STRUT_1A, "Port1", STRUT_1A, "Port2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                        JointHelper.CreateRevoluteJoint(STRUCTATT_1, "Pin", STRUT_1A, "Port1", Axis.Y, Axis.Y);
                        JointHelper.CreatePlanarJoint(STRUCTATT_1, "Structure", "-1", "Structure", structAtt1Plane[0], structAtt1Plane[1], 0);

                        break;

                    case 2:
                        //Configuration 2: Angles are set with respect to Global Vertical
                        //Get the Angle From Route Z to Global Z about the Route X Axis.

                        hangerPort = new Matrix4X4();
                        hangerPort = RefPortHelper.PortLCS("Route");
                        routeX = new Vector(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                        routeZ = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);

                        hangerPort = new Matrix4X4();
                        hangerPort = RefPortHelper.PortLCS("Structure");
                        structZ = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);

                        globalZ = new Vector(0, 0, 1);

                        plane = globalZ.Cross(routeX);
                        routeY = routeZ.Cross(routeX);

                        double routeAngleFromVertical = routeY.Angle(plane, routeX);

                        //Determine the Angle to Rotate the Strut Around the Pipe-Clamp Pin
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();

                        routeZ.Length = 1.0;
                        matrix.SetIndexValue(0, routeZ.X);
                        matrix.SetIndexValue(4, routeZ.Y);
                        matrix.SetIndexValue(8, routeZ.Z);
                        matrix.SetIndexValue(12, 1);

                        matrix.Rotate(strut1Angle1 + routeAngleFromVertical, routeX);
                        pz = new Vector(matrix.GetIndexValue(0), matrix.GetIndexValue(4), matrix.GetIndexValue(8));

                        py = pz.Cross(routeX);
                        py.Length = 1.0;

                        //Project Pz into a plane defined by -Sz
                        negSz = new Vector(-structZ.X, -structZ.Y, -structZ.Z);
                        planeNormal = negSz.Cross(negSz.Cross(routeX));
                        planeNormal.Length = 1.0;

                        perpComponent = new Vector(planeNormal.X * pz.Dot(planeNormal), planeNormal.Y * pz.Dot(planeNormal), planeNormal.Z * pz.Dot(planeNormal));

                        projection = pz.Subtract(perpComponent);

                        strut1CorrectionAngle = pz.Angle(projection, py);

                        componentDictionary[PIPEATT_1].SetPropertyValue(strut1Angle1 + routeAngleFromVertical, "IJUAhsAngle3", "Angle3");

                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct && withSecondStrut != true)
                        {
                            JointHelper.CreatePrismaticJoint(PIPECONN, "Connection", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                            JointHelper.CreatePointOnAxisJoint(STRUCTATT_1, "Structure", "-1", "Structure", Axis.X);
                        }
                        else
                            JointHelper.CreateRigidJoint("-1", "Route", PIPECONN, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreateRigidJoint(PIPECONN, "Connection", PIPEATT_1, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, strut1Offset1 + pipeAttOffset1);
                        componentDictionary[STRUT1CONN].SetPropertyValue(strut1Angle2 + strut1CorrectionAngle, "IJUAhsConnectionEx", "FlexPortRotY");

                        //Previously AngularRigidJoint was not applicable.So instead of the below two RigidJoint,AngularRigidJoint is implemented.
                        //JointHelper.CreateRigidJoint(nPipeAtt1, "Wing", nStrut1Conn, "FlexPort", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        //JointHelper.CreateRigidJoint(nStrut1Conn, "Connection", nStrut1B, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        JointHelper.CreateAngularRigidJoint(STRUT_1B, "Port1", PIPEATT_1, "Wing", new Vector(0, 0, 0), new Vector(0, strut1Angle2 + strut1CorrectionAngle, 0));

                        if (strut1EndOrientation == 1)
                            JointHelper.CreateRevoluteJoint(STRUT_1B, "Port2", STRUT_1A, "Port2", Axis.Z, Axis.Z);
                        else if (strut1EndOrientation == 2)
                            JointHelper.CreateRigidJoint(STRUT_1A, "Port2", STRUT_1B, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(STRUT_1A, "Port2", STRUT_1B, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                        JointHelper.CreatePrismaticJoint(STRUT_1A, "Port1", STRUT_1A, "Port2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                        JointHelper.CreateRevoluteJoint(STRUCTATT_1, "Pin", STRUT_1A, "Port1", Axis.Y, Axis.Y);
                        JointHelper.CreatePlanarJoint(STRUCTATT_1, "Structure", "-1", "Structure", Plane.XY, Plane.XY, 0);

                        break;
                }

                if (withSecondStrut)
                {
                    string structReferencePort = string.Empty;
                    string structPart = string.Empty;
                    int structIndex;
                    double routeStruct2Angle, angle;

                    if (twoStructures)
                    {
                        structReferencePort = "Struct_2";
                        structIndex = 1;
                        structPart = STRUT_1B;
                    }
                    else
                    {
                        structReferencePort = "Structure";
                        structIndex = 0;
                        structPart = STRUT_1A;
                    }

                    switch (strut2Config)
                    {
                        case 1:
                            //Configuration 1: Angles are set with respect to the line from the Route Reference Port to the Structure

                            //Get the adjustment angle such that the second strut's zero is towards the second structure

                            hangerPort = new Matrix4X4();
                            hangerPort = RefPortHelper.PortLCS("Route");
                            routeX = new Vector(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                            routeZ = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);
                            locationRoute = new Vector(hangerPort.Origin.X, hangerPort.Origin.Y, hangerPort.Origin.Z);

                            hangerPort = new Matrix4X4();
                            hangerPort = RefPortHelper.PortLCS(structReferencePort);
                            structX = new Vector(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                            structZ = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);
                            structY = structZ.Cross(structX);
                            locationStruct2 = new Vector(hangerPort.Origin.X, hangerPort.Origin.Y, hangerPort.Origin.Z);

                            //Vector From Route to Struct_2
                            vector = new Vector(locationStruct2.X - locationRoute.X, locationStruct2.Y - locationRoute.Y, locationStruct2.Z - locationRoute.Z);

                            //If Route and Struct_2 are Parralel, then Add the WBA Offset to the Vector from route to struct
                            //This will ensure that the StructAngle is calculated such that the StructAttachment is centered on the
                            //structure

                            routeStruct2Angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, structReferencePort, PortAxisType.X, OrientationAlong.Direct);
                            if (routeStruct2Angle < (Math.Atan(1) * 4.0) / 4.0 || routeStruct2Angle > 3 * Math.Atan(1) * 4.0 / 4.0)
                            {
                                BusinessObject part = componentDictionary[STRUCTATT_2].GetRelationship("madeFrom", "part").TargetObjects[0];
                                double struct2AttOffset = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsHeight2", "Height2")).PropValue;

                                if (SupportHelper.SupportingObjects.Count != 0)
                                {
                                    //Get the Vector Normal to the Structure Face (May Not be the Z Axis in certain Cases)
                                    IPlane iplane = (IPlane)support.SupportingFaces[structIndex];

                                    struct2FaceNormal = new Vector(iplane.Normal.X, iplane.Normal.Y, iplane.Normal.Z);
                                    struct2FaceNormal.Length = struct2AttOffset;

                                    double struct2NormalAngle = Math.Acos(struct2FaceNormal.Dot(structZ) / (struct2FaceNormal.Length * structZ.Length));
                                    if (struct2NormalAngle <= Math.Atan(1) * 4.0 / 4.0)
                                    {
                                        //Structure Z Axis is Pointing along the Face Normal
                                        structAtt2Plane[0] = Plane.XY;
                                        structAtt2Plane[1] = Plane.XY;
                                    }
                                    else if (struct2NormalAngle > Math.Atan(1) * 4.0 / 4.0 && struct2NormalAngle < 3 * Math.Atan(1) * 4.0 / 4.0)
                                    {
                                        //Structure Z Axis is Perpendicular to Face Normal (Therefore, Use Y Instead of Z when defining the joints)
                                        struct2NormalAngle = Math.Acos(struct2FaceNormal.Dot(structY) / (struct2FaceNormal.Length * structY.Length));
                                        if (struct2NormalAngle >= Math.Atan(1) * 4.0 / 2.0)
                                        {
                                            structAtt1Plane[0] = Plane.XY;
                                            structAtt1Plane[1] = Plane.NegativeZX;
                                        }
                                        else
                                        {
                                            structAtt1Plane[0] = Plane.XY;
                                            structAtt1Plane[1] = Plane.ZX;
                                        }
                                    }
                                    else
                                    {
                                        //Structure Z Axis is Pointing opposite to the Face Normal (Flip the Config Index)
                                        structAtt1Plane[0] = Plane.XY;
                                        structAtt1Plane[1] = Plane.NegativeXY;
                                    }
                                }
                                else            //Place By Reference
                                {
                                    //Structure Z Axis is Pointing opposite to the Face Normal (Flip the Config Index)
                                    structAtt1Plane[0] = Plane.XY;
                                    structAtt1Plane[1] = Plane.NegativeXY;
                                }

                                vector = vector.Add(struct2FaceNormal);

                                //If Vector and FaceNormal are Perpendicular, then Add an Extra Constraint

                                if (HgrCompareDoubleService.cmpdbl(Math.Round(vector.Dot(struct2FaceNormal), 5), 0) == true)
                                    JointHelper.CreatePointOnAxisJoint(STRUCTATT_2, "Structure", structPart, structReferencePort, Axis.X);
                            }

                            plane = vector.Cross(routeX);
                            routeY = routeZ.Cross(routeX);

                            double structAngle = routeY.Angle(plane, routeX);

                            //Determine the Angle to Rotate the Strut Around the Pipe-Clamp Pin
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();

                            routeZ.Length = 1.0;
                            matrix.SetIndexValue(0, routeZ.X);
                            matrix.SetIndexValue(4, routeZ.Y);
                            matrix.SetIndexValue(8, routeZ.Z);
                            matrix.SetIndexValue(12, 1);

                            matrix.Rotate(strut2Angle1 + structAngle, routeX);

                            pz = new Vector(matrix.GetIndexValue(0), matrix.GetIndexValue(4), matrix.GetIndexValue(8));

                            py = pz.Cross(routeX);
                            py.Length = 1.0;

                            //Project Pz into a plane defined by -Sz
                            negSz = new Vector(-structZ.X, -structZ.Y, -structZ.Z);

                            planeNormal = negSz.Cross(negSz.Cross(routeX));

                            planeNormal.Length = 1.0;

                            perpComponent = new Vector(planeNormal.X * pz.Dot(planeNormal), planeNormal.Y * pz.Dot(planeNormal), planeNormal.Z * pz.Dot(planeNormal));

                            projection = pz.Subtract(perpComponent);

                            strut2CorrectionAngle = pz.Angle(projection, py);

                            //Attach Pipe Clamp to Pipe
                            componentDictionary[PIPEATT_2].SetPropertyValue(strut2Angle1 + structAngle, "IJUAhsAngle3", "Angle3");

                            JointHelper.CreateRigidJoint(PIPECONN, "Connection", PIPEATT_2, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, strut2Offset1 - pipeAttOffset2);

                            componentDictionary[STRUT2CONN].SetPropertyValue(strut2Angle2 + strut2CorrectionAngle, "IJUAhsConnectionEx", "FlexPortRotY");

                            //Previously AngularRigidJoint was not applicable.So instead of the below two RigidJoint,AngularRigidJoint is implemented.
                            //JointHelper.CreateRigidJoint(nPipeAtt2, "Wing", nStrut2Conn, "FlexPort", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            //JointHelper.CreateRigidJoint(nStrut2Conn, "Connection", nStrut2B, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            JointHelper.CreateAngularRigidJoint(STRUT_2B, "Port1", PIPEATT_2, "Wing", new Vector(0, 0, 0), new Vector(0, strut2Angle2 + strut2CorrectionAngle, 0));

                            if (strut2EndOrientation == 1)
                                JointHelper.CreateRevoluteJoint(STRUT_2B, "Port2", STRUT_2A, "Port2", Axis.Z, Axis.Z);
                            else if (strut2EndOrientation == 2)
                                JointHelper.CreateRigidJoint(STRUT_2A, "Port2", STRUT_2B, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(STRUT_2A, "Port2", STRUT_2B, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                            JointHelper.CreatePrismaticJoint(STRUT_2A, "Port1", STRUT_2A, "Port2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                            JointHelper.CreateRevoluteJoint(STRUCTATT_2, "Pin", STRUT_2A, "Port1", Axis.Y, Axis.Y);

                            if (routeStruct2Angle > Math.Atan(1) * 4.0 / 4.0 && routeStruct2Angle < 3 * Math.Atan(1) * 4.0 / 4.0)
                            {
                                angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, structReferencePort, PortAxisType.Z, OrientationAlong.Direct);

                                if (Math.Round(angle * 180.0 / Math.PI, 0) > 170 || Math.Round(angle * 180.0 / Math.PI, 0) < 10)
                                    JointHelper.CreatePlanarJoint(STRUCTATT_2, "Structure", "-1", structReferencePort, Plane.XY, Plane.ZX, 0);
                                else
                                    JointHelper.CreatePlanarJoint(STRUCTATT_2, "Structure", "-1", structReferencePort, structAtt2Plane[0], structAtt2Plane[1], 0);
                            }
                            else
                                JointHelper.CreatePlanarJoint(STRUCTATT_2, "Structure", "-1", structReferencePort, structAtt2Plane[0], structAtt2Plane[1], 0);
                            break;
                        case 2:
                            //Configuration 2: Angles are set with respect to Global Vertical

                            hangerPort = new Matrix4X4();
                            hangerPort = RefPortHelper.PortLCS("Route");
                            routeX = new Vector(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                            routeZ = new Vector(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);

                            globalZ = new Vector(0, 0, 1);

                            plane = globalZ.Cross(routeX);
                            routeY = routeZ.Cross(routeX);

                            double routeAngleFromVertical = routeY.Angle(plane, routeX);

                            //Determine the Angle to Rotate the Strut Around the Pipe-Clamp Pin
                            matrix = new Matrix4X4();
                            matrix.SetIdentity();

                            routeZ.Length = 1.0;
                            matrix.SetIndexValue(0, routeZ.X);
                            matrix.SetIndexValue(4, routeZ.Y);
                            matrix.SetIndexValue(8, routeZ.Z);
                            matrix.SetIndexValue(12, 1);

                            matrix.Rotate(strut2Angle1 + routeAngleFromVertical, routeX);
                            pz = new Vector(matrix.GetIndexValue(0), matrix.GetIndexValue(4), matrix.GetIndexValue(8));

                            py = pz.Cross(routeX);
                            py.Length = 1.0;

                            //Project Pz into a plane defined by -Sz
                            negSz = new Vector(-structZ.X, -structZ.Y, -structZ.Z);
                            planeNormal = negSz.Cross(negSz.Cross(routeX));
                            planeNormal.Length = 1.0;

                            perpComponent = new Vector(planeNormal.X * pz.Dot(planeNormal), planeNormal.Y * pz.Dot(planeNormal), planeNormal.Z * pz.Dot(planeNormal));

                            projection = pz.Subtract(perpComponent);

                            strut2CorrectionAngle = pz.Angle(projection, py);

                            componentDictionary[PIPEATT_2].SetPropertyValue(strut2Angle1 + routeAngleFromVertical, "IJUAhsAngle3", "Angle3");

                            JointHelper.CreateRigidJoint(PIPECONN, "Connection", PIPEATT_2, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, strut2Offset1 - pipeAttOffset2);

                            componentDictionary[STRUT2CONN].SetPropertyValue(strut2Angle2 + strut2CorrectionAngle, "IJUAhsConnectionEx", "FlexPortRotY");

                            //Previously AngularRigidJoint was not applicable.So instead of the below two RigidJoint,AngularRigidJoint is implemented.
                            //JointHelper.CreateRigidJoint(nPipeAtt2, "Wing", nStrut2Conn, "FlexPort", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            //JointHelper.CreateRigidJoint(nStrut2Conn, "Connection", nStrut2B, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            JointHelper.CreateAngularRigidJoint(STRUT_2B, "Port1", PIPEATT_2, "Wing", new Vector(0, 0, 0), new Vector(0, strut2Angle2 + strut2CorrectionAngle, 0));

                            if (strut2EndOrientation == 1)
                                JointHelper.CreateRevoluteJoint(STRUT_2B, "Port2", STRUT_2A, "Port2", Axis.Z, Axis.Z);
                            else if (strut2EndOrientation == 2)
                                JointHelper.CreateRigidJoint(STRUT_2A, "Port2", STRUT_2B, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(STRUT_2A, "Port2", STRUT_2B, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                            JointHelper.CreatePrismaticJoint(STRUT_2A, "Port1", STRUT_2A, "Port2", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                            JointHelper.CreateRevoluteJoint(STRUCTATT_2, "Pin", STRUT_2A, "Port1", Axis.Y, Axis.Y);
                            routeStruct2Angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, structReferencePort, PortAxisType.X, OrientationAlong.Direct);
                            if (routeStruct2Angle > Math.Atan(1) * 4.0 / 4.0 && routeStruct2Angle < 3 * Math.Atan(1) * 4.0 / 4.0)
                            {
                                angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, structReferencePort, PortAxisType.Z, OrientationAlong.Direct);

                                if (Math.Round(angle * 180.0 / Math.PI, 0) > 170 || Math.Round(angle * 180.0 / Math.PI, 0) < 10)
                                    JointHelper.CreatePlanarJoint(STRUCTATT_2, "Structure", "-1", structReferencePort, Plane.XY, Plane.ZX, 0);
                                else
                                    JointHelper.CreatePlanarJoint(STRUCTATT_2, "Structure", "-1", structReferencePort, Plane.XY, Plane.XY, 0);
                            }
                            else
                                JointHelper.CreatePlanarJoint(STRUCTATT_2, "Structure", "-1", structReferencePort, Plane.XY, Plane.XY, 0);

                            break;
                    }
                }
                // Joints for Weld
                double width2 = 0;
                if (weld == null)
                    weld = "No Value";
                if (weld != "No Value")
                {
                    BusinessObject structAttach1 = componentDictionary[STRUCTATT_1].GetRelationship("madeFrom", "part").TargetObjects[0];
                    if (structAttach1!=null)
                        width2 = (double)((PropertyValueDouble)structAttach1.GetPropertyValue("IJUAhsWidth2", "Width2")).PropValue;
                    JointHelper.CreateRigidJoint(STRUCTATT_1, "Structure", sWeld, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, -width2/2);
                }


                //#################################################
                // Setup my dimension points
                //#################################################

                Note note1 = CreateNote("Route", componentDictionary[PIPEATT_1], "Route");
                note1.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList1 = note1.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note1.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");
                note1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note2 = CreateNote("StrutEnd", componentDictionary[STRUT_1B], "Port1");
                note2.SetPropertyValue("StrutEnd", "IJGeneralNote", "Text");
                CodelistItem codeList2 = note2.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note2.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");
                note2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note3 = CreateNote("StrutEnd1", componentDictionary[STRUT_1A], "Port1");
                note3.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList3 = note3.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note3.SetPropertyValue(codeList3, "IJGeneralNote", "Purpose");
                note3.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note4 = CreateNote("Structure", componentDictionary[STRUCTATT_1], "Structure");
                note4.SetPropertyValue("", "IJGeneralNote", "Text");
                CodelistItem codeList4 = note4.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                note4.SetPropertyValue(codeList4, "IJGeneralNote", "Purpose");
                note4.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                if (withSecondStrut)
                {
                    Note note12 = CreateNote("Route2", componentDictionary[PIPEATT_2], "Route");
                    note12.SetPropertyValue("", "IJGeneralNote", "Text");
                    CodelistItem codeList12 = note12.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    note12.SetPropertyValue(codeList12, "IJGeneralNote", "Purpose");
                    note12.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                    Note note22 = CreateNote("StrutEnd2", componentDictionary[STRUT_2B], "Port1");
                    note22.SetPropertyValue("", "IJGeneralNote", "Text");
                    CodelistItem codeList22 = note22.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    note22.SetPropertyValue(codeList22, "IJGeneralNote", "Purpose");
                    note22.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                    Note note32 = CreateNote("StrutEnd3", componentDictionary[STRUT_2A], "Port1");
                    note32.SetPropertyValue("", "IJGeneralNote", "Text");
                    CodelistItem codeList32 = note32.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    note32.SetPropertyValue(codeList32, "IJGeneralNote", "Purpose");
                    note32.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                    Note note42 = CreateNote("Structure2", componentDictionary[STRUCTATT_2], "Structure");
                    note42.SetPropertyValue("", "IJGeneralNote", "Text");
                    CodelistItem codeList42 = note42.GetPropertyValue("IJGeneralNote", "Purpose").PropertyInfo.CodeListInfo.GetCodelistItem(3);
                    note42.SetPropertyValue(codeList42, "IJGeneralNote", "Purpose");
                    note42.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");
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

                    routeConnections.Add(new ConnectionInfo(PIPEATT_1, 1));       //partindex, routeindex

                    if (withSecondStrut)
                        routeConnections.Add(new ConnectionInfo(PIPEATT_2, 1));       //partindex, routeindex

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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(STRUCTATT_1, 1));      //partindex, routeindex

                    if (withSecondStrut)
                    {
                        if (twoStructures == true)
                            structConnections.Add(new ConnectionInfo(STRUCTATT_2, 2));      //partindex, routeindex
                        else
                            structConnections.Add(new ConnectionInfo(STRUCTATT_2, 1));      //partindex, routeindex
                    }
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

        /// <summary>
        /// This method will be called to check property is available on the supportcomponent and support return its value.
        /// </summary>
        /// <param name="part">supportcomponent or support is of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="propertyValue">Returns Property Value</param>
        /// <returns></returns>
        /// <code>
        /// GetDoublePropertyValue(part, "IJUAhsHeight1","Height1",ref propertyValue)
        /// </code>
        public void GetDoublePropertyValueFromOccurrenceOrPart(Ingr.SP3D.Support.Middle.Support support, string interfaceName, string propertyName, ref double propertyValue)
        {
            BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
            try
            {
                if (support.SupportsInterface(interfaceName))
                    propertyValue = (double)((PropertyValueDouble)support.GetPropertyValue(interfaceName, propertyName)).PropValue;
                else if (part.SupportsInterface(interfaceName))
                    propertyValue = (double)((PropertyValueDouble)part.GetPropertyValue(interfaceName, propertyName)).PropValue;                 
            }
            catch
            {

            }
        }     

        /// <summary>
        /// This method will be called to check property is available on the supportcomponent and support return its value.
        /// </summary>
        /// <param name="part">supportcomponent or support is of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="propertyValue">Returns Property Value</param>
        /// <returns></returns>
        /// <code>
        /// GetStringPropertyValue(part, "IJUAhsSectionType","SectionType",ref propertyValue)
        /// </code>
        public void GetStringPropertyValueFromOccurrenceOrPart(Ingr.SP3D.Support.Middle.Support support, string interfaceName, string propertyName, ref string propertyValue)
        {
            BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
            try
            {  
                if (support.SupportsInterface(interfaceName))
                    propertyValue = (string)((PropertyValueString)support.GetPropertyValue(interfaceName, propertyName)).PropValue;
                else if (part.SupportsInterface(interfaceName))
                    propertyValue = (string)((PropertyValueString)part.GetPropertyValue(interfaceName, propertyName)).PropValue;

                if (propertyValue == null)
                    propertyValue = "No Value";
            }
            catch
            {
                propertyValue = "No Value";
            }
        }
       
        /// <summary>
        /// This method will be called to check property is available on the supportcomponent and support return its value.
        /// </summary>
        /// <param name="part">supportcomponent or support is of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="propertyValue">Returns Property Value</param>
        /// <returns></returns>
        /// <code>
        /// GetStringPropertyValue(part, "IJUAhsSectionType", "SectionType", ref propertyValue)
        /// </code>
        public void GetCodeListPropertyValueFromOccurrenceOrPart(Ingr.SP3D.Support.Middle.Support support, string interfaceName, string propertyName, ref int propertyValue)
        {
            BusinessObject part = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
          
                try
                {
                    if (support.SupportsInterface(interfaceName))
                        propertyValue = (int)((PropertyValueCodelist)support.GetPropertyValue(interfaceName, propertyName)).PropValue;
                     else if (part.SupportsInterface(interfaceName))
                        propertyValue = (int)((PropertyValueCodelist)part.GetPropertyValue(interfaceName, propertyName)).PropValue;
                }
                catch
                {
                    propertyValue = -1;
                }           
        }
        /// <summary>
        /// This method will Check either its Part or PartClass and return PSL(PartSelectionRule) if it is PartClass or return empty string for Part.
        /// </summary>
        /// <param name="partOrPartClassName">Name of the PartClass</param>
        /// <param name="partSelectionRule">Return the PartSelectionRule</param>
        /// <returns></returns>
        /// <code>
        /// GetPartClassValue(partClassName, ref partClassValue)
        /// </code>
        public void GetPartClassValue(string partOrPartClassName, ref string partSelectionRule)
        {
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            BusinessObject partclass = catalogBaseHelper.GetPartClass(partOrPartClassName);
            if (partclass is PartClass)
            {
                partSelectionRule = partclass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
            }
            else
            {
                partSelectionRule = "";
            }
        }
    }
}
