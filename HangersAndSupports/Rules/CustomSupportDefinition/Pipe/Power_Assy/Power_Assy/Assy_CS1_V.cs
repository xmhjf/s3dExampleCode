//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_CS1_V.cs
//   Power_Assy,Ingr.SP3D.Content.Support.Rules.Assy_CS1_V
//   Author       :  Manikanth
//   Creation Date:  20/03/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  20/03/2013    Manikanth  CR-CP-224472-Initial Creation
//  22/02/2015    PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class Assy_CS1_V : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants

        public const string pipeClamp = "nPipeClamp";
        private const string rod1 = "nRod1";
        private const string rod2 = "nRod2";
        private const string beamAtt1 = "nBeamAtt1";
        private const string beamAtt2 = "nBeamAtt2";
        private const string eyeNut1 = "nEyeNut1";
        private const string eyeNut2 = "nEyeNut2";
        private const string eyeNut3 = "nEyeNut3";
        private const string eyeNut4 = "nEyeNut4";
        private const string connObj1 = "nConnObj1";
        private const string connObj2 = "nConnObj2";

        private const string spring1 = "Spring1";
        private const string spring2 = "Spring2";
        private const string rod3 = "Rod3";
        private const string rod4 = "Rod4";

        string clampPart, loadType, supportType, rodSize;
        double clampWidth, clampHeight, rodSpacing, maxTemp, maxLoad, springHeight, springDia,  pipeDia;
        int rotAngle, showSpring, spring1Pos, spring2Pos;
        private string connection;

        private const string Steel = "Steel";
        private const string Slab = "Slab";
        private const string SlabSteel = "SlabSteel";
        private const string SteelSlab = "SteelSlab";

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //Get the attributes from assembly
                    if (support.SupportsInterface("IJUAHgrPowAssyClamp"))
                    {
                        clampWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyClamp", "ClampWidth")).PropValue;
                    }
                    else
                    {
                        clampWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyClamp_V", "ClampWidth")).PropValue;
                    }
                    if (support.SupportsInterface("IJUAHgrPowAssyClamp_V"))
                    {
                        clampHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyClamp_V", "ClampHeight")).PropValue;
                    }
                    else
                    {
                        clampHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyClamp", "ClampHeight")).PropValue;
                    }
                    if (support.SupportsInterface("IJUAHgrPowAssyRodSP_V"))
                    {

                        rodSpacing = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyRodSP_V", "RodSpacing")).PropValue;
                    }
                    else
                    {
                        rodSpacing = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyRodSP", "RodSpacing")).PropValue;
                    }



                    supportType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrPowAssyCSType", "SupType")).PropValue;
                    try
                    {
                        maxTemp = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyMaxTemp", "MaxTemp")).PropValue;
                    }
                    catch
                    {
                        maxTemp = 0;
                    }
                    try
                    {
                        maxLoad = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyMaxLoad", "MaximumLoad")).PropValue;
                    }
                    catch
                    {
                        maxLoad = 0;
                    }
                    showSpring = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssySpring", "ShowSpring")).PropValue;
                    spring1Pos = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssySpring", "Spring1Pos")).PropValue;
                    spring2Pos = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssySpring", "Spring2Pos")).PropValue;
                    springHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssySpring", "Height")).PropValue;
                    springDia = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssySpring", "Dia")).PropValue;
                    rotAngle = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssyRotAngle", "RotAngle")).PropValue;
                    PropertyValueCodelist rodSizeCodelist = null;


                    if (support.SupportsInterface("IJUAHgrPowAssyCSRod_V"))
                    {
                        rodSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssyCSRod_V", "RodSize");
                    }
                    else
                    {
                        rodSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrPowAssyCSRod", "RodSize");
                    }
                    int rodSizeValue = (int)rodSizeCodelist.PropValue;
                    rodSize = rodSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)rodSizeValue).ShortDisplayName;



                    CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                    double pipeDia1 = 0;
                    string strUnit = "mm";
                    if (supportType.ToUpper() == "FIXED")
                    {
                        if (support.SupportsInterface("IJUAHgrPowAssyCSRod_V"))
                        {
                            pipeDia = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyCS2PartSel", "PipeOutDia")).PropValue;
                        }
                        else
                        {
                            pipeDia = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrPowAssyClampDW", "PipeOutDia")).PropValue;
                        }
                        pipeDia = pipeDia / 1000;
                        strUnit = "mm";
                    }

                    else if (supportType.ToUpper() == "VARIABLE")
                    {
                        //Get the Pipe Clamp
                        PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                        NominalDiameter nominalPipeDia = new NominalDiameter();
                        nominalPipeDia = pipeInfo.NominalDiameter;
                        pipeDia = pipeInfo.OutsideDiameter;
                        strUnit = nominalPipeDia.Units;
                    }

                    if (strUnit == "mm")
                    {
                        pipeDia1 = pipeDia * 1000;
                    }

                    else if (strUnit == "in")
                    {
                        pipeDia1 = (pipeDia * 39.37008);
                        strUnit = "in";
                    }


                    PartClass auxTable = (PartClass)cataloghelper.GetPartClass("Assy_CS1PartSel");
                    ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;


                    if (HgrCompareDoubleService.cmpdbl(maxTemp , 0)==true || HgrCompareDoubleService.cmpdbl(maxLoad , 0)==true)
                    {
                        foreach (BusinessObject classItem in classItems)
                        {
                            if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAssyCS1PartSel", "PipeOutDia")).PropValue > pipeDia1 - 0.01) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAssyCS1PartSel", "PipeOutDia")).PropValue < pipeDia1 + 0.01) && ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAssyCS1PartSel", "UnitType")).PropValue == strUnit))
                            {
                                clampPart = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAssyCS1PartSel", "PartNo")).ToString();
                                break;
                            }
                        }
                    }
                    else
                    {
                        foreach (BusinessObject classItem in classItems)
                        {
                            loadType = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAssyCS1LoadType", "LoadType")).ToString();
                            clampPart = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAssyCS1PartSel", "PartNo")).ToString();
                        }
                    }


                    parts.Add(new PartInfo(pipeClamp, clampPart));
                    parts.Add(new PartInfo(rod1, "Anvil_FIG140_" + rodSize));
                    parts.Add(new PartInfo(rod2, "Anvil_FIG140_" + rodSize));
                    parts.Add(new PartInfo(beamAtt1, "Anvil_FIG66_" + rodSize));
                    parts.Add(new PartInfo(beamAtt2, "Anvil_FIG66_" + rodSize));
                    parts.Add(new PartInfo(eyeNut1, "Anvil_FIG290_" + rodSize));
                    parts.Add(new PartInfo(eyeNut2, "Anvil_FIG290_" + rodSize));
                    parts.Add(new PartInfo(eyeNut3, "Anvil_FIG290_" + rodSize));
                    parts.Add(new PartInfo(eyeNut4, "Anvil_FIG290_" + rodSize));
                    parts.Add(new PartInfo(connObj1, "Log_Conn_Part_1"));   //Rotational Connection Object
                    parts.Add(new PartInfo(connObj2, "Log_Conn_Part_1"));   //Rotational Connection Object

                    if (showSpring == 1)
                    {

                        parts.Add(new PartInfo(spring1, "Spring-1"));
                        parts.Add(new PartInfo(spring2, "Spring-1"));
                        parts.Add(new PartInfo(rod3, "Anvil_FIG140_" + rodSize));
                        parts.Add(new PartInfo(rod4, "Anvil_FIG140_" + rodSize));

                    }
                    return parts;    //Return the collection of Catalog Parts


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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(pipeClamp, 1));
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
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections.Add(new ConnectionInfo(pipeClamp, 1));
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();
                string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
                string rightStructPort = idxStructPort[0];
                string leftStructPort = idxStructPort[1];

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent pipe_clamp = componentDictionary[pipeClamp];

                //Determine whether connecting to Steel or a Slab
                if (SupportHelper.SupportingObjects.Count != 0)
                {
                    if (SupportHelper.SupportingObjects.Count > 1)
                    {
                        connection = Steel;
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                        {
                            connection = Slab;          //Two Slabs
                        }
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Member))
                        {
                            connection = SlabSteel;    //Slab then Steel
                        }
                        if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && (SupportingHelper.SupportingObjectInfo(2).SupportingObjectType == SupportingObjectType.Slab))
                        {
                            connection = SteelSlab;    //Steel then Slab
                        }
                    }
                    else
                    {
                        if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                        {
                            connection = Steel;
                        }
                        else if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Slab)
                        {
                            connection = Slab;
                        }
                    }
                }
                else
                    connection = Slab;

                if (PW_IsStructureSlopedAcrossPipe(connection))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "The structure is sloped. Please check.", "", "Assy_CS1_V.cs", 336);
                    return;
                }


                if (clampWidth <= rodSpacing)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Rod Spacing should be less than the Clamp Width. Please check the value", "", "Assy_CS1_V.cs", 343);
                    return;
                }

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();

                PartClass assy_CS1_VPartClass = (PartClass)cataloghelper.GetPartClass("Assy_CS1_V");
                ReadOnlyCollection<BusinessObject> parts = assy_CS1_VPartClass.Parts;

                // if user doesn't edit the value, set the default rod spacing to three times the Pipe Outside diameter
                double tempRodSpacing = (double)((PropertyValueDouble)parts[0].GetPropertyValue("IJUAHgrPowAssyRodSP_V", "RodSpacing")).PropValue;

                if (rodSpacing >= (tempRodSpacing - Math3d.DistanceTolerance) && (rodSpacing <= (tempRodSpacing + Math3d.DistanceTolerance)))
                {
                    rodSpacing = 3 * pipeDia;
                    clampWidth = rodSpacing + 0.1;
                }
                pipe_clamp.SetPropertyValue(rodSpacing, "IJUAHgrPowerCS", "C");
                pipe_clamp.SetPropertyValue(clampHeight, "IJUAHgrPowerCS", "ClampHeight");
                pipe_clamp.SetPropertyValue(clampWidth, "IJUAHgrPowerCS", "L");

                //Spring 1 position
                if (showSpring == 1)
                {
                    double eyenutG, eyenutE, springG, beamAttE, beamAttHt, beamAttTakeout, rodDia;

                    SupportComponent beamAtt = componentDictionary[beamAtt1];
                    BusinessObject beamAtt_1 = beamAtt.GetRelationship("madeFrom", "part").TargetObjects[0];

                    SupportComponent spring_1 = componentDictionary[spring1];
                    SupportComponent spring_2 = componentDictionary[spring2];
                    SupportComponent rod_3 = componentDictionary[rod3];
                    SupportComponent rod_4 = componentDictionary[rod4];
                    beamAttHt = (double)((PropertyValueDouble)beamAtt_1.GetPropertyValue("IJUAHgrAnvil_fig66", "B")).PropValue;
                    beamAttTakeout = (double)((PropertyValueDouble)beamAtt_1.GetPropertyValue("IJUAHgrTake_Out", "TAKE_OUT")).PropValue;
                    beamAttE = (double)((PropertyValueDouble)beamAtt_1.GetPropertyValue("IJUAHgrAnvil_fig66", "E")).PropValue;

                    SupportComponent rod = componentDictionary[rod1];
                    BusinessObject rod_1 = rod.GetRelationship("madeFrom", "part").TargetObjects[0];

                    rodDia = (double)((PropertyValueDouble)rod_1.GetPropertyValue("IJUAHgrRod_Dia", "ROD_DIA")).PropValue;

                    SupportComponent eye1 = componentDictionary[eyeNut1];
                    BusinessObject eyenut = eye1.GetRelationship("madeFrom", "part").TargetObjects[0];

                    eyenutG = (double)((PropertyValueDouble)eyenut.GetPropertyValue("IJUAHgrAnvil_FIG290", "G")).PropValue;
                    eyenutE = (double)((PropertyValueDouble)eyenut.GetPropertyValue("IJUAHgrAnvil_FIG290", "E")).PropValue;


                    SupportComponent nSpring = componentDictionary[spring1];
                    BusinessObject spri = nSpring.GetRelationship("madeFrom", "part").TargetObjects[0];

                    springG = (double)((PropertyValueDouble)spri.GetPropertyValue("IJUAHgrPowerCSSpring", "G")).PropValue;

                    spring_1.SetPropertyValue(springHeight, "IJUAHgrPowerCSSpring", "B");
                    spring_1.SetPropertyValue(springDia, "IJUAHgrPowerCSSpring", "C");
                    spring_1.SetPropertyValue(springDia + springDia / 5, "IJUAHgrPowerCSSpring", "D");
                    spring_1.SetPropertyValue(springHeight, "IJUAHgrPowTakeOut", "TakeOut");
                    spring_1.SetPropertyValue(rodDia, "IJUAHgrPowerCSSpring", "A");

                    spring_2.SetPropertyValue(springHeight, "IJUAHgrPowerCSSpring", "B");
                    spring_2.SetPropertyValue(springDia, "IJUAHgrPowerCSSpring", "C");
                    spring_2.SetPropertyValue(springDia + springDia / 5, "IJUAHgrPowerCSSpring", "D");
                    spring_2.SetPropertyValue(springHeight, "IJUAHgrPowTakeOut", "TakeOut");
                    spring_2.SetPropertyValue(rodDia, "IJUAHgrPowerCSSpring", "A");

                    double routeStrVertDist, spring1Dist = 0, spring2Dist = 0;
                    routeStrVertDist = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);



                    //Spring 1 position
                    if (spring1Pos == 1)
                    {
                        spring1Dist = eyenutG + springG + rodDia / 2;
                    }
                    else if (spring1Pos ==  2)
                    {
                        spring1Dist = routeStrVertDist / 2 - springHeight / 2 - beamAttHt;
                    }
                    else if (spring1Pos ==  3)
                    {
                        spring1Dist = routeStrVertDist - (beamAttTakeout + eyenutE - rodDia / 2) + clampHeight / 2 + springG;
                    }

                    //Spring 2 position
                    if (spring2Pos == 1)
                    {
                        spring2Dist = eyenutG + springG + rodDia / 2;
                    }
                    else if (spring2Pos == 2)
                    {
                        spring2Dist = routeStrVertDist / 2 - springHeight / 2 - beamAttHt;
                    }
                    else if (spring2Pos == 3)
                    {
                        spring2Dist = routeStrVertDist - (beamAttTakeout + eyenutE - rodDia / 2) + clampHeight / 2 + springG;
                    }

                    rod_3.SetPropertyValue(spring1Dist, "IJUAHgrOccLength", "Length");
                    rod_4.SetPropertyValue(spring2Dist, "IJUAHgrOccLength", "Length");
                }

                if (HgrCompareDoubleService.cmpdbl(maxTemp , 0 )== false || HgrCompareDoubleService.cmpdbl(maxLoad , 0 )== false)
                {
                    //Set properties on Assembly
                    double loadTyped = 0;
                    if (loadType == "1")
                    {
                        loadTyped = 1;
                    }
                    if (loadType == "2")
                    {
                        loadTyped = 2;
                    }
                    support.SetPropertyValue(loadTyped, "IJUAHgrPowAssyCSLT", "LoadType");
                }

                double routeStructConfigAng, structDirAng;
                Plane routePlane1, routePlane2, leftStructPortPlane1, leftStructPortPlane2, rightStructPortPlane1, rightStructPortPlane2;
                routePlane1 = routePlane2 = leftStructPortPlane1 = leftStructPortPlane2 = rightStructPortPlane1 = rightStructPortPlane2 = Plane.XY;

                if (connection == Steel)
                {
                    if (leftStructPort == rightStructPort)
                    {
                        leftStructPort = "Structure";
                        rightStructPort = "Structure";
                        routePlane1 = Plane.ZX;
                        routePlane2 = Plane.YZ;
                        leftStructPortPlane1 = rightStructPortPlane1 = Plane.XY;
                        leftStructPortPlane2 = rightStructPortPlane2 = Plane.NegativeXY;

                    }
                    else
                    {
                        routeStructConfigAng = RefPortHelper.PortConfigurationAngle("Struct_2", "Route", PortAxisType.Y);
                        structDirAng = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_X);


                        leftStructPortPlane1 = rightStructPortPlane1 = Plane.XY;
                        leftStructPortPlane2 = rightStructPortPlane2 = Plane.XY;

                        if ((structDirAng > (0 - 0.0001) && structDirAng < (0 + 0.0001) || structDirAng > (Math.PI - 0.0001) && structDirAng < (Math.PI + 0.0001)))
                        {
                            if (routeStructConfigAng < Math.PI / 2)
                            {
                                routePlane1 = Plane.YZ;
                                routePlane2 = Plane.NegativeYZ;
                            }
                            else
                            {
                                routePlane1 = Plane.YZ;
                                routePlane2 = Plane.YZ;
                            }
                        }
                        else
                        {
                            if (routeStructConfigAng > Math.PI / 2)
                            {
                                routePlane1 = Plane.YZ;
                                routePlane2 = Plane.NegativeYZ;
                            }
                            else
                            {
                                routePlane1 = Plane.YZ;
                                routePlane2 = Plane.YZ;
                            }
                        }
                    }
                }
                else if (connection == Slab)
                {
                    leftStructPortPlane1 = rightStructPortPlane1 = Plane.XY;
                    leftStructPortPlane2 = rightStructPortPlane2 = Plane.NegativeXY;
                    leftStructPort = "Structure";
                    rightStructPort = "Structure";
                }
                else if (connection == SlabSteel)
                {
                    leftStructPortPlane1 = Plane.XY;
                    leftStructPortPlane2 = Plane.NegativeXY;
                    rightStructPortPlane1 = Plane.XY;
                    rightStructPortPlane2 = Plane.XY;
                }
                else if (connection == SteelSlab)
                {
                    leftStructPortPlane1 = Plane.XY;
                    leftStructPortPlane2 = Plane.XY;
                    rightStructPortPlane1 = Plane.XY;
                    rightStructPortPlane2 = Plane.NegativeXY;
                }

                if (connection != Steel)
                {
                    if (rotAngle == 1)
                    {
                        routePlane1 = Plane.YZ;
                        routePlane2 = Plane.NegativeYZ;
                    }
                    else if (rotAngle == 2)
                    {
                        routePlane1 = Plane.ZX;
                        routePlane2 = Plane.YZ;
                    }
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "This support should be placed by Point.", "", "Assy_CS1_V.cs", 552);
                    return;
                }
                else
                {
                    JointHelper.CreateSphericalJoint(pipeClamp, "Route", "-1", "Route");
                    JointHelper.CreatePlanarJoint(pipeClamp, "Route", "-1", "Structure", routePlane1, routePlane2, 0);
                    //Vertical Joint
                    JointHelper.CreateGlobalAxesAlignedJoint(pipeClamp, "Route", Axis.Z, Axis.Z);
                }

                //Add a Joint between the Pipe Clamps and and Eye Nuts
                JointHelper.CreateRigidJoint(pipeClamp, "LeftPin", eyeNut1, "Eye", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(pipeClamp, "RightPin", eyeNut2, "Eye", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                //Create the Flexible (Prismatic) Joint between the ports of the top rods
                JointHelper.CreatePrismaticJoint(rod1, "TopExThdRH", rod1, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                JointHelper.CreatePrismaticJoint(rod2, "TopExThdRH", rod2, "BotExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);
                //Add a Joint between the beam attachment and eye nut
                JointHelper.CreateRigidJoint(eyeNut3, "Eye", beamAtt1, "Pin", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                JointHelper.CreateRigidJoint(eyeNut4, "Eye", beamAtt2, "Pin", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                //Add a Joint between the Rods and Eye Nuts
                JointHelper.CreateRigidJoint(rod1, "TopExThdRH", eyeNut1, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(rod2, "TopExThdRH", eyeNut2, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                if (showSpring == 1)
                {
                    //Spring 1 joints
                    JointHelper.CreateRigidJoint(rod1, "BotExThdRH", spring1, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(spring1, "TopInThdRH", rod3, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(rod3, "BotExThdRH", eyeNut3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    //Spring 2 joints
                    JointHelper.CreateRigidJoint(rod2, "BotExThdRH", spring2, "BotInThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(spring2, "TopInThdRH", rod4, "TopExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(rod4, "BotExThdRH", eyeNut4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else
                {
                    JointHelper.CreateRigidJoint(rod1, "BotExThdRH", eyeNut3, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(rod2, "BotExThdRH", eyeNut4, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                if (connection == Steel)
                {
                    if (leftStructPort != rightStructPort)
                    {
                        //Add a Joint between the Beam Attachement and Structure
                        JointHelper.CreateSphericalJoint(connObj1, "Connection", "-1", leftStructPort);
                        //Add a Joint between the Beam Attachement and Structure
                        JointHelper.CreateSphericalJoint(connObj2, "Connection", "-1", rightStructPort);
                        //Add a Joint between the Beam Attachement and Structure
                        JointHelper.CreateGlobalAxesAlignedJoint(connObj1, "Connection", Axis.X, Axis.X);
                        //Add a Joint between the Beam Attachement and Structure
                        JointHelper.CreateGlobalAxesAlignedJoint(connObj2, "Connection", Axis.X, Axis.X);
                        //Vertical Joint
                        JointHelper.CreateGlobalAxesAlignedJoint(connObj1, "Connection", Axis.Z, Axis.Z);
                        //Vertical Joint
                        JointHelper.CreateGlobalAxesAlignedJoint(connObj2, "Connection", Axis.Z, Axis.Z);
                        //Add a Joint between the Beam Attachement and Structure
                        JointHelper.CreatePlanarJoint(beamAtt1, "Structure", connObj1, "Connection", leftStructPortPlane1, leftStructPortPlane2, 0);
                        JointHelper.CreatePlanarJoint(beamAtt2, "Structure", connObj2, "Connection", rightStructPortPlane1, rightStructPortPlane2, 0);
                    }
                    else
                    {
                        //Add a Joint between the Beam Attachement and Structure
                        JointHelper.CreatePlanarJoint(beamAtt1, "Structure", "-1", leftStructPort, leftStructPortPlane1, leftStructPortPlane2, 0);
                        JointHelper.CreatePlanarJoint(beamAtt2, "Structure", "-1", rightStructPort, rightStructPortPlane1, rightStructPortPlane2, 0);
                    }
                }
                else if (connection == SlabSteel)
                {
                    //Add a Joint between the Beam Attachement and Structure
                    JointHelper.CreatePlanarJoint(beamAtt2, "Structure", "-1", leftStructPort, leftStructPortPlane1, leftStructPortPlane2, 0);
                    //Add a Joint between the Beam Attachement and Structure
                    JointHelper.CreateSphericalJoint(connObj1, "Connection", "-1", rightStructPort);
                    //Add a Joint between the Beam Attachement and Structure
                    JointHelper.CreateGlobalAxesAlignedJoint(connObj1, "Connection", Axis.X, Axis.X);
                    //Vertical Joint
                    JointHelper.CreateGlobalAxesAlignedJoint(connObj1, "Connection", Axis.Z, Axis.Z);
                    JointHelper.CreatePlanarJoint(beamAtt1, "Structure", connObj1, "Connection", rightStructPortPlane1, rightStructPortPlane2, 0);
                }
                else if (connection == SteelSlab)
                {
                    //Add a Joint between the Beam Attachement and Structure
                    JointHelper.CreatePlanarJoint(beamAtt2, "Structure", "-1", rightStructPort, rightStructPortPlane1, rightStructPortPlane2, 0);
                    //Add a Joint between the Beam Attachement and Structure
                    JointHelper.CreateSphericalJoint(connObj1, "Connection", "-1", leftStructPort);
                    //Add a Joint between the Beam Attachement and Structure
                    JointHelper.CreateGlobalAxesAlignedJoint(connObj1, "Connection", Axis.X, Axis.X);
                    //Vertical Joint
                    JointHelper.CreateGlobalAxesAlignedJoint(connObj1, "Connection", Axis.Z, Axis.Z);
                    JointHelper.CreatePlanarJoint(beamAtt1, "Structure", connObj1, "Connection", leftStructPortPlane1, leftStructPortPlane2, 0);
                }
                else if (connection == Slab)
                {
                    //Add a Joint between the Beam Attachement and Structure
                    JointHelper.CreatePlanarJoint(beamAtt1, "Structure", "-1", leftStructPort, leftStructPortPlane1, leftStructPortPlane2, 0);
                    JointHelper.CreatePlanarJoint(beamAtt2, "Structure", "-1", rightStructPort, rightStructPortPlane1, rightStructPortPlane2, 0);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        public Boolean PW_IsStructureSlopedAcrossPipe(string connection, Boolean bOnTurn = false)
        {
            try
            {
                bool PW_IsStructureSlopedAcrossPipe = true;

                double refAngle1;
                double refAngle2;
                double angle1 = 89.99999999999 / 180 * 3.14159265358979;
                double angle2 = 90.0000000001 / 180 * 3.14159265358979;

                refAngle1 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.Y, OrientationAlong.Global_Z);
                refAngle2 = RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, OrientationAlong.Global_Z);


                if (connection == Slab)
                {
                    if (refAngle1 > angle1 && refAngle2 < angle2)
                    {
                        PW_IsStructureSlopedAcrossPipe = false;
                    }
                    else
                    {
                        PW_IsStructureSlopedAcrossPipe = true;
                    }
                }
                else
                {
                    if (refAngle2 > angle1 && refAngle2 < angle2)
                    {
                        PW_IsStructureSlopedAcrossPipe = false;
                    }
                    else
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct || bOnTurn)
                        {
                            PW_IsStructureSlopedAcrossPipe = true;
                        }
                        else
                        {
                            PW_IsStructureSlopedAcrossPipe = false;
                        }
                    }
                }


                return PW_IsStructureSlopedAcrossPipe;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in PW_IsStructureSlopedAcrossPipe Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                string bom_Desc = "";
                string supportType;

                supportType = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrPowAssyCSType", "SupType")).PropValue;

                if (supportType.ToUpper() == "FIXED")
                {
                    bom_Desc = "CS1 Assembly with fixed clamp size (overall width and depth) and rod spacing";
                }
                else if (supportType.ToUpper() == "VARIABLE")
                {
                    bom_Desc = "CS1 Assembly with variable clamp size (overall width and depth) and rod spacing";
                }

                return bom_Desc;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOMDescription. Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
        /// <summary>
        /// This method differentiate multiple structure input object based on relative position.
        /// </summary>
        /// <param name="IsOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
        /// </code>
        public String[] GetIndexedStructPortName(Boolean[] IsOffsetApplied)
        {
            try
            {
                String[] structurePort = new String[2];
                int structureCount = SupportHelper.SupportingObjects.Count;
                int i;

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

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    structurePort[0] = "Structure";
                    structurePort[1] = "Structure";
                }
                else
                {
                    structurePort[0] = "Structure";
                    structurePort[1] = "Struct_2";

                    if (structureCount > 1)
                    {
                        if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        {
                            for (i = 0; i <= 1; i++)
                            {
                                double angle = 0;
                                if ((supportingType == "Steel") && IsOffsetApplied[i] == false)
                                {
                                    angle = RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
                                }
                                //the port is the right structure port
                                if (Math.Abs(angle) < Math.PI / 2.0)
                                {
                                    if (i == 0)
                                    {
                                        structurePort[0] = "Struct_2";
                                        structurePort[1] = "Structure";
                                    }
                                }
                                //the port is the left structure port
                                else
                                {
                                    if (i == 1)
                                    {
                                        structurePort[0] = "Struct_2";
                                        structurePort[1] = "Structure";
                                    }
                                }
                            }
                        }
                    }
                    else
                        structurePort[1] = "Structure";
                }
                //switch the OffsetApplied flag
                if (structurePort[0] == "Struct_2")
                {
                    Boolean flag = IsOffsetApplied[0];
                    IsOffsetApplied[0] = IsOffsetApplied[1];
                    IsOffsetApplied[1] = flag;
                }

                return structurePort;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIndexedStructPortName Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();  
        ///</code>
        public Boolean[] GetIsLugEndOffsetApplied()
        {
            try
            {
                Collection<BusinessObject> StructureObjects;
                Boolean[] isOffsetApplied = new Boolean[2];

                //first route object is set as primary route object
                StructureObjects = SupportHelper.SupportingObjects;
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

                isOffsetApplied[0] = true;
                isOffsetApplied[1] = true;

                if (StructureObjects != null)
                {
                    if (StructureObjects.Count >= 2)
                    {
                        for (int index = 0; index <= 1; index++)
                        {
                            //check if offset is to be applied
                            const double RouteStructAngle = Math.PI / 180.0;
                            Boolean varRuleApplied = true;

                            if (supportingType == "Steel")
                            {
                                //if angle is within 1 degree, regard as parallel case
                                //Also check for Sloped structure                                
                                MemberPart memberPart = (MemberPart)SupportHelper.SupportingObjects[index];
                                ICurve memberCurve = memberPart.Axis;

                                Vector supportedVector = new Vector();
                                Vector supportingVector = new Vector();

                                if (SupportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                                {
                                    Position startLocation = new Position(SupportedHelper.SupportedObjectInfo(1).StartLocation);
                                    Position endLocation = new Position(SupportedHelper.SupportedObjectInfo(1).EndLocation);
                                    supportedVector = new Vector(endLocation - startLocation);
                                }
                                if (memberCurve is ILine)
                                {
                                    ILine line = (ILine)memberCurve;
                                    supportingVector = line.Direction;
                                }

                                double angle = GetAngleBetweenVectors(supportingVector, supportedVector);
                                double refAngle1 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Y, OrientationAlong.Global_X) - Math.PI / 2;
                                double refAngle2 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.X, OrientationAlong.Global_X);

                                if (angle < (refAngle1 + 0.001) && angle > (refAngle1 - 0.001))
                                    angle = angle - Math.Abs(refAngle1);
                                else if (angle < (refAngle2 + 0.001) && angle > (refAngle2 - 0.001))
                                    angle = angle - Math.Abs(refAngle2);
                                else
                                    angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                if (Math.Abs(angle) < RouteStructAngle || Math.Abs(angle - Math.PI) < RouteStructAngle)
                                    varRuleApplied = false;
                            }

                            isOffsetApplied[index] = varRuleApplied;
                        }
                    }
                }

                return isOffsetApplied;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied Method of Power_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="vector1">The vector1(Vector).</param>
        /// <param name="vector2">The vector2(Vector).</param>
        /// <returns>double</returns>        
        /// <code>
        ///       ContentHelper contentHelper = new ContentHelper();
        ///       double value;
        ///       value = contentHelper. GetAngleBetweenVectors(vector1, vector2 );
        ///</code>
        public double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double arcCos = 0.0;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dotProd), 1) == false)
                {
                    arcCos = Math.PI / 2 - Math.Atan(dotProd / Math.Sqrt(1 - dotProd * dotProd));
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, -1) == true)
                {
                    arcCos = Math.PI;
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, 1) == true)
                {
                    arcCos = 0;
                }
                return arcCos;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied  Method of Power_Assy.." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }

        }
    }

}