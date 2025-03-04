//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_D2.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_D2
//   Author       :  Vijaya
//   Creation Date: 1-Jul-2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   1-Jul-2013     Vijaya   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//  11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5397_D2 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSTEEL5397D2 = "HORSTEEL5397_D2";
        private const string VERTSTEEL5397D2 = "VERTSTEEL5397_D2";

        private string size3;
        PipeObjectInfo pipeInfo;
        double pipeDiameter = 0.0, shoeHeight = 0.0, overHang = 0.0, horizontalLength;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");
                    size3 = sizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeCodeList.PropValue).DisplayName;

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;

                    horizontalLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal) + overHang;

                    parts.Add(new PartInfo(HORSTEEL5397D2, "Utility_GENERIC_L_1"));
                    if (horizontalLength >= 500.00 / 1000.00)// with clamps
                        parts.Add(new PartInfo(VERTSTEEL5397D2, "Utility_CUTBACK_L2_1"));


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
                return 2;
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
                // To get Pipe Nominal Diameter
                pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                pipeDiameter = pipeInfo.OutsideDiameter + pipeInfo.InsulationThickness * 2;

                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                NominalDiameter minNominalDiameter = new NominalDiameter(), maxNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";

                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_D2.cs", 111);
                    return;
                }
                double maxLength = 0.0, MA = 0.0, MP = 0.0;
                PropertyValue LSize;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size3", size3);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_D2_Dim", "IJUAFINLSrv_SFS5397_D2_Dim", "L_Size", parameter, out LSize);
                string lSize = LSize.ToString();


                maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D2_Dim", "IJUAFINLSrv_SFS5397_D2_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_D2_Dim", "Size3", size3);
                MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D2_Dim", "IJUAFINLSrv_SFS5397_D2_Dim", "MA", "IJUAFINLSrv_SFS5397_D2_Dim", "Size3", size3);
                MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D2_Dim", "IJUAFINLSrv_SFS5397_D2_Dim", "MP", "IJUAFINLSrv_SFS5397_D2_Dim", "Size3", size3);

                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                double lWidth, lFlangeThickness, lWebThickness, lDepth;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", lSize, out  lWidth, out   lFlangeThickness, out  lWebThickness, out  lDepth);


                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_D2.cs", 187);
                    return;
                }

                double theta = Math.Abs(Math.PI / 2 - RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "", PortAxisType.Y, OrientationAlong.Global_Z));
                double angleOffset = Math.Tan(theta) * lWidth;


                double structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                horizontalLength = structDistance + overHang;


                if (structDistance + overHang > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Maximum allowable length of " + (maxLength * 1000).ToString() + "mm exceeded.", "", "SFS5397_D2.cs", 147);


                if (overHang < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must exceed pipe outside dia:" + (pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_D2.cs", 151);


                double verticalLength = (550.00 / 1000.00 - 1.0 / 2.0 * (lDepth / Math.Tan(Math.PI / 6) + (lDepth / Math.Tan(Math.PI / 3))));
                double verticalOffset = 550.00 / 1000.00 * Math.Cos(Math.PI / 6) - horizontalLength;
                double lCut = lDepth / Math.Sin(Math.PI / 6);

                string inputBOM = lSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[HORSTEEL5397D2].SetPropertyValue(lWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                componentDictionary[HORSTEEL5397D2].SetPropertyValue(lDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                componentDictionary[HORSTEEL5397D2].SetPropertyValue(horizontalLength, "IJOAHgrUtility_GENERIC_L", "L");
                componentDictionary[HORSTEEL5397D2].SetPropertyValue(lFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                componentDictionary[HORSTEEL5397D2].SetPropertyValue(inputBOM, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");


                if (horizontalLength >= 500.00 / 1000.00)
                {

                    inputBOM = lSize + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                    componentDictionary[VERTSTEEL5397D2].SetPropertyValue(lWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                    componentDictionary[VERTSTEEL5397D2].SetPropertyValue(lDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                    componentDictionary[VERTSTEEL5397D2].SetPropertyValue(verticalLength, "IJOAHgrUtility_GENERIC_L", "L");
                    componentDictionary[VERTSTEEL5397D2].SetPropertyValue(lFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                    componentDictionary[VERTSTEEL5397D2].SetPropertyValue(-Math.PI / 6, "IJOAHgrUtility_CUTBACK", "ANGLE");//rotae about y-axis
                    componentDictionary[VERTSTEEL5397D2].SetPropertyValue(-Math.PI / 3, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[VERTSTEEL5397D2].SetPropertyValue(inputBOM, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                }

                double flipAngle = 0.0;
                int sign = 1;
                if (Configuration == 2)
                {
                    flipAngle = Math.PI;
                    sign = -1;
                }

                double verticalDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                double theta1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);
                double theta2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Z, OrientationAlong.Global_Z);


                if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                {
                    if (Symbols.HgrCompareDoubleService.cmpdbl(flipAngle, Math.PI) == true)
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D2, "Structure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, (structDistance + overHang), -(-sign * pipeDiameter / 2 - sign * shoeHeight + angleOffset / 2 - verticalDistance), -lDepth / 2);
                    else
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D2, "Structure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY, (structDistance + overHang), -(-sign * pipeDiameter / 2 - sign * shoeHeight + angleOffset / 2 - verticalDistance), -lDepth / 2);

                    if (horizontalLength >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397D2, "Structure", VERTSTEEL5397D2, "StartStructure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, lFlangeThickness, 0, -verticalOffset + lCut);

                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (Math.Abs(theta1) < Math.PI / 2)
                        {
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D2, "Structure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, structDistance + overHang, -lDepth / 2, -pipeDiameter / 2 - shoeHeight - angleOffset / 2);
                            if (horizontalLength >= 500.00 / 1000.00)
                                JointHelper.CreateRigidJoint(HORSTEEL5397D2, "Structure", VERTSTEEL5397D2, "StartStructure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, lFlangeThickness, 0, -verticalOffset + lCut);

                        }
                        else
                        {//flip
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D2, "Structure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, structDistance + overHang, lDepth / 2, (pipeDiameter / 2 + shoeHeight + angleOffset / 2));
                            if (horizontalLength >= 500.00 / 1000.00)
                                JointHelper.CreateRigidJoint(HORSTEEL5397D2, "Structure", VERTSTEEL5397D2, "StartStructure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, lFlangeThickness, 0, -verticalOffset + lCut);

                        }
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of FINL_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    routeConnections.Add(new ConnectionInfo(HORSTEEL5397D2, 1));

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

                    structConnections.Add(new ConnectionInfo(HORSTEEL5397D2, 1));

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
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomString = "";
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");

                bomString = "Console support D2 -" + sizeCodeList.PropValue + " SFS 5397";

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5397_D2" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}


