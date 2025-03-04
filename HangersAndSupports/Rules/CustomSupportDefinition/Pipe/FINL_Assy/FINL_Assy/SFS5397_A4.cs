//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_A4.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_A4
//   Author       :  Vijaya
//   Creation Date:  27.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27-Jun-2013     Vijaya   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;

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
    public class SFS5397_A4 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSTEEL15397A4 = "HORSTEEL15397_A4";
        private const string HORSTEEL25397A4 = "HORSTEEL25397_A4";
        private const string L5397A4 = "L5397_A4";
        private const string CSECTION15397A4 = "CSECTION15397_A4";
        private const string CSECTION25397A4 = "CSECTION25397_A4";
        private const string VERTSTEEL15397A4 = "VERTSTEEL15397_A4";
        private const string VERTSTEEL25397A4 = "VERTSTEEL25397_A4";
        private string size3;
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

                    if (horizontalLength >= 500.00 / 1000.00)// with clamps
                    {
                        parts.Add(new PartInfo(HORSTEEL15397A4, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(HORSTEEL25397A4, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(L5397A4, "Utility_GENERIC_L_1"));
                        parts.Add(new PartInfo(CSECTION15397A4, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION25397A4, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(VERTSTEEL15397A4, "Utility_CUTBACK_L2_1"));
                        parts.Add(new PartInfo(VERTSTEEL25397A4, "Utility_CUTBACK_L2_1"));
                    }
                    else
                    {
                        parts.Add(new PartInfo(HORSTEEL15397A4, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(HORSTEEL25397A4, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(L5397A4, "Utility_GENERIC_L_1"));
                        parts.Add(new PartInfo(CSECTION15397A4, "Utility_GENERIC_C_1"));
                        parts.Add(new PartInfo(CSECTION25397A4, "Utility_GENERIC_C_1"));
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
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                pipeDiameter = pipeInfo.OutsideDiameter + pipeInfo.InsulationThickness * 2;

                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                NominalDiameter minNominalDiameter = new NominalDiameter(), maxNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";

                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_A4.cs", 129);
                    return;
                }
                double maxLength = 0.0, MA = 0.0, MP = 0.0, FQ = 0.0, FU = 0.0;
                PropertyValue CSize, LSize, BraceSize;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size3", size3);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_A4_Dim", "IJUAFINLSrv_SFS5397_A4_Dim", "C_Size", parameter, out CSize);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_A4_Dim", "IJUAFINLSrv_SFS5397_A4_Dim", "L_Size", parameter, out LSize);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_A4_Dim", "IJUAFINLSrv_SFS5397_A4_Dim", "Brace_Size", parameter, out BraceSize);
                string cSize = CSize.ToString(), lSize = LSize.ToString(), braceSize = BraceSize.ToString();


                maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A4_Dim", "IJUAFINLSrv_SFS5397_A4_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_A4_Dim", "Size3", size3);
                MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A4_Dim", "IJUAFINLSrv_SFS5397_A4_Dim", "MA", "IJUAFINLSrv_SFS5397_A4_Dim", "Size3", size3);
                MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A4_Dim", "IJUAFINLSrv_SFS5397_A4_Dim", "MP", "IJUAFINLSrv_SFS5397_A4_Dim", "Size3", size3);
                FQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A4_Dim", "IJUAFINLSrv_SFS5397_A4_Dim", "FQ", "IJUAFINLSrv_SFS5397_A4_Dim", "Size3", size3);
                FU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A4_Dim", "IJUAFINLSrv_SFS5397_A4_Dim", "FU", "IJUAFINLSrv_SFS5397_A4_Dim", "Size3", size3);


                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(FQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(FU, "IJUAFINLFQFU", "FU");

                double cWidth, cFlangeThickness, cWebThickness, cDepth, lWidth, lFlangeThickness, lWebThickness, lDepth, bWidth, bFlangeThickness, bWebThickness, bDepth;
                FINLAssemblyServices.GetCrossSectionDimensions("EURO-OTUA-2002", "C", cSize, out  cWidth, out   cFlangeThickness, out  cWebThickness, out  cDepth);
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "L", lSize, out  lWidth, out   lFlangeThickness, out  lWebThickness, out  lDepth);
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "L", braceSize, out  bWidth, out   bFlangeThickness, out  bWebThickness, out  bDepth);


                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_A4cs", 162);
                    return;
                }
                double structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                horizontalLength = structDistance + overHang;

                if (structDistance + overHang > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + (maxLength * 1000).ToString() + "mm exceeded.", "", "SFS5397_A4cs", 169);

                if (overHang < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + (pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_A4.cs", 172);

                double verticalLength = (400.00 / 1000.00 - 1.0 / 2.0 * Math.Round((bDepth / Math.Tan(Math.PI / 4) + (bDepth / Math.Tan(Math.PI / 2))), 2));
                double lCut = bDepth / Math.Sin(Math.PI / 4);
                double VertSteelOffset = (400.00 / 1000.00 - lCut * Math.Cos(Math.PI / 4)) * Math.Cos(Math.PI / 4);


                string inputBOM = cSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[HORSTEEL15397A4].SetPropertyValue(cWidth, "IJOAHgrUtility_Width", "Width");
                componentDictionary[HORSTEEL15397A4].SetPropertyValue(cDepth, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[HORSTEEL15397A4].SetPropertyValue(horizontalLength, "IJOAHgrUtility_L", "L");
                componentDictionary[HORSTEEL15397A4].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[HORSTEEL15397A4].SetPropertyValue(cWebThickness, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[HORSTEEL15397A4].SetPropertyValue(inputBOM, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                componentDictionary[HORSTEEL25397A4].SetPropertyValue(cWidth, "IJOAHgrUtility_Width", "Width");
                componentDictionary[HORSTEEL25397A4].SetPropertyValue(cDepth, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[HORSTEEL25397A4].SetPropertyValue(horizontalLength, "IJOAHgrUtility_L", "L");
                componentDictionary[HORSTEEL25397A4].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[HORSTEEL25397A4].SetPropertyValue(cWebThickness, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[HORSTEEL25397A4].SetPropertyValue(inputBOM, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                double lLength = 400.00 / 1000.00;
                inputBOM = lSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[L5397A4].SetPropertyValue(lWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                componentDictionary[L5397A4].SetPropertyValue(lDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                componentDictionary[L5397A4].SetPropertyValue(lLength, "IJOAHgrUtility_GENERIC_L", "L");
                componentDictionary[L5397A4].SetPropertyValue(lFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                componentDictionary[L5397A4].SetPropertyValue(inputBOM, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");


                //additional C - section
                Double lC1 = 550.00 / 1000.00 - lDepth / 2, lC2 = cDepth;
                String plateBOM = cSize + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lC1, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                //four hole plate
                componentDictionary[CSECTION15397A4].SetPropertyValue(cWidth, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION15397A4].SetPropertyValue(cDepth, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION15397A4].SetPropertyValue(lC1, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION15397A4].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION15397A4].SetPropertyValue(cWebThickness, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION15397A4].SetPropertyValue(plateBOM, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                plateBOM = cSize + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lC2, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[CSECTION25397A4].SetPropertyValue(cWidth, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION25397A4].SetPropertyValue(cDepth, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION25397A4].SetPropertyValue(lC2, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION25397A4].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION25397A4].SetPropertyValue(cWebThickness, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION25397A4].SetPropertyValue(plateBOM, "IJOAHgrUtility_BomDesc", "InputBomDesc");



                if (horizontalLength >= 500.00 / 1000.00)
                {
                    //vertical steel is 2 L cut back steels
                    string bom = lSize + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                    componentDictionary[VERTSTEEL15397A4].SetPropertyValue(bWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                    componentDictionary[VERTSTEEL15397A4].SetPropertyValue(bDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                    componentDictionary[VERTSTEEL15397A4].SetPropertyValue(verticalLength, "IJOAHgrUtility_GENERIC_L", "L");
                    componentDictionary[VERTSTEEL15397A4].SetPropertyValue(bFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                    componentDictionary[VERTSTEEL15397A4].SetPropertyValue(-Math.PI / 4, "IJOAHgrUtility_CUTBACK", "ANGLE");//rotae about y-axis
                    componentDictionary[VERTSTEEL15397A4].SetPropertyValue(-Math.PI / 2, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[VERTSTEEL15397A4].SetPropertyValue(bom, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                    componentDictionary[VERTSTEEL25397A4].SetPropertyValue(bWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                    componentDictionary[VERTSTEEL25397A4].SetPropertyValue(bDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                    componentDictionary[VERTSTEEL25397A4].SetPropertyValue(verticalLength, "IJOAHgrUtility_GENERIC_L", "L");
                    componentDictionary[VERTSTEEL25397A4].SetPropertyValue(bFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                    componentDictionary[VERTSTEEL25397A4].SetPropertyValue(-Math.PI / 2, "IJOAHgrUtility_CUTBACK", "ANGLE");//rotae about y-axis
                    componentDictionary[VERTSTEEL25397A4].SetPropertyValue(-Math.PI / 4, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[VERTSTEEL25397A4].SetPropertyValue(bom, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                }

                double theta1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                if (Math.Abs(theta1) < Math.PI / 2)
                {
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL15397A4, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Z, -overHang, cDepth / 2 - (cDepth / 2 + pipeDiameter / 2 + shoeHeight), cDepth / 2);
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL25397A4, "EndCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, -overHang, cDepth / 2 - (cDepth / 2 + pipeDiameter / 2 + shoeHeight), -cDepth / 2);
                    JointHelper.CreateRigidJoint(HORSTEEL15397A4, "EndCap", CSECTION15397A4, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeZ, 0, 0, 0);
                    JointHelper.CreateRigidJoint(HORSTEEL15397A4, "BeginCap", CSECTION25397A4, "BeginCap", Plane.XY, Plane.ZX, Axis.X, Axis.Z, -cDepth, 0, 0);
                    JointHelper.CreateRigidJoint(CSECTION15397A4, "BeginCap", L5397A4, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, lLength / 2 + cDepth / 2, 0);

                    if (horizontalLength >= 500.00 / 1000.00)
                    {
                        JointHelper.CreateRigidJoint(HORSTEEL15397A4, "BeginCap", VERTSTEEL15397A4, "StartStructure", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, 0, cDepth, horizontalLength - VertSteelOffset);
                        JointHelper.CreateRigidJoint(HORSTEEL25397A4, "EndCap", VERTSTEEL25397A4, "EndStructure", Plane.XY, Plane.ZX, Axis.X, Axis.Z, 0, cDepth, -horizontalLength + VertSteelOffset);

                    }
                }
                else
                {
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL15397A4, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, -overHang, -cDepth / 2 + (cDepth / 2 + pipeDiameter / 2 + shoeHeight), -cDepth / 2);
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL25397A4, "EndCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, -overHang, -cDepth / 2 + (cDepth / 2 + pipeDiameter / 2 + shoeHeight), cDepth / 2);
                    JointHelper.CreateRigidJoint(HORSTEEL15397A4, "EndCap", CSECTION15397A4, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.NegativeZ, 0, 0, 0);
                    JointHelper.CreateRigidJoint(HORSTEEL15397A4, "BeginCap", CSECTION25397A4, "BeginCap", Plane.XY, Plane.ZX, Axis.X, Axis.Z, -cDepth, 0, 0);
                    JointHelper.CreateRigidJoint(CSECTION15397A4, "BeginCap", L5397A4, "Structure", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, lLength / 2 + cDepth / 2, 0);

                    if (horizontalLength >= 500.00 / 1000.00)
                    {
                        JointHelper.CreateRigidJoint(HORSTEEL15397A4, "BeginCap", VERTSTEEL15397A4, "StartStructure", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, 0, cDepth, horizontalLength - VertSteelOffset);//0, -90, -90  3308
                        JointHelper.CreateRigidJoint(HORSTEEL25397A4, "EndCap", VERTSTEEL25397A4, "EndStructure", Plane.XY, Plane.ZX, Axis.X, Axis.Z, 0, cDepth, -horizontalLength + VertSteelOffset);//0, 90, 90 11500

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

                    routeConnections.Add(new ConnectionInfo(HORSTEEL15397A4, 1));

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

                    structConnections.Add(new ConnectionInfo(HORSTEEL15397A4, 1));

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

                bomString = "Console support A4 -" + sizeCodeList.PropValue + " SFS 5397";

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5397_A4" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}




