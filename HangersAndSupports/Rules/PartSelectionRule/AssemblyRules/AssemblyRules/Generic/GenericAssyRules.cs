//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   GenericAssyRules.cs
//   Author       :Vijaya
//   Creation Date:30.July.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  23.Sep.2013     Vijaya   CR-CP-233077  Convert HgrGenAssyRules to C# .Net  
//  31.Oct.2014     PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Linq;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

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

    public static class AssemblyRulesServices
    {
        /// <summary>
        /// This method Gets the largest Nominal pipe diameter.
        /// </summary>
        /// <param name="numberOfPipe">Number of pipes - integer</param>
        /// <param name="support">Support object..</param>
        /// <returns>NominalDiameter</returns>
        /// <code>  
        ///  Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
        ///  int numberOfPipe = supportHelper.SupportedObjects.Count
        /// NominalDiameter largePipeDiameter = AssemblyRulesServices.GetLargePipeDiameter(numberOfPipe, support);
        /// </code>
        public static NominalDiameter GetLargePipeDiameter(int numberOfPipe, Ingr.SP3D.Support.Middle.Support support)
        {
            try
            {
                double[] pipeDiameter = new double[numberOfPipe];
                SupportedHelper supportedHelper = new SupportedHelper(support);
                string unitType = string.Empty;
                for (int i = 0; i < numberOfPipe; i++)
                {
                    PipeObjectInfo pipe = (PipeObjectInfo)supportedHelper.SupportedObjectInfo(i + 1);
                    pipeDiameter[i] = pipe.NominalDiameter.Size;
                    unitType = pipe.NominalDiameter.Units;
                }
                NominalDiameter largePipeDiameter = new NominalDiameter();

                largePipeDiameter.Size = pipeDiameter.Max();
                largePipeDiameter.Units = unitType;

                return largePipeDiameter;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetLargePipeDiameter." +"AssemblyRules" + "," + "Ingr.SP3D.Content.Support.Rules" + "." + "GenericAssyRules" + ". Error:" + e.Message, e);
                throw e1;
            }
        }

    }
    //----------------------------------------------------------------------
    //This Rule retuns array of width and height offset values.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.FrmOffsets
    //----------------------------------------------------------------------
    public class FrmOffsets : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("GenServ_HgrFrmOffsets");
            ReadOnlyCollection<BusinessObject> classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

            ReadOnlyCollection<PropertyValue> allProperties;
            double[] attributeValues = new double[2];
            string offsetSectionSize = string.Empty, sectionSize = string.Empty, currentProperty = string.Empty;
            double widthOffset = 0, heightOffset = 0;
            PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySection", "SectionSize");
            sectionSize = sizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeCodeList.PropValue).ShortDisplayName;
            foreach (BusinessObject item in classItems)
            {
                allProperties = item.GetAllProperties();
                foreach (PropertyValue property in allProperties)
                {
                    currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                    switch (currentProperty)
                    {
                        case ("IJUAHgrGenServFrmOffsets:SectionSize"):
                            offsetSectionSize = ((PropertyValueString)property).PropValue;
                            break;
                        case ("IJUAHgrGenServFrmOffsets:WidthOffset"):
                            widthOffset = (double)((PropertyValueDouble)property).PropValue;
                            break;
                        case ("IJUAHgrGenServFrmOffsets:HeightOffset"):
                            heightOffset = (double)((PropertyValueDouble)property).PropValue;
                            break;
                    }
                }
                if (offsetSectionSize.ToUpper().Equals(sectionSize.ToUpper()))
                {
                    attributeValues[0] = widthOffset;
                    attributeValues[1] = heightOffset;
                    break;
                }
            }
            return attributeValues;
        }
    }
     //----------------------------------------------------------------------
    //This Rule retuns 2.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.IncludePad
    //----------------------------------------------------------------------
    public class IncludePad : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            int[] attributeValues = new int[]{2};
            return attributeValues;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns array of length values.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.Length
    //----------------------------------------------------------------------
    public class Length : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportedHelper supportedHelper = new SupportedHelper(support);
            PipeObjectInfo pipeInfo = (PipeObjectInfo)supportedHelper.SupportedObjectInfo(1);
            NominalDiameter pipeDiameter = pipeInfo.NominalDiameter;

            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("GenServ_HgrOverhang");
            ReadOnlyCollection<BusinessObject> classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            ReadOnlyCollection<PropertyValue> allProperties;
            double[] attributeValues = new double[2];
            double leftOverHang = 0.0, rightOverHang = 0.0;
            string overHangNumOfPipes = string.Empty, overHangNPDUnitType = string.Empty, sectionName = string.Empty, currentProperty = string.Empty;
            double sectionNPD = 0, overHangA = 0, overHangB = 0;
            foreach (BusinessObject item in classItems)
            {
                allProperties = item.GetAllProperties();
                foreach (PropertyValue property in allProperties)
                {
                    currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                    switch (currentProperty)
                    {
                        case ("IJUAHgrGenServHgrOH:NPDUnitType"):
                            overHangNPDUnitType = ((PropertyValueString)property).PropValue;
                            break;
                        case ("IJUAHgrGenServHgrOH:NPD"):
                            sectionNPD = (double)((PropertyValueDouble)property).PropValue;
                            break;
                        case ("IJUAHgrGenServHgrOH:A"):
                            overHangA = (double)((PropertyValueDouble)property).PropValue;
                            break;
                        case ("IJUAHgrGenServHgrOH:B"):
                            overHangB = (double)((PropertyValueDouble)property).PropValue;
                            break;
                    }
                }
                if (sectionNPD < pipeDiameter.Size + 0.001 && sectionNPD > pipeDiameter.Size - 0.001 && overHangNPDUnitType.ToUpper().Equals(pipeDiameter.Units.ToUpper()))
                {
                    leftOverHang = overHangA;
                    rightOverHang = overHangB;
                    break;
                }
            }
            attributeValues[0] = 2 * leftOverHang;
            attributeValues[1] = 2 * rightOverHang;

            return attributeValues;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns array of left over Hang and right over Hang values.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.Overhang
    //----------------------------------------------------------------------
    public class Overhang : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportedHelper supportedHelper = new SupportedHelper(support);
            SupportHelper supportHelper = new SupportHelper(support);
            PipeObjectInfo pipeInfo = (PipeObjectInfo)supportedHelper.SupportedObjectInfo(1);
            int numberOfPipe = supportHelper.SupportedObjects.Count;

            NominalDiameter largePipeDiameter = AssemblyRulesServices.GetLargePipeDiameter(numberOfPipe, support);

            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("GenServ_HgrOverhang");
            ReadOnlyCollection<BusinessObject> classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            ReadOnlyCollection<PropertyValue> allProperties;
            double[] attributeValues = new double[2];
            string overHangNumOfPipes = string.Empty, overHangNPDUnitType = string.Empty, sectionName = string.Empty, currentProperty = string.Empty;
            double sectionNPD = 0, overHangA = 0, overHangB = 0;
            foreach (BusinessObject item in classItems)
            {
                allProperties = item.GetAllProperties();
                foreach (PropertyValue property in allProperties)
                {
                    currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                    switch (currentProperty)
                    {
                        case ("IJUAHgrGenServHgrOH:NPDUnitType"):
                            overHangNPDUnitType = ((PropertyValueString)property).PropValue;
                            break;
                        case ("IJUAHgrGenServHgrOH:NPD"):
                            sectionNPD = (double)((PropertyValueDouble)property).PropValue;
                            break;
                        case ("IJUAHgrGenServHgrOH:A"):
                            overHangA = (double)((PropertyValueDouble)property).PropValue;
                            break;
                        case ("IJUAHgrGenServHgrOH:B"):
                            overHangB = (double)((PropertyValueDouble)property).PropValue;
                            break;
                    }
                }
                if (sectionNPD < largePipeDiameter.Size + 0.001 && sectionNPD > largePipeDiameter.Size - 0.001 && overHangNPDUnitType.ToUpper().Equals(largePipeDiameter.Units.ToUpper()))
                {
                    attributeValues[0] = overHangA;
                    attributeValues[1] = overHangB;
                    break;
                }
            }

            return attributeValues;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns section size value.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.GenericSectionSize
    //----------------------------------------------------------------------
    public class GenericSectionSize : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportHelper supportHelper = new SupportHelper(support);
            GenericHelper genericHelper = new GenericHelper(support);
            PropertyValueCodelist locationCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySupLoc", "SupLocation");

            string numOfPipes = string.Empty, steelStadard = string.Empty, section = string.Empty;
            int pipeCount = supportHelper.SupportedObjects.Count;
            NominalDiameter largePipeDiameter = AssemblyRulesServices.GetLargePipeDiameter(pipeCount, support);
            if (largePipeDiameter.Units == "in")
            {
                largePipeDiameter.Size = 25.4 * largePipeDiameter.Size;
                largePipeDiameter.Units = "mm";
            }
            if (pipeCount > 1)
                numOfPipes = "1+";
            else if (pipeCount == 1)
                numOfPipes = "1";

            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("GenAssy_HgrSectionSize");
            ReadOnlyCollection<BusinessObject> classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            ReadOnlyCollection<PropertyValue> allProperties;
            string sectionNumOfPipes = string.Empty, sectionlNPDUnitType = string.Empty, sectionName = string.Empty, currentProperty = string.Empty;
            int sectionSupLocation = 0;
            double sectionNPDMin = 0, sectionNPDMax = 0;
            foreach (BusinessObject item in classItems)
            {
                allProperties = item.GetAllProperties();
                foreach (PropertyValue property in allProperties)
                {
                    currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                    switch (currentProperty)
                    {
                        case ("IJUAHgrGenSrvSectSize:NumOfPipes"):
                            sectionNumOfPipes = ((PropertyValueString)property).PropValue;
                            break;
                        case ("IJUAHgrGenSrvSectSize:NPDUnitType"):
                            sectionlNPDUnitType = ((PropertyValueString)property).PropValue;
                            break;
                        case ("IJUAHgrGenSrvSectSize:SupLocation"):
                            sectionSupLocation = (int)((PropertyValueCodelist)property).PropValue;
                            break;
                        case ("IJUAHgrGenSrvSectSize:NPDMin"):
                            sectionNPDMin = (double)((PropertyValueDouble)property).PropValue;
                            break;
                        case ("IJUAHgrGenSrvSectSize:NPDMax"):
                            sectionNPDMax = (double)((PropertyValueDouble)property).PropValue;
                            break;
                        case ("IJUAHgrGenSrvSectSize:Section"):
                            sectionName = ((PropertyValueString)property).PropValue;
                            break;
                    }
                }
                if (sectionNumOfPipes.ToUpper().Equals(numOfPipes.ToUpper()) && sectionlNPDUnitType.ToUpper().Equals(largePipeDiameter.Units.ToUpper()) && sectionSupLocation == locationCodeList.PropValue && sectionNPDMin < largePipeDiameter.Size && sectionNPDMax > largePipeDiameter.Size)
                {
                    section = sectionName;
                    break;
                }
            }
            bool value = genericHelper.GetDataByRule("HgrSteelStandardName", (BusinessObject)support, out steelStadard);

            partClass = (PartClass)catalogBaseHelper.GetPartClass("HgrSteelCorrespondence");
            classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            string[] SectionSize = new string[1];
            string steelSize = string.Empty, steelStdName = string.Empty, setionSizeName = string.Empty;
            foreach (BusinessObject classItem in classItems)
            {
                allProperties = classItem.GetAllProperties();
                foreach (PropertyValue property in allProperties)
                {
                    currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                    switch (currentProperty)
                    {
                        case ("IJUAHgrStCorrespondence:Size"):
                            steelSize = ((PropertyValueString)property).PropValue;
                            break;
                        case ("IJUAHgrStCorrespondence:StdName"):
                            steelStdName = ((PropertyValueString)property).PropValue;
                            break;
                        case ("IJUAHgrStCorrespondence:SectionSize"):
                            setionSizeName = ((PropertyValueString)property).PropValue;
                            break;
                    }
                }
                if (steelSize.ToUpper().Equals(section.ToUpper()) && steelStdName.ToUpper().Equals(steelStadard.ToUpper()))
                {
                    SectionSize[0] = setionSizeName;
                    break;
                }
            }
            return SectionSize;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns uBolt part number.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.UBoltSelection
    //----------------------------------------------------------------------
    public class UBoltSelection : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportHelper supportHelper = new SupportHelper(support);
            GenericHelper genericHelper = new GenericHelper(support);
            SupportedHelper supportedHelper = new SupportedHelper(support);
            PropertyValueCodelist locationCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySupLoc", "SupLocation");
            int pipeCount = supportHelper.SupportedObjects.Count;
            string[] uBolt = new string[pipeCount], uBoltType = new string[pipeCount];
            string specificationName = string.Empty, catalogName = string.Empty, uBoltSelSpecName = string.Empty, currentProperty = string.Empty;
            for (int index = 0; index < pipeCount; index++)
            {
                PipeObjectInfo pipeInfo = (PipeObjectInfo)supportedHelper.SupportedObjectInfo(index + 1);
                SpecificationBase pipeSpec = pipeInfo.Spec;
                specificationName = pipeSpec.SpecificationName;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("GenServ_HgrUBoltSel");
                ReadOnlyCollection<BusinessObject> classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                ReadOnlyCollection<PropertyValue> allProperties;

                foreach (BusinessObject item in classItems)
                {
                    allProperties = item.GetAllProperties();
                    foreach (PropertyValue property in allProperties)
                    {
                        currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                        switch (currentProperty)
                        {
                            case ("IJUAHgrGenSrvUBoltSel:SpecName"):
                                uBoltSelSpecName = ((PropertyValueString)property).PropValue;
                                break;
                        }
                    }
                    if (uBoltSelSpecName.ToUpper().Equals(specificationName.ToUpper()))
                    {
                        catalogName = uBoltSelSpecName;
                        break;
                    }
                }
                if (catalogName.Equals(specificationName))
                {
                    string uBoltSelNomDiaUnitType = string.Empty, uBoltSelScheduleThick = string.Empty, uBoltSeluBoltType = string.Empty, schedule = string.Empty;
                    PropertyValueCodelist scheduleCodeList = pipeInfo.Schedule;
                    if (scheduleCodeList.PropValue == 0)
                        scheduleCodeList.PropValue = 100;
                    schedule = scheduleCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(scheduleCodeList.PropValue).ShortDisplayName;

                    int uBoltSelArea = 0;
                    double uBoltSelNomDiaFrom = 0, uBoltSelNomDiaTo = 0;
                    foreach (BusinessObject item in classItems)
                    {
                        allProperties = item.GetAllProperties();
                        foreach (PropertyValue property in allProperties)
                        {
                            currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                            switch (currentProperty)
                            {

                                case ("IJUAHgrGenSrvUBoltSel:SpecName"):
                                    uBoltSelSpecName = ((PropertyValueString)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:NomDiaUnitType"):
                                    uBoltSelNomDiaUnitType = ((PropertyValueString)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:ScheduleThick"):
                                    uBoltSelScheduleThick = ((PropertyValueString)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:Area"):
                                    uBoltSelArea = (int)((PropertyValueCodelist)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:NomDiaFrom"):
                                    uBoltSelNomDiaFrom = (double)((PropertyValueDouble)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:NomDiaTo"):
                                    uBoltSelNomDiaTo = (double)((PropertyValueDouble)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:UBolt"):
                                    uBoltSeluBoltType = ((PropertyValueString)property).PropValue;
                                    break;
                            }
                        }
                        if (uBoltSelSpecName.ToUpper().Equals(specificationName.ToUpper()) && uBoltSelNomDiaUnitType.ToUpper().Equals(pipeInfo.NominalDiameter.Units.ToUpper()) && uBoltSelScheduleThick.ToUpper().Equals(schedule.ToUpper()) && uBoltSelArea == locationCodeList.PropValue && uBoltSelNomDiaFrom < pipeInfo.NominalDiameter.Size && uBoltSelNomDiaTo > pipeInfo.NominalDiameter.Size)
                        {
                            uBoltType[index] = uBoltSeluBoltType;
                            break;
                        }
                    }
                    partClass = (PartClass)catalogBaseHelper.GetPartClass("GenServ_HgrUBoltType");
                    classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    string uBoltNPDUnitType = string.Empty, srvUBoltType = string.Empty, srvUBoltName = string.Empty;
                    double uBoltPipeDia = 0;
                    foreach (BusinessObject item in classItems)
                    {
                        allProperties = item.GetAllProperties();
                        foreach (PropertyValue property in allProperties)
                        {
                            currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                            switch (currentProperty)
                            {
                                case ("IJUAHgrGenServUBolt:NPDUnitType"):
                                    uBoltNPDUnitType = ((PropertyValueString)property).PropValue;
                                    break;
                                case ("IJUAHgrGenServUBolt:Type"):
                                    srvUBoltType = ((PropertyValueString)property).PropValue;
                                    break;
                                case ("IJUAHgrGenServUBolt:PipeDia"):
                                    uBoltPipeDia = (double)((PropertyValueDouble)property).PropValue;
                                    break;
                                case ("IJUAHgrGenServUBolt:UBolt"):
                                    srvUBoltName = ((PropertyValueString)property).PropValue;
                                    break;
                            }
                        }
                        if (uBoltPipeDia < pipeInfo.NominalDiameter.Size + 0.001 && uBoltPipeDia > pipeInfo.NominalDiameter.Size - 0.001 && uBoltNPDUnitType.ToUpper().Equals(pipeInfo.NominalDiameter.Units.ToUpper()) && srvUBoltType.ToUpper().Equals(uBoltType[index].ToUpper()))
                        {
                            uBolt[index] = srvUBoltName;
                            break;
                        }
                    }
                }
                else
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Data is not available for the Pipe Spec " + specificationName, "", "Assy_SS3_V.cs", 1);
                    return null;
                }
            }
            return uBolt;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns uBolt type.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.UBoltType
    //----------------------------------------------------------------------
    public class UBoltType : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportHelper supportHelper = new SupportHelper(support);
            GenericHelper genericHelper = new GenericHelper(support);
            SupportedHelper supportedHelper = new SupportedHelper(support);
            PropertyValueCodelist locationCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrGenAssySupLoc", "SupLocation");
            int pipeCount = supportHelper.SupportedObjects.Count;
            string[] uBolt = new string[pipeCount];
            string specificationName = string.Empty, catalogName = string.Empty;
            for (int index = 0; index < pipeCount; index++)
            {
                PipeObjectInfo pipeInfo = (PipeObjectInfo)supportedHelper.SupportedObjectInfo(index + 1);
                SpecificationBase pipeSpec = pipeInfo.Spec;
                specificationName = pipeSpec.SpecificationName;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("GenServ_HgrUBoltSel");
                ReadOnlyCollection<BusinessObject> classItems = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                ReadOnlyCollection<PropertyValue> allProperties;
                string uBoltSelSpecName = string.Empty, currentProperty = string.Empty;
                foreach (BusinessObject item in classItems)
                {
                    allProperties = item.GetAllProperties();
                    foreach (PropertyValue property in allProperties)
                    {
                        currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                        switch (currentProperty)
                        {
                            case ("IJUAHgrGenSrvUBoltSel:SpecName"):
                                uBoltSelSpecName = ((PropertyValueString)property).PropValue;
                                break;
                        }
                    }
                    if (uBoltSelSpecName.ToUpper().Equals(specificationName.ToUpper()))
                    {
                        catalogName = uBoltSelSpecName;
                        break;
                    }
                }
                if (catalogName.Equals(specificationName))
                {
                    string uBoltSelNomDiaUnitType = string.Empty, uBoltSelScheduleThick = string.Empty, uBoltSeluBoltType = string.Empty, schedule = string.Empty;
                    PropertyValueCodelist scheduleCodeList = pipeInfo.Schedule;
                    if (scheduleCodeList.PropValue == 0)
                        scheduleCodeList.PropValue = 100;
                    schedule = scheduleCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(scheduleCodeList.PropValue).ShortDisplayName;

                    int uBoltSelArea = 0;
                    double uBoltSelNomDiaFrom = 0, uBoltSelNomDiaTo = 0;
                    foreach (BusinessObject item in classItems)
                    {
                        allProperties = item.GetAllProperties();
                        foreach (PropertyValue property in allProperties)
                        {
                            currentProperty = property.PropertyInfo.InterfaceInfo.Name + ":" + property.PropertyInfo.Name;
                            switch (currentProperty)
                            {
                                case ("IJUAHgrGenSrvUBoltSel:SpecName"):
                                    uBoltSelSpecName = ((PropertyValueString)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:NomDiaUnitType"):
                                    uBoltSelNomDiaUnitType = ((PropertyValueString)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:ScheduleThick"):
                                    uBoltSelScheduleThick = ((PropertyValueString)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:Area"):
                                    uBoltSelArea = (int)((PropertyValueCodelist)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:NomDiaFrom"):
                                    uBoltSelNomDiaFrom = (double)((PropertyValueDouble)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:NomDiaTo"):
                                    uBoltSelNomDiaTo = (double)((PropertyValueDouble)property).PropValue;
                                    break;
                                case ("IJUAHgrGenSrvUBoltSel:UBolt"):
                                    uBoltSeluBoltType = ((PropertyValueString)property).PropValue;
                                    break;
                            }
                        }
                        if (uBoltSelSpecName.ToUpper().Equals(specificationName.ToUpper()) && uBoltSelNomDiaUnitType.ToUpper().Equals(pipeInfo.NominalDiameter.Units.ToUpper()) && uBoltSelScheduleThick.ToUpper().Equals(schedule.ToUpper()) && uBoltSelArea == locationCodeList.PropValue && uBoltSelNomDiaFrom < pipeInfo.NominalDiameter.Size && uBoltSelNomDiaTo > pipeInfo.NominalDiameter.Size)
                        {
                            uBolt[index] = uBoltSeluBoltType;
                            break;
                        }
                    }

                }
                else
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Data is not available for the Pipe Spec " + specificationName, "", "Assy_SS3_V.cs", 1);
                    return null;
                }
            }
            return uBolt;
        }
    }
}
