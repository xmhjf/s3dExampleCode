//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   SingleWeld class 
//                
//                 
//                    
//
//      History:
//      Sept 25th, 2014   Created by VIBGYOR
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using System.Runtime.InteropServices;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// Single Weld Rule Implements Planning Joint Rule Base Class
    /// </summary>
    public class SingleWeld : PlanningJointRuleBase
    {
        # region Global Variables

        S3D_COMPONENT rulesXML;
        private bool xmlParsed = false;
        private string sharedContentPath = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.SymbolShare;
        private double blockIntersectionJointLength = 0.0;
        # endregion

        #region Methods

        /// <summary>
        /// Deserialize input XML 
        /// </summary>
        /// <param name="xmlPath"></param>
        private void DeserializeXML(string xmlPath)
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Parse;
            settings.ValidationType = ValidationType.Schema;
            settings.IgnoreComments = true;
            settings.CloseInput = true;
            XmlReader xmlreader = XmlReader.Create(xmlPath, settings);
            XmlSerializer myserializer = new XmlSerializer(typeof(S3D_COMPONENT));
            rulesXML = (S3D_COMPONENT)myserializer.Deserialize(xmlreader);
            xmlParsed = true;
        }

        /// <summary>
        /// GenerateGeometries method creates Planning Joint geometry from the input physical connection
        /// </summary>
        /// <param name="physicalConn"></param>
        /// <returns>ReadOnlyCollection<ComplexString3d></returns>
        public override ReadOnlyCollection<ComplexString3d> GenerateGeometries(BusinessObject physicalConn)
        {
            #region Input Error Handling
            if (physicalConn == null)
            {
                throw new ArgumentNullException("Physical connection");
            }
            #endregion //Input Error Handling
            ReadOnlyCollection<ComplexString3d> geometryCollections = null;
            try
            {
                if (xmlParsed == false)
                {
                    string xmlPath = string.Empty;

                    if (!string.IsNullOrEmpty(base.Argument) && (File.Exists(sharedContentPath + base.Argument) == true))
                    {
                        xmlPath = sharedContentPath + base.Argument;
                    }
                    else
                    {
                        xmlPath = sharedContentPath + "\\Planning\\PlanningJoint\\PlanningJointRules.xml";
                    }

                    DeserializeXML(xmlPath);

                    // BlockIntersectionJointLength > 0 then - creates BlockIntersectionJoints with length specified in as BlockIntersectionJointLength.
                    //else full planning joints welds will be created.
                    S3D_ARG arg = GetArgument(rulesXML, "PlanningJointRules", "General", "BlockIntersectionJointLength");
                    if (arg != null)
                    {
                        UnitType unitType;
                        UnitName unitName;
                        Enum.TryParse<UnitType>(Convert.ToString(arg.UNIT_TYPE), true, out unitType);
                        Enum.TryParse<UnitName>(Convert.ToString(arg.UNIT), true, out unitName);
                        double xmlJointLength;
                        Double.TryParse(arg.VALUE, out xmlJointLength);
                        blockIntersectionJointLength = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(unitType, xmlJointLength, unitName);
                    }
                }

                SetPhysicalConnection(physicalConn);
                SetPlnJointSplitParameters(false, 0.0);

                // singlePJForMultiLumpPC => True - If we need to generate a single Planning Joint for parts, which have edge features.
                // singlePJForMultiLumpPC => False - If we need default behaviour(i.e Multiple segments of Planning Joints).
                string argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "General", "SinglePJForMultiLumpPC");
                bool singlePJForMultiLumpPC = Convert.ToBoolean(argValue);                
               
                // Get the interface name related to PC parameters
                string paramInterfaceName = GetInterfaceRelatedToConnectionParamters();
                Ingr.SP3D.Planning.Middle.WeldType weldType = GetWeldType();
                Ingr.SP3D.Planning.Middle.WeldCategory weldCategory = GetWeldCategory(paramInterfaceName, weldType);                

                //weldCategoryName is Note needed  for default Weld case
                geometryCollections = GeneratePlnJointGeometry(WeldGeometry.SingleWeld, singlePJForMultiLumpPC, weldType, weldCategory, blockIntersectionJointLength);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }

            return geometryCollections;
        }

        /// <summary>
        ///  EvaluateProperties updates the planning joint properties
        /// </summary>
        /// <param name="plnJoint"></param>
        /// <param name="refPlnJoint"></param>
        /// <param name="updateProdPropsOnly"></param>
        /// <returns></returns>
        public override void EvaluateProperties(PlanningJoint plnJoint, PlanningJoint refPlnJoint, bool updatePlanningProps, bool updateProductionProps)
        {
            #region Input Error Handling
            if (plnJoint == null)
            {
                throw new ArgumentNullException("Planning Joint");
            }
            #endregion //Input Error Handling
            try
            {
                if (xmlParsed == false)
                {
                    string xmlPath = string.Empty;

                    if (!string.IsNullOrEmpty(base.Argument) && (Directory.Exists(sharedContentPath + "\\" + base.Argument) == true) && (File.Exists(sharedContentPath + "\\" + base.Argument) == true))
                    {
                        xmlPath = sharedContentPath + "\\" + base.Argument;
                    }
                    else
                    {
                        xmlPath = sharedContentPath + "\\" + "Planning\\PlanningJoint\\PlanningJointRules.xml";
                    }

                    DeserializeXML(xmlPath);
                }

                SetPlnJoint(plnJoint);

                if (updatePlanningProps == true && plnJoint.PlanningPropertieByRule == 1)
                {
                    string paramInterfaceName = string.Empty;

                    SetPhysicalConnection();
                    
                    // Property Weld Type and Weld Category
                    if (refPlnJoint == null)
                    {
                        // Get the interface name related to PC parameters
                        paramInterfaceName = GetInterfaceRelatedToConnectionParamters();

                        Ingr.SP3D.Planning.Middle.WeldType weldType = GetWeldType();
                        Ingr.SP3D.Planning.Middle.WeldCategory weldCategory = GetWeldCategory(paramInterfaceName, weldType);
                        
                        plnJoint.WeldType = weldType;
                        plnJoint.WeldCategory = weldCategory;
                    }
                    else
                    {
                        plnJoint.WeldType = refPlnJoint.WeldType;
                        plnJoint.WeldCategory = refPlnJoint.WeldCategory;
                    }

                    // Property Weld Length
                    string argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "General", "SinglePJForMultiLumpPC");
                    bool singlePJForMultiLumpPC = Convert.ToBoolean(argValue);
                    plnJoint.WeldLength = GetLength(singlePJForMultiLumpPC);

                    // Weld Field Fit Length - use the weld length as default
                    plnJoint.WeldFieldFitLength = plnJoint.WeldLength;

                     // Property Weld Type and Weld Category
                    if (refPlnJoint == null)
                    {
                        if (plnJoint.WeldType != WeldType.ButtWeld) // Butt weld does not have fillet measure method
                        {
                            try
                            {
                                // Property Weld Leg Length and Weld Throat Length
                                PropertyValue propertyVal = GetFilletMeasureMethod(paramInterfaceName, "FilletMeasureMethod");

                                if (propertyVal != null)
                                {
                                    PropertyValueDouble propValDouble = (PropertyValueDouble)propertyVal;
                                    long filletMeasureMethod = Convert.ToInt64(propValDouble.PropValue);

                                    if (filletMeasureMethod == 65536L) // Leg
                                    {
                                        plnJoint.WeldFilletMeasureMethod = WeldFilletMeasureMethod.Leg;
                                    }
                                    else // 65537 - Throat
                                    {
                                        plnJoint.WeldFilletMeasureMethod = WeldFilletMeasureMethod.Throat;
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                MiddleServiceProvider.ErrorLogger.Log("Failed to fill the WeldFilletMeasureMethod property because of " + paramInterfaceName + " does not exist." + e.Message);
                            }
                            
                        }
                    }
                    else
                    {
                        plnJoint.WeldFilletMeasureMethod = refPlnJoint.WeldFilletMeasureMethod;
                    }

                    // Property Weld Side
                    WeldSide weldSide = GetWeldSide();
                    plnJoint.WeldSide = weldSide;

                    if (plnJoint.WeldType != WeldType.ButtWeld)
                    {
                        string strAttributeName = "MoldedFillet";

                        if (weldSide == WeldSide.Molded || weldSide == WeldSide.Reference)
                        {
                            strAttributeName = "MoldedFillet";
                        }
                        else
                        {
                            strAttributeName = "AntiMoldedFillet";
                        }

                        plnJoint.WeldFilletSize = GetWeldDimension(strAttributeName);
                    }

                    // Property Weld RootGap
                    if (plnJoint.WeldType != WeldType.LapWeld)
                    {
                        // Property Weld RootGap
                        if (refPlnJoint == null)
                        {
                            plnJoint.WeldRootGap = GetWeldRootGap();
                        }
                        else
                        {
                            plnJoint.WeldRootGap = refPlnJoint.WeldRootGap;
                        }
                    }

                    // Property Weld Inclination
                    plnJoint.WeldInclination = GetInclination();

                    // Property Weld Rotation
                    plnJoint.WeldRotation = GetRotation();

                    // Property Weld Position
                    plnJoint.WeldPosition = GetWeldPosition();

                    //property weld anti-ref-position
                    if (plnJoint.WeldType == WeldType.ButtWeld)
                    {
                        if (plnJoint.WeldPosition == WeldPosition.Overhead)
                        {
                            plnJoint.WeldAntiRefPosition = WeldPosition.Flat;
                        }
                        else if (plnJoint.WeldPosition == WeldPosition.Flat)
                        {
                            plnJoint.WeldAntiRefPosition = WeldPosition.Overhead;
                        }
                        else
                        {
                            plnJoint.WeldAntiRefPosition = plnJoint.WeldPosition;
                        }
                    }
                    else if (plnJoint.WeldType == WeldType.TeeWeld)
                    {
                        plnJoint.WeldAntiRefPosition = plnJoint.WeldPosition;
                    }
                    else
                    {
                        plnJoint.WeldAntiRefPosition = WeldPosition.NotApplicable;
                    }

                    // Property Weld Shape
                    plnJoint.WeldShape = GetWeldShape();

                    if (refPlnJoint == null)
                    {
                        // Property Weld Edge Prep
                        plnJoint.WeldEdgePrep = GetEdgePrep();

                        // Property Weld Note
                        plnJoint.WeldNote = GetTailNotes();

                        // Property Weld Ratio
                        plnJoint.WeldRatio = 100;

                        try
                        {
                            //property Weld thickness
                            plnJoint.WeldThickness = GetWeldThickness();
                        }
                        catch (Exception)
                        {
                            // Do nothing
                        }
                    }
                    else
                    {
                        // Property Weld Edge Prep
                        plnJoint.WeldEdgePrep = refPlnJoint.WeldEdgePrep;

                        // Property Weld Note
                        plnJoint.WeldNote = refPlnJoint.WeldNote;

                        // Property Weld Ratio
                        plnJoint.WeldRatio = refPlnJoint.WeldRatio;

                        //property Weld thickness
                        plnJoint.WeldThickness = refPlnJoint.WeldThickness;
                    }
                }

                if (updateProductionProps == true && plnJoint.ProductionPropertieByRule == 1)
                {
                    string argValue = string.Empty;

                    SetPhysicalConnection();

                    // Property Weld Direction
                    plnJoint.WeldDirection = GetWeldDirection();

                    if (refPlnJoint == null)
                    {
                        // Property Weld Process
                        argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldProcess");
                        plnJoint.WeldProcess = (WeldProcess)Convert.ToInt32(argValue);

                        // Property WeldStage
                        argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldStage");
                        plnJoint.WeldStage = (WeldStage)Convert.ToInt32(argValue);

                        // Property Weld Equipment
                        AssemblyBase assembly = plnJoint.AssemblyParent as AssemblyBase;
                        if (assembly != null)
                        {
                            Ingr.SP3D.ReferenceData.Middle.ProductionEquipment weldEquipment = assembly.GetEquipment(AssemblyEquipmentType.WeldEquipment);
                            if (weldEquipment != null && !string.IsNullOrEmpty(weldEquipment.Name))
                            {
                                plnJoint.WeldEquipment = weldEquipment.Name;
                            }
                            else
                            {
                                plnJoint.WeldEquipment = string.Empty;
                            }
                        }
                        else
                        {
                            plnJoint.WeldEquipment = string.Empty;
                        }

                        // Property Weld Assembly Type
                        plnJoint.WeldAssemblyType = GetAssemblyType();
                    }
                    else
                    {
                        // Property Weld Process
                        plnJoint.WeldProcess = refPlnJoint.WeldProcess;

                        // Property WeldStage
                        plnJoint.WeldStage = refPlnJoint.WeldStage;

                        AssemblyBase assembly = plnJoint.AssemblyParent as AssemblyBase;
                        if (assembly != null)
                        {
                            // Property Weld Equipment
                            plnJoint.WeldEquipment = refPlnJoint.WeldEquipment;

                            // Property Weld Assembly Type
                            plnJoint.WeldAssemblyType = refPlnJoint.WeldAssemblyType;
                        }
                        else //If the Planning joint parent is Project root.
                        {
                            // Property Weld Equipment
                            plnJoint.WeldEquipment = string.Empty;

                            // Property Weld Assembly Type
                            plnJoint.WeldAssemblyType = 0;
                        }
                    }

                    // Property Weld Accessibility
                    argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldAccessibility");
                    plnJoint.WeldAccessibility = (WeldAccessibility) Convert.ToInt32(argValue);

                    // Property WPS Number
                    argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WPSNumber");
                    plnJoint.WPSNumber = argValue;

                    // Property Weld Classification
                    argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldClassification");
                    plnJoint.WeldClassification = argValue;

                    // Property Number of Weld Passes
                    argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldPasses");
                    plnJoint.WeldPasses = Convert.ToInt32(argValue);

                    // Property Weld location
                    argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldLocation");
                    plnJoint.WeldLocation = (WeldLocation)Convert.ToInt32(argValue);

                    if (refPlnJoint == null)
                    {
                        // Property Weld Material Name
                        argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldMaterialName");
                        plnJoint.WeldMaterialName = argValue;

                        // Property Weld Material Grade
                        argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldMaterialGrade");
                        plnJoint.WeldMaterialGrade = argValue;

                        // Property Weld Testing Type
                        argValue = GetArgumentValue(rulesXML, "PlanningJointRules", "Production", "WeldTestingType");
                        plnJoint.WeldTestingType = (WeldTestingType) Convert.ToInt32(argValue);
                    }
                    else
                    {
                        // Property Weld Material Name
                        plnJoint.WeldMaterialName = refPlnJoint.WeldMaterialName;

                        // Property Weld Material Grade
                        plnJoint.WeldMaterialGrade = refPlnJoint.WeldMaterialGrade;

                        // Property Weld Testing Type
                        plnJoint.WeldTestingType = refPlnJoint.WeldTestingType;
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        /// <summary>
        /// Gets The Weld Category
        /// </summary>
        /// <param name="weldType"></param>
        /// <param name="propertyVal">weld category property value</param>
        private Ingr.SP3D.Planning.Middle.WeldCategory GetWeldCategory(string paramInterfaceName, WeldType weldType)
        {
            Ingr.SP3D.Planning.Middle.WeldCategory weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.Default;

            try
            {
                PropertyValue propertyVal;
                long lCategory = 0;
                
                
                if (weldType == WeldType.ButtWeld)
                {
                    propertyVal = GetWeldCategory(paramInterfaceName, "ButtCategory");

                    PropertyValueDouble propValDouble = (PropertyValueDouble)propertyVal;
                    lCategory = Convert.ToInt64(propValDouble.PropValue);
                    
                    // For BUTT weld the code list values are different
                    if (lCategory == 65536L)
                    {
                        weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.OneSidedBevel;
                    }
                    else // 65537
                    {
                        weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.TwoSidedBevel;
                    }
                }
                else
                {
                    propertyVal = GetWeldCategory(paramInterfaceName, "Category");
                    PropertyValueDouble propValDouble = (PropertyValueDouble)propertyVal;

                    lCategory = Convert.ToInt64(propValDouble.PropValue);

                    if (lCategory != 0)
                    {
                        if (lCategory == 65536L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.Default;
                        }
                        else if (lCategory == 65537L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.Normal;
                        }
                        else if (lCategory == 65538L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.Deep;
                        }
                        else if (lCategory == 65539L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.Full;
                        }
                        else if (lCategory == 65540L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.Chain;
                        }
                        else if (lCategory == 65541L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.Staggered;
                        }
                        else if (lCategory == 65542L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.OneSidedBevel;
                        }
                        else if (lCategory == 65543L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.TwoSidedBevel;
                        }
                        else if (lCategory == 65544L)
                        {
                            weldCategory = Ingr.SP3D.Planning.Middle.WeldCategory.Chill;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }

            return weldCategory;
        }

        /// <summary>
        /// Gets The Interface related to PC Paramters
        /// </summary>
        private string GetInterfaceRelatedToConnectionParamters()
        {
            try
            {
                string itemName = GetConnectionSmartItemName();
                return "IASMPhysConnRules_" + itemName + "Parm";
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return string.Empty;
        }

        /// <summary>
        /// 1.if it is create mode
        ///     a.create root folder and add relation to assembly.
        ///     b.then create leaf folder and add relation to root folder then return leaf folder.
        /// 2. if it update mode
        ///     a.compute the folder and return it.
        /// </summary>
        /// <param name="computemode"></param>
        /// <param name="planningJoint"></param>
        /// <param name="planningJointFolder"></param>
        public override void GenerateHierarchy(PlnJointComputeMode computemode, PlanningJoint planningJoint, out PlanningJointFolder planningJointFolder)
        {
            planningJointFolder = null;

            #region Input Error Handling
            if (planningJoint == null)
            {
                throw new ArgumentNullException("Input argument is null");
            }
            #endregion //Input Error Handling

            try
            {
                FindFolderForPlnJoint(planningJoint, computemode, out planningJointFolder);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        # endregion
    }
}
