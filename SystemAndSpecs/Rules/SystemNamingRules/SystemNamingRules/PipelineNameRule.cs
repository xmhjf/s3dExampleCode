/*******************************************************************************
  Copyright (C) 2001, Intergraph Corporation.  All rights reserved.  
  Project: SystemsAndSpecs SystemNameRules
  Class:   PipelineNameRule
  Abstract: The file contains a sample implementation of a naming rule for Pipeline systems.
  
  10/05/2014   Mishra,Anamay  CR-CP-270276  System and Specs VB6 Rules need to be replaced;
 *****************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Systems.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.System.Rules
{
    /// <summary>
    /// Class implementing the name rule for PipeLine system
    /// </summary>
    public class PipelineNameRule : NameRuleBase
    {

        string duplicateName = string.Empty;
        string hasChanged = string.Empty;
        string blankName = string.Empty;
        string systemName = string.Empty;
        string newName = string.Empty;
        string genericSystem = string.Empty;
        string conduitSystem = string.Empty;
        string ductingSystem = string.Empty;
        string equipmentSystem = string.Empty;
        string pipelineSystem = string.Empty;
        string pipingSystem = string.Empty;
        string structuralSystem = string.Empty;
        string unitSystem = string.Empty;
        string areaSystem = string.Empty;       

        public PipelineNameRule()
        {
            InitStrings();
        }

        /// <summary>
        /// Computes the name for the given entity. 
        /// </summary>
        /// <param name="entity">The business object whose name is being computed.</param>
        /// <param name="parents">The parentsNaming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("The entity to be named is null");
            }

            if (parents == null)
            {
                throw new ArgumentNullException("The naming parents of the business object to be named are null");
            }

            if (activeEntity == null)
            {
                throw new ArgumentNullException("The name rule active entity associated to the business object to be named is null");
            }

            try
            {

                int fluidCode;
                string newName = string.Empty;
                string oldName = string.Empty;


                INamedItem parentNamedItem;
                INamedItem childNamedItem;

                //check the name of the entity
                if (IsNameBlank(entity, ref newName))
                {
                  //Since this is Triggered in the Case of Default Name Rule thus no need to notify 
                  //user even the name is not specified ,autogenereated name will be assigned

                }
                //entity casted to pipeline
                Pipeline pipeline = (Pipeline)entity;

                string childName;
                //this means system is placed under any other system
                parentNamedItem = (INamedItem)parents[0];
                childNamedItem = (INamedItem)entity;

                if (parentNamedItem != null)
                {
                    childName = parentNamedItem.Name + pipeline.SequenceNumber;
                    fluidCode = pipeline.FluidCode;
                    GenerateNameAndTestDuplicate(entity, pipeline, activeEntity, childName, fluidCode, oldName, newName);

                }
            }
            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNamingRules.PipelineNameRule.ComputeName");
            }
        }
        /// <summary>
        /// Method to generate name for an entity placed under Heirarchical Root or System Parent
        /// </summary>
        /// <param name="entity">BO for which name has to be compute</param>
        /// <param name="pipeline">pipeline system</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed</param>
        /// <param name="childName">name of the child</param>
        /// <param name="fluidCode">fluid code for pipeline system</param>
        /// <param name="oldName">original name of the system</param>
        /// <param name="newName">new name of the system</param>
        private void GenerateNameAndTestDuplicate(BusinessObject entity, Pipeline pipeline, BusinessObject activeEntity, string childName, int fluidCode, string oldName, string newName)
        {
            string namingParentsString = string.Empty;
            string equivalentfluidCodeString = string.Empty;
            PropertyValue propVal = pipeline.GetPropertyValue("IJPipelineSystem", "FluidCode");

            if (propVal.PropertyInfo.PropertyType == SP3DPropType.PTCodelist)
            {
                //get Short Display Name of the Property
                PropertyValueCodelist oPropertyValuecodelist = (PropertyValueCodelist)propVal;
                CodelistItem item = propVal.PropertyInfo.CodeListInfo.GetCodelistItem(oPropertyValuecodelist.PropValue);
                if (item != null)
                {
                    equivalentfluidCodeString = item.ShortDisplayName;
                }
            }

            namingParentsString = base.GetNamingParentsString(activeEntity);
			// Concatenating childName with fluidCode
			string childPlusFluidCodeName = childName + fluidCode;
			if (!childPlusFluidCodeName.Equals(namingParentsString))
            {
				base.SetNamingParentsString(activeEntity, childPlusFluidCodeName);
                childName = childName + "-" + equivalentfluidCodeString;
                base.SetName(entity, childName);
                TestForDuplicateName(entity, ref oldName, ref newName);
            }
        }

        /// <summary>
        /// Gets the naming parents from naming rule.
        /// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="entity">BusinessObject for which naming parents are required.</param>
        /// <returns>Collection of BusinessObjects.</returns>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("The entity to be named is null");
            }
            Collection<BusinessObject> parents = new Collection<BusinessObject>();
            // Get system parent
            BusinessObject parentSystem = base.GetParent(HierarchyTypes.System, entity);
            if (null != parentSystem)
            {
                // Add business object as parent
                parents.Add(parentSystem);
            }
            return parents;
        }

        /// <summary>
        /// Compare the name of the system whose name is being checked and rename if there is
        /// another system who shares the same parent system that has the same name.
        /// </summary>
        /// <param name="system">Input.  System whose name is being checked.</param>
        /// <param name="origName">Output.  Original name of the system.</param>
        /// <param name="newName">Output.  Modified name of the system.</param>
        /// <returns>boolean that indicates if the name is a duplicate among the
        /// systems that share the same parent.</returns>
        private bool TestForDuplicateName(BusinessObject system, ref string origName, ref string newName)
        {
            bool bIsDuplicate = false;

            try
            {
                BusinessObject oParent = base.GetParent(HierarchyTypes.System, system);
                ReadOnlyCollection<BusinessObject> oChildren = GetSystemChildren(oParent);

                origName = base.GetName(system);
                newName = origName;

                long lHighSubscript = 0;

                // Loop through each named child of the target parent system to see if there
                // is already a child object with the same name.
                int iNdx;
                if (oChildren != null)
                {
                    for (iNdx = 0; iNdx < oChildren.Count; iNdx++)
                    {
                        BusinessObject oChild = oChildren[iNdx];
                        string strChildName = base.GetName(oChild);

                        long lCurrSubscript = 0;

                        // Be sure not to compare the object with itself.
                        if (!oChild.Equals(system))
                        {
                            if (CheckNamesMatch(strChildName, ref origName, ref lCurrSubscript))
                            {
                                bIsDuplicate = true;
                                if (lCurrSubscript > 0 && lCurrSubscript > lHighSubscript)
                                {
                                    lHighSubscript = lCurrSubscript;
                                }
                                else if (lCurrSubscript == -1 && lHighSubscript == 0)
                                {
                                    lHighSubscript = -1;
                                }
                            }
                        }
                    }
                }

                if (lHighSubscript == -1)
                {
                    // The name of this system matches the name of at least one other system,
                    // but its subscript is unique.  Leave it as it is.
                    bIsDuplicate = false;
                }

                if (bIsDuplicate)
                {
                    // The same name already exists
                    // Test to see if the new name already has a subscript
                    long lNextSubscript = lHighSubscript + 1;
                    string strNextSubscript = lNextSubscript.ToString();

                    string strNameWithoutSubscript = string.Empty;
                    long lTmpNbr;
                    lTmpNbr = NameSubscript(origName, ref strNameWithoutSubscript);
                    if (origName.Equals(strNameWithoutSubscript))
                    {
                        newName = origName + "(" + lNextSubscript + ")";
                    }
                    else
                    {
                        newName = strNameWithoutSubscript + "(" + lNextSubscript + ")";
                    }
                    system.SetPropertyValue(newName, "IJNamedItem", "Name");

                }

            }
            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNamingRules.UserDefinedNameRule.TestForDuplicateName");
            }

            return bIsDuplicate;
        }

        /// <summary>
        /// Examine the name to see if it is blank.  If it is blank, generate a new name
        /// based on its system type and the number of systems present as children of its
        /// parent system.
        /// </summary>
        /// <param name="system">Input.  System whose name is being checked.</param>
        /// <param name="generatedName">Output.  Generated new name.</param>
        /// <returns>boolean that indicates if the name was blank.</returns>
        private bool IsNameBlank(BusinessObject system, ref string generatedName)
        {
            bool bIsBlank = false;

            try
            {
                generatedName = string.Empty;

                string strCurName;
                strCurName = base.GetName(system);
                if (strCurName == null)
                {
                    strCurName = string.Empty;
                }

                if (strCurName == string.Empty)
                {
                    bIsBlank = true;

                    //Get the collection of child systems under the parent system
                    BusinessObject oParentSystem = base.GetParent(HierarchyTypes.System, system);
                    ReadOnlyCollection<BusinessObject> oChildren = GetSystemChildren(oParentSystem);

                    //Create a starting new name based on the system type and number of children
                    // systems already present under the parent system.
                    if (oChildren != null)
                    {
                        long lCount = oChildren.Count;
                        string sType = GetSystemType(system);
                        generatedName = sType + lCount;

                        bool bComparedAllChildren = false;
                        while (!bComparedAllChildren)
                        {
                            // Loop through all the sibling systems to see that the proposed
                            // name is unique.
                            bool bFoundMatch = false;
                            int iNdx;
                            for (iNdx = 0; iNdx < oChildren.Count; iNdx++)
                            {
                                BusinessObject oChild = oChildren[iNdx];
                                string strSysName = base.GetName(oChild);
                                if (generatedName == strSysName)
                                {
                                    bFoundMatch = true;
                                    break;
                                }
                            }
                            if (bFoundMatch)
                            {
                                // Found another child system with the same name.  Increment the
                                // count of the proposed name and compare to all siblings again.
                                lCount = lCount + 1;
                                generatedName = newName + sType + systemName + lCount;
                                bFoundMatch = false;
                            }
                            else
                            {
                                // The proposed name is unique among the sibling systems.
                                bComparedAllChildren = true;
                            }
                        }
                        // Set "Name" property value on IJNamedItem interface
                        system.SetPropertyValue(generatedName, "IJNamedItem", "Name");
                    }
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNamingRules.UserDefinedNameRule.IsnameBlank");
            }

            return bIsBlank;
        }

        /// <summary>
        /// Get a collection of all the system children of a parent system.
        /// </summary>
        /// <param name="parent">Input.  Parent system.</param>
        /// <returns>Collection of the system children.</returns>
        private ReadOnlyCollection<BusinessObject> GetSystemChildren(BusinessObject parent)
        {
            ReadOnlyCollection<BusinessObject> oChildrenCollection = null;

            try
            {
                // Get relation collection for "SystemHierarchy" relationship with "SystemChildren" rolename
                RelationCollection oRelationCollection = parent.GetRelationship("SystemHierarchy", "SystemChildren");

                if (oRelationCollection != null)
                {
                    oChildrenCollection = oRelationCollection.TargetObjects;
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNamingRules.UserDefinedNameRule.GetSystemChildren");
            }

            return oChildrenCollection;
        }

        /// <summary>
        /// Given two strings, check to see if they are the same.  If they are the same,
        /// create a new name based on the original name with a subscript appended.
        /// </summary>
        /// <param name="oldName">Input.  Original name of system.</param>
        /// <param name="newName">Output.  New name of the system.</param>
        /// <param name="outSub">Output.  Subscript used in creating the unique name, if needed.</param>
        /// <returns>boolean flag indicating if the name of the system was unique.</returns>
        /// <remarks>
        /// If dealing with system names with subscripts, two names are determined to be
        /// equal if the base parts of their names are the same.  So A, A(1), and A(10)
        /// would all be identified as having the same name.
        /// 
        /// The subscript is determined by examining all other sibling systems and identifying
        /// all with the same root name, findign the highest subscript, and adding one.
        /// </remarks>
        private bool CheckNamesMatch(string oldName, ref string newName, ref long outSub)
        {
            bool bNamesMatch = false;

            try
            {
                string strOldNameWOSub = string.Empty;
                long lOldSub = NameSubscript(oldName, ref strOldNameWOSub);

                string strNewNameWOSub = string.Empty;
                long lNewSub = NameSubscript(newName, ref strNewNameWOSub);

                if (newName.Equals(oldName))
                {
                    bNamesMatch = true;
                    if (oldName.Equals(strOldNameWOSub))
                    {
                        // The names match and they don't have subscripts.
                        outSub = 0;
                    }
                    else
                    {
                        // The names match and they do have subscripts.
                        outSub = lOldSub;
                    }
                }
                else
                {
                    if (newName.Equals(strOldNameWOSub))
                    {
                        // The existing name has a subscript and the new name does
                        // not.  When the subscript is removed from the existing name,
                        // it matches the new name.
                        bNamesMatch = true;
                        outSub = lOldSub;
                    }
                    else
                    {
                        if (strNewNameWOSub.Equals(oldName))
                        {
                            // The new name has a subscript and the existing name
                            // does not.  When the subscript is removed from the
                            // new name, it matches the existing name.
                            bNamesMatch = true;
                            if (lOldSub > lNewSub)
                            {
                                outSub = lOldSub;
                            }
                            else
                            {
                                outSub = lNewSub;
                            }
                        }
                        else
                        {
                            if (strNewNameWOSub.Equals(strOldNameWOSub))
                            {
                                bNamesMatch = true;
                                if (lOldSub == lNewSub)
                                {
                                    // Both the new name and the existing name have
                                    // subscripts.  The names match and the subscripts
                                    // also match.
                                    outSub = lOldSub;
                                }
                                else
                                {
                                    // Both the new name and the existing name have
                                    // subscripts.  The names match but the subscripts
                                    // do not match.
                                    if (lOldSub > lNewSub)
                                    {
                                        outSub = lOldSub;
                                    }
                                    else
                                    {
                                        outSub = lNewSub;
                                    }
                                }
                            }
                            else
                            {
                                //The new name is not the same as the existing name.
                                bNamesMatch = false;
                            }
                        }
                    }
                }

                if (outSub == 0)
                {
                    outSub = 1;
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNamingRules.UserDefinedNameRule.ChecknamesMatch");
            }

            return bNamesMatch;
        }

        /// <summary>
        /// Given a system name, determine if it has a subscript and if so, strip that
        /// subscript from the original name.
        /// </summary>
        /// <param name="originalName">Input.  Original name.</param>
        /// <param name="strippedName">Output.  Original name stripped of its subscript.</param>
        /// <returns>Numeric value of the subscript.</returns>
        private long NameSubscript(string originalName, ref string strippedName)
        {
            long lSubscript = 0;

            try
            {
                strippedName = originalName;
                string strWorkingName = originalName;

                // Look for an open parenthesis that would indicate that we might have
                // a subscript
                char cTest = '(';
                long lNbrOpen = CountOccurrences(strWorkingName, cTest);
                cTest = ')';
                long lNbrClose = CountOccurrences(strWorkingName, cTest);

                if ((lNbrOpen == 0) || (lNbrOpen != lNbrClose))
                {
                    // We have either no open parenthesis or an uneven number of open and close
                    // parenthesis.  In the first case, there is no subscript.  In the second
                    // case, we don't know what to do so we will treat it as if there is no subscript.
                }
                else
                {
                    //string strDelim = "(";
                    //char[] cDelim = strDelim.ToCharArray();
                    char cDelim = '(';

                    string[] strSplit = null;
                    strSplit = strWorkingName.Split(cDelim);

                    long lNbrWords = strSplit.GetUpperBound(0) - strSplit.GetLowerBound(0) + 1;
                    if (lNbrWords > 1)
                    {
                        // We found the open parenthesis and its not the last character in
                        // the string.
                        strWorkingName = strSplit[0];
                        long lNdx;
                        for (lNdx = 1; lNdx < lNbrWords - 2; ++lNdx)
                        {
                            strWorkingName = strWorkingName + "(" + strSplit[lNdx];
                        }

                        // Let's see if we have a close
                        // parenthesis at the end of the second string.
                        string strWorkingSubscript = strSplit[lNbrWords - 1];

                        // Verify that there is a closed parenthesis at the end of the
                        // second string
                        long lLong = strWorkingSubscript.Length;
                        cTest = ')';
                        lNbrClose = CountOccurrences(strWorkingSubscript, cTest);
                        if (lNbrClose == 1)
                        {
                            // There is a closed parenthesis somewhere in the second string.
                            //  Is it at the end of the string?
                            cDelim = ')';
                            strSplit = strWorkingSubscript.Split(cDelim);
                            long lShort = strSplit[0].Length;
                            if (lLong == (lShort + 1))
                            {
                                // We have the index - in string form.  Convert it.
                                bool bIsNumeric = true;
                                int iNdx;
                                for (iNdx = 0; iNdx < strSplit[0].Length; ++iNdx)
                                {
                                    if (!char.IsNumber(strSplit[0], iNdx))
                                    {
                                        bIsNumeric = false;
                                        break;
                                    }
                                }
                                if (bIsNumeric)
                                {
                                    lSubscript = long.Parse(strSplit[0]);
                                    // Strip the subscript from the name
                                    strippedName = strWorkingName;
                                }

                            }
                        }
                    }
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNamingRules.UserDefinedNameRule.NameSubscript");
            }

            return lSubscript;
        }

        /// <summary>
        /// Given a string and a character, count how many times the character is present
        /// in the string.
        /// </summary>
        /// <param name="text">Input.  String to be examined.</param>
        /// <param name="find">Input.  Character being searched for.</param>
        /// <returns>Number of times the character being searched for appears in
        /// the input string.</returns>
        private long CountOccurrences(string text, char find)
        {
            int iCount = 0;

            try
            {
                int iNdx;
                for (iNdx = 0; iNdx < text.Length; iNdx++)
                {
                    if (text[iNdx] == find)
                    {
                        ++iCount;
                    }
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNamingRules.UserDefinedNameRule.CountOccurrences");
            }

            return iCount;
        }
        private void InitStrings()
        {
            duplicateName = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_DUPLICATENAME, "Duplicate system name:  ");
            hasChanged = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_HASCHANGED, " has been changed to ");
            blankName = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_BLANKNAME, "Blank system names are not allowed.  Name has been changed to:  ");
            systemName = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_SYSTEM, "System");
            newName = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_NEW, "New");
            genericSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_GENERIC, "Generic");
            conduitSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_CONDUIT, "Conduit");
            ductingSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_DUCTING, "Ducting");
            equipmentSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_MACHINERY, "Equipment");
            pipelineSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_PIPELINE, "Pipeline");
            pipingSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_PIPING, "Piping");
            structuralSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_STRUCTURAL, "Structural");
            unitSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_UNIT, "Unit");
            areaSystem = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_AREA, "Area");
        }


        private string GetSystemType(BusinessObject systemType)
        {
            string systemCategory = string.Empty;

            if (systemType.SupportsInterface("IJGenericSystem"))
            {
                systemCategory = genericSystem;
            }
            else if (systemType.SupportsInterface("IJConduitSystem"))
            {
                systemCategory = conduitSystem;
            }
            else if (systemType.SupportsInterface("IJDuctingSystem"))
            {
                systemCategory = ductingSystem;
            }
            else if (systemType.SupportsInterface("IJMachinerySystem"))
            {
                systemCategory = equipmentSystem;
            }
            else if (systemType.SupportsInterface("IJPipelineSystem"))
            {
                systemCategory = pipelineSystem;
            }
            else if (systemType.SupportsInterface("IJPipingSystem"))
            {
                systemCategory = pipingSystem;
            }
            else if (systemType.SupportsInterface("IJStructuralSystem"))
            {
                systemCategory = structuralSystem;
            }
            else if (systemType.SupportsInterface("IJUnitSystem"))
            {
                systemCategory = unitSystem;
            }
            else if (systemType.SupportsInterface("IJAreaSystem"))
            {
                systemCategory = areaSystem;
            }

            return systemCategory;

        }

    }


}
