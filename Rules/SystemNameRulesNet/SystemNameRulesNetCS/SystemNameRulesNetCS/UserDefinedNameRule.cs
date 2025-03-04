using System;
using System.Collections.Generic;
using System.Text;
using System.Collections.ObjectModel;

using Ingr.SP3D.Common.Middle;

namespace SystemNameRulesNetCS
{
    /// <summary>
    /// Copyright (C) 2007 Intergraph Corporation.  All rights reserved.
    /// 
    /// Class <c>UserDefinedNameRule</c> provides an example of a .Net name rule.
    /// This rule examines a system and takes action to assure that its name is
    /// unique among the system children sharing the same parent system.
    /// 
    /// Author:  T. Merchant
    /// Date created:  Aug 2, 2007
    /// 
    /// History:
    /// </summary>
    /// <remarks>
    /// If the name is blank a new one is generated based on the system type and
    /// the number of child systems already present.
    /// 
    /// If the name is non-blank a search is made for another systems under the same
    /// parent system with the same name.  If one is found the name of the system being
    /// examined by adding a subscript to the name.
    /// 
    /// Examples:
    /// Original name is blank and there are three other systems under the parent system.
    ///     System is of type Area.  New name:  NewAreaSystem4.
    /// Original name:  A.  Another system named A already exists.  New name:  A(1).
    /// Original name:  A.  Systems named A, A(1), A(2) already exist.  New name:  A(3).
    /// Original name:  A(2).  Systems named A, A(1), A(2) already exist.  New name:  A(3).
    /// Original name:  A(2).  Systems named A, A(1), A(3) already exist.  New name:  A(4).
    /// 
    /// To use this name rule, modify CatalogData\Bulkload\DataFiles\GenericNamingRules.xls.
    /// Add a new line to assign this name rule to a system class or modify an existing assignment.
    /// For example, you could replace the standard VB name rule assigned to a system class with
    /// this .Net naming rule.  Do do so, enter:
    /// "SystemNameRulesNetCS,SystemNameRulesNetCS.UserDefinedNameRule|SystemNameRulesNetCSCab.Cab"
    /// in the SolverProgID column.
    /// </remarks>
    public class UserDefinedNameRule : NameRuleBase
    {
        const int E_DUPLICATENAME = -2147220988;
        const int E_BLANKNAME = -2147220987;

        /// <summary>
        /// Check the name of this sytem for uniqueness among the systems that share its parent system.
        /// Create a new name if it has no name or if its name is not unique.
        /// </summary>
        /// <param name="oEntity">Input.  System object whose name is being checked.</param>
        /// <param name="oParents">Input.  Parent system of oEntity.</param>
        /// <param name="oActiveEntity">Input.  Naming active entity.</param>
        public override void ComputeName(BusinessObject oEntity, ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                string strNewName = "";
                if (IsNameBlank(oEntity, ref strNewName))
                {
                    // A new default name was created for the system.  Raise an error to
                    // let the user know if this change.
                    NotifyUser(E_BLANKNAME, "ComputeName",
                        "Blank system names are not allowed.  Name has been changed to:  " + strNewName,
                        "NAMING");
                }
                else
                {
                    string strOrigName = "";
                    if (TestForDuplicateName(oEntity, ref strOrigName, ref strNewName))
                    {
                        // The name of the system being checked is already in use by another
                        // system that shares the same parent.  A subscripted version of the
                        // original name has been created.  raise an error to let the user know
                        // of this change.
                        NotifyUser(E_DUPLICATENAME, "ComputeName",
                            "Duplicate system name:  " + strOrigName + "  has been changed to:  " + strNewName,
                            "NAMING");
                    }
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNameRulesNetCS.UserDefinedNameRule.ComputeName");
            }
        }

        /// <summary>
        /// Get the parent object of the system being named.
        /// </summary>
        /// <param name="oEntity">Input.  System object whose name is being checked.</param>
        /// <returns>null</returns>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oParents = null;

            return oParents;
        }

        /// <summary>
        /// Compare the name of the system whose name is being checked and rename if there is
        /// another system who shares the same parent system that has the same name.
        /// </summary>
        /// <param name="oSystem">Input.  System whose name is being checked.</param>
        /// <param name="strOrigName">Output.  Original name of the system.</param>
        /// <param name="strNewName">Output.  Modified name of the system.</param>
        /// <returns>boolean that indicates if the name is a duplicate among the
        /// systems that share the same parent.</returns>
        private bool TestForDuplicateName(BusinessObject oSystem, ref string strOrigName, ref string strNewName)
        {
            bool bIsDuplicate = false;

            try
            {
                BusinessObject oParent = base.GetParent(HierarchyTypes.System, oSystem);
                ReadOnlyCollection<BusinessObject> oChildren = GetSystemChildren(oParent);

                strOrigName = base.GetName(oSystem);
                strNewName = strOrigName;

                long lHighSubscript = 0;

                // Loop through each named child of the target parent system to see if there
                // is already a child object with the same name.
                int iNdx;

                //Precaution. GetSystemChildren may return null value
                if (oChildren != null)
                {
                    for (iNdx = 0; iNdx < oChildren.Count; iNdx++)
                    {
                        BusinessObject oChild = oChildren[iNdx];
                        string strChildName = base.GetName(oChild);

                        long lCurrSubscript = 0;

                        // Be sure not to compare the object with itself.
                        if (!oChild.Equals(oSystem))
                        {
                            if (CheckNamesMatch(strChildName, ref strOrigName, ref lCurrSubscript))
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

                    string strNameWithoutSubscript = "";
                    long lTmpNbr;
                    lTmpNbr = NameSubscript(strOrigName, ref strNameWithoutSubscript);
                    if (strOrigName.Equals(strNameWithoutSubscript))
                    {
                        strNewName = strOrigName + "(" + lNextSubscript + ")";
                    }
                    else
                    {
                        strNewName = strNameWithoutSubscript + "(" + lNextSubscript + ")";
                    }
                    oSystem.SetPropertyValue(strNewName, "IJNamedItem", "Name");

                }

            }
            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNameRulesNetCS.UserDefinedNameRule.TestForDuplicateName");
            }

            return bIsDuplicate;
        }

        /// <summary>
        /// Examine the name to see if it is blank.  If it is blank, generate a new name
        /// based on its system type and the number of systems present as children of its
        /// parent system.
        /// </summary>
        /// <param name="oSystem">Input.  System whose name is being checked.</param>
        /// <param name="strGeneratedName">Output.  Generated new name.</param>
        /// <returns>boolean that indicates if the name was blank.</returns>
        private bool IsNameBlank(BusinessObject oSystem, ref string strGeneratedName)
        {
            bool bIsBlank = false;

            try
            {
                strGeneratedName = "";

                string strCurName;
                strCurName = base.GetName(oSystem);
                if (strCurName == null)
                {
                    strCurName = "";
                }

                if (strCurName == "")
                {
                    bIsBlank = true;

                    //Get the collection of child systems under the parent system
                    BusinessObject oParentSystem = base.GetParent(HierarchyTypes.System, oSystem);
                    ReadOnlyCollection<BusinessObject> oChildren = GetSystemChildren(oParentSystem);

                    //Create a starting new name based on the system type and number of children
                    // systems already present under the parent system.

                    //Precaution. GetSystemChildren may return null value
                    if (oChildren != null)
                    {
                        long lCount = oChildren.Count;
                        string sType = base.GetTypeString(oSystem);
                        strGeneratedName = "New" + sType + "System" + lCount;

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
                                if (strGeneratedName == strSysName)
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
                                strGeneratedName = "New" + sType + "System" + lCount;
                                bFoundMatch = false;
                            }
                            else
                            {
                                // The proposed name is unique among the sibling systems.
                                bComparedAllChildren = true;
                            }
                        }
                        // Set "Name" property value on IJNamedItem interface
                        oSystem.SetPropertyValue(strGeneratedName, "IJNamedItem", "Name");
                    }
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNameRulesNetCS.UserDefinedNameRule.IsnameBlank");
            }

            return bIsBlank;
        }

        /// <summary>
        /// Add an error to the errors collection that indicates a unacceptable name has been
        /// corrected.
        /// </summary>
        /// <param name="iErrNum">Input.  Error number that indicates either a blank name or 
        /// a duplicate name has been corrected.</param>
        /// <param name="strSource">Input.  Name of calling subroutine.</param>
        /// <param name="strDescription">Input.  Description of the problem.</param>
        /// <param name="strContext">Input.  Error context.</param>
        /// <remarks>
        /// Input parameter strContext should be set to "NAMING" for callers in SystemsAndSpecs
        /// to be able to handle this correctly - by notifying the user but not raising
        /// the error any higher.
        /// </remarks>
        private void NotifyUser(int iErrNum, string strSource, string strDescription, string strContext)
        {
            // On initial creation we're waiting for a logging service to be provided by the CommonApp
            // team.  See DI CP122309 - Create a logger service. 
        }

        /// <summary>
        /// Get a collection of all the system children of a parent system.
        /// </summary>
        /// <param name="oParent">Input.  Parent system.</param>
        /// <returns>Collection of the system children.</returns>
        private ReadOnlyCollection<BusinessObject> GetSystemChildren(BusinessObject oParent)
        {
            ReadOnlyCollection<BusinessObject> oChildrenCollection = null;

            try
            {
                // Get relation collection for "SystemHierarchy" relationship with "SystemChildren" rolename
                RelationCollection oRelationCollection = oParent.GetRelationship("SystemHierarchy", "SystemChildren");

                if (oRelationCollection != null)
                {
                    oChildrenCollection = oRelationCollection.TargetObjects;
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNameRulesNetCS.UserDefinedNameRule.GetSystemChildren");
            }

            return oChildrenCollection;
        }

        /// <summary>
        /// Given two strings, check to see if they are the same.  If they are the same,
        /// create a new name based on the original name with a subscript appended.
        /// </summary>
        /// <param name="strOldName">Input.  Original name of system.</param>
        /// <param name="strNewName">Output.  New name of the system.</param>
        /// <param name="lOutSub">Output.  Subscript used in creating the unique name, if needed.</param>
        /// <returns>boolean flag indicating if the name of the system was unique.</returns>
        /// <remarks>
        /// If dealing with system names with subscripts, two names are determined to be
        /// equal if the base parts of their names are the same.  So A, A(1), and A(10)
        /// would all be identified as having the same name.
        /// 
        /// The subscript is determined by examining all other sibling systems and identifying
        /// all with the same root name, findign the highest subscript, and adding one.
        /// </remarks>
        private bool CheckNamesMatch(string strOldName, ref string strNewName, ref long lOutSub)
        {
            bool bNamesMatch = false;

            try
            {
                string strOldNameWOSub = "";
                long lOldSub = NameSubscript(strOldName, ref strOldNameWOSub);

                string strNewNameWOSub = "";
                long lNewSub = NameSubscript(strNewName, ref strNewNameWOSub);

                if (strNewName.Equals(strOldName))
                {
                    bNamesMatch = true;
                    if (strOldName.Equals(strOldNameWOSub))
                    {
                        // The names match and they don't have subscripts.
                        lOutSub = 0;
                    }
                    else
                    {
                        // The names match and they do have subscripts.
                        lOutSub = lOldSub;
                    }
                }
                else
                {
                    if (strNewName.Equals(strOldNameWOSub))
                    {
                        // The existing name has a subscript and the new name does
                        // not.  When the subscript is removed from the existing name,
                        // it matches the new name.
                        bNamesMatch = true;
                        lOutSub = lOldSub;
                    }
                    else
                    {
                        if (strNewNameWOSub.Equals(strOldName))
                        {
                            // The new name has a subscript and the existing name
                            // does not.  When the subscript is removed from the
                            // new name, it matches the existing name.
                            bNamesMatch = true;
                            if (lOldSub > lNewSub)
                            {
                                lOutSub = lOldSub;
                            }
                            else
                            {
                                lOutSub = lNewSub;
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
                                    lOutSub = lOldSub;
                                }
                                else
                                {
                                    // Both the new name and the existing name have
                                    // subscripts.  The names match but the subscripts
                                    // do not match.
                                    if (lOldSub > lNewSub)
                                    {
                                        lOutSub = lOldSub;
                                    }
                                    else
                                    {
                                        lOutSub = lNewSub;
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

                if (lOutSub == 0)
                {
                    lOutSub = 1;
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNameRulesNetCS.UserDefinedNameRule.ChecknamesMatch");
            }

            return bNamesMatch;
        }

        /// <summary>
        /// Given a system name, determine if it has a subscript and if so, strip that
        /// subscript from the original name.
        /// </summary>
        /// <param name="strOriginalName">Input.  Original name.</param>
        /// <param name="strStrippedName">Output.  Original name stripped of its subscript.</param>
        /// <returns>Numeric value of the subscript.</returns>
        private long NameSubscript(string strOriginalName, ref string strStrippedName)
        {
            long lSubscript = 0;

            try
            {
                strStrippedName = strOriginalName;
                string strWorkingName = strOriginalName;

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
                                    strStrippedName = strWorkingName;
                                }

                            }
                        }
                    }
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNameRulesNetCS.UserDefinedNameRule.NameSubscript");
            }

            return lSubscript;
        }

        /// <summary>
        /// Given a string and a character, count how many times the character is present
        /// in the string.
        /// </summary>
        /// <param name="strText">Input.  String to be examined.</param>
        /// <param name="cFind">Input.  Character being searched for.</param>
        /// <returns>Number of times the character being searched for appears in
        /// the input string.</returns>
        private long CountOccurrences(string strText, char cFind)
        {
            int iCount = 0;

            try
            {
                int iNdx;
                for (iNdx = 0; iNdx < strText.Length; iNdx++)
                {
                    if (strText[iNdx] == cFind)
                    {
                        ++iCount;
                    }
                }
            }

            catch (Exception)
            {
                throw new Exception("Unexpected error:  SystemNameRulesNetCS.UserDefinedNameRule.CountOccurrences");
            }

            return iCount;
        }
    }
}

