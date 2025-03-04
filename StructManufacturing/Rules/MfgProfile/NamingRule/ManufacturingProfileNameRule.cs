//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      Manufacturing Profile Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the Manufacturing Profile. 
//
//      History:
//      March 27, 2013    Manasa Jaisetty       Creation.
//-----------------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// ManufacturingProfileNameRule computes name for ManufacturingProfile according to its parent part name and location.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class ManufacturingProfileNameRule : ManufacturingNameRuleBase
    {
        /// <summary>
        /// Computes name for the ManufacturingProfile. 
        /// </summary>
        /// <param name="entity">Manufacturing Profile whose name is being computed.</param>
        /// <param name="parents">Naming parents of the ManufacturingProfile. They can be business objects which control the naming of the ManufacturingProfile whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the ManufacturingProfile whose name is being computed.</param>
        /// <exception cref="ArgumentNullException">Raised when the entity is null.</exception>
        public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }
                if (parents == null)
                {
                    throw new ArgumentNullException("parents");
                }
                if (activeEntity == null)
                {
                    throw new ArgumentNullException("activeEntity");
                }

                string manufacturingProfileName = string.Empty, parentName = string.Empty, namingParentsString = string.Empty;

                //Get the running count for the Manufacturing Profile and location ID from the NameGeneratorService. 
                string locationID = string.Empty;
                long counter = 0;
                GetCountAndLocationID("GSCADStrMfgNamingRule_PlateDefaultGenerator", out counter, out locationID);

                if (parents.Count > 0)
                {
                    foreach (BusinessObject parent in parents)
                    {
                        if (parent is INamedItem)
                        {
                            INamedItem parentNamed = (INamedItem)parent;
                            parentName = parentNamed.Name;
                        }
                    }

                    // This case means parentName already contains the block name. Either a true blockname or "NoBlk"
                    // and it assumes that strParentName has always <B0.1.1.1.1>-xxxx-yyyy or <NoBlk>-xxxx-yyyy convention.
                    //Take out the blockinformation regardless of the true blockname or NoBlk as well as the "-".  
                    
                    // No Block Case
                    if (parentName.IndexOf('<') == -1 || parentName.IndexOf('>') == -1)
                    {
                        if (locationID != string.Empty)
                        {
                            manufacturingProfileName = parentName + "-" + locationID;
                        }
                        else
                        {
                            manufacturingProfileName = parentName;
                        }
                    }
                    else
                    {
                        string[] parentPartNameSplit = parentName.Split('>');

                        if (locationID != string.Empty)
                        {
                            manufacturingProfileName = parentPartNameSplit[parentPartNameSplit.GetUpperBound(0)].Substring(1) + "-" + locationID;
                        }
                        else
                        {
                            manufacturingProfileName = parentPartNameSplit[parentPartNameSplit.GetUpperBound(0)].Substring(1);
                        }
                    }

                    namingParentsString = parentName;
                }
                else
                {
                    if (locationID != string.Empty)
                    {
                        manufacturingProfileName = "nXXn-" + locationID + "-" + counter.ToString().TrimStart() + "A";
                    }
                    else
                    {
                           manufacturingProfileName = "nXXn-" + counter.ToString().TrimStart() + "A";
                    }
                    namingParentsString = "ProfileDefaultName";
                }

                // Set the name of the Manufacturing Profile
                SetName(entity, manufacturingProfileName);

                // Set the NamingParentsString of the Name Rule Active Entity.
                SetNamingParentsString(activeEntity, namingParentsString);         
       
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 2018, "Call to ManufacturingProfile Name Rule failed with the error" + e.Message);
            }
        }

        /// <summary>
        /// Gets the naming parents of the ManufacturingProfile.
        /// All the naming parents that need to participate in the ManufacturingProfile naming are added to the BusinessObject collection. 
        /// The parents added are used in computing the name of the ManufacturingProfile in ComputeName(). 
        /// </summary>
        /// <param name="entity">ManufacturingProfile for which the naming parents are required.</param>
        /// <returns>Returns the collection of naming parents.</returns>
        /// <exception cref="ArgumentNullException">Raised when the entity is null.</exception>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            Collection<BusinessObject> namingParents = new Collection<BusinessObject>();
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }
                ManufacturingBase manfacturingEntity = entity as ManufacturingBase;
                if (manfacturingEntity != null)
                {
                    // Get the parent part of the Manufacturing Profile
                    BusinessObject profilePart = (BusinessObject)manfacturingEntity.DetailedPart;
                    if (profilePart != null)
                    {
                        //Add the parent part to the naming parents
                        namingParents.Add(profilePart);
                    }                    
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 2018, "Call to ManufacturingProfile Name Rule failed with the error" + e.Message);
            }
            return namingParents;
        }
    }
}