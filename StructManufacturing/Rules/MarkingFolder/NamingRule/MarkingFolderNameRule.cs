//-----------------------------------------------------------------------------
//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      Marking Folder Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the Marking Folder in Manufacturing.
//
//      History:
//      February 04, 2013    Harsh Sutraway       Creation.
//-----------------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// MarkingFolderNameRule computes name for Marking Folder according to its parent part name and location.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary>  
    public class MarkingFolderNameRule : ManufacturingNameRuleBase
    {
        private const string DEFAULTBASE = "MKFR";

        /// <summary>
        /// Computes name for the Marking Folder. 
        /// </summary>
        /// <param name="entity">Marking Folder whose name is being computed.</param>
        /// <param name="parents">Naming parents of the Marking Folder. They can be business objects which control the naming of the Marking Folder whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the Marking Folder whose name is being computed.</param>
        /// <exception cref="ArgumentNullException">Raised when the entity or parents or active antity is null.</exception>   
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

                string parentName = string.Empty, namingParentsString = string.Empty;
                string markingFolderName = string.Empty, markingFolderCurrentName = string.Empty, markingFolderTruncatedName = string.Empty;
                
                long counter = 0;
                string locationID = string.Empty;
                //Get the running count and location ID from the NameGeneratorService. 
                GetCountAndLocationID("GSCADStrMfgNamingRule_MarkingDefaultGenerator", out counter, out locationID);

                foreach (BusinessObject parent in parents)
                {
                    if (parent is INamedItem)
                    {
                        INamedItem namedParent = (INamedItem)parent;                    
                        parentName = namedParent.Name;

                        if (locationID != string.Empty)
                        {
                            markingFolderName = DEFAULTBASE + "_" + parentName + "-" + locationID;
                        }
                        else
                        {
                            markingFolderName = DEFAULTBASE + "_" + parentName;
                        }

                        namingParentsString = parentName;
                    }
                }

                //Assign the default name if there was no name taken from the Naming Parent object.
                if(markingFolderName.Length == 0)
                {
                    if (locationID != string.Empty)
                    {
                        markingFolderName = DEFAULTBASE + "_NoParent" + "-" + locationID;
                    }
                    else
                    {
                        markingFolderName = DEFAULTBASE + "_NoParent";
                    }
                    namingParentsString = "MarkingFolderDefaultName";
                }

                if (entity is INamedItem)
                {
                    INamedItem namedMarkingFolder = (INamedItem)entity;
                    markingFolderCurrentName = namedMarkingFolder.Name;

                    if (markingFolderCurrentName.LastIndexOf("-") > 0)
                    {
                        // Get name of the object without Count.So truncate count part of the string.
                        markingFolderTruncatedName = markingFolderCurrentName.Substring(0, (markingFolderCurrentName.LastIndexOf("-") - 1));
                    }
                }

                if (markingFolderTruncatedName != markingFolderName)
                {
                    //Marking Folder having User defined Name OR without Name comes here and default name is assigned here.
                    markingFolderName = markingFolderName + "-" + counter.ToString();
                    SetName(entity, markingFolderName); 
                }
                else
                {
                    SetName(entity, markingFolderCurrentName);
                }


                // Set the NamingParentsString of the Name Rule Active Entity.
                SetNamingParentsString(activeEntity, namingParentsString);
               
             }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 4002, "Call to Marking Folder Name Rule failed with error " + e.Message);
            }
        }

        /// <summary>
        /// Gets the naming parents of the Marking Folder.
        /// All the naming parents that need to participate in the Marking Folder naming are added to the BusinessObject collection. 
        /// The parents added are used in computing the name of the Marking Folder in ComputeName(). 
        /// </summary>
        /// <param name="entity">Marking Folder for which the naming parents are required.</param>
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
                    // Get the parent part of the Marking Folder
                    BusinessObject parentPart = (BusinessObject)manfacturingEntity.DetailedPart;
                    if (parentPart != null)
                    {
                        //Add the parent part to the naming parents
                        namingParents.Add(parentPart);
                    }
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 4002, "Call to Marking Folder Name Rule failed with error " + e.Message);
            }
            return namingParents;
        }
    }
}