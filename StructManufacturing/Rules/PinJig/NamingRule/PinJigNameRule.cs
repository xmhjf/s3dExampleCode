//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      PinJig Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the PinJig in Manufacturing.
//
//      History:
//      February 04, 2013    Harsh Sutraway       Creation.
//-----------------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Structure.Middle;


namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// PinjigNameRule computes name for a PinJig based on its parent's name and location.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class PinJigNameRule : ManufacturingNameRuleBase
    {
        /// <summary>
        /// Computes name for the pinjig.
        /// </summary>
        /// <param name="entity">The entity.</param>
        /// <param name="parents">Naming parents of the pinjig. They are the business objects which control the naming of the pinjig whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the pinjig whose name is being computed.</param>
        /// <exception cref="System.ArgumentNullException">
        /// entity
        /// or
        /// activeEntity
        /// </exception>
        /// <exception cref="ArgumentNullException">Raised when the entity or the activeEntity is null.</exception>
        public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }
                if (activeEntity == null)
                {
                    throw new ArgumentNullException("activeEntity");
                }  

                Collection<ManufacturingBase> siblings = null;                
                string pinJigNameToSet = string.Empty;
                string parentName = string.Empty;
                long counter = 0;
                string location = string.Empty;
                int numSiblings = 0;

                if (parents.Count > 0)
                {
                    foreach (BusinessObject parent in parents)
                    {
                        if (parent is INamedItem)
                        {
                            INamedItem namedParent = (INamedItem)parent;
                            parentName = namedParent.Name;                            
                            if (parent is PlatePartBase)
                            {
                                pinJigNameToSet = "PJ" + "_" + parentName + "-";
                            }
                            else
                            {
                                pinJigNameToSet = "AJ" + "_" + parentName + "-";
                            }

                            GetCountAndLocationID(pinJigNameToSet, out counter, out location);
                            if (location != string.Empty)
                            {
                                pinJigNameToSet = pinJigNameToSet + location + "-";
                            }

                            siblings = EntityService.GetManufacturingEntity(parent, ManufacturingEntityType.PinJig);
                            numSiblings = siblings.Count;

                            if (numSiblings == 0)
                                break;

                            INamedItem namedPinJig = (INamedItem)entity;
                            string pinJigOriginalName = namedPinJig.Name;
                            int newNameLength = pinJigNameToSet.Length;
                            if (pinJigOriginalName.Length > newNameLength)
                            {
                                if (pinJigOriginalName.Substring(0, newNameLength) == pinJigNameToSet)
                                {
                                    if (Convert.ToInt64(pinJigOriginalName.Substring(newNameLength + 1)) <= (numSiblings + 1))
                                    {
                                        //Name is already well-formed.
                                        pinJigNameToSet = pinJigOriginalName;
                                        break;
                                    }
                                }
                            }

                            //Iterate through all the siblings to find the first name that is not already taken.
                            string[] siblingNames = null;
                            Array.Resize(ref siblingNames, numSiblings);
                            int i, j;                            
                            for (i = 0; i < numSiblings; i++)
                            {
                                if (siblings[i] != entity)
                                {   
                                    siblingNames[i] = siblings[i].Name;                                    
                                }
                            }

                                                  
                            for (i = 0; i < numSiblings; i++)
                            {
                                string candidateName = pinJigNameToSet + Convert.ToString(i);
                                bool doesCandidateNameExist= false; 
                                for (j = 0; j < numSiblings; j++)
                                {
                                    if (candidateName == siblingNames[j])
                                    {
                                        doesCandidateNameExist = true;
                                        break;
                                    }
                                }
                                if (doesCandidateNameExist == false)
                                {
                                    pinJigNameToSet = candidateName;
                                    break;
                                }
                            }                          

                        }//if (parent is INamedItem)

                    }//end of for
                }
                else
                {
                    GetCountAndLocationID("GSCADStrMfgNamingRule_PinJigDefaultGenerator", out counter, out location);
                    if (location != string.Empty)
                    {
                        pinJigNameToSet = "J" + "_NoParent-" + "-" + location + "-" + counter.ToString("000000");
                    }
                    else
                    {
                        pinJigNameToSet = "J" + "_NoParent-" + counter.ToString("000000");
                    }
                    parentName = "JigDefaultName";
                }

                SetNamingParentsString(activeEntity, parentName);
                SetName(entity, pinJigNameToSet);
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 5001, "Call to PinJig Name Rule failed." + e.Message);                
            }
        }

        /// <summary>
        /// Gets the naming parents of the pinjig.
        /// All the naming parents that need to participate in the pinjig naming are added to the BusinessObject collection. 
        /// The parents added are used in computing the name of the pinjig in ComputeName(). 
        /// </summary>
        /// <param name="entity">Pinjig for which the naming parents are required.</param>
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

                //Add the parent plate or the parent assembly.
                ManufacturingBase manfacturingBase = (ManufacturingBase)entity;
                namingParents.Add((BusinessObject)manfacturingBase.AssemblyParent);                          
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 5001, "Call to PinJig Name Rule failed." + e.Message);                
            }

            return namingParents;      
        } 
    }
}
