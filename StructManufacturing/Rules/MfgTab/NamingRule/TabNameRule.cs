//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      Tab Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the Tab in Manufacturing.
//
//      History:
//      February 04, 2013    Harsh Sutraway       Creation.
//-----------------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// TabNameRule computes name for Tab according to its parent part name and location.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary>  
    public class TabNameRule : ManufacturingNameRuleBase
    {      
        /// <summary>
        /// Computes name for the Tab. 
        /// </summary>
        /// <param name="entity">Tab whose name is being computed.</param>
        /// <param name="parents">Naming parents of the Tab. They can be business objects which control the naming of the Tab whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the Tab whose name is being computed.</param>
        /// <exception cref="ArgumentNullException">Raised when the entity is null.</exception>
        public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }                                 

                Tab tab = entity as Tab;
                if (tab != null)
                {
                    Part tabPart = tab.Part;

                    if (tabPart != null)
                    {
                        long counter = 0;
                        string location = string.Empty;
                        GetCountAndLocationID(tabPart.PartNumber, out counter, out location);

                        string tabName = string.Empty;
                        if (location != string.Empty)
                        {
                            tabName = tabPart.PartNumber + "-" + location + "-" + counter.ToString();
                        }
                        else
                        {
                            tabName = tabPart.PartNumber + "-" + counter.ToString();
                        }

                        SetName(entity, tabName);
                    }
                }
                
            }               
            catch( Exception e )
            {    
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 9001, "Call to Tab Name Rule failed with the error" + e.Message);
            }
              
        }

        /// <summary>
        /// Gets the naming parents of the Tab.
        /// All the naming parents that need to participate in the Tab naming are added to the BusinessObject collection. 
        /// The parents added are used in computing the name of the Tab in ComputeName(). 
        /// </summary>
        /// <param name="entity">Tab for which the naming parents are required.</param>
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
                    // Get the parent part of the Tab
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
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 9001, "Call to Tab Name Rule failed with the error" + e.Message);

            }

            return namingParents;
        } 
    }
}
