//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      Marking Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the Marking in Manufacturing.
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
    /// MarkingLineNameRule computes name for a Marking Line according to its parent part name property and location. 
    /// ManufacturingNameRuleBase Class contains common implementation across all name rules and provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class MarkingNameRule : ManufacturingNameRuleBase
    {   
        #region private members
        private const string DEFAULTBASE = "MK";
        #endregion private members

        /// <summary>
        /// Computes a name for the marking line.
        /// </summary>
        /// <param name="entity">The entity.</param>
        /// <param name="parents">Naming parents of the marking line. They can be business objects which control the naming of the marking line whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the marking line whose name is being computed.</param>
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
                if(entity == null)
                {
                    throw new ArgumentNullException("entity");
                }
                if(activeEntity == null)
                {
                    throw new ArgumentNullException("activeEntity");
                }

                long count = 0;
                string location = string.Empty;
                GetCountAndLocationID("GSCADStrMfgNamingRule_MarkingDefaultGenerator", out count, out location);

                Marking marking = (Marking)entity;
                string markingName = string.Empty;
                string parentName = string.Empty; 
                if (parents.Count > 0)
                {
                    BusinessObject parent = parents[0];
                    INamedItem namedParent = (INamedItem)parent;
                    parentName = namedParent.Name;
                    if (location != string.Empty)
                    {
                        markingName = DEFAULTBASE + "_" + parentName + "-" + location;
                    }
                    else
                    {
                        markingName = DEFAULTBASE + "_" + parentName;
                    }
                }

                // If markingName is empty then apply the default.
                if (markingName.Length == 0)
                {
                    //Default name will be (ex; "MK_NoParent000001", "MK_NoParent000001", ..., "MK_NoParent999999")                   
                    if (location != string.Empty)
                    {
                        markingName = DEFAULTBASE + "_NoParent" + "-" + location;
                    }
                    else
                    {
                        markingName = DEFAULTBASE + "_NoParent";
                    }
                    parentName = "MarkingDefaultName";
                }

                SetNamingParentsString(activeEntity, parentName);                              

                // Get current name of the object
                string currentMarkingName = marking.Name;
                string truncatedMarkingName = string.Empty;

                // Get name of the object without Count.So truncate count part of the string
                if (currentMarkingName.LastIndexOf("-") > 0)
                {
                    truncatedMarkingName = currentMarkingName.Substring((currentMarkingName.LastIndexOf("-") - 1));
                }
                                
                if (truncatedMarkingName != markingName)
                {
                    //Object having user-defined name OR without name comes here and default name is assigned here.
                    markingName = markingName + "-" + count.ToString();
                    SetName(entity, markingName);
                }
                else
                {
                    SetName(entity, currentMarkingName);
                }
            }
            catch (Exception e)
            {
                string exception = e.ToString();
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 4001, "Call to Tab Name Rule failed");
            }
        }

        /// <summary>
        /// Gets the naming parents of the marking line.
        /// All the naming parents that need to participate in the marking line naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the marking line in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="entity">Marking line for which naming parents are required.</param>
        /// <returns>Returns the collection of naming parents.</returns>
        /// <exception cref="ArgumentNullException">Raised when the entity is null.</exception>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            Collection<BusinessObject> parentsColl = new Collection<BusinessObject>();
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }

                ManufacturingBase manufacturingEntity = entity as ManufacturingBase;
                if (manufacturingEntity != null)
                {
                    BusinessObject parentPart = (BusinessObject)manufacturingEntity.DetailedPart;                    
                    parentsColl.Add(parentPart);
                }                
            }
            catch (Exception e)
            {                
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 4001, "Call to marking line name rule failed." + e.Message);
            }
            return parentsColl;
        }            
    }
}
