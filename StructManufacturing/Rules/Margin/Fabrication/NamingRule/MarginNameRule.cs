//-----------------------------------------------------------------------------
//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      Margin Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the Fabrication Margin in Manufacturing.
//
//      History:
//      February 04, 2013    Harsh Sutraway       Creation.
//-----------------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Planning.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// MarginNameRule computes name for the Margin according to its parent part's name.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class MarginNameRule : ManufacturingNameRuleBase
    {
        /// <summary>
        /// Margin type will be handle as part of the DI-CP-307384 .
        /// </summary>
        private const double WELDCOMPENSATIONTYPE = 1000;
        
        /// <summary>
        /// Computes name for the margin. 
        /// </summary>
        /// <param name="entity">Margin whose name is being computed.</param>
        /// <param name="parents">Naming parents of the margin. They can be business objects which control the naming of the margin whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the margin whose name is being computed.</param>
        /// <exception cref="ArgumentNullException">Raised when the entity is null.</exception>
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

                string marginName = string.Empty;
                Margin margin = entity as Margin;
                if (margin != null)
                {
                    if (margin.Mode == MarginMode.Constant)
                    {
                        if (margin.Type == WELDCOMPENSATIONTYPE)
                        {
                            marginName = "Weld Compensation";
                        }
                        else
                        {
                            marginName = "Constant margin";
                        }
                    }

                    if (margin.Mode == MarginMode.Oblique)
                    {
                        marginName = "Oblique margin";
                    }

                    string location = string.Empty;
                    long count = 0;
                    GetCountAndLocationID(marginName, out count, out location);

                    string namedParentString = string.Empty;
                    if (location != string.Empty)
                    {
                        namedParentString = marginName + " " + "-" + location + "-" + count.ToString("########");
                    }
                    else
                    {
                        namedParentString = marginName + " " + "-" + count.ToString("########");
                    }

                    if (GetNamingParentsString(activeEntity) != namedParentString)
                    {
                        SetNamingParentsString(activeEntity, namedParentString);
                    }

                    SetName(entity, namedParentString);
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 6001, "Call to margin name rule failed." + e.Message);
            }             
        }

        /// <summary>
        /// Gets the naming parents of the margin.
        /// All the Naming Parents that need to participate in the margin naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the margin in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="entity">Margin for which naming parents are required.</param>
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
                    BusinessObject parentBlock = (BusinessObject)manfacturingEntity.GetParent(AssemblyObjectType.Block);
                    if (parentBlock != null)
                    {
                        //Add the parent block as the first parent to the naming parents collection.
                        namingParents.Add(parentBlock);
                    }

                    BusinessObject parentPart = (BusinessObject)manfacturingEntity.AssemblyParent;
                    if (parentPart != null)
                    {
                        //Add the parent plate part or the profile part to the naming parents collection.
                        namingParents.Add(parentPart);
                    }
                }            
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 6001, "Call to margin name rule failed." + e.Message);
            }
            return namingParents;
        }
    }
}

