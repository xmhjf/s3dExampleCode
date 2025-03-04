//-----------------------------------------------------------------------------
//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      Margin Assembly Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the Assembly Margin in Manufacturing.
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
    /// MarginAssemblyNameRule computes name for the Margin Assembly.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary> 
    public class MarginAssemblyNameRule : ManufacturingNameRuleBase
    {
        /// <summary>
        /// Computes name for the margin assembly. 
        /// </summary>
        /// <param name="entity">Margin assembly whose name is being computed.</param>
        /// <param name="parents">Naming parents of the margin assembly. They can be business objects which control the naming of the margin assembly whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the margin assembly whose name is being computed.</param>
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

                string marginName = string.Empty;                
                string namingParent = string.Empty;                
                if (entity is MarginAssembly)
                {
                    marginName = "Assembly margin parent";
                }
                else if(entity is Margin)
                {
                    Margin childmargin = (Margin)entity;
                    if (childmargin.CreationType == MarginCreationType.Assembly)
                    {   
                        //Assembly margin child.
                        marginName = "Assembly Margin Child";
                    }
                }

                long count = 0;
                string location = string.Empty;
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

                // Set the NamingParentsString property of the Active Entity.
                if (GetNamingParentsString(activeEntity) != namedParentString)
                {
                    SetNamingParentsString(activeEntity, namedParentString);
                }

                SetName(entity, namedParentString);                
            }
            catch( Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 6002, "Call to Assembly Margin Name Rule failed with error" + e.Message);
            }
        }
         
        /// <summary>
        /// Gets the naming parents of the Margin Assembly.
        /// All the naming parents that need to participate in the Margin Assembly naming are added to the BusinessObject collection. 
        /// The parents added are used in computing the name of the Margin Assembly in ComputeName(). 
        /// </summary>
        /// <param name="entity">Margin Assembly for which the naming parents are required.</param>
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

                    BusinessObject parentPartOrAssembly = (BusinessObject)manfacturingEntity.AssemblyParent;
                    if (parentPartOrAssembly != null)
                    {
                        //Add the parent part or the parent assembly to the naming parents collection.
                        namingParents.Add(parentPartOrAssembly);
                    }
                }                
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 6002, "Call to Assembly Margin Name Rule failed." + e.Message);
            }
            return namingParents;
        }
    }
}
