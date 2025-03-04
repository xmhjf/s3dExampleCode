//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      Shrinkage Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the Shrinkage in Manufacturing.
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
using Ingr.SP3D.Planning.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// ShrinkageNameRule computes name for the shrinkage according to its parent part name and location.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary>  
    public class ShrinkageNameRule : ManufacturingNameRuleBase
    {
        /// <summary>
        /// Computes name for the Shrinkage. 
        /// </summary>
        /// <param name="entity">Shrinkage whose name is being computed.</param>
        /// <param name="parents">Naming parents of the Shrinkage. They can be business objects which control the naming of the Shrinkage whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the Shrinkage whose name is being computed.</param>
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
                    throw new ArgumentNullException("entity");
                }

                Shrinkage shrinkage = (Shrinkage)entity;
                string parentName = string.Empty;
                string shrinkageName = string.Empty;

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
                }
               
                if (shrinkage.Mode == ShrinkageMode.Assembly)
                {
                    AssemblyBase assy = (AssemblyBase)parents[0];                    
                    int assyType = assy.AssemblyType;
                    string codeListTable = "AssemblyType";
                    string nameSpace = "PLANNG";
                    shrinkageName = CatalogService.GetCodeListStringValue(codeListTable, assyType, nameSpace, true);                   
                   
                }
                else if (shrinkage.Mode == ShrinkageMode.AssemblyToPart)
                {
                    shrinkageName = "AssemblyBased -";
                }
                else if (shrinkage.Mode == ShrinkageMode.Part)
                {
                    shrinkageName = "Part -";
                }

                if (shrinkage.ScalingType == ShrinkageType.ByEdge)
                {
                    shrinkageName = parentName + "-" + shrinkageName + " " + "Edge";
                }
                else if (shrinkage.ScalingType == ShrinkageType.ByAxis)
                {
                    shrinkageName = parentName + "-" + shrinkageName + " " + "Axis";
                }
                else
                {
                    shrinkageName = parentName + "-" + shrinkageName + " " + "Vector";
                }                     
               
                long count = 0;
                string location = string.Empty;
                GetCountAndLocationID(shrinkageName, out count, out location);

                string namedParentString = string.Empty;
                if (location != string.Empty)
                {
                    namedParentString = shrinkageName + "-" + location + "-" + count.ToString("########");
                }
                else
                {
                    namedParentString = shrinkageName + "-" + count.ToString("########");
                }

                if (!GetNamingParentsString(activeEntity).Equals(namedParentString))
                {
                    SetNamingParentsString(activeEntity, namedParentString);
                }

                SetName(entity, namedParentString);
            }
            catch (Exception e)
            {                
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 7001, "Call to Shrinkage Name Rule failed" + e.Message);
            }
        }

        /// <summary>
        /// Gets the naming parents of the Shrinkage.
        /// All the naming parents that need to participate in the Shrinkage naming are added to the BusinessObject collection. 
        /// The parents added are used in computing the name of the Shrinkage in ComputeName(). 
        /// </summary>
        /// <param name="entity">Shrinkage for which the naming parents are required.</param>
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
                string exception = e.ToString();
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 7001, "Call to Shrinkage Name Rule failed" + e.Message);
            }

            return namingParents;
        } 
    }
}
