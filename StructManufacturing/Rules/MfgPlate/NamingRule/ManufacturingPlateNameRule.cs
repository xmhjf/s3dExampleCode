﻿//      Copyright (C) 2013, Intergraph Corporation. All rights reserved.
//
//      Manufacturing Plate Name Rule Class Implementation.
//
//      Abstract: The file contains an implementation of the default naming rule
//                for the Manufacturing Plate in Manufacturing.
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
    /// ManufacturingPlateNameRule computes name for ManufacturingPlate according to its parent part name and location.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    public class ManufacturingPlateNameRule : ManufacturingNameRuleBase
    {
        /// <summary>
        /// Computes name for the ManufacturingPlate. 
        /// </summary>
        /// <param name="entity">Manufacturing Plate whose name is being computed.</param>
        /// <param name="parents">Naming parents of the ManufacturingPlate. They can be business objects which control the naming of the ManufacturingPlate whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the ManufacturingPlate whose name is being computed.</param>
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

                string manufacturingPlateName = string.Empty, parentName = string.Empty, namingParentsString = string.Empty;

                //Get the running count for the Manufacturing Plate and location ID from the NameGeneratorService.
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

                    string[] parentPartNameSplit = parentName.Split('>');
                    // This case means parentName already contains the block name. Either a true blockname or "NoBlk"
                    // and it assumes that strParentName has always <B0.1.1.1.1>-xxxx-yyyy or <NoBlk>-xxxx-yyyy convention.
                    //Take out the blockinformation regardless of the true blockname or NoBlk as well as the "-".                  
                    if (locationID != string.Empty)
                    {
                        manufacturingPlateName = parentPartNameSplit[parentPartNameSplit.GetUpperBound(0)].Substring(1) + "-" + locationID;
                    }
                    else
                    {
                        manufacturingPlateName = parentPartNameSplit[parentPartNameSplit.GetUpperBound(0)].Substring(1);
                    }

                    namingParentsString = parentName;
                }
                else
                {
                    if (locationID != string.Empty)
                    {
                        manufacturingPlateName = locationID + "-" + counter.ToString().TrimStart() + "P";
                    }
                    else
                    {
                        manufacturingPlateName = counter.ToString().TrimStart() + "P";
                    }
                    namingParentsString = "PlateDefaultName";
                }

                // Set the name of the Manufacturing Plate
                SetName(entity, manufacturingPlateName);

                // Set the NamingParentsString of the Name Rule Active Entity.
                SetNamingParentsString(activeEntity, namingParentsString);                

            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 1034, "Call to ManufacturingPlate Name Rule failed with the error" + e.Message);
            }
        }

        /// <summary>
        /// Gets the naming parents of the ManufacturingPlate.
        /// All the naming parents that need to participate in the ManufacturingPlate naming are added to the BusinessObject collection. 
        /// The parents added are used in computing the name of the ManufacturingPlate in ComputeName(). 
        /// </summary>
        /// <param name="entity">Manufacturing Plate for which the naming parents are required.</param>
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
                    // Get the parent part of the Manufacturing Plate
                    BusinessObject platePart = (BusinessObject)manfacturingEntity.DetailedPart;
                    if (platePart != null)
                    {
                        //Add the parent part to the naming parents
                        namingParents.Add(platePart);
                    }                    
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 1034, "Call to ManufacturingPlate Name Rule failed with the error" + e.Message);
            }
            return namingParents;
        }
    }
}