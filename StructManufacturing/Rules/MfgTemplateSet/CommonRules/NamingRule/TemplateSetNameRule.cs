using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// TemplateSetNameRule computes name for Template Set according to its parent part name and location.
    /// ManufacturingNameRuleBase class contains common implementation across all the manufacturing name rules. It provides properties and methods for use in implementation of custom name rules. 
    /// </summary>  
    public class TemplateSetNameRule : ManufacturingNameRuleBase
    {
        private const string DEFAULTBASE = "Bxx";

        /// <summary>
        /// Computes name for the Template Set. 
        /// </summary>
        /// <param name="entity">Template Set whose name is being computed.</param>
        /// <param name="parents">Naming parents of the Template Set. They can be business objects which control the naming of the Template Set whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the Template Set whose name is being computed.</param>
        /// <exception cref="ArgumentNullException">Raised when the entity or parents or active entity is null.</exception>   
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

                string parentPartName = string.Empty, parentBlockName = string.Empty;
                string templateSetName = string.Empty, templateSetNewName = string.Empty, namingParentsString = string.Empty;

                Collection<ManufacturingBase> allTemplateSetsOnPart = null;
                int allTemplateSetsCount = 0;

                //Get the running count and location ID from the NameGeneratorService. 
                long counter = 0;
                string locationID = string.Empty;
                GetCountAndLocationID("GSCADStrMfgNamingRule_TemplateDefaultGenerator", out counter, out locationID);

                if (parents.Count > 0)
                {
                    foreach (BusinessObject parent in parents)
                    {
                        if (parent is INamedItem)
                        {
                            INamedItem namedParent = (INamedItem)parent;
                            if (parent is Block)
                            {
                                parentBlockName = namedParent.Name;
                            }
                            else
                            {
                                if (parent is IManufacturable)
                                {
                                    parentPartName = namedParent.Name;

                                    // Get all the Template Sets under this part.
                                    allTemplateSetsOnPart = EntityService.GetManufacturingEntity(parent, ManufacturingEntityType.TemplateSet);

                                    if (allTemplateSetsOnPart != null)
                                    {
                                        allTemplateSetsCount = allTemplateSetsOnPart.Count;
                                    }
                                }
                            }
                        }
                    }

                    //Assign the default name if there was no name taken from the Naming Parent object.
                    if (parentBlockName.Length == 0)
                    {
                        parentBlockName = DEFAULTBASE;
                    }

                    if (parentPartName.Length == 0)
                    {
                        parentPartName = "X99-9XXX-9P";
                    }

                    //PATTERN: TBBB-SSS  (T: An initial letter of Template, BBB: Block name, SSS: Single part name)  
                    // Get the last portion of the parent part name after '>'. Remove the "-" also.
                    string[] parentPartNameSplit = parentPartName.Split('>');

                    if (parentPartNameSplit.Length > 1)
                    {
                        templateSetName = "T" + parentBlockName + "-" + parentPartNameSplit[parentPartNameSplit.GetUpperBound(0)].Substring(1);
                    }
                    else
                    {
                        templateSetName = "T" + parentBlockName + "-" + parentPartNameSplit[parentPartNameSplit.GetUpperBound(0)];
                    }
                    templateSetNewName = templateSetName + "_" + Convert.ToString(allTemplateSetsCount);

                    bool setNewName = false;
                    if (allTemplateSetsOnPart != null)
                    {
                        for (int templateSetsCount = 0; templateSetsCount < allTemplateSetsOnPart.Count; templateSetsCount++)
                        {
                            INamedItem namedItem = allTemplateSetsOnPart[templateSetsCount];

                            int index = 0;
                            index = namedItem.Name.IndexOf(templateSetNewName, StringComparison.InvariantCultureIgnoreCase);

                            if (index > 0)
                            {
                                allTemplateSetsCount = allTemplateSetsCount + 1;
                                templateSetName = templateSetName + "_" + Convert.ToString(allTemplateSetsCount);
                                setNewName = true;
                            }
                        }
                    }

                    if (setNewName == false && allTemplateSetsOnPart != null)
                    {
                        if (allTemplateSetsOnPart.Count > 0)
                        {
                            templateSetName = templateSetName + "_" + Convert.ToString(allTemplateSetsCount);
                        }
                    }

                    if (locationID != string.Empty)
                    {
                        templateSetName = templateSetName + "-" + locationID;
                    }
                    namingParentsString = parentBlockName + "/" + parentPartName;
                }
                else
                {
                    if (locationID != string.Empty)
                    {
                        templateSetName = "T" + DEFAULTBASE + "-" + locationID + "-" + counter.ToString().TrimStart() + "P";
                    }
                    else
                    {
                        templateSetName = "T" + DEFAULTBASE + counter.ToString().TrimStart() + "P";
                    }
                    namingParentsString = "TemplateDefaultName";
                }

                // Set the name of the Template Set
                SetName(entity, templateSetName);

                // Set the NamingParentsString of the Name Rule Active Entity.
                SetNamingParentsString(activeEntity, namingParentsString);


            }
            catch (Exception e)
            {
                System.Diagnostics.StackFrame sf = new System.Diagnostics.StackFrame(true);
                string methodName = sf.GetMethod().Name;
                int lineNumber = sf.GetFileLineNumber();
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 3001, "Call to Template Set Name Rule failed with error " + e.Message );
            }              
        }

        /// <summary>
        /// Gets the naming parents of the Template Set.
        /// All the naming parents that need to participate in the Template Set naming are added to the BusinessObject collection. 
        /// The parents added are used in computing the name of the Template Set in ComputeName(). 
        /// </summary>
        /// <param name="entity">Template Set for which the naming parents are required.</param>
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
                    // Get the parent block of the Template Set
                    BusinessObject parentBlock = (BusinessObject)manfacturingEntity.GetParent(AssemblyObjectType.Block);
                    if (parentBlock != null)
                    {
                        //Add the parent block to the naming parents
                        namingParents.Add(parentBlock);
                    }

                    // Get the parent part of the Template Set
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
                LogForToDoList(ToDoMessageTypes.ToDoMessageWarning, methodName, lineNumber, "SMCustomWarningMessages", 3001, "Call to Template Set Name Rule failed with error " + e.Message);
            }
            return namingParents;
        }
       
    }    
}
