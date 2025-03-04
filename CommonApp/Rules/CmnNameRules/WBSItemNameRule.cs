//'*******************************************************************
//'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
//'
//'  Project: NameRules
//'  File:    WBSItemNameRule.cls
//'
//'  Abstract: The file contains a sample implementation of a naming rule
//'       If the object is a WBSItem object, then the name rule is
//'       "WBS Project parent Name" + "Unique counter"
//'
//'  Author: David Kelley
//'
//'  24-Apr-2006  Created
//'  17.MAR.2015  Vijay  CR-CP-269668  Rewriting the WBSNameRules using 3DAPI
//'
//'******************************************************************

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Common.Rules
{
	/// <summary>
	/// WBSItemNameRule computes name for WBSItem
	/// NameRuleBase Class contains common implementation across all name rules and provides properties and methods for use in implementation of custom name rules.
	/// </summary>
	public class WBSItemNameRule : NameRuleBase
	{
		/// <summary>
		/// Computes the name for the given entity.
		/// </summary>
		/// <param name="entity">The business object whose name is being computed.</param>
		/// <param name="parents">The parentsNaming parents of the business object,
		/// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
		/// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
		public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
		{
			if (entity == null)
			{
				throw new ArgumentNullException("The entity to be named is null");
			}

			if (parents == null)
			{
				throw new ArgumentNullException("The naming parents of the business object to be named are null");
			}

			if (activeEntity == null)
			{
				throw new ArgumentNullException("The name rule active entity associated to the business object to be named is null");
			}

			try
			{
				string childName = string.Empty;

				INamedItem parentNamedItem = null;

				INamedItem childNamedItem = entity as INamedItem;
				IWBSChild WBSChild = entity as IWBSChild;
				if (WBSChild != null)
				{
					parentNamedItem = WBSChild.WBSParent as INamedItem;
				}
				if (parentNamedItem != null)
				{
					childName = parentNamedItem.Name;
				}
				else
				{
					childName = childNamedItem.Name;
				}
				string namingParentsString = base.GetNamingParentsString(activeEntity);
				if (childName != null && namingParentsString != null)
				{
					//Name only needs to be recomputed and if the naming parent string is different from the existing one.
					if (!childName.Equals(namingParentsString))
					{
						//Set the naming parents string to obtained child name.
						base.SetNamingParentsString(activeEntity, childName);
						long counter = 0;
						string location = string.Empty;

						//Get the running count for the business object and location ID from the NameGeneratorService.
						base.GetCountAndLocationID(childName, out counter, out location);

						//Set the child name accordingly.
						//Counter has to be padded with zeros for correct formatting.
						if (!string.IsNullOrEmpty(location))
						{
							childName = childName + "-" + "-" + location + "-" + String.Format("{0:0000}", counter);
						}
						else
						{
							childName = childName + "-" + "-" + String.Format("{0:0000}", counter);
						}
						//Set name on entity.
						base.SetName(entity, childName + counter.ToString());
					}
				}
			}
			catch (Exception)
			{
				throw new Exception("Unexpected error:  CmnNameRules.WBSItemNameRule.ComputeName");
			}
		}

		/// <summary>
		/// Gets the naming parents from naming rule.
		/// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection.
		/// </summary>
		/// <param name="entity">BusinessObject for which naming parents are required.</param>
		/// <returns>ReadOnlyCollection of BusinessObjects</returns>
		public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
		{
			if (entity == null)
			{
				throw new ArgumentNullException("The entity to be named is null");
			}

			Collection<BusinessObject> parentsColl = new Collection<BusinessObject>();
			BusinessObject parentBO = base.GetParent(HierarchyTypes.WBS, entity);
			parentsColl.Add(parentBO);
			return parentsColl;
		}
	}
}