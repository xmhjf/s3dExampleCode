/*******************************************************************
  Copyright (C) 1999, Intergraph Corporation.  All rights reserved.
'
'  Project: SystemsAndSpecs SystemNameRules
'  Class:   DefaultNameRule
'
'  Abstract: The file contains a sample implementation of a naming rule for systems.
'
             Mishra,Anamay               CR-CP-270276  System&Specs VB6 Rules need to be replaced
'
'******************************************************************/


using System;
using System.Collections.Generic;
using System.Text;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.System.Rules
{
	/// <summary>
	/// Class implementing the default Name Rule for Systems
	/// </summary>
	public class DefaultNameRule : NameRuleBase
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
				string partName = string.Empty;
				// Get part from part occurrence
				BusinessObject part = base.GetPart(entity);

				if (null != part)
				{
					// Get part number from part
					partName = base.GetPartNumber(part);
				}

				if (string.IsNullOrEmpty(partName))
				{
					if (entity.SupportsInterface("IJNamedItem"))
					{
						// Get "TypeString" property value from IJDNamedItem interface
						partName = base.GetTypeString(entity);
					}
				}

				// Remove spaces from name
				partName = partName.Replace(" ", "").Trim();
				// Get naming parents string from active entity
				string namingParentsString = base.GetNamingParentsString(activeEntity);

				string parentName = string.Empty;
				string childName = string.Empty;
				int range = 1;
				if (parents.Count > 0)
				{
					// Check if new naming parent string and existing naming parent string are same
					// If they are same, we do not generate a new sequence number and there is no change to the name of the object.
					if (!partName.Equals(namingParentsString))
					{
						// Set "NamingParentsString" property value on IJNameRuleAE interface
						base.SetNamingParentsString(activeEntity, partName);
						long count;
						string locationID;
						//Get count and location ID using the range specified from NameGeneratorService
						base.GetCountAndLocationID(partName, range, out count, out locationID);
						string name = string.Empty;
						if (!string.IsNullOrEmpty(locationID))
						{
							name = partName + "-" + locationID + "-" + count.ToString("D4");
						}
						else
						{
							name = partName + "-" + count.ToString("D4");
						}
						// Set name on entity
						base.SetName(entity, name);
					}
				}
			}
			catch (Exception)
			{
				throw new Exception("Unexpected error:  SystemNamingRule.DefaultNameRule.ComputeName");
			}

		}

		/// <summary>
		/// Gets the naming parents from naming rule.
		/// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
		/// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
		/// Both these methods are called from naming rule semantic. 
		/// </summary>
		/// <param name="entity">BusinessObject for which naming parents are required.</param>
		/// <returns>Collection of BusinessObjects.</returns>
		public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
		{
			if (entity == null)
			{
				throw new ArgumentNullException("The entity to be named is null");
			}
			Collection<BusinessObject> parents = new Collection<BusinessObject>();
			// Get system parent
			BusinessObject parentSystem = base.GetParent(HierarchyTypes.System, entity);
			if (null != parentSystem)
			{
				// Add business object as parent
				parents.Add(parentSystem);
			}
			return parents;
		}
	}
}
