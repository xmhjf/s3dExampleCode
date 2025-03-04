//**************************************************************************************************
//  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
//
//  Project: EquipmentNamingRules
//
//  Abstract: The file contains naming rule implementation for Shape
//
//  Author: Samba
//
//   History:
//       21 Jul, 2005    Samba           Initial Creation
//       17 Mar, 2015    Aditya          CR-CP-269668 : Re-writing the existing EquipNamingRules name rule from vb6 to 3dapi
//**************************************************************************************************

using System;
using System.Collections.Generic;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using Ingr.SP3D.Equipment.Middle;

namespace Ingr.SP3D.Content.Equipment.Rules
{

	public class DatumShapeNameRule : NameRuleBase
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

			if(parents == null)
			{
				throw new ArgumentNullException("The naming parents of the business object to be named are null");
			}

			if (activeEntity == null)
			{
				throw new ArgumentNullException("The name rule active entity associated to the business object to be named is null");
			}

			try
			{
				long count;
				string name;

				string partName = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_STRING_DATUMSHAPE, "DP");
				INamedItem namedItem = entity as INamedItem;

				if (null != namedItem)
				{
					string existingName = namedItem.Name;
					bool isNameGenRequired = IsExistingNameBad(existingName, partName);

					if (isNameGenRequired)
					{
						string nameBasis = base.GetNamingParentsString(activeEntity);
						if (!(partName.Equals(nameBasis)))
						{
							base.SetNamingParentsString(activeEntity, partName);
							count = GetRelatedObjects(entity);
							name = partName + count.ToString();
							base.SetName((BusinessObject)namedItem, name);
						}
					}
					namedItem = null;
				}
			}
			catch(Exception)
			{
				throw new Exception("Unexpected error:  EquipmentNamingRules.DatumShapeNameRule.ComputeName");
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
			if (entity.SupportsInterface("IJShape") == true)
			{
				// Get system parent
				SystemChildHelper child = new SystemChildHelper(entity);
				BusinessObject parentSystem = (BusinessObject)child.SystemParent;
				if (null != parentSystem)
				{
					parents.Add(parentSystem);
				}

			}
			return parents;
		}


		/// <summary>
		/// Checks if the already any object with that name exists. 
		/// </summary>
		/// <param name="currName">Objects new name is appended with this its existing name</param>
		/// <param name="preFix">Objects new name is appended with part name</param>
		/// <returns>bool value</returns>
		private bool IsExistingNameBad(string currName, string preFix)
		{
			bool val = true;
			if (string.IsNullOrEmpty(currName))
			{
				return val;
			}

			//Check with the base name Name of DP in this rule is DP + Integer number of DP shape.
			bool stringContains = currName.Contains(preFix);
			if (stringContains)
			{
				string found = currName.Substring(preFix.Length + 1);
				int result;
				if(int.TryParse(found, out result))
					return val;
				else
					val = false;
			}
			return val;
		}


		/// <summary>
		/// Gets the naming parents from naming rule.
		/// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
		/// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
		/// Both these methods are called from naming rule semantic. 
		/// </summary>
		/// <param name="oEntity">BusinessObject for which naming parents are required.</param>
		/// <returns>ReadOnlyCollection of BusinessObjects.</returns>
		private long GetRelatedObjects(BusinessObject businessObj)
		{

			GenericShape genericShape = (GenericShape)businessObj; // Get the System Parent
			ISystem sysParent = genericShape.SystemParent;
			long count = 0;
			long val = 1;
			String datumShape = LocalizerResourceIDS.GetString(LocalizerResourceIDS.IDS_STRING_DATUMSHAPE, "DP");
			ReadOnlyCollection<ISystemChild> sysChildColl = sysParent.SystemChildren;
			foreach (BusinessObject item in sysChildColl)
			{
				if (item is GenericShape)
				{
					INamedItem namedItem = (INamedItem)item;
					string name = namedItem.Name;
					if (!(string.IsNullOrEmpty(name)))
					{
						int firstFound = name.IndexOf(datumShape);
						if (firstFound == 0)
						{
							long temp = long.Parse(name.Substring(2, name.Length - 2));
							if (count < temp)
							{
								count = temp;
							}
						}
					}
					namedItem = null;
				}
			}
			val = count + 1;
			return val;
		}
	}
}
