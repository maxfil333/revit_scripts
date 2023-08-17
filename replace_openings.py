# -*- coding: utf-8 -*-

import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *

clr.AddReference('System')
from System.Collections.Generic import List

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument


import os
import time
import sys


# --- config ---

class PARAMS:
	def __init__(self, error_family_file, workset_name):
		self.error_family_file = error_family_file
		self.name_type_family_of_error = os.path.basename(error_family_file).replace('.rfa', '')
		self.workset_name = workset_name
		
params = PARAMS(error_family_file="your_opening_family.rfa",
				workset_name='some workset name used in project')
	
# FUNCTIONS

def alert(msg):
	TaskDialog.Show('Скрипт закончил свою работу!', msg)

def mm_to_feet (MTF_mm):
	MTF_feet = MTF_mm / 304.8
	return MTF_feet
	
def feet_to_mm (MTF_feet):
	MTF_mm = MTF_feet * 304.8
	return MTF_mm

def get_bb_props(el):

	bb_max = el.get_BoundingBox(doc.ActiveView).Max
	bb_min = el.get_BoundingBox(doc.ActiveView).Min
	bb_max_x, bb_max_y, bb_max_z = bb_max[0], bb_max[1], bb_max[2]
	bb_min_x, bb_min_y, bb_min_z = bb_min[0], bb_min[1], bb_min[2]
	bb_cen_x = (bb_max[0] + bb_min[0]) / 2
	bb_cen_y = (bb_max[1] + bb_min[1]) / 2
	bb_cen_z = (bb_max[2] + bb_min[2]) / 2
	bb_max = (bb_max_x, bb_max_y, bb_max_z)
	bb_min = (bb_min_x, bb_min_y, bb_min_z)
	bb_cen = (bb_cen_x, bb_cen_y, bb_cen_z)
	
	return XYZ(bb_cen_x, bb_cen_y, bb_cen_z)
	
def get_parameter(elem, param_name):

	element = UnwrapElement(elem)
	instance_parameters_list = list(map(lambda x: x.Definition.Name, element.Parameters)) # список параметров экземпляра
	symbol_parameters_list = list(map(lambda x: x.Definition.Name, element.Symbol.Parameters)) # список параметров типа

	if param_name in instance_parameters_list:
		value = element.LookupParameter(param_name).AsDouble()
	elif param_name in symbol_parameters_list:
		value = element.Symbol.LookupParameter(param_name).AsDouble()
	else:
		msg = "В объекте <{0}>, id = {1} отсутствует параметр <{2}>".format(element.Name, element.Id.ToString(), param_name)
		raise ValueError(msg)
	return value


# collecting data

windows_list = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements()
doors_list = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsNotElementType().ToElements()

collect_list = []
collect_list.extend(windows_list)
collect_list.extend(doors_list)

list_inst = [] # список FamilyInstance
list_noti = [] # список системных семейств (для отчетов/дебага)

for el in collect_list:
	if isinstance(el, FamilyInstance):
		list_inst.append(el)
	else:
		list_noti.append(el)

# select elements in model
selection = uidoc.Selection
collection = List[ElementId](elem.Id for elem in list_inst)
selection.SetElementIds(collection)


# Загрузка / активация семейства

error_family_file = params.error_family_file
name_type_family_of_error = params.name_type_family_of_error
data_list = list_inst

TransactionManager.Instance.ForceCloseTransaction()

all_family_types_of_generic_model = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsElementType().ToElements()

flag = "NOT_IN_MODEL"
for family_type_of_generic_model in all_family_types_of_generic_model:
	temp_name = family_type_of_generic_model.LookupParameter("Имя типа").AsString()
	if temp_name == name_type_family_of_error:
	
		with Transaction(doc, "Activate old family") as TR01:
			TR01.Start()
			error_type_family = family_type_of_generic_model
			family_type_of_generic_model.Activate()
			TR01.Commit()
		flag = "IN_MODEL"

if flag == "NOT_IN_MODEL":
	with Transaction(doc, "Upload family") as TR02:
		TR02.Start()
		doc.LoadFamily(error_family_file)
		TR02.Commit()
	
	all_family_types_of_generic_model = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsElementType().ToElements()
	for family_type_of_generic_model in all_family_types_of_generic_model:
		temp_name = family_type_of_generic_model.LookupParameter("Имя типа").AsString()
		if temp_name == name_type_family_of_error:
			with Transaction(doc, "Activate uploaded family") as TR03:
				TR03.Start()
				error_type_family = family_type_of_generic_model
				family_type_of_generic_model.Activate()
				TR03.Commit()

# выбор рабочего набора

ws = None
worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
for workset in worksets:
	if params.workset_name in workset.Name:
		ws = workset
		break
assert ws != None, 'Отсутствует Рабочий Набор: {}'.format(params.workset_name)

# размещение новых семейств

with Transaction(doc, "replace openings") as TR04:
	TR04.Start()
	for element in data_list:
	
		element = UnwrapElement(element)
		wall	= element.Host
		levelId = element.LevelId
		level   = doc.GetElement(levelId)
		
		old_height = get_parameter(element, "Высота")
		old_width = get_parameter(element, "Ширина")
			
		XYZ_of_object = get_bb_props(element)
		doc.Delete(element.Id)
		new_object = doc.Create.NewFamilyInstance(XYZ_of_object, error_type_family, wall, level, Structure.StructuralType.NonStructural)
	
		new_object.LookupParameter("ADSK_Размер_Высота").Set(old_height)
		new_object.LookupParameter("ADSK_Размер_Ширина").Set(old_width)
		new_object.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(ws.Id.IntegerValue)
								
	TR04.Commit()
