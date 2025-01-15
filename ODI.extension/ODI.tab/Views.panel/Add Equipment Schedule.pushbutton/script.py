#pylint: disable=import-error,invalid-name,broad-except
"""Creates an equipment schedule with specified fields."""

__title__ = "Add Equipment\nSchedule"
__doc__ = "Creates a new equipment schedule with Type, Type Name, Type Comments, Count, and Level fields"
__author__ = "ODI"

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

# Get current document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def schedule_exists(doc, schedule_name):
    """Check if schedule with given name exists."""
    for schedule in FilteredElementCollector(doc).OfClass(ViewSchedule):
        if schedule.Name == schedule_name:
            return True
    return False

def get_unique_schedule_name(doc, base_name):
    """Generate unique schedule name."""
    if not schedule_exists(doc, base_name):
        return base_name
    
    counter = 1
    while schedule_exists(doc, f"{base_name} {counter}"):
        counter += 1
    return f"{base_name} {counter}"

def create_equipment_schedule(doc, schedule_name):
    """Create new equipment schedule with required fields."""
    try:
        with revit.Transaction("Create Equipment Schedule") as t:
            # Create schedule
            schedule = ViewSchedule.CreateSchedule(
                doc, 
                ElementId(BuiltInCategory.OST_MechanicalEquipment)
            )
            schedule.Name = schedule_name
            
            # Get schedule definition
            sched_def = schedule.Definition
            
            # Add required fields
            fields = [
                ("Type", lambda s: ScheduleField.CreateElementType(s)),
                ("Type Name", ElementId(BuiltInParameter.ELEM_TYPE_PARAM)),
                ("Type Comments", ElementId(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)),
                ("Count", lambda s: ScheduleField.CreateCount(s)),
                ("Level", ElementId(BuiltInParameter.SCHEDULE_LEVEL_PARAM))
            ]
            
            for field_name, field_creator in fields:
                if callable(field_creator):
                    field = field_creator(schedule)
                    sched_def.AddField(field)
                else:
                    param = SchedulableField(ScheduleFieldType.Instance, field_creator)
                    sched_def.AddField(param)
            
            return schedule
            
    except Exception as e:
        forms.alert(f"Error creating schedule: {str(e)}")
        return None

def main():
    base_name = "Equipment Schedule"
    
    # Check for existing schedule
    if schedule_exists(doc, base_name):
        if not forms.alert("An Equipment Schedule already exists. Create another one?",
                          yes=True, no=True):
            script.exit()
    
    # Get unique name and create schedule
    schedule_name = get_unique_schedule_name(doc, base_name)
    new_schedule = create_equipment_schedule(doc, schedule_name)
    
    if new_schedule:
        forms.alert(f"Schedule '{schedule_name}' created successfully!")

if __name__ == '__main__':
    main()
