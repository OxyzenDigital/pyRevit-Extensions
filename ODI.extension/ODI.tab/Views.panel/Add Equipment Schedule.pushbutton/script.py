"""Creates a specialty equipment schedule."""

__title__ = "Specialty\nEquipment\nSchedule"
__doc__ = "Creates a schedule of all specialty equipment with Type, Comments, Count and Level"
__author__ = "ODI"

from Autodesk.Revit.DB import *
from pyrevit import revit, forms

doc = __revit__.ActiveUIDocument.Document

def create_specialty_schedule(doc, schedule_name="Specialty Equipment Schedule"):
    """Creates a new specialty equipment schedule."""
    try:
        with revit.Transaction("Create Specialty Equipment Schedule"):
            # Create schedule
            schedule = ViewSchedule.CreateSchedule(
                doc,
                ElementId(BuiltInCategory.OST_SpecialityEquipment)
            )
            schedule.Name = schedule_name
            
            # Get schedule definition
            sched_def = schedule.Definition
            
            # Get all available fields and print them for debugging
            available_fields = schedule.Definition.GetSchedulableFields()
            added_param_ids = set()  # Track parameter IDs to prevent duplicates
            
            # First pass - add only the specific fields we want
            desired_fields = {
                'Family and Type': 'Type',  # Map the actual parameter name to desired heading
                'Level': 'Level',
                'Type Comments': 'Type Comments'
            }
            
            # Add fields one by one with validation
            for field in available_fields:
                param_name = field.GetName(doc)
                param_id = field.ParameterId
                
                # Skip if we already added this parameter
                if param_id in added_param_ids:
                    continue
                
                # Only add if it's in our desired fields
                if param_name in desired_fields:
                    try:
                        added_field = sched_def.AddField(field)
                        if added_field:
                            added_field.ColumnHeading = desired_fields[param_name]
                            added_param_ids.add(param_id)
                    except Exception as e:
                        forms.alert("Could not add field {}: {}".format(param_name, str(e)))
            
            # Add count field last
            try:
                count_field = sched_def.AddField(ScheduleFieldType.Count)
                count_field.ColumnHeading = "Count"
            except Exception as e:
                forms.alert("Could not add Count field: {}".format(str(e)))
            
            # Make schedule show all instances
            sched_def.IsItemized = True
            
            return schedule
            
    except Exception as e:
        forms.alert("Failed to create schedule: {}".format(str(e)))
        return None

def get_unique_schedule_name(doc, base_name):
    """Generate unique schedule name with increment if needed."""
    schedule_names = set()
    
    # Get all existing schedule names
    for schedule in FilteredElementCollector(doc).OfClass(ViewSchedule):
        schedule_names.add(schedule.Name)
    
    if base_name not in schedule_names:
        return base_name
        
    counter = 1
    while "{}{}".format(base_name, counter) in schedule_names:
        counter += 1
    
    return "{}{}".format(base_name, counter)

def main():
    if not doc:
        forms.alert("No active document found.")
        return
    
    base_name = "Specialty Equipment Schedule"
    schedule_name = get_unique_schedule_name(doc, base_name)
    
    new_schedule = create_specialty_schedule(doc, schedule_name)
    
    if new_schedule:
        forms.alert("Schedule '{}' created successfully!".format(schedule_name))
    else:
        forms.alert("Failed to create schedule.")

if __name__ == '__main__':
    main()
