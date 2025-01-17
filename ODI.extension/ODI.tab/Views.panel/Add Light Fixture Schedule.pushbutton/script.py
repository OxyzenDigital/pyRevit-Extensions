"""Creates a Light Fixture schedule."""

__title__ = "Light\nFixture\nSchedule"
__doc__ = "Creates a schedule of all light fixture with Type, Comments, Count and Level"
__author__ = "ODI"

from Autodesk.Revit.DB import *
from pyrevit import revit, forms

doc = __revit__.ActiveUIDocument.Document

def create_lightfixture_schedule(doc, schedule_name="Light Fixture Schedule"):
    """Creates a new light fixture schedule."""
    try:
        with revit.Transaction("Create Light Fixture Schedule"):
            # Create schedule with correct category name
            schedule = ViewSchedule.CreateSchedule(
                doc,
                ElementId(BuiltInCategory.OST_LightingFixtures)  # Fixed: Changed from OST_LightFixture to OST_LightingFixtures
            )
            schedule.Name = schedule_name
            
            # Get schedule definition
            sched_def = schedule.Definition
            
            # Get all available fields and print them for debugging
            available_fields = schedule.Definition.GetSchedulableFields()
            added_param_ids = set()  # Track parameter IDs to prevent duplicates
            
            # Updated fields with correct order
            desired_fields = {
                'Type Mark': 'Type Mark',
                'Family and Type': 'Type',  # Will be hidden
                'Level': 'Level',
                'Type Comments': 'Comments'
            }
            
            # Ensure specific field order
            field_order = [
                'Type Mark',
                'Family and Type',
                'Level',
                'Type Comments'
            ]
            
            # Add fields in specified order
            for field_name in field_order:
                for field in available_fields:
                    param_name = field.GetName(doc)
                    param_id = field.ParameterId
                    
                    if param_id in added_param_ids:
                        continue
                    
                    if param_name == field_name:
                        try:
                            added_field = sched_def.AddField(field)
                            if added_field:
                                added_field.ColumnHeading = desired_fields[param_name]
                                # Hide Family and Type field
                                if param_name == 'Family and Type':
                                    added_field.IsHidden = True
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
    
    base_name = "Light Fixture Schedule"
    schedule_name = get_unique_schedule_name(doc, base_name)
    
    new_schedule = create_lightfixture_schedule(doc, schedule_name)
    
    if new_schedule:
        forms.alert("Schedule '{}' created successfully!".format(schedule_name))
    else:
        forms.alert("Failed to create schedule.")

if __name__ == '__main__':
    main()
