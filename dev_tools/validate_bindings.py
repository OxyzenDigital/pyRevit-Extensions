import os
import re

def validate():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    managesheets_dir = os.path.join(base_dir, 'ODI.extension', 'lib', 'managesheets')
    
    ui_xaml = os.path.join(managesheets_dir, 'ui.xaml')
    data_model = os.path.join(managesheets_dir, 'data_model.py')
    
    if not os.path.exists(ui_xaml) or not os.path.exists(data_model):
        print("Could not find ui.xaml or data_model.py")
        return
        
    with open(ui_xaml, 'r', encoding='utf-8') as f:
        xaml_content = f.read()
        
    with open(data_model, 'r', encoding='utf-8') as f:
        python_content = f.read()
        
    # Extract bindings from XAML
    # e.g., Binding Path=CollectionAndSeries
    # or {Binding CollectionAndSeries}
    bindings = set()
    matches1 = re.findall(r'Binding\s+Path=([a-zA-Z0-9_]+)', xaml_content)
    matches2 = re.findall(r'\{Binding\s+([a-zA-Z0-9_]+)', xaml_content)
    
    bindings.update(matches1)
    bindings.update(matches2)
    
    # Extract properties from data_model.py
    properties = set()
    matches_prop = re.findall(r'@property\s+def\s+([a-zA-Z0-9_]+)', python_content)
    matches_setter = re.findall(r'@([a-zA-Z0-9_]+)\.setter', python_content)
    
    properties.update(matches_prop)
    properties.update(matches_setter)
    
    # Check for missing properties
    missing = []
    for b in bindings:
        if b not in properties and b not in ['IsChecked', 'IsExpanded', 'IsSelected', 'RelativeSource', 'ElementName']:
            # Ignore some common built-in WPF properties or ones mapped dynamically if any
            missing.append(b)
            
    print(f"--- Binding Validation Report ---")
    print(f"Found {len(bindings)} bindings in ui.xaml")
    print(f"Found {len(properties)} properties in data_model.py")
    
    if missing:
        print("\n[WARNING] The following bindings appear in XAML but have no explicit @property in data_model.py:")
        for m in sorted(set(missing)):
            print(f" - {m}")
        print("\nNote: Some of these may be built-in properties of the bound objects (e.g. from Revit API) or dynamically injected.")
    else:
        print("\n[SUCCESS] All XAML bindings appear to have matching properties.")

if __name__ == '__main__':
    validate()
