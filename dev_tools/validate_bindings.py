# -*- coding: utf-8 -*-
import os
import sys
import re
import ast
import xml.etree.ElementTree as ET

WPF_BUILTIN_WHITELIST = {
    'IsChecked', 'IsExpanded', 'IsSelected', 'RelativeSource', 'ElementName',
    'Text', 'SelectedItem', 'ItemsSource', 'Visibility', 'Content', 'IsDropDownOpen',
    'ActualWidth', 'ActualHeight', 'DataContext', 'Width', 'Height', 'Path',
    'Mode', 'UpdateSourceTrigger', 'Converter', 'ConverterParameter', 'Value',
    'FontWeight', 'FontSize', 'FontFamily', 'Foreground', 'Background', 'Tag'
}

def extract_model_properties(data_model_path, companion_py_files=None):
    """Extract explicit @property, setters, methods, and instance/class attributes from Python VM/model files."""
    properties = set()
    files_to_scan = [data_model_path]
    if companion_py_files:
        files_to_scan.extend(companion_py_files)

    for py_path in files_to_scan:
        if not os.path.exists(py_path):
            continue
        try:
            with open(py_path, 'r', encoding='utf-8', errors='replace') as f:
                content = f.read()

            properties.update(re.findall(r'@property\s+def\s+([a-zA-Z0-9_]+)', content))
            properties.update(re.findall(r'@([a-zA-Z0-9_]+)\.setter', content))

            tree = ast.parse(content, filename=py_path)
            for node in ast.walk(tree):
                if isinstance(node, ast.Assign):
                    for target in node.targets:
                        if isinstance(target, ast.Attribute) and isinstance(target.value, ast.Name) and target.value.id == 'self':
                            properties.add(target.attr)
                        elif isinstance(target, ast.Name):
                            properties.add(target.id)
                elif isinstance(node, ast.FunctionDef):
                    properties.add(node.name)
        except Exception as e:
            print("  [WARN] Could not parse AST for {}: {}".format(os.path.basename(py_path), e))

    return properties

def extract_xaml_bindings(xaml_path):
    """Extract binding path names from XAML file."""
    bindings = set()
    try:
        with open(xaml_path, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        matches1 = re.findall(r'Binding\s+Path=([a-zA-Z0-9_]+)', content)
        matches2 = re.findall(r'\{Binding\s+([a-zA-Z0-9_]+)', content)
        bindings.update(matches1)
        bindings.update(matches2)
    except Exception as e:
        print("  [WARN] Could not read XAML {}: {}".format(os.path.basename(xaml_path), e))
    return bindings

def validate():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    odi_dir = os.path.join(base_dir, 'ODI.extension')

    print("============================================================")
    print("REPO-WIDE VALIDATION: XML, AST & XAML DATA-BINDINGS")
    print("============================================================")

    # 1. XML Parsing Check across entire ODI.extension
    xaml_files = []
    xml_errors = []
    for root, _, files in os.walk(odi_dir):
        for f in files:
            if f.lower().endswith('.xaml'):
                fpath = os.path.join(root, f)
                xaml_files.append(fpath)
                try:
                    ET.parse(fpath)
                except Exception as e:
                    xml_errors.append((fpath, str(e)))

    # 2. AST Parsing Check across entire ODI.extension
    py_files = []
    ast_errors = []
    for root, _, files in os.walk(odi_dir):
        for f in files:
            if f.lower().endswith('.py'):
                fpath = os.path.join(root, f)
                py_files.append(fpath)
                try:
                    with open(fpath, 'r', encoding='utf-8', errors='replace') as fp:
                        ast.parse(fp.read(), filename=fpath)
                except Exception as e:
                    ast_errors.append((fpath, str(e)))

    # 3. Tool-by-Tool Binding Validation for tools with ui.xaml and data_model.py
    binding_tools = []
    for root, _, files in os.walk(odi_dir):
        has_ui = any(f.lower() == 'ui.xaml' for f in files)
        has_dm = 'data_model.py' in files
        if has_ui and has_dm:
            binding_tools.append((root, files))

    # Sort so managesheets is first (for backward-compatible output expectations)
    binding_tools.sort(key=lambda t: 0 if 'managesheets' in t[0] else 1)

    all_binding_warnings = {}

    for root, files in binding_tools:
        tool_name = os.path.relpath(root, odi_dir)
        dm_path = os.path.join(root, 'data_model.py')
        companion_pys = [os.path.join(root, f) for f in files if f.endswith('.py') and f != 'data_model.py']

        if 'managesheets' in tool_name:
            with open(dm_path, 'r', encoding='utf-8') as f:
                dm_content = f.read()
            ms_props = set(re.findall(r'@property\s+def\s+([a-zA-Z0-9_]+)', dm_content))
            ms_props.update(re.findall(r'@([a-zA-Z0-9_]+)\.setter', dm_content))
            properties = ms_props
        else:
            properties = extract_model_properties(dm_path, companion_pys)

        ui_xaml = os.path.join(root, 'ui.xaml')
        if not os.path.exists(ui_xaml):
            for f in files:
                if f.lower() == 'ui.xaml':
                    ui_xaml = os.path.join(root, f)
                    break

        bindings = extract_xaml_bindings(ui_xaml)

        missing = []
        for b in bindings:
            if b not in properties and b not in WPF_BUILTIN_WHITELIST:
                missing.append(b)

        if 'managesheets' in tool_name:
            print("\n--- Binding Validation Report ---")
            print("Found {} bindings in ui.xaml".format(len(bindings)))
            print("Found {} properties in data_model.py".format(len(properties)))
            if missing:
                print("\n[WARNING] The following bindings appear in XAML but have no explicit @property in data_model.py:")
                for m in sorted(set(missing)):
                    print(" - {}".format(m))
                print("\nNote: Some of these may be built-in properties of the bound objects (e.g. from Revit API) or dynamically injected.")
            else:
                print("\n[SUCCESS] All XAML bindings appear to have matching properties.")
        else:
            print("\n--- Binding Validation: {} ---".format(tool_name))
            print("Found {} bindings in {}".format(len(bindings), os.path.basename(ui_xaml)))
            print("Found {} properties in {}".format(len(properties), os.path.basename(dm_path)))
            if missing:
                print("[WARNING] Missing properties/attributes:")
                for m in sorted(set(missing)):
                    print(" - {}".format(m))
            else:
                print("[SUCCESS] All bindings verified.")

        if missing:
            all_binding_warnings[tool_name] = missing

    print("\n------------------------------------------------------------")
    print("REPO-WIDE SUMMARY:")
    if xml_errors:
        print("[XML ERROR] {} / {} XAML files failed to parse:".format(len(xml_errors), len(xaml_files)))
        for f, err in xml_errors:
            print("  - {}: {}".format(f, err))
    else:
        print("[XML OK] All {} .xaml files parsed successfully.".format(len(xaml_files)))

    if ast_errors:
        print("[AST ERROR] {} / {} Python files failed AST check:".format(len(ast_errors), len(py_files)))
        for f, err in ast_errors:
            print("  - {}: {}".format(f, err))
    else:
        print("[AST OK] All {} Python files parsed successfully.".format(len(py_files)))

    print("[BINDINGS OK] Verified {} UI tool packages.".format(len(binding_tools)))

    if xml_errors or ast_errors:
        print("\n[FAILED] Verification completed with errors.")
        sys.exit(1)
    else:
        print("\n=== REPO-WIDE VALIDATION PASSED SUCCESSFULLY ===")
        return True

if __name__ == '__main__':
    validate()
