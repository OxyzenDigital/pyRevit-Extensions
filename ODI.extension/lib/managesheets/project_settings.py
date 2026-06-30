# -*- coding: utf-8 -*-
import json
import System
import clr
from Autodesk.Revit.DB.ExtensibleStorage import SchemaBuilder, Schema, Entity, FieldBuilder, DataStorage, AccessLevel
from Autodesk.Revit.DB import FilteredElementCollector, Transaction

# A unique GUID for our naming schemes Extensible Storage Schema
SCHEMA_GUID = System.Guid("F034C698-13AE-4054-9FEA-7E5C3EFE2EBE")
CLASSIFICATION_SCHEMA_GUID = System.Guid("B1B6D208-44A3-40A1-B88E-3C5EF5A1E523")
PROJECT_SETUP_SCHEMA_GUID = System.Guid("F9E8D7C6-B5A4-3021-9876-1A2B3C4D5E6F")

def get_or_create_schema():
    schema = Schema.Lookup(SCHEMA_GUID)
    if not schema:
        builder = SchemaBuilder(SCHEMA_GUID)
        builder.SetReadAccessLevel(AccessLevel.Public)
        builder.SetWriteAccessLevel(AccessLevel.Public)
        builder.SetSchemaName("ManageSheetsSettings")
        
        # Add a string field to hold the JSON dict
        builder.AddSimpleField("NamingSchemesJson", clr.GetClrType(System.String))
        
        schema = builder.Finish()
    return schema

def load_naming_schemes(doc):
    """Loads custom naming schemes from Extensible Storage in the document."""
    if not doc:
        return None
        
    schema = get_or_create_schema()
    collector = FilteredElementCollector(doc).OfClass(DataStorage)
    
    for ds in collector:
        entity = ds.GetEntity(schema)
        if entity.IsValid():
            json_str = entity.Get[System.String]("NamingSchemesJson")
            if json_str:
                try:
                    return json.loads(json_str)
                except Exception:
                    return None
    return None

def save_naming_schemes(doc, schemes_dict):
    """Saves custom naming schemes to Extensible Storage in an isolated transaction."""
    schema = get_or_create_schema()
    
    collector = FilteredElementCollector(doc).OfClass(DataStorage)
    target_ds = None
    for ds in collector:
        entity = ds.GetEntity(schema)
        if entity.IsValid():
            target_ds = ds
            break
            
    json_str = json.dumps(schemes_dict)
    
    def _save():
        with Transaction(doc, "Save Manage Sheets Settings") as t:
            t.Start()
            if not target_ds:
                new_ds = DataStorage.Create(doc)
                entity = Entity(schema)
                entity.Set[System.String]("NamingSchemesJson", json_str)
                new_ds.SetEntity(entity)
            else:
                entity = Entity(schema)
                entity.Set[System.String]("NamingSchemesJson", json_str)
                target_ds.SetEntity(entity)
            t.Commit()
            
    from pyrevit.revit.events import execute_in_revit_context
    execute_in_revit_context(_save)

def get_or_create_classification_schema():
    schema = Schema.Lookup(CLASSIFICATION_SCHEMA_GUID)
    if not schema:
        builder = SchemaBuilder(CLASSIFICATION_SCHEMA_GUID)
        builder.SetReadAccessLevel(AccessLevel.Public)
        builder.SetWriteAccessLevel(AccessLevel.Public)
        builder.SetSchemaName("ManageSheetsClassification")
        builder.AddSimpleField("ClassificationJson", clr.GetClrType(System.String))
        schema = builder.Finish()
    return schema

def load_classification_dict(doc):
    """Loads custom classification dict from Extensible Storage."""
    if not doc:
        return None
        
    schema = get_or_create_classification_schema()
    collector = FilteredElementCollector(doc).OfClass(DataStorage)
    
    for ds in collector:
        entity = ds.GetEntity(schema)
        if entity.IsValid():
            json_str = entity.Get[System.String]("ClassificationJson")
            if json_str:
                try:
                    return json.loads(json_str)
                except Exception:
                    return None
    return None

def save_classification_dict(doc, class_dict):
    """Saves custom classification dict to Extensible Storage."""
    schema = get_or_create_classification_schema()
    
    collector = FilteredElementCollector(doc).OfClass(DataStorage)
    target_ds = None
    for ds in collector:
        entity = ds.GetEntity(schema)
        if entity.IsValid():
            target_ds = ds
            break
            
    json_str = json.dumps(class_dict)
    
    def _save():
        with Transaction(doc, "Save Manage Sheets Classification") as t:
            t.Start()
            if not target_ds:
                new_ds = DataStorage.Create(doc)
                entity = Entity(schema)
                entity.Set[System.String]("ClassificationJson", json_str)
                new_ds.SetEntity(entity)
            else:
                entity = Entity(schema)
                entity.Set[System.String]("ClassificationJson", json_str)
                target_ds.SetEntity(entity)
            t.Commit()
            
    from pyrevit.revit.events import execute_in_revit_context
    execute_in_revit_context(_save)

def get_or_create_project_setup_schema():
    schema = Schema.Lookup(PROJECT_SETUP_SCHEMA_GUID)
    if not schema:
        builder = SchemaBuilder(PROJECT_SETUP_SCHEMA_GUID)
        builder.SetReadAccessLevel(AccessLevel.Public)
        builder.SetWriteAccessLevel(AccessLevel.Public)
        builder.SetSchemaName("ManageSheetsProjectSetup")
        builder.AddSimpleField("SetupJson", clr.GetClrType(System.String))
        schema = builder.Finish()
    return schema

def load_project_setup(doc):
    """Loads project setup dict from Extensible Storage."""
    if not doc:
        return None
        
    schema = get_or_create_project_setup_schema()
    collector = FilteredElementCollector(doc).OfClass(DataStorage)
    
    for ds in collector:
        entity = ds.GetEntity(schema)
        if entity.IsValid():
            json_str = entity.Get[System.String]("SetupJson")
            if json_str:
                try:
                    return json.loads(json_str)
                except Exception:
                    return None
    return None

def save_project_setup(doc, setup_dict):
    """Saves project setup dict to Extensible Storage."""
    schema = get_or_create_project_setup_schema()
    
    collector = FilteredElementCollector(doc).OfClass(DataStorage)
    target_ds = None
    for ds in collector:
        entity = ds.GetEntity(schema)
        if entity.IsValid():
            target_ds = ds
            break
            
    json_str = json.dumps(setup_dict)
    
    def _save():
        with Transaction(doc, "Save Manage Sheets Project Setup") as t:
            t.Start()
            if not target_ds:
                new_ds = DataStorage.Create(doc)
                entity = Entity(schema)
                entity.Set[System.String]("SetupJson", json_str)
                new_ds.SetEntity(entity)
            else:
                entity = Entity(schema)
                entity.Set[System.String]("SetupJson", json_str)
                target_ds.SetEntity(entity)
            t.Commit()
            
    from pyrevit.revit.events import execute_in_revit_context
    execute_in_revit_context(_save)

def clear_project_setup(doc):
    """Deletes project setup dict from Extensible Storage."""
    schema = get_or_create_project_setup_schema()
    
    collector = FilteredElementCollector(doc).OfClass(DataStorage)
    target_ds = None
    for ds in collector:
        entity = ds.GetEntity(schema)
        if entity.IsValid():
            target_ds = ds
            break
            
    if target_ds:
        def _delete():
            with Transaction(doc, "Clear Manage Sheets Project Setup") as t:
                t.Start()
                doc.Delete(target_ds.Id)
                t.Commit()
        from pyrevit.revit.events import execute_in_revit_context
        execute_in_revit_context(_delete)
        return True
    return False
