# -*- coding: utf-8 -*-
import json
import System
import clr
from Autodesk.Revit.DB.ExtensibleStorage import SchemaBuilder, Schema, Entity, FieldBuilder
from Autodesk.Revit.DB import DataStorage, FilteredElementCollector, Transaction

# A unique GUID for our naming schemes Extensible Storage Schema
SCHEMA_GUID = System.Guid("F034C698-13AE-4054-9FEA-7E5C3EFE2EBE")

def get_or_create_schema():
    schema = Schema.Lookup(SCHEMA_GUID)
    if not schema:
        builder = SchemaBuilder(SCHEMA_GUID)
        builder.SetReadAccessLevel(Autodesk.Revit.DB.ExtensibleStorage.AccessLevel.Public)
        builder.SetWriteAccessLevel(Autodesk.Revit.DB.ExtensibleStorage.AccessLevel.Public)
        builder.SetSchemaName("ManageSheetsSettings")
        
        # Add a string field to hold the JSON dict
        builder.AddSimpleField("NamingSchemesJson", clr.GetClrType(System.String))
        
        schema = builder.Finish()
    return schema

def load_naming_schemes(doc):
    """Loads custom naming schemes from Extensible Storage in the document."""
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
    
    # We open a separate dedicated transaction just for saving settings
    with Transaction(doc, "Save Manage Sheets Settings") as t:
        t.Start()
        if not target_ds:
            target_ds = DataStorage.Create(doc)
            
        entity = Entity(schema)
        entity.Set[System.String]("NamingSchemesJson", json_str)
        target_ds.SetEntity(entity)
        t.Commit()
