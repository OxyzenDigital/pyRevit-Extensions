# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import (
    Transaction, ElementId, XYZ, BuiltInCategory, FilteredElementCollector, BuiltInParameter,
    DirectShape, Line, Point
)
from System.Collections.Generic import List
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
import data_model

class PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, elem):
        if not elem.Category: return False
        
        # Safe ID Retrieval
        cat_id = -1
        eid = elem.Category.Id
        try:
            if hasattr(eid, "Value"): cat_id = eid.Value
            elif hasattr(eid, "IntegerValue"): cat_id = eid.IntegerValue
        except: return False

        # Allowed Categories
        allowed_cats = [
            int(BuiltInCategory.OST_PipeCurves),
            int(BuiltInCategory.OST_DuctCurves),
            int(BuiltInCategory.OST_Conduit),
            int(BuiltInCategory.OST_CableTray),
            int(BuiltInCategory.OST_FabricationPipework),
            int(BuiltInCategory.OST_FabricationDuctwork),
            int(BuiltInCategory.OST_FabricationContainment)
        ]
        
        return cat_id in allowed_cats

    def AllowReference(self, ref, pt): return True

class RevitService(object):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc

    def pick_element(self, prompt):
        ref = self.uidoc.Selection.PickObject(ObjectType.Element, PipeSelectionFilter(), prompt)
        return self.doc.GetElement(ref)

    def get_element_data(self, element_id):
        if not element_id: return None
        el = self.doc.GetElement(element_id)
        if not el: return None
        
        # Extract geometry data for Logic engine
        try:
            curve = el.Location.Curve
            
            # Safe ID
            eid_val = -1
            if hasattr(el.Id, "Value"): eid_val = el.Id.Value
            elif hasattr(el.Id, "IntegerValue"): eid_val = el.Id.IntegerValue

            # Get Diameter/Width
            diam = 0.0
            param = el.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
            if param: diam = param.AsDouble()
            else:
                # Try Duct Width/Height? For now just default 0.0
                pass

            return {
                "p1": curve.GetEndPoint(0),
                "p2": curve.GetEndPoint(1),
                "id": eid_val,
                "diameter": diam,
                "level_id": el.LevelId
            }
        except:
            return None

    def clear_preview(self, app_state):
        if not app_state.preview_ids: return
        
        t = Transaction(self.doc, "Clear Preview")
        t.Start()
        try:
            for eid in app_state.preview_ids:
                try: self.doc.Delete(eid)
                except: pass
            app_state.preview_ids = []
            t.Commit()
        except: t.RollBack()

    def visualize_solution(self, solution, app_state):
        """
        Draws temporary DirectShape lines for the solution.
        """
        self.clear_preview(app_state) # Clear old first
        
        if not solution or not solution.context: return
        
        ctx = solution.context
        pt_a = ctx.get("meet_pt_a")
        pt_b = ctx.get("meet_pt_b")
        id_a = ctx.get("id_a")
        id_b = ctx.get("id_b")
        
        if not pt_a or not pt_b: return

        # Get start points of existing pipes to draw the "extension"
        p_start_a = None
        p_start_b = None
        
        el_a = self.doc.GetElement(ElementId(id_a))
        if el_a:
            # Find the end farthest from the meeting point to draw the full line?
            # Or just draw the segment from current closest end to meet point?
            # Let's draw the FULL proposed centerline for clarity.
            c = el_a.Location.Curve
            d0 = c.GetEndPoint(0).DistanceTo(pt_a)
            d1 = c.GetEndPoint(1).DistanceTo(pt_a)
            fixed_pt = c.GetEndPoint(1) if d0 < d1 else c.GetEndPoint(0)
            p_start_a = fixed_pt
            
        el_b = self.doc.GetElement(ElementId(id_b))
        if el_b:
            c = el_b.Location.Curve
            d0 = c.GetEndPoint(0).DistanceTo(pt_b)
            d1 = c.GetEndPoint(1).DistanceTo(pt_b)
            fixed_pt = c.GetEndPoint(1) if d0 < d1 else c.GetEndPoint(0)
            p_start_b = fixed_pt

        t = Transaction(self.doc, "Draw Preview")
        t.Start()
        try:
            new_ids = []
            
            # Helper to create DS
            def create_ds_line(p1, p2):
                if p1.DistanceTo(p2) < 0.01: return # Skip short lines
                try:
                    ds = DirectShape.CreateElement(self.doc, ElementId(BuiltInCategory.OST_PipeCurves))
                    line = Line.CreateBound(p1, p2)
                    ds.SetShape(List[GeometryObject]([line]))
                    new_ids.append(ds.Id)
                except: pass

            from Autodesk.Revit.DB import GeometryObject
            
            # Draw Path A
            if p_start_a: create_ds_line(p_start_a, pt_a)
            
            # Draw Path B
            if p_start_b: create_ds_line(p_start_b, pt_b)
            
            # Draw Bridge (if distinct)
            if pt_a.DistanceTo(pt_b) > 0.01:
                create_ds_line(pt_a, pt_b)
                
            app_state.preview_ids = new_ids
            self.doc.Regenerate() # Force update
            t.Commit()
            self.uidoc.RefreshActiveView()
        except Exception as e:
            t.RollBack()
            print("Preview Error: " + str(e))
    
    def commit_solution(self, solution):
        """
        Executes the actual modeling transaction.
        """
        if not solution.is_valid: return False
        
        t = Transaction(self.doc, "Join Pipes")
        t.Start()
        try:
            # 1. Resize/Move existing pipes to the new endpoints
            # Assuming solution.steps contains specific instructions or points
            # For this MVP, let's assume solution.context stores the calculated points:
            # context = {"meet_pt_a": XYZ, "meet_pt_b": XYZ}
            
            ctx = solution.context
            if not ctx: 
                t.RollBack(); return False
                
            pt_a = ctx.get("meet_pt_a")
            pt_b = ctx.get("meet_pt_b")
            
            el_a = self.doc.GetElement(ElementId(ctx.get("id_a")))
            el_b = self.doc.GetElement(ElementId(ctx.get("id_b")))
            
            if not el_a or not el_b:
                t.RollBack(); return False

            # Helper to move closest end to target
            def adjust_end(element, new_pt):
                c = element.Location.Curve
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                
                # Use SetEndPoint? Only works for some curves.
                # Safer: Create new curve? Or use LocationCurve.set_Curve?
                # Simplest for Linear: Determine which end is closer and move it.
                
                dist0 = p0.DistanceTo(new_pt)
                dist1 = p1.DistanceTo(new_pt)
                
                # Check if we need to modify
                if dist0 < 0.01: pass # Already there
                if dist1 < 0.01: pass
                
                # Try extending
                if dist0 < dist1:
                    # Move p0 to new_pt
                    new_c = c.Clone()
                    # Revit API Line binding...
                    try:
                        # For bound lines, we just need to replace the curve
                        from Autodesk.Revit.DB import Line
                        new_line = Line.CreateBound(new_pt, p1)
                        element.Location.Curve = new_line
                    except: pass
                else:
                    # Move p1 to new_pt
                    try:
                        from Autodesk.Revit.DB import Line
                        new_line = Line.CreateBound(p0, new_pt)
                        element.Location.Curve = new_line
                    except: pass
                    
            if not el_a or not el_b:
                t.RollBack(); return False

            # Helper to move closest end to target
            def adjust_end(element, new_pt):
                c = element.Location.Curve
                p0 = c.GetEndPoint(0)
                p1 = c.GetEndPoint(1)
                
                # Use SetEndPoint? Only works for some curves.
                # Safer: Create new curve? Or use LocationCurve.set_Curve?
                # Simplest for Linear: Determine which end is closer and move it.
                
                dist0 = p0.DistanceTo(new_pt)
                dist1 = p1.DistanceTo(new_pt)
                
                # Check if we need to modify
                if dist0 < 0.01: pass # Already there
                if dist1 < 0.01: pass
                
                # Try extending
                if dist0 < dist1:
                    # Move p0 to new_pt
                    new_c = c.Clone()
                    # Revit API Line binding...
                    try:
                        # For bound lines, we just need to replace the curve
                        from Autodesk.Revit.DB import Line
                        new_line = Line.CreateBound(new_pt, p1)
                        element.Location.Curve = new_line
                    except: pass
                else:
                    # Move p1 to new_pt
                    try:
                        from Autodesk.Revit.DB import Line
                        new_line = Line.CreateBound(p0, new_pt)
                        element.Location.Curve = new_line
                    except: pass
            
            def get_conn_at(elem, pt):
                try:
                    m = elem.ConnectorManager
                    if not m: return None
                    for c in m.Connectors:
                        if c.Origin.DistanceTo(pt) < 0.1: # 0.1 ft tolerance
                            return c
                except: pass
                return None

            adjust_end(el_a, pt_a)
            adjust_end(el_b, pt_b)
            
            self.doc.Regenerate() # CRITICAL: Update connectors after moving geometry
            
            # 2. Create Connectors/Elbows
            # If pt_a == pt_b, create elbow
            # If dist > 0, create pipe + 2 elbows
            
            dist = pt_a.DistanceTo(pt_b)
            
            if dist < 0.01: # Intersection
                # Find connectors at this location
                # This is tricky without connector manager. 
                # Ideally use doc.Create.NewElbowFitting(connector1, connector2)
                
                c1 = get_conn_at(el_a, pt_a)
                c2 = get_conn_at(el_b, pt_b)
                
                if c1 and c2:
                    try:
                        self.doc.Create.NewElbowFitting(c1, c2)
                    except Exception as e:
                        print("Fitting Error: " + str(e))
                        t.RollBack(); return False
                else:
                    print("Connectors not found at intersection.")
                    t.RollBack(); return False
            
            else: # Rolling Offset / Connecting Pipe
                # Create intermediate pipe
                # We need SystemTypeId and LevelId and PipeTypeId
                
                # Copy properties from el_a
                sys_id = el_a.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsElementId()
                type_id = el_a.GetTypeId()
                level_id = el_a.LevelId
                
                new_pipe = None
                try:
                    # Create Pipe
                    # doc.Create.NewPipe(start, end, type_id) # older API
                    # Pipe.Create(doc, system_type_id, pipe_type_id, level_id, start, end) # 2018+
                    from Autodesk.Revit.DB.Plumbing import Pipe
                    new_pipe = Pipe.Create(self.doc, sys_id, type_id, level_id, pt_a, pt_b)
                except Exception as e:
                    print("Pipe Create Error: " + str(e))
                
                if new_pipe:
                    # Connect Elbows
                    # Refetch connectors because geometry changed
                    self.doc.Regenerate()
                    
                    # Connect A -> New
                    c1 = get_conn_at(el_a, pt_a)
                    c_new_1 = get_conn_at(new_pipe, pt_a)
                    if c1 and c_new_1: 
                        try: self.doc.Create.NewElbowFitting(c1, c_new_1)
                        except: pass
                        
                    # Connect B -> New
                    c2 = get_conn_at(el_b, pt_b)
                    c_new_2 = get_conn_at(new_pipe, pt_b)
                    if c2 and c_new_2: 
                        try: self.doc.Create.NewElbowFitting(c2, c_new_2)
                        except: pass

            t.Commit()
            return True
        except Exception as e:
            print("Commit Error: " + str(e)) # Debug
            t.RollBack()
            return False