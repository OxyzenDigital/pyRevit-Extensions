from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc

def draw_box(sheet, min_pt, max_pt, color=None):
    try:
        p1 = DB.XYZ(min_pt.X, min_pt.Y, 0)
        p2 = DB.XYZ(max_pt.X, min_pt.Y, 0)
        p3 = DB.XYZ(max_pt.X, max_pt.Y, 0)
        p4 = DB.XYZ(min_pt.X, max_pt.Y, 0)
        
        c1 = doc.Create.NewDetailCurve(sheet, DB.Line.CreateBound(p1, p2))
        c2 = doc.Create.NewDetailCurve(sheet, DB.Line.CreateBound(p2, p3))
        c3 = doc.Create.NewDetailCurve(sheet, DB.Line.CreateBound(p3, p4))
        c4 = doc.Create.NewDetailCurve(sheet, DB.Line.CreateBound(p4, p1))
        
        if color:
            ogs = DB.OverrideGraphicSettings()
            ogs.SetProjectionLineColor(color)
            doc.ActiveView.SetElementOverrides(c1.Id, ogs)
            doc.ActiveView.SetElementOverrides(c2.Id, ogs)
            doc.ActiveView.SetElementOverrides(c3.Id, ogs)
            doc.ActiveView.SetElementOverrides(c4.Id, ogs)
    except Exception as e:
        print("Failed to draw box: {}".format(e))

def main():
    if not isinstance(doc.ActiveView, DB.ViewSheet):
        forms.alert("Please run this tool from an active Sheet view.")
        return
        
    sheet = doc.ActiveView
    
    selection = revit.get_selection()
    vp_ids = [el.Id for el in selection if isinstance(el, DB.Viewport)]
    
    if not vp_ids:
        vp_ids = sheet.GetAllViewports()
        
    if not vp_ids:
        forms.alert("No viewports found on this sheet.")
        return

    with DB.Transaction(doc, "Draw Debug Bounds") as t:
        t.Start()
        
        for vp_id in vp_ids:
            vp = doc.GetElement(vp_id)
            view = doc.GetElement(vp.ViewId)
            
            # 1. Viewport.GetBoxOutline() - Red
            try:
                box = vp.GetBoxOutline()
                draw_box(sheet, box.MinimumPoint, box.MaximumPoint, DB.Color(255, 0, 0))
                print("Viewport.GetBoxOutline drawn in RED")
            except Exception as e:
                print("Failed to get Viewport Box Outline: {}".format(e))
                
            # 2. View.Outline (Scaled to Sheet) - Green
            try:
                outline = view.Outline
                scale = float(view.Scale)
                if scale <= 0: scale = 1.0
                w = (outline.Max.U - outline.Min.U) / scale
                h = (outline.Max.V - outline.Min.V) / scale
                
                # To place this on the sheet, we assume it's centered at vp.GetBoxCenter()
                center = vp.GetBoxCenter()
                min_pt = DB.XYZ(center.X - w/2, center.Y - h/2, 0)
                max_pt = DB.XYZ(center.X + w/2, center.Y + h/2, 0)
                
                draw_box(sheet, min_pt, max_pt, DB.Color(0, 255, 0))
                print("View.Outline (Scaled) drawn in GREEN")
            except Exception as e:
                print("Failed to get View Outline: {}".format(e))
                
            # 3. Viewport.GetLabelOutline() - Blue
            try:
                label_box = vp.GetLabelOutline()
                draw_box(sheet, label_box.MinimumPoint, label_box.MaximumPoint, DB.Color(0, 0, 255))
                print("Viewport.GetLabelOutline drawn in BLUE")
            except Exception as e:
                print("Failed to get Label Outline: {}".format(e))

        # 4. Create a dummy Drafting View and place it to test logic
        try:
            view_family_types = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType).ToElements()
            drafting_type = next((v for v in view_family_types if v.ViewFamily == DB.ViewFamily.Drafting), None)
            
            if drafting_type:
                drafting_view = DB.ViewDrafting.Create(doc, drafting_type.Id)
                drafting_view.Name = "DEBUG DRAFTING VIEW"
                
                # Draw a rectangle inside the drafting view so it has geometry
                p1 = DB.XYZ(0, 0, 0)
                p2 = DB.XYZ(10, 0, 0)
                p3 = DB.XYZ(10, 5, 0)
                p4 = DB.XYZ(0, 5, 0)
                doc.Create.NewDetailCurve(drafting_view, DB.Line.CreateBound(p1, p2))
                doc.Create.NewDetailCurve(drafting_view, DB.Line.CreateBound(p2, p3))
                doc.Create.NewDetailCurve(drafting_view, DB.Line.CreateBound(p3, p4))
                doc.Create.NewDetailCurve(drafting_view, DB.Line.CreateBound(p4, p1))
                
                # Place on sheet
                new_vp = DB.Viewport.Create(doc, sheet.Id, drafting_view.Id, DB.XYZ(1, 1, 0))
                print("Created dummy Drafting View to test Viewport geometry.")
        except Exception as e:
            print("Failed to create dummy Drafting View: {}".format(e))
            
        t.Commit()
        
    print("Done drawing debug boundaries.")

if __name__ == '__main__':
    main()
