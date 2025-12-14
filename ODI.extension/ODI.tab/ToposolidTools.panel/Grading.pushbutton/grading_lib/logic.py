from Autodesk.Revit.DB import Transaction, SubTransaction

def apply_points_to_toposolid(doc, toposolid, points):
    """
    Safely adds points to a Toposolid/Floor using a SubTransaction.
    """
    if not hasattr(toposolid, "GetSlabShapeEditor"):
        return False, "Element does not support Shape Editing."

    t = Transaction(doc, "Carve Toposolid")
    t.Start()
    st = SubTransaction(doc)
    st.Start()
    
    try:
        editor = toposolid.GetSlabShapeEditor()
        for pt in points:
            editor.AddPoint(pt)
        
        st.Commit()
        t.Commit()
        return True, "Success! Added {} points.".format(len(points))
        
    except Exception as e:
        st.RollBack()
        t.Commit()
        return False, "Geometry Error: {}".format(e)