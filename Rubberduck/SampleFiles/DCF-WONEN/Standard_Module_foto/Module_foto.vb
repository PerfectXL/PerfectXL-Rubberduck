Attribute VB_Name = "Module_foto"
Sub mcrFotos()
Dim ImgFileFormat As String, picObject As Variant, picLocatie As Variant
On Error Resume Next
ImgFileFormat = "Image Files jpg (*.jpg),*.jpg,(*.bmp),others, tif (*.tif),*.tif"

With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Stap 1: Kies foto object"
    .AllowMultiSelect = False
    If .Show = True Then
        picObject = .SelectedItems(1)
    Else
        Exit Sub
    End If

End With

With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Stap 2: Kies foto locatie"
    .AllowMultiSelect = False
    If .Show = True Then
        picLocatie = .SelectedItems(1)
    Else
        Exit Sub
    End If
End With

On Error GoTo 0

Sheets("NRVT rapport (ENG)").Select
ActiveSheet.Unprotect "MVGM"
Sheets("NRVT rapport (ENG)").Shapes.Range(Array("NRVT_object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picObject
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue

ActiveSheet.Protect "MVGM"

End With

Sheets("NRVT rapport (ENG)").Select
ActiveSheet.Unprotect "MVGM"

Sheets("NRVT rapport (ENG)").Select
ActiveSheet.Unprotect "MVGM"
Sheets("NRVT rapport (ENG)").Shapes.Range(Array("NRVT_locatie")).Select

With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picLocatie
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue

ActiveSheet.Protect "MVGM"

End With

Sheets("NRVT rapport (NL)").Select
ActiveSheet.Unprotect "MVGM"
Sheets("NRVT rapport (NL)").Shapes.Range(Array("NRVT_NL_object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picObject
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue

ActiveSheet.Protect "MVGM"

End With

Sheets("NRVT rapport (NL)").Select
ActiveSheet.Unprotect "MVGM"

Sheets("NRVT rapport (NL)").Select
ActiveSheet.Unprotect "MVGM"
Sheets("NRVT rapport (NL)").Shapes.Range(Array("NRVT_NL_locatie")).Select

With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picLocatie
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue

ActiveSheet.Protect "MVGM"

End With

Sheets("Voorblad Altera").Select
ActiveSheet.Unprotect "MVGM"
Sheets("Voorblad Altera").Shapes.Range(Array("Voorblad_Altera_object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picLocatie
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue

ActiveSheet.Protect "MVGM"

End With

Sheets("Voorblad CBRE").Select
ActiveSheet.Unprotect "MVGM"
Sheets("Voorblad CBRE").Shapes.Range(Array("Voorblad_CBRE_object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picObject
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue
ActiveSheet.Protect "MVGM"

End With

Sheets("Summary Amvest").Select
ActiveSheet.Unprotect "MVGM"
Sheets("Summary Amvest").Shapes.Range(Array("Summary_Amvest_object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picObject
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue
ActiveSheet.Protect "MVGM"

End With

Sheets("VOORBLAD").Select
ActiveSheet.Unprotect "MVGM"
Sheets("VOORBLAD").Shapes.Range(Array("VOORBLAD_object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picObject
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue
ActiveSheet.Protect "MVGM"
End With

Sheets("Engels (print)").Select
ActiveSheet.Unprotect "MVGM"
Sheets("Engels (print)").Shapes.Range(Array("Engels_(print)_object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picObject
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue
ActiveSheet.Protect "MVGM"

End With

Sheets("Voorblad NRVT rapport").Select
ActiveSheet.Unprotect "MVGM"
Sheets("Voorblad NRVT rapport").Shapes.Range(Array("txtVoorblad_Object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picObject
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue
ActiveSheet.Protect "MVGM"

End With

Sheets("FACTSHEET Achmea").Select
ActiveSheet.Unprotect "MVGM"
Sheets("FACTSHEET Achmea").Shapes.Range(Array("Foto_object")).Select
With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picObject
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue

ActiveSheet.Protect "MVGM"

End With

Sheets("FACTSHEET Achmea").Select
ActiveSheet.Unprotect "MVGM"

Sheets("FACTSHEET Achmea").Select
ActiveSheet.Unprotect "MVGM"
Sheets("FACTSHEET Achmea").Shapes.Range(Array("Foto_locatie")).Select

With Selection.ShapeRange.Fill
    .Visible = msoTrue
    .UserPicture picLocatie
    .TextureTile = msoFalse
    .RotateWithObject = msoTrue

ActiveSheet.Protect "MVGM"

End With

End Sub

