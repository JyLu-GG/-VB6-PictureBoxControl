Attribute VB_Name = "mPictureBoxControl"
'' ============================================================
'' This is pictrue box control module.
'' Have some simply function for picturebox control practice.
''
'' Writer is J.Y.L by 2020/07/24
'' ============================================================

Option Explicit


Dim fs As New FileSystemObject


'' =================== Functions Area ===================

'' pbObj is picturebox object that do draw line B
'' leftupX is rectangle left X point
'' leftupY is rectangle left Y point
'' rightdownX is rectangle right X point
'' rightdownY is rectangle right Y point
'' color is pen color
'' Return True is work success; False is work fail
Public Function DrawRectangle(ByVal pbObj As PictureBox, ByVal leftupX As Double, ByVal leftupY As Double, ByVal rightdownX As Double, ByVal rightdownY As Double, _
ByVal color As Long)

On Error GoTo Err

    pbObj.Line (leftupX, leftupY)-(rightdownX, rightdownY), color, B
    
    DrawRectangle = True
    
    Exit Function
    
Err:
    DrawRectangle = False
    MsgBox ("DrawRectangle is fail" & vbCrLf & Err.Description)
    
End Function


'' pbObj is picturebox object that do draw line B
'' filepath is will load image path
'' Return True is work success; False is work fail
Public Function LoadImage(ByVal pbObj As PictureBox, ByVal filepath As String)

On Error GoTo Err
    
    If (TypeName(fs) <> "FileSystemObject") Then Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(filepath) = False Then LoadImage = False: MsgBox ("The file is not exists with : " & filepath): Exit Function
    
    pbObj.Picture = LoadPicture(filepath)
    
    LoadImage = True
    
    Exit Function
    
Err:
    LoadImage = False
    MsgBox ("LoadImage is fail" & vbCrLf & Err.Description)
    
End Function
    

'' img is picturebox display image
'' folderPath is will save file path, here will check folder exists and create it
'' filename is image file name, need have extensions
'' Return True is work success; False is work fail
Public Function SaveImage(Optional ByVal img As IPictureDisp, Optional ByVal folderPath As String, Optional ByVal filename As String)

On Error GoTo Err
    
    If (TypeName(fs) <> "FileSystemObject") Then Set fs = CreateObject("Scripting.FileSystemObject")
    
    If folderPath = "" Then SaveImage = False: MsgBox ("The folder path is empty"): Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    If fs.FolderExists(folderPath) = False Then fs.CreateFolder (folderPath)
    
    If img = 0 Or img.Type = 0 Then SaveImage = False: MsgBox ("The image is empty"): Exit Function
    If filename = "" Then SaveImage = False: MsgBox ("The file name is empty"): Exit Function
    
    
    Call SavePicture(img, folderPath & filename)
    
    SaveImage = True
    
    Exit Function
    
Err:
    SaveImage = False
    MsgBox ("SaveImage is fail" & vbCrLf & Err.Description)
    
End Function


'' pbObj is picturebox will show picturebox.picture
'' img is picturebox display image
'' Return True is work success; False is work fail
Public Function GetPictureIntoPictureBox(ByVal pbObj As PictureBox, ByVal img As IPictureDisp)

On Error GoTo Err
    
    pbObj.Picture = img
    
    GetPictureIntoPictureBox = True
    
    Exit Function
    
Err:
    GetPictureIntoPictureBox = False
    MsgBox ("GetPictureIntoPictureBox is fail" & vbCrLf & Err.Description)
    
End Function


'' pbObj is picturebox will show picturebox.image
'' img is picturebox display image
'' Return True is work success; False is work fail
Public Function GetImageIntoPictureBox(ByVal pbObj As PictureBox, ByVal img As IPictureDisp)

On Error GoTo Err
    
    '' pbObj.Image = img '' <= object doesn't support property or method
    pbObj.Picture = img
    
    GetImageIntoPictureBox = True
    
    Exit Function
    
Err:
    GetImageIntoPictureBox = False
    MsgBox ("GetImageIntoPictureBox is fail" & vbCrLf & Err.Description)
    
End Function

'' =================== Functions Area ===================
