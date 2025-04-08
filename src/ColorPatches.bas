Attribute VB_Name = "ColorPatches"
'===============================================================================
'   Макрос          : ColorPatches
'   Версия          : 2025.04.08
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "ColorPatches"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_FILEBASENAME As String = "elvin_" & APP_NAME
Public Const APP_VERSION As String = "2025.04.08"
Public Const APP_URL As String = "https://vk.com/elvin_macro/" & APP_NAME

'===============================================================================
' # Globals

Private Const COLOR_WELL_TAG As String = "*COLOR_WELL"
Private Const COLOR_TEXT_TAG As String = "*COLOR_TEXT"

'===============================================================================
' # Entry points

Sub ProcessPatches()

    #If DEV = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Shapes As ShapeRange
    If Not InputData.ExpectShapes.Ok(Shapes) Then GoTo Finally
    
    BoostStart APP_DISPLAYNAME
    
    TagValidPatchGroups Shapes
    If Shapes.Count = 0 Then
        Warn "Не найдено подходящих групп."
        GoTo Finally
    End If
    
    Dim Source As ShapeRange: Set Source = ActiveSelectionRange
    
    PatchesRoutine Shapes
    
    If IsSome(Source) Then Source.CreateSelection
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers
 
Private Sub PatchesRoutine(ByVal ValidShapes As ShapeRange)
    Dim GroupShape As Shape
    For Each GroupShape In ValidShapes
        ProcessPatchGroup GroupShape
    Next GroupShape
End Sub

Private Sub ProcessPatchGroup(ByVal PatchGroup As Shape)
    Dim Well As Shape: Set Well = PatchGroup.Shapes(COLOR_WELL_TAG)
    Dim Text As Shape: Set Text = PatchGroup.Shapes(COLOR_TEXT_TAG)
    Dim Color As Color: Set Color = Well.Fill.UniformColor.GetCopy
    
    Text.Text.Story.Text = ColorToShowable(Color)
    
    If IsOverlapBox(Well, Text) Then
        If GetColorLightness(Color) > 127 Then
            Text.Fill.ApplyUniformFill GrayBlack
        Else
            Text.Fill.ApplyUniformFill GrayWhite
        End If
    End If
End Sub
 
Private Sub TagValidPatchGroups(ByRef Shapes As ShapeRange)
    Dim ValidShapes As New ShapeRange
    Dim Shape As Shape
    For Each Shape In Shapes
        If Shape.Type = cdrGroupShape Then
            If TagValidPatchGroup(Shape) = Ok Then ValidShapes.Add Shape
        End If
    Next Shape
    Set Shapes = ValidShapes
End Sub

Private Function TagValidPatchGroup(ByRef GroupShape As Shape) As BooleanResult
    With GroupShape.Shapes
        If Not .Count = 2 Then Exit Function
        Dim TextIndex As Long: TextIndex = TextShapeIndex(.All)
        Dim WellIndex As Long
        If TextIndex = 1 Then
            WellIndex = 2
        ElseIf TextIndex = 2 Then
            WellIndex = 1
        Else
            Exit Function
        End If
        If Not ShapeHasUniformFill(.Item(WellIndex)) Then Exit Function
        .Item(TextIndex).Name = COLOR_TEXT_TAG
        .Item(WellIndex).Name = COLOR_WELL_TAG
    End With
    TagValidPatchGroup = Ok
End Function

Private Property Get TextShapeIndex(ByVal Shapes As ShapeRange) As Long
    Dim i As Long
    For i = 1 To Shapes.Count
        If Shapes(i).Type = cdrTextShape Then
            TextShapeIndex = i
            Exit For
        End If
    Next i
End Property

'===============================================================================
' # Cfg helpers

Public Function ShowColorCaptionView(ByRef Cfg As Dictionary) As BooleanResult
    Dim FileBinder As JsonFileBinder: Set FileBinder = BindConfig
    Set Cfg = FileBinder.GetOrMakeSubDictionary("ColorCaption")
    Dim View As New ColorCaptionView
    Dim ViewBinder As ViewToDictionaryBinder: Set ViewBinder = _
        ViewToDictionaryBinder.New_( _
            Dictionary:=Cfg, _
            View:=View, _
            ControlNames:=Pack("Offset") _
        )
    View.Show vbModal
    ViewBinder.RefreshDictionary
    ShowColorCaptionView = View.IsOk
End Function

Private Function BindConfig() As JsonFileBinder
    Set BindConfig = JsonFileBinder.New_(APP_FILEBASENAME)
End Function

'===============================================================================
' # Tests

Private Sub TestSomething()
    Show Application.FontList(1)
End Sub
