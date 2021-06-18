'========================================================================
' @description:
' ReplaceBoxCharacterWithCheckboxElement()
'  replace empty checkbox character ('ChrW(&H25a1)) with MS Word checkbox element
' @Author:  Anatoliy Sakhno   Anatoly.Sakhno@gmail.com
'  @Date   :    6/18/2021
'========================================================================
Sub ReplaceBoxCharacterWithCheckboxElement()

    Dim Rng As Range
    Dim Fnd As Boolean
    
    Application.DisplayStatusBar = True
    Set Rng = Selection.Range
    With Rng.Find
        .ClearFormatting
        .Execute FindText:=ChrW(&H25A1), Forward:=True, _
                 Format:=False, Wrap:=wdFindStop
        Fnd = .Found
    End With
    Let Index = 0
    Do While Fnd = True
        With Rng
            .ContentControls.Add (wdContentControlCheckBox)
        End With
        Set Rng = Selection.Range
        With Rng.Find
            .ClearFormatting
            .Execute FindText:=ChrW(&H25A1), Forward:=True, _
                     Format:=False, Wrap:=wdFindStop
            Fnd = .Found
        End With
        DoEvents
        Index = Index + 1
        
        Application.StatusBar = "Replacing checkbox characters: " & Index & " replacements made..."
        
    Loop
End Sub

