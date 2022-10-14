# Replacing or Substituting Strings in VBA

[TOC]

## The Replace Function

Sub *ReplacingStrings*()

    Dim s As String
    
    s = Range("A4").Value
    Debug.Print s 'King Kong


    s = Replace(s, "K", "D")
    Debug.Print s 'Ding Dong

End Sub

## Replacing Single Characters

`s = Replace(s, "K", "D")`

## Replacing Multiple Characters

`s = Replace(s, "K", "Dilly D")`

`s = Replace(s, "Gr", "")`

`s = Replace(s, "wilight", "errible")`

## Controlling the Number of Replacements

`s = Replace("Fast & Furious 6", "F", "L",,1)` => *"Last & Furious 6"*

## Setting the Start Position

Sub *ReplacingStrings*()

    Dim s As String
    Dim i As Integer
    
    s = "Harry Potter and the Goblet of Fire"
     
    i = InStr(1, s, "o")
    s = Left(s, i) & Replace(s, "o", "i", i + 1, 1) 'Harry Potter and the Giblet of Fire
    Debug.Print s

End Sub

## Dealing with Case-Sensitivity

Sub *CaseSensitivity*() 'Case-Sensitive by default

    Dim s As String
    
    s = "The Lord of the Rings: Return of the King"
    
    s=Replace(s, "R", "W") => 'The Lord of the Wings: Weturn of the King
    s=Replace(s, "R", "W", , , vbTextCompare) => 'The LoWd of the Wings: WetuWn of the King
    s=StrConv(s, vbProperCase) => 'The Lowd Of The Wings: Wetuwn Of The Kin
    
    s =  "Ant-Man"
    
    s=Replace(s, "a", "e", , , vbTextCompare) => 'ent-Men '
    s=StrConv(s, vbProperCase) => 'Ent-men 
    s=Replace(StrConv(Replace(s, "-", " "), vbProperCase), " ", "-") => Ent-Men

End Sub

## A Practical Example for Generating Valid File Names

Sub *CreateFilmFiles*()

    Dim fso As New Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim r As Range
    Dim FolderPath As String
    Dim FileName As String
    Dim Cols As Integer
    
    Dim IllegalCharts()
    Dim v As Variant
    
    IllegalCharts = Array("<", ">", ":", """", "/", "\", "|", "?", "*")
    
    FolderPath = Environ("UserProfile") & "\Desktop\Films"
    
    If fso.FolderExists(FolderPath) Then fso.DeleteFolder FolderPath
    
    fso.CreateFolder FolderPath
    
    Cols = Range("A1").CurrentRegion.Columns.Count
    
    For Each r In Range("A2", Range("A1").End(xlDown))
    
        FileName = r.Value
        
        For Each v In IllegalCharts
            FileName = Replace(FileName, v, "")
        Next v
        
        FileName = FileName & ".txt"
        
        Set ts = fso.OpenTextFile(FolderPath & "\" & FileName, ForAppending, True)
        
        'write to text
        ts.WriteLine _
            Join(Application.Transpose(Application.Transpose(r.Resize(1, Cols).Value)), vbTab)
        
        ts.Close
        
        Set ts = Nothing
        
    Next r

End Sub