Attribute VB_Name = "Container_Module"

Public Function Contains(str As String, rng As Range) As Variant
    
    Dim r As Range
    
    Dim myarray() As String
    
    Dim n As Long
    
    n = 0
    
    If LCase(str) = "all" Then
    
        Contains = rng
    
    Else
    
        ReDim Preserve myarray(n)
    
        Contains = ""
    
        For Each r In rng
        
            If InStr(1, LCase(r.Value), str) > 0 Then
        
                ReDim Preserve myarray(0 To n)
            
                myarray(n) = r.Value
            
                n = n + 1
        
            End If
        
        Next r
    
        If IsError(Application.Match("*", myarray(), 0)) Then
 
            Contains = "No results found"
            
        Else
                  
            
            Contains = Application.Unique(Application.Transpose(myarray()))
            
    
        End If

    End If
    
End Function
