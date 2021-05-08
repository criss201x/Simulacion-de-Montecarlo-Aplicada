Attribute VB_Name = "Módulo1"
Function in_area(valor_x, valor_y, punto_x, punto_y) As Boolean

Dim N As Integer, J As Integer, C As Integer
Dim YC As Double

N = valor_x.Count

'la figura tiene cierre?
 
If valor_x(1) <> valor_x(N) Or valor_y(1) <> valor_y(N) Then
    in_area = CVErr(xlErrVa1ue): Exit Function
End If

    
    
For J = 1 To N - 1
    If valor_x(J).Formula = "" Or valor_y(J).Formula = "" Then
        in_area = CVErr(xlErrVa1ue): Exit Function
    End If
    
    If punto_x >= valor_x(J) And punto_x > valor_x(J + 1) Then
        GoTo EOL
    End If
    
    If punto_x <= valor_x(J) And punto_x < valor_x(J + 1) Then
        GoTo EOL
    End If
    
    If punto_y >= valor_y(J) And punto_y > valor_y(J + 1) Then
        GoTo EOL
    End If
    
    
    YC = valor_y(J + 1) + (valor_y(J) - valor_y(J + 1)) * (punto_x - valor_x(J + 1)) / (valor_x(J) - valor_x(J + 1))
        
    If YC - punto_y > 0 Then
        C = C + 1
    End If
    
EOL: Next J

in_area = C Mod 2

End Function
