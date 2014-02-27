'****h* Tech/FreekConversie
' NAME
'   FreekTextFileTools
' COPYRIGHT
'   (c) Freek van Gilst
'   Zie voor gebruiksvoorwaarden http://www.vangilst.com/Voorwaarden.html
'   Onder andere houdt dat in dat er geen garantie wordt gegeven en het niet
'   zonder meer is toegestaan om dit te kopiëren.
'   Voor sommige organisaties zijn uitzonderingen mogelijk, maar ALLEEN WANNEER
'   DEZE TEKST OVER COPYRIGHT WORDT MEEGEKOPIEERD wanneer deze module of delen
'   ervan wordt gebruikt!
'   Lees voor gebruik hiervan eerst de genoemde webpagina.
' AUTHOR
'   Freek van Gilst, e-mail mailto:freek@vangilst.com
' FUNCTION
'   Deze module bevat functies om te converteren tussen verschillende dataypes.
'****

Option Explicit

'****f* FreekConversie/CADbl
' FUNCTION
'   Maakt een array met Double waardes van de invoer.
'   De array begint op index 1.
'   Als het een Range is dan wordt het een 1-dimensionale array.
'   Als het een scalar is dan wordt het een array met lengte 1.
'   Als het een meerdimensionale matrix is dan is het resultaat alleen de
'   eerste dimensie
' SYNOPSIS
Public Function CADbl(V As Variant) As Double()
' INPUTS
'   V: Variant, de te converteren vector
' RESULT
'   Double(): de resulterende vector
'*****
Dim Ix As Long         'Index in de resultaatvector
Dim TweeDim As Boolean 'Is de vector tweedimensionaal (of meer)?
Dim Rslt() As Double   'Het resultaat
Dim Ntl As Long        'Het aantal elementen
Dim IxV As Long        'Index in de vector V

    If TypeName(V) = "Range" Then V = V.Value2
        
    '2-dimensionaal?
    TweeDim = False
    On Error Resume Next
    TweeDim = IsNumeric(UBound(V, 2))
    
    If TweeDim Then
        Ntl = UBound(V, 1) - LBound(V, 1) + 1
        'Eventueel transponeren
        If Ntl = 1 Then
            Ntl = UBound(V, 2) - LBound(V, 2) + 1
            ReDim Rslt(1 To Ntl)
            IxV = LBound(V, 2)
            For Ix = 1 To Ntl
                Rslt(Ix) = CDbl(V(1, IxV))
                IxV = IxV + 1
            Next Ix
        Else 'Hoeft niet te transponeren want eerste dimensie langer dan 1
            ReDim Rslt(1 To Ntl)
            IxV = LBound(V, 1)
            For Ix = 1 To Ntl
                Rslt(Ix) = CDbl(V(IxV, 1))
                IxV = IxV + 1
            Next Ix
        End If
    Else 'Niet meerdimensionaal
        If Not IsArray(V) Then 'Geen array dus array van maken
            ReDim Rslt(1 To 1)
            Rslt(1) = V
        Else
            Ntl = UBound(V) - LBound(V) + 1
            If Ntl > 0 Then
                ReDim Rslt(1 To Ntl)
                IxV = LBound(V)
                For Ix = 1 To Ntl
                    Rslt(Ix) = CDbl(V(IxV))
                    IxV = IxV + 1
                Next Ix
            End If
        End If
    End If
    CADbl = Rslt
End Function


'****f* FreekConversie/CALng
' FUNCTION
'   Maakt een array met Long waardes van de invoer.
'   De array begint op index 1.
'   Als het een Range is dan wordt het een 1-dimensionale array.
'   Als het een scalar is dan wordt het een array met lengte 1.
'   Als het een meerdimensionale matrix is dan is het resultaat alleen de
'   eerste dimensie
' SYNOPSIS
Public Function CALng(V As Variant) As Long()
' INPUTS
'   V: Variant, de te converteren vector
' RESULT
'   Long(): de resulterende vector
'*****
Dim Ix As Long         'Index in de resultaatvector
Dim TweeDim As Boolean 'Is de vector tweedimensionaal (of meer)?
Dim Rslt() As Long   'Het resultaat
Dim Ntl As Long        'Het aantal elementen
Dim IxV As Long        'Index in de vector V

    If TypeName(V) = "Range" Then V = V.Value2
        
    '2-dimensionaal?
    TweeDim = False
    On Error Resume Next
    TweeDim = IsNumeric(UBound(V, 2))
    
    If TweeDim Then
        Ntl = UBound(V, 1) - LBound(V, 1) + 1
        'Eventueel transponeren
        If Ntl = 1 Then
            Ntl = UBound(V, 2) - LBound(V, 2) + 1
            ReDim Rslt(1 To Ntl)
            IxV = LBound(V, 2)
            For Ix = 1 To Ntl
                Rslt(Ix) = CDbl(V(1, IxV))
                IxV = IxV + 1
            Next Ix
        Else 'Hoeft niet te transponeren want eerste dimensie langer dan 1
            ReDim Rslt(1 To Ntl)
            IxV = LBound(V, 1)
            For Ix = 1 To Ntl
                Rslt(Ix) = CDbl(V(IxV, 1))
                IxV = IxV + 1
            Next Ix
        End If
    Else 'Niet meerdimensionaal
        If Not IsArray(V) Then 'Geen array dus array van maken
            ReDim Rslt(1 To 1)
            Rslt(1) = V
        Else
            Ntl = UBound(V) - LBound(V) + 1
            If Ntl > 0 Then
                ReDim Rslt(1 To Ntl)
                IxV = LBound(V)
                For Ix = 1 To Ntl
                    Rslt(Ix) = CDbl(V(IxV))
                    IxV = IxV + 1
                Next Ix
            End If
        End If
    End If
    CALng = Rslt
End Function



