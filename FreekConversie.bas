Attribute VB_Name = "FreekConversie"
'****h* Tech/FreekConversie
' NAME
'   FreekConversie
' COPYRIGHT
'   (c) Freek van Gilst
'   Zie voor gebruiksvoorwaarden http://www.vangilst.com/Voorwaarden.html
'   Onder andere houdt dat in dat er geen garantie wordt gegeven en het niet
'   zonder meer is toegestaan om dit te kopiëren.
'   Voor sommige organisaties zijn uitzonderingen mogelijk, maar ALLEEN WANNEER
'   DEZE TEKST OVER COPYRIGHT WORDT MEEGEKOPIEERD wanneer deze module of een
'   deel ervan wordt gebruikt!
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
'   De array begint default op index 0.
'   Als het een Range is dan wordt het een 1-dimensionale array.
'   Als het een scalar is dan wordt het een array met lengte 1.
'   Als het een meerdimensionale matrix is dan is het resultaat alleen de
'   eerste dimensie.
'   Als een element zelf ook een array is, geeft dan het eerste element daarvan
'   (eventueel recursief).
'   Als een element een object is, dan wordt het de waarde van de default
'   property.
'   Als een element een type is of anders niet geconverteerd kan worden, geeft
'   dan een foutmelding.
' SYNOPSIS
Public Function CADbl(V As Variant, Optional Base As Long = 0) As Double()
' INPUTS
'   V: Variant, de te converteren vector
'   Base: Long, optional, de index waarop de array begint, default is 0.
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
    On Error GoTo 0
    
    If TweeDim Then
        Ntl = UBound(V, 1) - LBound(V, 1) + 1
        'Eventueel transponeren
        If Ntl = 1 Then
            Ntl = UBound(V, 2) - LBound(V, 2) + 1
            ReDim Rslt(Base To Ntl + Base - 1)
            IxV = LBound(V, 2)
            For Ix = Base To Ntl + Base - 1
                Rslt(Ix) = CDbl(V(1, IxV))
                IxV = IxV + 1
            Next Ix
        Else 'Hoeft niet te transponeren want eerste dimensie langer dan 1
            ReDim Rslt(Base To Ntl + Base - 1)
            IxV = LBound(V, 1)
            For Ix = Base To Ntl + Base - 1
                Rslt(Ix) = CDbl(V(IxV, 1))
                IxV = IxV + 1
            Next Ix
        End If
    Else 'Niet meerdimensionaal
        If Not IsArray(V) Then 'Geen array dus array van maken
            ReDim Rslt(Base To Base)
            Rslt(Base) = CDbl(V)
        Else
            Ntl = UBound(V) - LBound(V) + 1
            If Ntl > 0 Then
                ReDim Rslt(Base To Ntl + Base - 1)
                IxV = LBound(V)
                For Ix = Base To Ntl + Base - 1
                    If IsArray(V(IxV)) Then
                        Rslt(Ix) = CADbl(V(IxV))(0)
                    Else
                        Rslt(Ix) = CDbl(V(IxV))
                    End If
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
'   De array begint default op index 0.
'   Als het een Range is dan wordt het een 1-dimensionale array.
'   Als het een scalar is dan wordt het een array met lengte 1.
'   Als het een meerdimensionale matrix is dan is het resultaat alleen de
'   eerste dimensie.
'   Als een element zelf ook een array is, geeft dan het eerste element daarvan
'   (eventueel recursief).
'   Als een element een object is, dan wordt het de waarde van de default
'   property.
'   Als een element een type is of anders niet geconverteerd kan worden, geeft
'   dan een foutmelding.
' SYNOPSIS
Public Function CALng(V As Variant, Optional Base As Long = 0) As Long()
' INPUTS
'   V: Variant, de te converteren vector
'   Base: Long, optional, de index waarop de array begint, default is 0.
' RESULT
'   Long(): de resulterende vector
'*****
Dim Ix As Long         'Index in de resultaatvector
Dim TweeDim As Boolean 'Is de vector tweedimensionaal (of meer)?
Dim Rslt() As Long     'Het resultaat
Dim Ntl As Long        'Het aantal elementen
Dim IxV As Long        'Index in de vector V

    If TypeName(V) = "Range" Then V = V.Value2
        
    '2-dimensionaal?
    TweeDim = False
    On Error Resume Next
    TweeDim = IsNumeric(UBound(V, 2))
    On Error GoTo 0
    
    If TweeDim Then
        Ntl = UBound(V, 1) - LBound(V, 1) + 1
        'Eventueel transponeren
        If Ntl = 1 Then
            Ntl = UBound(V, 2) - LBound(V, 2) + 1
            ReDim Rslt(Base To Ntl + Base - 1)
            IxV = LBound(V, 2)
            For Ix = Base To Ntl + Base - 1
                Rslt(Ix) = CLng(V(1, IxV))
                IxV = IxV + 1
            Next Ix
        Else 'Hoeft niet te transponeren want eerste dimensie langer dan 1
            ReDim Rslt(Base To Ntl + Base - 1)
            IxV = LBound(V, 1)
            For Ix = Base To Ntl + Base - 1
                Rslt(Ix) = CLng(V(IxV, 1))
                IxV = IxV + 1
            Next Ix
        End If
    Else 'Niet meerdimensionaal
        If Not IsArray(V) Then 'Geen array dus array van maken
            ReDim Rslt(Base To Base)
            Rslt(Base) = CLng(V)
        Else
            Ntl = UBound(V) - LBound(V) + 1
            If Ntl > 0 Then
                ReDim Rslt(Base To Ntl + Base - 1)
                IxV = LBound(V)
                For Ix = Base To Ntl + Base - 1
                    If IsArray(V(IxV)) Then
                        Rslt(Ix) = CALng(V(IxV))(0)
                    Else
                        Rslt(Ix) = CLng(V(IxV))
                    End If
                    IxV = IxV + 1
                Next Ix
            End If
        End If
    End If
    CALng = Rslt
End Function



