'****h* UnitTest/FreekConversie
' NAME
'   FreekConversie
' COPYRIGHT
'   (c) Freek van Gilst
'   Zie voor gebruiksvoorwaarden http://www.vangilst.com/Voorwaarden.html
'   Onder andere houdt dat in dat het niet zonder meer is toegestaan om dit te
'   kopiëren. Voor sommige organisaties zijn uitzonderingen mogelijk, maar
'   ALLEEN WANNEER DEZE TEKST OVER COPYRIGHT WORDT MEEGEKOPIEERD wanneer deze
'   module of delen ervan wordt gebruikt!
'   Lees voor gebruik hiervan eerst de genoemde webpagina.
' AUTHOR
'   Freek van Gilst, e-mail mailto:freek@vangilst.com
' FUNCTION
'   Deze module bevat de unittests voor de module FreekConversie
'****

Option Explicit

Public Sub TestConversie()
    TestCADbl
    TestCALng
End Sub

Public Sub TestCADbl()
    Debug.Print "TestCADbl Begonnen"
    
    If CADbl(Array("1"))(1) <> 1 Then Debug.Print "CADbl(Array(""1""))(1) <> 1"
    If TypeName(CADbl(Array("1"))(1)) <> "Double" Then Debug.Print "TypeName(CADbl(Array(""1""))(1)) <> ""Double"""
    If UBound(CADbl(Array(1))) <> 1 Then Debug.Print "UBound(CADbl(Array(1))) <> 1"
    If LBound(CADbl(Array(1))) <> 1 Then Debug.Print "LBound(CADbl(Array(1))) <> 1"
    If CADbl("3")(1) <> 3 Then Debug.Print CADbl("3")(1) <> 3
    If TypeName(CADbl("1")(1)) <> "Double" Then Debug.Print "TypeName(CADbl(""1"")(1)) <> ""Double"""
    If UBound(CADbl(1)) <> 1 Then Debug.Print "UBound(CADbl(1)) <> 1"
    If LBound(CADbl(1)) <> 1 Then Debug.Print "LBound(CADbl(1)) <> 1"
    If TypeName(CADbl(Array())) <> "Double()" Then Debug.Print "TypeName(CADbl(Array())) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!A1:A6"))) <> 6 Then Debug.Print "UBound(CADbl(Range(""UTCA!A1:A6""))) <> 6"
    If LBound(CADbl(Range("UTCA!A1:A6"))) <> 1 Then Debug.Print "LBound(CADbl(Range(""UTCA!A1:A6""))) <> 1"
    If CADbl(Range("UTCA!A1:A6"))(6) <> 2 Then Debug.Print "CADbl(Range(""UTCA!A1:A6""))(6) <> 2"
    If TypeName(CADbl(Range("UTCA!A1:A6"))) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!A1:A6""))) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!C1:H1"))) <> 6 Then Debug.Print "UBound(CADbl(Range(""UTCA!C1:H1""))) <> 6"
    If LBound(CADbl(Range("UTCA!C1:H1"))) <> 1 Then Debug.Print "LBound(CADbl(Range(""UTCA!C1:H1""))) <> 1"
    If CADbl(Range("UTCA!C1:H1"))(6) <> 2 Then Debug.Print "CADbl(Range(""UTCA!C1:H1""))(6) <> 2"
    
    If UBound(CADbl(Range("UTCA!H1:H1"))) <> 1 Then Debug.Print "UBound(CADbl(Range(""UTCA!H1:H1""))) <> 1"
    If LBound(CADbl(Range("UTCA!H1:H1"))) <> 1 Then Debug.Print "LBound(CADbl(Range(""UTCA!H1:H1""))) <> 1"
    If CADbl(Range("UTCA!H1:H1"))(1) <> 2 Then Debug.Print "CADbl(Range(""UTCA!H1:H1""))(1) <> 2"
    If TypeName(CADbl(Range("UTCA!H1:H1"))) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!H1:H1""))) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!C8:H12"))) <> 5 Then Debug.Print "UBound(CADbl(Range(""UTCA!C8:H12""))) <> 5"
    If LBound(CADbl(Range("UTCA!C8:H12"))) <> 1 Then Debug.Print "LBound(CADbl(Range(""UTCA!C8:H12""))) <> 1"
    If CADbl(Range("UTCA!C8:H12"))(3) <> 1 / 12 Then Debug.Print "CADbl(Range(""UTCA!C8:H12""))(1) <> 1/12"
    If TypeName(CADbl(Range("UTCA!C8:H12"))) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!C8:H12""))) <> ""Double()"""
    
    Debug.Print "TestCADbl Afgerond"
End Sub

Public Sub TestCALng()
    Debug.Print "TestCALng Begonnen"
    
    If CALng(Array("1"))(1) <> 1 Then Debug.Print "CALng(Array(""1""))(1) <> 1"
    If TypeName(CALng(Array("1"))(1)) <> "Long" Then Debug.Print "TypeName(CALng(Array(""1""))(1)) <> ""Long"""
    If UBound(CALng(Array(1))) <> 1 Then Debug.Print "UBound(CALng(Array(1))) <> 1"
    If LBound(CALng(Array(1))) <> 1 Then Debug.Print "LBound(CALng(Array(1))) <> 1"
    If CALng("3")(1) <> 3 Then Debug.Print CALng("3")(1) <> 3
    If TypeName(CALng("1")(1)) <> "Long" Then Debug.Print "TypeName(CALng(""1"")(1)) <> ""Long"""
    If UBound(CALng(1)) <> 1 Then Debug.Print "UBound(CALng(1)) <> 1"
    If LBound(CALng(1)) <> 1 Then Debug.Print "LBound(CALng(1)) <> 1"
    If TypeName(CALng(Array())) <> "Long()" Then Debug.Print "TypeName(CALng(Array())) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!A1:A6"))) <> 6 Then Debug.Print "UBound(CALng(Range(""UTCA!A1:A6""))) <> 6"
    If LBound(CALng(Range("UTCA!A1:A6"))) <> 1 Then Debug.Print "LBound(CALng(Range(""UTCA!A1:A6""))) <> 1"
    If CALng(Range("UTCA!A1:A6"))(6) <> 2 Then Debug.Print "CALng(Range(""UTCA!A1:A6""))(6) <> 2"
    If TypeName(CALng(Range("UTCA!A1:A6"))) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!A1:A6""))) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!C1:H1"))) <> 6 Then Debug.Print "UBound(CALng(Range(""UTCA!C1:H1""))) <> 6"
    If LBound(CALng(Range("UTCA!C1:H1"))) <> 1 Then Debug.Print "LBound(CALng(Range(""UTCA!C1:H1""))) <> 1"
    If CALng(Range("UTCA!C1:H1"))(6) <> 2 Then Debug.Print "CALng(Range(""UTCA!C1:H1""))(6) <> 2"
    
    If UBound(CALng(Range("UTCA!H1:H1"))) <> 1 Then Debug.Print "UBound(CALng(Range(""UTCA!H1:H1""))) <> 1"
    If LBound(CALng(Range("UTCA!H1:H1"))) <> 1 Then Debug.Print "LBound(CALng(Range(""UTCA!H1:H1""))) <> 1"
    If CALng(Range("UTCA!H1:H1"))(1) <> 2 Then Debug.Print "CALng(Range(""UTCA!H1:H1""))(1) <> 2"
    If TypeName(CALng(Range("UTCA!H1:H1"))) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!H1:H1""))) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!C8:H12"))) <> 5 Then Debug.Print "UBound(CALng(Range(""UTCA!C8:H12""))) <> 5"
    If LBound(CALng(Range("UTCA!C8:H12"))) <> 1 Then Debug.Print "LBound(CALng(Range(""UTCA!C8:H12""))) <> 1"
    If CALng(Range("UTCA!C8:H12"))(3) <> 0 Then Debug.Print "CALng(Range(""UTCA!C8:H12""))(1) <> 0"
    If TypeName(CALng(Range("UTCA!C8:H12"))) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!C8:H12""))) <> ""Long()"""
    
    Debug.Print "TestCALng Afgerond"
End Sub

