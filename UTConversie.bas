Attribute VB_Name = "UTConversie"
'****h* UnitTest/FreekConversie
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
'   Deze module bevat de unittests voor de module FreekConversie
'****

Option Explicit

Public Sub TestConversie()
    TestCADbl
    TestCALng
End Sub

Public Sub TestCADbl()
    Debug.Print "TestCADbl Begonnen"
    'Met arrays die beginnen op index 1
    If CADbl(Array("1"), 1)(1) <> 1 Then Debug.Print "CADbl(Array(""1""), 1)(1) <> 1"
    If TypeName(CADbl(Array("1"), 1)(1)) <> "Double" Then Debug.Print "TypeName(CADbl(Array(""1""), 1)(1)) <> ""Double"""
    If UBound(CADbl(Array(1), 1)) <> 1 Then Debug.Print "UBound(CADbl(Array(1), 1)) <> 1"
    If LBound(CADbl(Array(1), 1)) <> 1 Then Debug.Print "LBound(CADbl(Array(1), 1)) <> 1"
    If CADbl("3", 1)(1) <> 3 Then Debug.Print CADbl("3", 1)(1) <> 3
    If TypeName(CADbl("1", 1)(1)) <> "Double" Then Debug.Print "TypeName(CADbl(""1"", 1)(1)) <> ""Double"""
    If UBound(CADbl(1, 1)) <> 1 Then Debug.Print "UBound(CADbl(1, 1)) <> 1"
    If LBound(CADbl(1, 1)) <> 1 Then Debug.Print "LBound(CADbl(1, 1)) <> 1"
    If TypeName(CADbl(Array(), 1)) <> "Double()" Then Debug.Print "TypeName(CADbl(Array(), 1)) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!A1:A6"), 1)) <> 6 Then Debug.Print "UBound(CADbl(Range(""UTCA!A1:A6""), 1)) <> 6"
    If LBound(CADbl(Range("UTCA!A1:A6"), 1)) <> 1 Then Debug.Print "LBound(CADbl(Range(""UTCA!A1:A6""), 1)) <> 1"
    If CADbl(Range("UTCA!A1:A6"), 1)(6) <> 2 Then Debug.Print "CADbl(Range(""UTCA!A1:A6""), 1)(6) <> 2"
    If TypeName(CADbl(Range("UTCA!A1:A6"), 1)) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!A1:A6""), 1)) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!C1:H1"), 1)) <> 6 Then Debug.Print "UBound(CADbl(Range(""UTCA!C1:H1""), 1)) <> 6"
    If LBound(CADbl(Range("UTCA!C1:H1"), 1)) <> 1 Then Debug.Print "LBound(CADbl(Range(""UTCA!C1:H1""), 1)) <> 1"
    If CADbl(Range("UTCA!C1:H1"), 1)(6) <> 2 Then Debug.Print "CADbl(Range(""UTCA!C1:H1""), 1)(6) <> 2"
    
    If UBound(CADbl(Range("UTCA!H1:H1"), 1)) <> 1 Then Debug.Print "UBound(CADbl(Range(""UTCA!H1:H1""), 1)) <> 1"
    If LBound(CADbl(Range("UTCA!H1:H1"), 1)) <> 1 Then Debug.Print "LBound(CADbl(Range(""UTCA!H1:H1""), 1)) <> 1"
    If CADbl(Range("UTCA!H1:H1"), 1)(1) <> 2 Then Debug.Print "CADbl(Range(""UTCA!H1:H1""), 1)(1) <> 2"
    If TypeName(CADbl(Range("UTCA!H1:H1"), 1)) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!H1:H1""), 1)) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!C8:H12"), 1)) <> 5 Then Debug.Print "UBound(CADbl(Range(""UTCA!C8:H12""), 1)) <> 5"
    If LBound(CADbl(Range("UTCA!C8:H12"), 1)) <> 1 Then Debug.Print "LBound(CADbl(Range(""UTCA!C8:H12""), 1)) <> 1"
    If CADbl(Range("UTCA!C8:H12"), 1)(3) <> 1 / 12 Then Debug.Print "CADbl(Range(""UTCA!C8:H12""), 1)(3) <> 1/12"
    If TypeName(CADbl(Range("UTCA!C8:H12"), 1)) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!C8:H12""), 1)) <> ""Double()"""

    If CADbl(Array(1, 2, Array(30, 40), 5), 1)(3) <> 30 Then Debug.Print "CADbl(Array(1, 2, Array(30, 40), 5), 1)(3) <> 30"
    ' Errorobject heeft Number als default property
    If CADbl(Err, 1)(1) <> 0 Then Debug.Print "CADbl(Err, 1)(1) <> 0"
    If CADbl(Array(Err), 1)(1) <> 0 Then Debug.Print "CADbl(Array(Err), 1)(1) <> 0"
    
    'Met arrays die beginnen op index 0
    If CADbl(Array("1"))(0) <> 1 Then Debug.Print "CADbl(Array(""1""))(0) <> 1"
    If TypeName(CADbl(Array("1"))(0)) <> "Double" Then Debug.Print "TypeName(CADbl(Array(""1""))(0)) <> ""Double"""
    If UBound(CADbl(Array(1))) <> 0 Then Debug.Print "UBound(CADbl(Array(1))) <> 0"
    If LBound(CADbl(Array(1))) <> 0 Then Debug.Print "LBound(CADbl(Array(1))) <> 0"
    If CADbl("3")(0) <> 3 Then Debug.Print CADbl("3")(0) <> 3
    If TypeName(CADbl("1")(0)) <> "Double" Then Debug.Print "TypeName(CADbl(""1"")(0)) <> ""Double"""
    If UBound(CADbl(1)) <> 0 Then Debug.Print "UBound(CADbl(1)) <> 0"
    If LBound(CADbl(1)) <> 0 Then Debug.Print "LBound(CADbl(1)) <> 0"
    If TypeName(CADbl(Array())) <> "Double()" Then Debug.Print "TypeName(CADbl(Array())) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!A1:A6"))) <> 5 Then Debug.Print "UBound(CADbl(Range(""UTCA!A1:A6""))) <> 5"
    If LBound(CADbl(Range("UTCA!A1:A6"))) <> 0 Then Debug.Print "LBound(CADbl(Range(""UTCA!A1:A6""))) <> 0"
    If CADbl(Range("UTCA!A1:A6"))(5) <> 2 Then Debug.Print "CADbl(Range(""UTCA!A1:A6""))(5) <> 2"
    If TypeName(CADbl(Range("UTCA!A1:A6"))) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!A1:A6""))) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!C1:H1"))) <> 5 Then Debug.Print "UBound(CADbl(Range(""UTCA!C1:H1""))) <> 5"
    If LBound(CADbl(Range("UTCA!C1:H1"))) <> 0 Then Debug.Print "LBound(CADbl(Range(""UTCA!C1:H1""))) <> 0"
    If CADbl(Range("UTCA!C1:H1"))(5) <> 2 Then Debug.Print "CADbl(Range(""UTCA!C1:H1""))(5) <> 2"
    
    If UBound(CADbl(Range("UTCA!H1:H1"))) <> 0 Then Debug.Print "UBound(CADbl(Range(""UTCA!H1:H1""))) <> 1"
    If LBound(CADbl(Range("UTCA!H1:H1"))) <> 0 Then Debug.Print "LBound(CADbl(Range(""UTCA!H1:H1""))) <> 1"
    If CADbl(Range("UTCA!H1:H1"))(0) <> 2 Then Debug.Print "CADbl(Range(""UTCA!H1:H1""))(0) <> 2"
    If TypeName(CADbl(Range("UTCA!H1:H1"))) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!H1:H1""))) <> ""Double()"""
    
    If UBound(CADbl(Range("UTCA!C8:H12"))) <> 4 Then Debug.Print "UBound(CADbl(Range(""UTCA!C8:H12""))) <> 4"
    If LBound(CADbl(Range("UTCA!C8:H12"))) <> 0 Then Debug.Print "LBound(CADbl(Range(""UTCA!C8:H12""))) <> 0"
    If CADbl(Range("UTCA!C8:H12"))(2) <> 1 / 12 Then Debug.Print "CADbl(Range(""UTCA!C8:H12""))(2) <> 1/12"
    If TypeName(CADbl(Range("UTCA!C8:H12"))) <> "Double()" Then Debug.Print "TypeName(CADbl(Range(""UTCA!C8:H12""))) <> ""Double()"""

    If CADbl(Array(1, 2, Array(30, 40), 5))(2) <> 30 Then Debug.Print "CADbl(Array(1, 2, Array(30, 40), 5))(2) <> 30"
    ' Errorobject heeft Number als default property
    If CADbl(Err)(0) <> 0 Then Debug.Print "CADbl(Err)(0) <> 0"
    If CADbl(Array(Err))(0) <> 0 Then Debug.Print "CADbl(Array(Err))(0) <> 0"
    
    Debug.Print "TestCADbl Afgerond"
End Sub

Public Sub TestCALng()
    Debug.Print "TestCALng Begonnen"
    'Met arrays die beginnen op index 1
    If CALng(Array("1"), 1)(1) <> 1 Then Debug.Print "CALng(Array(""1""), 1)(1) <> 1"
    If TypeName(CALng(Array("1"), 1)(1)) <> "Long" Then Debug.Print "TypeName(CALng(Array(""1""), 1)(1)) <> ""Long"""
    If UBound(CALng(Array(1), 1)) <> 1 Then Debug.Print "UBound(CALng(Array(1), 1)) <> 1"
    If LBound(CALng(Array(1), 1)) <> 1 Then Debug.Print "LBound(CALng(Array(1), 1)) <> 1"
    If CALng("3", 1)(1) <> 3 Then Debug.Print CALng("3", 1)(1) <> 3
    If TypeName(CALng("1", 1)(1)) <> "Long" Then Debug.Print "TypeName(CALng(""1"", 1)(1)) <> ""Long"""
    If UBound(CALng(1, 1)) <> 1 Then Debug.Print "UBound(CALng(1, 1)) <> 1"
    If LBound(CALng(1, 1)) <> 1 Then Debug.Print "LBound(CALng(1, 1)) <> 1"
    If TypeName(CALng(Array(), 1)) <> "Long()" Then Debug.Print "TypeName(CALng(Array(), 1)) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!A1:A6"), 1)) <> 6 Then Debug.Print "UBound(CALng(Range(""UTCA!A1:A6""), 1)) <> 6"
    If LBound(CALng(Range("UTCA!A1:A6"), 1)) <> 1 Then Debug.Print "LBound(CALng(Range(""UTCA!A1:A6""), 1)) <> 1"
    If CALng(Range("UTCA!A1:A6"), 1)(6) <> 2 Then Debug.Print "CALng(Range(""UTCA!A1:A6""), 1)(6) <> 2"
    If TypeName(CALng(Range("UTCA!A1:A6"), 1)) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!A1:A6""), 1)) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!C1:H1"), 1)) <> 6 Then Debug.Print "UBound(CALng(Range(""UTCA!C1:H1""), 1)) <> 6"
    If LBound(CALng(Range("UTCA!C1:H1"), 1)) <> 1 Then Debug.Print "LBound(CALng(Range(""UTCA!C1:H1""), 1)) <> 1"
    If CALng(Range("UTCA!C1:H1"), 1)(6) <> 2 Then Debug.Print "CALng(Range(""UTCA!C1:H1""), 1)(6) <> 2"
    
    If UBound(CALng(Range("UTCA!H1:H1"), 1)) <> 1 Then Debug.Print "UBound(CALng(Range(""UTCA!H1:H1""), 1)) <> 1"
    If LBound(CALng(Range("UTCA!H1:H1"), 1)) <> 1 Then Debug.Print "LBound(CALng(Range(""UTCA!H1:H1""), 1)) <> 1"
    If CALng(Range("UTCA!H1:H1"), 1)(1) <> 2 Then Debug.Print "CALng(Range(""UTCA!H1:H1""), 1)(1) <> 2"
    If TypeName(CALng(Range("UTCA!H1:H1"), 1)) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!H1:H1""), 1)) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!C8:H12"), 1)) <> 5 Then Debug.Print "UBound(CALng(Range(""UTCA!C8:H12""), 1)) <> 5"
    If LBound(CALng(Range("UTCA!C8:H12"), 1)) <> 1 Then Debug.Print "LBound(CALng(Range(""UTCA!C8:H12""), 1)) <> 1"
    If CALng(Range("UTCA!C8:H12"), 1)(3) <> 0 Then Debug.Print "CALng(Range(""UTCA!C8:H12""), 1)(3) <> 0"
    If TypeName(CALng(Range("UTCA!C8:H12"), 1)) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!C8:H12""), 1)) <> ""Long()"""

    If CALng(Array(1, 2, Array(30, 40), 5), 1)(3) <> 30 Then Debug.Print "CALng(Array(1, 2, Array(30, 40), 5), 1)(3) <> 30"
    ' Errorobject heeft Number als default property
    If CALng(Err, 1)(1) <> 0 Then Debug.Print "CALng(Err, 1)(1) <> 0"
    If CALng(Array(Err), 1)(1) <> 0 Then Debug.Print "CALng(Array(Err), 1)(1) <> 0"
    
    'Met arrays die beginnen op index 0
    If CALng(Array("1"))(0) <> 1 Then Debug.Print "CALng(Array(""1""))(0) <> 1"
    If TypeName(CALng(Array("1"))(0)) <> "Long" Then Debug.Print "TypeName(CALng(Array(""1""))(0)) <> ""Long"""
    If UBound(CALng(Array(1))) <> 0 Then Debug.Print "UBound(CALng(Array(1))) <> 0"
    If LBound(CALng(Array(1))) <> 0 Then Debug.Print "LBound(CALng(Array(1))) <> 0"
    If CALng("3")(0) <> 3 Then Debug.Print CALng("3")(0) <> 3
    If TypeName(CALng("1")(0)) <> "Long" Then Debug.Print "TypeName(CALng(""1"")(0)) <> ""Long"""
    If UBound(CALng(1)) <> 0 Then Debug.Print "UBound(CALng(1)) <> 0"
    If LBound(CALng(1)) <> 0 Then Debug.Print "LBound(CALng(1)) <> 0"
    If TypeName(CALng(Array())) <> "Long()" Then Debug.Print "TypeName(CALng(Array())) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!A1:A6"))) <> 5 Then Debug.Print "UBound(CALng(Range(""UTCA!A1:A6""))) <> 5"
    If LBound(CALng(Range("UTCA!A1:A6"))) <> 0 Then Debug.Print "LBound(CALng(Range(""UTCA!A1:A6""))) <> 0"
    If CALng(Range("UTCA!A1:A6"))(5) <> 2 Then Debug.Print "CALng(Range(""UTCA!A1:A6""))(5) <> 2"
    If TypeName(CALng(Range("UTCA!A1:A6"))) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!A1:A6""))) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!C1:H1"))) <> 5 Then Debug.Print "UBound(CALng(Range(""UTCA!C1:H1""))) <> 5"
    If LBound(CALng(Range("UTCA!C1:H1"))) <> 0 Then Debug.Print "LBound(CALng(Range(""UTCA!C1:H1""))) <> 0"
    If CALng(Range("UTCA!C1:H1"))(5) <> 2 Then Debug.Print "CALng(Range(""UTCA!C1:H1""))(5) <> 2"
    
    If UBound(CALng(Range("UTCA!H1:H1"))) <> 0 Then Debug.Print "UBound(CALng(Range(""UTCA!H1:H1""))) <> 1"
    If LBound(CALng(Range("UTCA!H1:H1"))) <> 0 Then Debug.Print "LBound(CALng(Range(""UTCA!H1:H1""))) <> 1"
    If CALng(Range("UTCA!H1:H1"))(0) <> 2 Then Debug.Print "CALng(Range(""UTCA!H1:H1""))(0) <> 2"
    If TypeName(CALng(Range("UTCA!H1:H1"))) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!H1:H1""))) <> ""Long()"""
    
    If UBound(CALng(Range("UTCA!C8:H12"))) <> 4 Then Debug.Print "UBound(CALng(Range(""UTCA!C8:H12""))) <> 4"
    If LBound(CALng(Range("UTCA!C8:H12"))) <> 0 Then Debug.Print "LBound(CALng(Range(""UTCA!C8:H12""))) <> 0"
    If CALng(Range("UTCA!C8:H12"))(2) <> 0 Then Debug.Print "CALng(Range(""UTCA!C8:H12""))(2) <> 0"
    If TypeName(CALng(Range("UTCA!C8:H12"))) <> "Long()" Then Debug.Print "TypeName(CALng(Range(""UTCA!C8:H12""))) <> ""Long()"""

    If CALng(Array(1, 2, Array(30, 40), 5))(2) <> 30 Then Debug.Print "CALng(Array(1, 2, Array(30, 40), 5))(2) <> 30"
    ' Errorobject heeft Number als default property
    If CALng(Err)(0) <> 0 Then Debug.Print "CALng(Err)(0) <> 0"
    If CALng(Array(Err))(0) <> 0 Then Debug.Print "CALng(Array(Err))(0) <> 0"
    
    Debug.Print "TestCALng Afgerond"
End Sub
