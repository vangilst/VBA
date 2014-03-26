'****h* Tech/FreekFileTools
' NAME
'   FreekFileTools
' COPYRIGHT
'   (c) Freek van Gilst
'   Zie voor gebruiksvoorwaarden http://www.vangilst.com/Voorwaarden.html
'   Onder andere houdt dat in dat er geen garantie wordt gegeven en het niet
'   zonder meer is toegestaan om dit te kopi�ren.
'   Voor sommige organisaties zijn uitzonderingen mogelijk, maar ALLEEN WANNEER
'   DEZE TEKST OVER COPYRIGHT WORDT MEEGEKOPIEERD wanneer deze module of delen
'   ervan wordt gebruikt!
'   Lees voor gebruik hiervan eerst de genoemde webpagina.
' AUTHOR
'   Freek van Gilst, e-mail mailto:freek@vangilst.com
' FUNCTION
'   Functies en procedures voor het makkelijk werken met files.
'****

Option Explicit

'****f* FreekFileTools/GarandeerDirectory
' FUNCTION
'   Garandeer het bestaan van een directory door deze desnoods (desnoods recursief)
'   aan te maken.
' SYNOPSIS
Public Sub GarandeerDirectory(ByVal D As String)
' INPUTS
'   D: De (volledige) naam van de directory.
'*****
Dim A
Dim i As Long: i = 1
Dim NwDir As String
    
    If Right(D, 1) = "\" Then D = Left(D, Len(D) - 1)
    A = Split(D, "\")
    NwDir = A(0)
    If Left(D, 2) = "\\" Then i = i + 3: NwDir = "\\" & NwDir & A(1) & A(2) & "\" & A(3)
    Do While i <= UBound(A)
        NwDir = NwDir & "\" & A(i)
        If Not FileBestaat(NwDir) Then
            MkDir NwDir
        End If
        i = i + 1
    Loop
    
End Sub

'****f* FreekFileTools/FileBestaat
' FUNCTION
'   Controleer of een bestand of directory bestaat
' SYNOPSIS
Public Function FileBestaat(ByRef F As String) As Boolean
' INPUTS
'   F: Volledige bestandsnaam van het bestand
' RESULT
'   Boolean True als het bestand bestaat, anders False
'*****
Dim Dummy

    FileBestaat = False
    On Error GoTo NotFound
    Dummy = FileDateTime(F)
    FileBestaat = True
    Exit Function
    
NotFound:
    FileBestaat = False
    
End Function


'****f* FreekFileTools/Subdirs
' FUNCTION
'   Maakt een array met alle subdirectories van de opgegeven directory
'   (niet het volledige pad maar alleen de naam van de subdirectory)
' SYNOPSIS
Public Function Subdirs(ByVal D As String)
' INPUTS
'   D: De (volledige) naam van de directory waaruit de subdirectories worden gevraagd.
' RESULT
'   Array met de namen van de subdirectories.
'   Als er geen subdirectories zijn dan is het een Array() As String
'*****
Dim Dirnaam
Dim FileSystem As FileSystemObject: Set FileSystem = New FileSystemObject
Dim Directories As Folders: Set Directories = FileSystem.GetFolder(D).Subfolders
Dim Resultaat() As String
Dim N As Long
    
    If Directories.Count = 0 Then
        Subdirs = Array()
    Else
        ReDim Resultaat(0 To Directories.Count - 1)
        For Each Dirnaam In Directories
            Resultaat(N) = Mid(Dirnaam, InStrRev(Dirnaam, "\") + 1)
            N = N + 1
        Next Dirnaam
        Subdirs = Resultaat
    End If

End Function

'****f* FreekFileTools/ZoekFileRecusief
' FUNCTION
'   Zoekt een bestand (of pattern) in directories en subdirectories en geeft het volledige pad (inclusief bestandsnaam)
'   Toegevoegd een check of de polissen staan in subdirectories met de eerste 8 cijfers van
'   het polisnummer als directorynaam. In dat geval stoppen we als de polis niet in die directory zit (want anders door-
'   zoeken we heel veel andere directories waarvan we al van te voren weten dat de polis daar niet inzit).
' SYNOPSIS
Public Function ZoekFileRecusief(Pattern As String, Directory As String, Optional Ondiep As Boolean = False) As String
' INPUTS
'   Pattern:   polisnummer gevolgd door bestandsextensie
'   Directory: de te doorzoeken directory
'   Ondiep:    true als we maar 1 niveau diep zoeken
' RESULT
'   Bestandsnaam, of een lege string als er geen bestanden zijn gevonden.
'*****
Dim Gevonden As String
    
    If Right$(Directory, 1) <> "\" Then Directory = Directory & "\"
    Gevonden = Dir$(Directory & Pattern)
    If Len(Gevonden) <> 0 Then
        ZoekFileRecusief = Directory & Gevonden
    Else
        Gevonden = Dir$(Directory & Left(Pattern, 8) & "\" & Pattern)
        If Len(Gevonden) <> 0 Then
            ZoekFileRecusief = Directory & Left(Pattern, 8) & "\" & Gevonden
        Else
            If Ondiep Then
                ZoekFileRecusief = ""
            Else
                Dim Subdirectory
                For Each Subdirectory In Subdirs(Directory)
                    DoEvents
                    Gevonden = ZoekFileRecusief(Pattern, Directory & Subdirectory)
                    If Len(Gevonden) > 0 Then ZoekFileRecusief = Gevonden: Exit Function
                Next Subdirectory
            End If
        End If
    End If
    
End Function

'****f* FreekFileTools/FileLijstInDir
' FUNCTION
'   Maakt een array met alle bestanden in de opgegeven directory
'   Bijeffecten: gebruikt de functie Dir$
' SYNOPSIS
Public Function FileLijstInDir(D As String, Pattern As String)
' INPUTS
'   D:       de te doorzoeken directory
'   Pattern: moet de bestandsnaam aan voldoen
' RESULT
'   Array met de namen van de bestanden.
'   Als er geen bestanden in D zijn die voldoen aan Pattern
'   dan is het resultaat een Array() As String
'*****
Dim Aantal As Long
Dim Filename As String
Dim Filelist() As String

    Filename = Dir$(D & Pattern)
    Aantal = 0
    Do While Len(Filename) <> 0
        Aantal = Aantal + 1
        ReDim Preserve Filelist(0 To Aantal - 1)
        Filelist(Aantal - 1) = Filename
        Filename = Dir$
    Loop
    FileLijstInDir = Filelist
    
End Function

'****f* FreekFileTools/TempDir
' FUNCTION
'   Bepaalt de lokatie van de TEMP directory.
'   Gebruik bij voorkeur de Environment variabelen.
'   Anders C:\Temp
' SYNOPSIS
Public Function TempDir() As String
' RESULT
'   String: Naam van de TEMP directory
'*****

    TempDir = Environ$("TEMP")
    
    If Len(TempDir) = 0 Then
        TempDir = Environ$("USERPROFILE")
        If Len(TempDir) = 0 Then
            'Zelf lokatie bepalen: C:\Temp
            MkDir "C:\Temp\"
            TempDir = "C:\Temp\"
        Else
            TempDir = TempDir & "\Local Settings\Temp\"
        End If
    Else 'TempDir = Environ$("TEMP")
        ' Afsluiten met backslash
        If Right$(TempDir, 1) <> "\" Then TempDir = TempDir & "\"
    End If

End Function

'****f* FreekFileTools/RandomTempDir
' FUNCTION
'   Maak een subdirectory met een unieke random naam in de Temp-directory.
' SYNOPSIS
Public Function RandomTempDir() As String
' RESULT
'   String: De random naam van de directory. Als side effect wordt de directory gemaakt!
'*****

    Randomize
    RandomTempDir = TempDir() & CStr(Int(1 + 9999 * Rnd())) & "_" & CStr(Int(1 + 9999 * Rnd())) & "_" & CStr(Int(1 + 9999 * Rnd())) & "\"
    GarandeerDirectory RandomTempDir
    
End Function

'****f* FreekFileTools/Backslash
' FUNCTION
'   Zorgt dat de string eindigt op een backslash.
' SYNOPSIS
Public Function Backslash(ByVal S As String) As String
' INPUTS
'   S: De string die, eventueel met toegevoegde backslash ook weer wordt teruggegeven.
' RESULT
'   String: De string met gegarandeerd een backslash aan het einde.
'*****

    If Right$(S, 1) = "\" Then
        Backslash = S
    Else
        Backslash = S & "\"
    End If
    
End Function

'****f* FreekFileTools/Quote
' FUNCTION
'   Zorgt dat de string begint en eindigt met ".
' SYNOPSIS
Public Function Quote(ByVal S As String) As String
' INPUTS
'   S: De string die, eventueel met toegevoegde "" ook weer wordt teruggegeven.
' RESULT
'   String: De string met gegarandeerd een " aan het begin en einde.
'*****

    If Left$(S, 1) <> """" Then
        If Right$(S, 1) <> """" Then
            Quote = """" & S & """"
        Else
            Quote = """" & S
        End If
    Else
        If Right$(S, 1) <> """" Then
            Quote = S & """"
        Else
            Quote = S
        End If
    End If
    
End Function

'****f* FreekFileTools/FileNmZonderFouteTekens
' FUNCTION
'   Zorgt dat de string geen tekens bevat die niet in een filenaam mogen
'   voorkomen. Met filenaam wordt hier bedoeld de bestandsnaam zonder pad.
'   Ook tekens als de backslash en de dubbele punt worden vertaald naar andere
'   tekens.
' SYNOPSIS
Public Function FileNmZonderFouteTekens(ByVal S As String) As String
' INPUTS
'   S: De string met eventueel foute tekens.
' RESULT
'   String: De string met 'vertaalde' tekens.
'*****

    FileNmZonderFouteTekens = Replace(S, "<", "{")
    FileNmZonderFouteTekens = Replace(FileNmZonderFouteTekens, ">", "}")
    FileNmZonderFouteTekens = Replace(FileNmZonderFouteTekens, "\", "`")
    FileNmZonderFouteTekens = Replace(FileNmZonderFouteTekens, "/", "´")
    FileNmZonderFouteTekens = Replace(FileNmZonderFouteTekens, "*", "·")
    FileNmZonderFouteTekens = Replace(FileNmZonderFouteTekens, "?", "!")
    FileNmZonderFouteTekens = Replace(FileNmZonderFouteTekens, ":", "÷")
    FileNmZonderFouteTekens = Replace(FileNmZonderFouteTekens, """", "¨")
    FileNmZonderFouteTekens = Replace(FileNmZonderFouteTekens, "|", "¦")

End Function

'****f* FreekFileTools/FileNmZonderPad
' FUNCTION
'   Bepaalt het deel van de filenaam zonder pad maar met extensie.
' SYNOPSIS
Public Function FileNmZonderPad(ByVal S As String) As String
' INPUTS
'   S: De string met de filenaam.
' RESULT
'   String: De string met filenaam zonder pad.
'*****

    Dim PosLaatsteBackslash As Long
    PosLaatsteBackslash = InStrRev(S, "\")
    If PosLaatsteBackslash < 1 Then
        PosLaatsteBackslash = InStrRev(S, ":")
        If PosLaatsteBackslash < 1 Then PosLaatsteBackslash = 0
    End If
    FileNmZonderPad = Mid$(S, PosLaatsteBackslash + 1)

End Function

'****f* FreekFileTools/PadNm
' FUNCTION
'   Bepaalt het deel van de filenaam waarin het pad staat maar niet de
'   specifieke file.
'   De padnaam eindigt met een backslash, tenzij het een driveaanduiding is,
'   (d.w.z. dat "C:data.txt" door deze functie wordt vertaald naar "C:") of
'   tenzij er geen pad in de string staat (bijv. "data.txt" wordt vertaald naar
'   de lege string "").
' SYNOPSIS
Public Function PadNm(ByVal S As String) As String
' INPUTS
'   S: De string met de filenaam.
' RESULT
'   String: De string met padnaam.
'*****

    Dim PosLaatsteBackslash As Long
    PosLaatsteBackslash = InStrRev(S, "\")
    If PosLaatsteBackslash < 1 Then
        PosLaatsteBackslash = InStrRev(S, ":")
        If PosLaatsteBackslash < 1 Then PosLaatsteBackslash = 0
    End If
    PadNm = Left$(S, PosLaatsteBackslash)

End Function
