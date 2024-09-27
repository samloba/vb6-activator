Attribute VB_Name = "Dialogs2"
'Option Explicit
Private Type BROWSEINFO    ' used by the function GetFolderName
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Sub PathStripPath Lib "shlwapi.dll" Alias "PathStripPathA" (ByVal pszPath As String)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
                                         "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

'Structure du fichier
Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Constantes
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000

Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0


'Boite de dialogue couleur
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias _
                                     "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Repertoires Spéciaux
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" _
                                                (ByVal hWndOwner As Long, ByVal lpszPath As String, _
                                                 ByVal nFolder As Long, ByVal fCreate As Long) As Long
Public Enum SpecialFolderEnum
    CurrentUserDeskTop = 0    'Or 16
    CurrentUserDocuments = 5
    SendTos = 9
    CurrentUserMusic = 13
    CurrentUserVideos = 14
    Fonts = 20
    ProgramsGroup = 23
    StartUp = 24
    WinDir = 36
    SysDir = 41
    PublicDocument = 46
End Enum

 
 'Constantes permettant de personnaliser le fonctionnement de BrowseForFolder
Const BIF_RETURNONLYFSDIRS = &H1 'pour chercher les fichiers systèmes seulement
                                    ' si le dossier sélectionné ne contient pas
                                    ' de fichier système alors le bouton "OK" est grisé
Const BIF_DONTGOBELOWDOMAIN = &H2 'interdit d'explorer en dehors du domaine 'For starting the Find Computer
Const BIF_STATUSTEXT = &H4 '&H4&
Const BIF_RETURNFSANCESTORS = &H8 'seulement des dossiers
Const BIF_EDITBOX = &H10 'Affiche une zone d'édition
Const BIF_VALIDATE = &H20 'Vérifie la saisie dans la zone d'édition
Const BIF_BROWSEFORCOMPUTER = &H1000 'Autorise le parcours réseau
Const BIF_BROWSEFORPRINTER = &H2000 'mes documents et bureau uniquemnet
Const BIF_BROWSEINCLUDEFILES = &H4000 'dossiers et fichiers
Const BIF_NONEWFOLDERBUTTON = &H200 'ne pas mettre le bouton Nouveau dossier
  ' Déclaration de l'API
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
        Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) _
        As Long


Function EnregistrerUnFichier(Handle As Long, Titre As String, _
                    NomFichier As String, chemin As String) As String

 ' EnregistrerUnFichier est la fonction à utiliser dans votre formulaire pour ouvrir _
la boîte de dialogue d'enregistrement d'un fichier.
 ' Explication des paramètres
    ' Handle = le handle de la fenêtre (Me.Hwnd)
    ' Titre = titre de la boîte de dialogue
    ' NomFichier = nom par défaut du fichier à enregistrer
    ' Chemin = chemin par défaut du fichier à enregistrer
       
Dim structSave As OPENFILENAME

With structSave
    .lStructSize = Len(structSave)
    .hWndOwner = Handle
    .nMaxFile = 255
    .lpstrFile = NomFichier & String$(255 - Len(NomFichier), 0)
    .lpstrInitialDir = chemin
    .lpstrFilter = "Tous (*.*)" & Chr$(0) & "*.*" & Chr$(0) ' Définition du filtre (aucun)
    .Flags = &H4  'Option de la boite de dialogue
End With

If (GetSaveFileName(structSave)) Then
    EnregistrerUnFichier = Mid$(structSave.lpstrFile, 1, InStr(1, structSave.lpstrFile, vbNullChar) - 1)
End If

End Function

Public Function GetSpecialFolder(StrDossier As SpecialFolderEnum) As String
    On Error Resume Next
    Dim Buffer As String

    GetSpecialFolder = ""
    Buffer = Space(256)
    SHGetSpecialFolderPath frm_activator.hWnd, Buffer, StrDossier, 0
    GetSpecialFolder = Left(Buffer, InStr(Buffer, Chr(0)) - 1) & "\"
End Function


Function GetFolderName(Msg As String) As String
    ' returns the name of the folder selected by the user
    Dim bInfo As BROWSEINFO, path As String, R As Long, X As Long, Pos As Integer
    bInfo.pidlRoot = 0&    ' Root folder = Desktop
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Selectionner un répertoire de travail"    ' the dialog title
    Else
        bInfo.lpszTitle = Msg    ' the dialog title
    End If
    bInfo.ulFlags = 0  ' Type of directory to return
    X = SHBrowseForFolder(bInfo)    ' display the dialog
    ' Parse the result
    path = Space$(512)
    R = SHGetPathFromIDList(ByVal X, ByVal path)
    If R Then
        Pos = InStr(path, Chr$(0))
        GetFolderName = Left(path, Pos - 1)
    Else
        GetFolderName = ""
    End If
End Function
Sub test123698()
'Debug.Print GetFolderName("TEST")
End Sub
Public Function OuvrirUnFichier(Handle As Long, _
                                Titre As String, _
                                TypeRetour As Byte, _
                                Optional TitreFiltre As String, _
                                Optional TypeFichier As String, _
                                Optional RepParDefaut As String) As String
    Dim StructFile As OPENFILENAME
    Dim sFiltre As String

    'Construction du filtre en fonction des arguments spécifiés
    If Len(TitreFiltre) > 0 And Len(TypeFichier) > 0 Then
        sFiltre = TitreFiltre & " (" & TypeFichier & ")" & Chr$(0) & "*." & TypeFichier & Chr$(0)
    End If
    sFiltre = sFiltre & "Tous (*.*)" & Chr$(0) & "*.*" & Chr$(0)


    'Configuration de la boîte de dialogue
    With StructFile
        .lStructSize = Len(StructFile)    'Initialisation de la grosseur de la structure
        .hWndOwner = Handle    'Identification du handle de la fenêtre
        .lpstrFilter = sFiltre    'Application du filtre
        .lpstrFile = String$(254, vbNullChar)    'Initialisation du fichier '0' x 254
        .nMaxFile = 254    'Taille maximale du fichier
        .lpstrFileTitle = String$(254, vbNullChar)    'Initialisation du nom du fichier '0' x 254
        .nMaxFileTitle = 254  'Taille maximale du nom du fichier
        .lpstrTitle = Titre    'Titre de la boîte de dialogue
        .Flags = OFN_HIDEREADONLY  'Option de la boite de dialogue
        If ((IsNull(RepParDefaut)) Or (RepParDefaut = "")) Then
            RepParDefaut = CurDir
            PathStripPath (RepParDefaut)
            .lpstrInitialDir = RepParDefaut
        Else: .lpstrInitialDir = RepParDefaut
        End If
    End With

    If (GetOpenFileName(StructFile)) Then    'Si un fichier est sélectionné
        Select Case TypeRetour
            Case 1: OuvrirUnFichier = Trim$(Left(StructFile.lpstrFile, InStr(1, StructFile.lpstrFile, vbNullChar) - 1))
            Case 2: OuvrirUnFichier = Trim$(Left(StructFile.lpstrFileTitle, InStr(1, StructFile.lpstrFileTitle, vbNullChar) - 1))
        End Select
    End If

End Function
Public Function ShowColor(Handle As Long, Optional DefaultColor = 16777215) As Long
    Dim CC As ChooseColor
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    'set the structure size
    CC.lStructSize = Len(CC)
    'Set the owner
    CC.hWndOwner = Handle
    'set the custom colors (converted to Unicode)
    CC.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    CC.Flags = 0

    'Show the 'Select Color'-dialog
    If ChooseColor(CC) <> 0 Then
        ShowColor = CC.rgbResult
        CustomColors = StrConv(CC.lpCustColors, vbFromUnicode)
    Else
        ShowColor = DefaultColor ' -1  'Nz(DefaultColor, 16777215) '16777215 blanc
    End If
End Function



