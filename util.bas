Attribute VB_Name = "util"
Public Const LOBA_HDD_ID = 1745409527       ' -2097825750
Public Const ROBY_HDD_ID = -2105601844

Private Const SPD_CRYPT_KEY = "le.travail.de.l'insensé.le.fatigue.sagesse.de.salomon"

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
                                              (ByVal lpRootPathName As String, _
                                               ByVal lpVolumeNameBuffer As String, _
                                               ByVal nVolumeNameSize As Long, _
                                               lpVolumeSerialNumber As Long, _
                                               lpMaximumComponentLength As Long, _
                                               lpFileSystemFlags As Long, _
                                               ByVal lpFileSystemNameBuffer As String, _
                                               ByVal nFileSystemNameSize As Long) As Long

Public fct_name As String
                                              
Private Const MAX_PATH As Integer = 255
Private Declare Function apiGetSystemDirectory& Lib "kernel32" _
        Alias "GetSystemDirectoryA" _
        (ByVal lpBuffer As String, ByVal nSize As Long)
        '
        
Public Property Get SPD_PWD_ARRAY()
    SPD_PWD_ARRAY = Array("Pwd2021!", "$ages2020", "B.E.2007", "dorsale")
End Property

Public Property Get SPD_OLD_DORSALE_PWD()
    SPD_OLD_DORSALE_PWD = SPD_PWD_ARRAY(1)
End Property

Public Property Get SPD_OLD_DORSALE_CONNECT()
    SPD_OLD_DORSALE_CONNECT = "MS Access;PWD=" & SPD_PWD_ARRAY(1)
End Property

Public Property Get SPD_DORSALE_PWD()
    SPD_DORSALE_PWD = SPD_PWD_ARRAY(0)
End Property

Public Property Get SPD_DORSALE_CONNECT()
    SPD_DORSALE_CONNECT = "MS Access;PWD=" & SPD_PWD_ARRAY(0)
End Property


Function fReturnSysDir() As String
    'Returns System Folder Name (C:\WinNT\System32)
    Dim strSysDirName As String
    Dim lngx As Long
    strSysDirName = String$(MAX_PATH, 0)
    lngx = apiGetSystemDirectory(strSysDirName, MAX_PATH)
    If lngx <> 0 Then
        fReturnSysDir = Left$(strSysDirName, lngx)
    Else
        fReturnSysDir = ""
    End If
End Function
        
        
        
Public Function GetDiskSerial() As String
    Dim Serial As Long
    Dim VName As String
    Dim FSName As String
    VName = String(255, Chr(0))
    FSName = String(255, Chr(0))
    GetVolumeInformation "C:\", VName, 255, Serial, 0, 0, FSName, 255
    GetDiskSerial = CStr(Serial)
    Debug.Print GetDiskSerial
End Function
Function get_activation_code(anScol As String, code_etab As String)
    Dim a, B, c
    a = Val(code_etab)
    B = Mid(anScol, 1, 4)
    c = Mid(anScol, 6, 4)
    get_activation_code = Fix((a * 21) * (B - 35) * (c / 3))
End Function
Function code_cert_etab(anScol As String, codeEtab As String)
    Dim a, B, c
    a = Val(codeEtab)
    B = Mid(anScol, 1, 4)
    c = Mid(anScol, 6, 4)
    code_cert_etab = (a * 28) * (B + 29) * (c - 3)
End Function
 Function get_reactivation_code(code_etab As String, anScol As String, my_case) As String
    Dim a, B, c, X
    a = Val(code_etab)
    B = Mid(anScol, 1, 4)
    c = Mid(anScol, 6, 4)
    X = Fix((a * 21) * (B - 35) * (c / 3))

    Select Case CLng(my_case)
        Case 0: X = X / 62
        Case 1: X = X / 158
        Case 2: X = X / 243
        Case 3: X = X / 325
    End Select
    get_reactivation_code = CStr(X * 5 - 1)
End Function
Function get_sms_activation_code(anScol As String, code_etab As String, imei As String)
    Dim a, B, c, D
    a = Val(code_etab)
    B = CLng(Mid(anScol, 1, 4))
    c = CLng(Mid(anScol, 6, 4))
    If imei = "" Then D = 0 Else D = Fix(Sqr(Val(imei)))
    get_sms_activation_code = Fix((a * 241) * (B - 73) * (c / 5)) + D
End Function
Function Crypt(TextNonCrypté As String, Optional forSQL As Integer = 0) As String
    If IsNull(TextNonCrypté) Or TextNonCrypté = "" Then
        Crypt = ""
        Exit Function
    End If

    Dim a, B, TextCrypté, Incr, TeXtec, okm, z, w
    z = 0
    For Incr = 1 To Len(TextNonCrypté)
        'C'a, c pour effectuer le code suivant le nombre de
        'caractère ke contient le texte
        TeXtec = Mid$(TextNonCrypté, Incr, 1)
        'Là, c pour lire le caractère du texte
        a = Asc(TeXtec)
        'Je transforme le caractère alphabétique en
        'caractère ASCII
        z = z + 1
        If z > Len(SPD_CRYPT_KEY) Then
            z = 1
            'Kan le mot de passe a été lu, on recommence à
            'son premier caractère
        End If

        okm = Mid$(SPD_CRYPT_KEY, z, 1)
        'On lit le caractère du mot de passe
        B = Asc(okm)
        'On prend son code ASCII
        w = a + B
        'On additionne le code ASCII du texte et du mot de passe
        If w > 255 Then
            w = w - 255
            'Au cas où l'addition dépasserai 255, on retire 255
        End If
        TextCrypté = TextCrypté + Chr(w)
        'On additionne les lettres cryptés ki sont
        'retransformées en caractère "lisible" mais crypté
    Next Incr
    'on fait de même pour la lettre suivante du texte à crypter

    If forSQL Then TextCrypté = Replace(TextCrypté, Chr(34), """""")

    Crypt = TextCrypté
    'Et on affiche le texte crypté ki ne ve plus rien dire
End Function
Function DeCrypt(TextCrypté As String) As String

    If Len(TextCrypté) = 0 Then
        DeCrypt = ""
        Exit Function
    End If


    Dim a, B, TextDecrypté, Incr, TeXtec, okm, z, w
    z = 0
    For Incr = 1 To Len(TextCrypté)
        TeXtec = Mid$(TextCrypté, Incr, 1)
        a = Asc(TeXtec)
        z = z + 1
        If z > Len(SPD_CRYPT_KEY) Then
            z = 1
        End If
        okm = Mid$(SPD_CRYPT_KEY, z, 1)
        B = Asc(okm)
        w = a - B
        If w < 0 Then
            w = w + 255
        End If
        TextDecrypté = TextDecrypté + Chr(w)
    Next Incr
    DeCrypt = TextDecrypté
End Function
Function ParamEtabBase(strInfo As String, myDB As DAO.Database) As String
    'On Error Resume Next
    Dim rst_Param As DAO.Recordset

    ParamEtabBase = ""
    Sql_Param = "SELECT * FROM Params_etab WHERE param_name=""" & strInfo & """"
    Set rst_Param = myDB.OpenRecordset(Sql_Param)
    If Not rst_Param.EOF Then ParamEtabBase = DeCrypt(Nz(rst_Param("param_value"), "")) Else ParamEtabBase = ""

CleanUp:
    If Not rst_Param Is Nothing Then
        rst_Param.Close: Set rst_Param = Nothing
    End If

End Function
Sub SaveParamEtabBase(strParam As String, myValue As String, myDB As DAO.Database)
    'On Error Resume Next
    Dim strConnect As String
    MyValueX = Crypt(myValue)
    MyValueX = Replace(MyValueX, Chr(34), """""")
    If ParamEtabBaseExist(strParam, myDB) Then
        SQL = "UPDATE Params_Etab SET param_value=""" & MyValueX & """ WHERE param_name=""" & strParam & """"
        myDB.Execute SQL
    Else
        SQL = "INSERT INTO Params_Etab (param_name, param_value) VALUES (""" & strParam & """,""" & MyValueX & """)"
        myDB.Execute SQL
    End If


    'En clair pour x-server
    Dim bExistIn_Param_Etab_js As Boolean
    Dim rst As DAO.Recordset

    Set rst = myDB.OpenRecordset("SELECT * FROM params_etab_js WHERE param_name=" & WithQuotes(strParam))
    If rst.EOF Then rst.AddNew Else rst.Edit
    rst!param_name = strParam
    rst!param_value = myValue
    rst.Update


    '            bExistIn_Param_Etab_js
    '        SQL = "UPDATE Params_Etab_js SET param_value=""" & myValue & """ WHERE param_name=""" & strParam & """"
    '        myDB.Execute SQL
    '
    '            'En clair pour x-server
    '        SQL = "INSERT INTO Params_Etab (param_name, param_value) VALUES (""" & strParam & """,""" & myValue & """)"
    '        myDB.Execute SQL


End Sub
Function ParamEtabBaseExist(strParam As String, myDB As DAO.Database) As Boolean
    Dim rst As DAO.Recordset
    Set rst = myDB.OpenRecordset("SELECT * FROM Params_etab WHERE param_name=""" & strParam & """")
    If Not rst.EOF Then ParamEtabBaseExist = -1 Else ParamEtabBaseExist = 0
    rst.Close: Set rst = Nothing
End Function
Function EstVoyelle(Caractère As String) As Boolean
    Select Case Caractère
        Case "a", "e", "y", "i", "o", "u"
            EstVoyelle = True
        Case Else
            EstVoyelle = False
    End Select
End Function



Sub fill_cbo_dren()
    Dim DB As DAO.Database
    Dim rst As DAO.Recordset
    Dim SQL As String

    'Récupérer le nom de l'etablissement
    Set DB = DBEngine.OpenDatabase("C:\SPIDER\SPIDER-APP.spdb", 0, 0, "MS Access;PWD=")

    SQL = "SELECT tbl_drenet.* " & vbCrLf & _
          "FROM tbl_drenet " & vbCrLf & _
          "ORDER BY tbl_drenet.Libelledden;"


    Set rst = DB.OpenRecordset(SQL)
    With frm_activator
        .cbo_dren.Clear
        .cbo_dden.Clear
        While Not rst.EOF
            .cbo_dren.AddItem rst("codeDren") & " | " & rst("libelleDden")
            rst.MoveNext
        Wend
    End With

    rst.Close: Set rst = Nothing
    DB.Close: Set DB = Nothing

End Sub


Sub fill_cbo_dden(my_id_dren)
    Dim DB As DAO.Database
    Dim rst As DAO.Recordset
    Dim SQL As String

    'Récupérer le nom de l'etablissement
    Set DB = DBEngine.OpenDatabase("C:\SPIDER\SPIDER-APP.spdb", 0, 0, "MS Access;PWD=")

    SQL = "SELECT * " & vbCrLf & _
          "FROM tbl_drenet " & vbCrLf & _
          "WHERE Dren_Rattach=" & WithQuotes(my_id_dren) & vbCrLf & _
          "ORDER BY tbl_drenet.Libelledden;"

    Set rst = DB.OpenRecordset(SQL)
    With frm_activator
        .cbo_dden.Clear
        While Not rst.EOF
            .cbo_dden.AddItem rst("codeDren") & " | " & rst("libelleDden")
            rst.MoveNext
        Wend
    End With

    rst.Close: Set rst = Nothing
    DB.Close: Set DB = Nothing

End Sub
Function WithQuotes(ByVal strTxt, Optional DblQuotes As Integer = -1)

    If strTxt = Null Then
        WithQuotes = ""
        Exit Function
    End If

    If DblQuotes Then
        WithQuotes = Chr(34) & strTxt & Chr(34)
    Else
        WithQuotes = Chr(39) & strTxt & Chr(39)
    End If

End Function
Function GetProcessesCount(strAppExeName As String) As Long
    Dim oSvc As Object
    Dim sQuery As String
    Dim oProcs As Object, lgCnt As Long
 
    Set oSvc = GetObject("winmgmts:root\cimv2")
    sQuery = "select * from win32_process where name='" & strAppExeName & "'"
    Set oProcs = oSvc.execquery(sQuery)
    lgCnt = oProcs.Count
    Set oProcs = Nothing
    Set oSvc = Nothing
 
    GetProcessesCount = lgCnt
End Function
Function app_is_running(strAppExeName As String) As Boolean
    app_is_running = (GetProcessesCount(strAppExeName) > 0)
End Function
Function activate_database(BaseAValider As String) As Boolean
          Dim myDB As DAO.Database
          Dim an_scol As String
          Dim code_etab As String
          Dim code_activation_normal
          Dim my_dren As String
          Dim my_dren_complet As String
          Dim my_dden As String
          Dim my_dden_complet As String
          Dim my_dren_ou_dden_complet As String
          Dim my_dren_ou_dden_court As String

10    fct_name = "activate_database"

20    On Error GoTo GestErr

30    Set myDB = DBEngine.OpenDatabase(BaseAValider, False, False, SPD_DORSALE_CONNECT)

          'Inscrire l'année scolaire
40    Call SaveParamEtabBase("AnScol1", frm_activator.zt_anScol.Text, myDB)

50    With frm_activator
60      If Len(.cbo_dren.Text) > 0 Then
70          my_dren = Mid(.cbo_dren.Text, 6)
80          my_dren_complet = "DIRECTION REGIONALE " & IIf(EstVoyelle(Mid(my_dren, 1, 1)), "D'", "DE ") & UCase(my_dren)
90          SaveParamEtabBase "CodeDRENET", Mid(.cbo_dren.Text, 1, 2), myDB
100         SaveParamEtabBase "DREN", my_dren, myDB
110         SaveParamEtabBase "DRENComplet", my_dren_complet, myDB
120     End If

130     If Len(.cbo_dden.Text) > 0 Then
140         my_dden = Mid(.cbo_dden.Text, 6)
150         my_dden_complet = "DIRECTION DEPARTEMENTALE " & IIf(EstVoyelle(Mid(my_dden, 1, 1)), "D'", "DE ") & UCase(my_dden)
160         SaveParamEtabBase "CodeDDENET", Mid(.cbo_dden.Text, 1, 2), myDB
170         SaveParamEtabBase "DDEN", my_dden, myDB
180         SaveParamEtabBase "DDENComplet", my_dden_complet, myDB
190     End If

200     my_dren_ou_dden_complet = IIf(my_dden_complet = "", my_dren_complet, my_dden_complet)
210     SaveParamEtabBase "DRENouDDEN", my_dren_ou_dden_complet, myDB

220     my_dren_ou_dden_court = IIf(my_dden = "", my_dren, my_dden)
230     SaveParamEtabBase "DRENouDDENCourt", my_dren_ou_dden_court, myDB

240     SaveParamEtabBase "CodeEtab", .txt_code_etab.Text, myDB
250     SaveParamEtabBase "NomEtab", .Label_libEtab.Caption, myDB
260     SaveParamEtabBase "NomCompletEtab", .Label_libEtab.Caption, myDB
270     SaveParamEtabBase "NomEtabAbr", .Label_libEtab.Caption, myDB

280     If .zt_activation.Text <> "" Then SaveParamEtabBase "CodeActivation", .zt_activation.Text, myDB
290     If .zt_photo.Text <> "" Then SaveParamEtabBase "code_cert_etab", .zt_photo.Text, myDB
300     If .zt_paie.Text <> "" Then SaveParamEtabBase "code_paie", .zt_paie.Text, myDB
310     If .zt_web_sms.Text <> "" Then SaveParamEtabBase "code_web_sms", .zt_web_sms.Text, myDB
320     If .zt_modem.Text <> "" Then SaveParamEtabBase "code_modem", .zt_modem.Text, myDB
330     If .zt_validity.Text <> "" Then SaveParamEtabBase "vKey", .zt_validity.Text, myDB

340     an_scol = .zt_anScol.Text
350     SaveParamEtabBase "AnScol2", Mid(an_scol, 3, 2) & "-" & Mid("an_scol", 8, 2), myDB
360     SaveParamEtabBase "AnScol3", Mid(an_scol, 6, 4), myDB
370     SaveParamEtabBase "AnScol4", Mid(an_scol, 8, 2), myDB
380     SaveParamEtabBase "AnScol5", "Année scolaire " & an_scol, myDB

390     SaveParamEtabBase "DecoupSemestres", "0", myDB

400     SaveParamEtabBase "migrated", "0", myDB

410     myDB.Execute "UPDATE UsysCtl SET c_e=""" & text_for_sql(Crypt(.txt_code_etab.Text)) & """"

420     myDB.Execute "INSERT INTO UsysDroits VALUES ('ADMINISTRATEUR', 'frm_droits', ""Gérer les utilisateurs et leurs droits d'accès"", 'ADMIN', 0, -1)"

430   End With

          'Met le compteur à zero pour les refELève
440   SaveParamEtabBase "last_ref_eleve", "0000", myDB

450   activate_database = -1


CleanUp:
460   myDB.Close: Set myDB = Nothing
470   Exit Function

DoNotActivate:
480   MsgBox "les paramètres fournis sont invalides!", vbCritical, APP_NAME
490   activate_database = 0
500   GoTo CleanUp


510   Exit Function

GestErr:
520   MsgBox Err.Number & " - " & Err.Description, vbCritical, "Erreur dans la méthode: " & fct_name & " - ligne: " & Erl
530   Resume CleanUp


End Function
Public Function CreateData() As Boolean
    Dim RetVal As Boolean
    Dim Fso As New FileSystemObject
    Dim NewDataFile As String
    Dim Shell_string As String

10  fct_name = "CreateData"

20  On Error GoTo GestErr

    If frm_activator.cbo_dren.Text = "" Then
        MsgBox "Vous devez choisir une DREN et éventuellement une DDEN!", vbCritical
        frm_activator.cbo_dren.SetFocus
    End If


30  NewDataFile = EnregistrerUnFichier(frm_activator.hWnd, "Enregistrer sous", "SPIDER-DATA.spdb", CurDir)

40  If NewDataFile = "" Then Exit Function

50  If Fso.GetExtensionName(NewDataFile) <> "spdb" Then NewDataFile = NewDataFile & ".spdb"

60  Fso.CopyFile "C:\SPIDER\Fichiers\SPIDER-DATA.spdb", NewDataFile

    'mise à jour
70  Shell_string = GetSpecialFolder(SysDir) & "spd_data_update.exe af/s/" & NewDataFile
80  Debug.Print Shell_string

90  Shell Shell_string, vbNormalFocus
100 While app_is_running("spd_data_update.exe")
110     DoEvents
120 Wend

    'Activer
130 Call activate_database(NewDataFile)





    'compacter la base
140 Call CompactDatabase(NewDataFile)

150 Call MsgBox("La base " & NewDataFile & " a été créée et activée avec succès", vbInformation, APP_NAME)


sortie:
160 Screen.MousePointer = 0

170 Exit Function

GestErr:
180 MsgBox Err.Number & " - " & Err.Description, vbCritical, "Erreur dans la méthode: " & fct_name & " - ligne: " & Erl


End Function
Function text_for_sql(original_text)
    text_for_sql = Replace(original_text, Chr(34), """""")
End Function
Public Function CompactDatabase(ByVal strFile As String)
    Dim Fso As New FileSystemObject

10  On Error GoTo GestErr

20  fct_name = "CompactDatabase"

30  f_comp = Replace(strFile, ".spdb", "_cmp.spdb")
40  If Fso.FileExists(f_comp) Then Call Fso.DeleteFile(f_comp, -1)
50  DBEngine.CompactDatabase strFile, f_comp, ";pwd=" & SPD_DORSALE_PWD, dbVersion120, SPD_DORSALE_CONNECT
60  Fso.DeleteFile strFile
70  Fso.MoveFile f_comp, strFile

80  Exit Function

GestErr:
90  MsgBox Err.Number & " - " & Err.Description, vbCritical, "Erreur dans la méthode: " & fct_name & " - ligne: " & Erl

End Function

Function arrayIndexOf(value, myArray As Variant) As Long
    Dim RetVal As Long
    RetVal = -1
    For i = LBound(myArray) To UBound(myArray)
        If myArray(i) = value Then    'Si valeur trouvée
            RetVal = i
            Exit For
        End If
    Next
    arrayIndexOf = RetVal
End Function














