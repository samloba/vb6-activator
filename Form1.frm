VERSION 5.00
Begin VB.Form frm_activator 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spider activator"
   ClientHeight    =   8775
   ClientLeft      =   11730
   ClientTop       =   6540
   ClientWidth     =   6585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   6585
   Begin VB.CommandButton btn_gen_data 
      Caption         =   "Générer Data"
      Height          =   480
      Left            =   4440
      TabIndex        =   54
      Top             =   7800
      Width           =   1515
   End
   Begin VB.ComboBox cbo_dden 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   53
      Text            =   "cbo_dren"
      Top             =   7200
      Width           =   5295
   End
   Begin VB.CheckBox opt_wave 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   51
      Top             =   5220
      Width           =   300
   End
   Begin VB.TextBox zt_wave 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   720
      TabIndex        =   50
      Top             =   5280
      Width           =   2235
   End
   Begin VB.CheckBox opt_RhControl 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   48
      Top             =   4260
      Width           =   300
   End
   Begin VB.TextBox zt_RhControl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   3720
      TabIndex        =   47
      Top             =   4320
      Width           =   2235
   End
   Begin VB.TextBox zt_schoolcontrol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   720
      TabIndex        =   45
      Top             =   4320
      Width           =   2235
   End
   Begin VB.CheckBox opt_schoolcontrol 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   44
      Top             =   4260
      Width           =   300
   End
   Begin VB.CheckBox opt_cinetpay 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   43
      Top             =   3480
      Width           =   300
   End
   Begin VB.TextBox zt_cinetpay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   3720
      TabIndex        =   41
      Top             =   3480
      Width           =   2235
   End
   Begin VB.CommandButton btn_extract_etabs_to_json 
      Caption         =   "extraire etabs pour l'appli mobile"
      Height          =   495
      Left            =   720
      TabIndex        =   40
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CheckBox opt_web_sms 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   38
      Top             =   2565
      Width           =   300
   End
   Begin VB.TextBox zt_web_sms 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   3720
      TabIndex        =   37
      Top             =   2640
      Width           =   2235
   End
   Begin VB.TextBox zt_photo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   405
      Left            =   3705
      TabIndex        =   13
      Top             =   1920
      Width           =   2235
   End
   Begin VB.TextBox zt_paie 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   720
      TabIndex        =   33
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CheckBox opt_paie 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   32
      Top             =   2565
      Width           =   300
   End
   Begin VB.CheckBox opt_modem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   345
      TabIndex        =   30
      Top             =   6075
      Width           =   300
   End
   Begin VB.CheckBox opt_validity 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   29
      Top             =   3420
      Width           =   300
   End
   Begin VB.TextBox zt_expireDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   450
      Left            =   4080
      TabIndex        =   26
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox cbo_dren 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   720
      TabIndex        =   25
      Text            =   "cbo_dren"
      Top             =   6720
      Width           =   5295
   End
   Begin VB.TextBox zt_modem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   3720
      TabIndex        =   24
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton btn_gen_all 
      Caption         =   "Extraire tout"
      Height          =   375
      Left            =   9000
      TabIndex        =   23
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton btn_copy_sms_code 
      Height          =   255
      Left            =   7725
      TabIndex        =   22
      Top             =   3405
      Width           =   255
   End
   Begin VB.CommandButton btn_copy_reactivation_code 
      Height          =   255
      Left            =   7725
      TabIndex        =   21
      Top             =   3030
      Width           =   255
   End
   Begin VB.CommandButton btn_copy_certif_code 
      Height          =   255
      Left            =   7725
      TabIndex        =   20
      Top             =   2670
      Width           =   255
   End
   Begin VB.CommandButton btn_copy_activation_code 
      Height          =   255
      Left            =   7725
      TabIndex        =   19
      Top             =   2325
      Width           =   255
   End
   Begin VB.CommandButton btn_copy 
      Caption         =   "Copier tout"
      Height          =   360
      Left            =   8400
      TabIndex        =   18
      Top             =   960
      Width           =   1065
   End
   Begin VB.CommandButton btn_fich_activ 
      Caption         =   "Créer fichier"
      Height          =   360
      Left            =   9480
      TabIndex        =   17
      Top             =   960
      Width           =   1035
   End
   Begin VB.TextBox txt_code_etab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   450
      Left            =   2640
      TabIndex        =   16
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox zt_imei 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   405
      Left            =   720
      TabIndex        =   15
      Top             =   6120
      Width           =   2835
   End
   Begin VB.TextBox zt_validity 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   720
      TabIndex        =   14
      Top             =   3480
      Width           =   2235
   End
   Begin VB.TextBox zt_activation 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   720
      TabIndex        =   12
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox zt_anScol 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   450
      Left            =   720
      MaxLength       =   9
      TabIndex        =   11
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame opt_fichier 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1155
      Left            =   7200
      TabIndex        =   7
      Top             =   240
      Width           =   3060
      Begin VB.CheckBox opt_sms_old 
         BackColor       =   &H00FFFFFF&
         Caption         =   "sms"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   360
         Width           =   825
      End
      Begin VB.CheckBox opt_data 
         BackColor       =   &H00FFFFFF&
         Caption         =   "data"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         TabIndex        =   9
         Top             =   105
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox opt_cert 
         BackColor       =   &H00FFFFFF&
         Caption         =   "photos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CheckBox opt_activation 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   28
      Top             =   1860
      Value           =   1  'Checked
      Width           =   225
   End
   Begin VB.CheckBox opt_photo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3345
      TabIndex        =   31
      Top             =   1860
      Width           =   420
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Wave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   52
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "RH-Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3720
      TabIndex        =   49
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label label_schoolcontrol 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SchoolControl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   46
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label label_cinetpay 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "cinetpay"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3720
      TabIndex        =   42
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code web sms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3720
      TabIndex        =   39
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label_libEtab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   720
      TabIndex        =   36
      Top             =   960
      Width           =   5295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code certif. photos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Etab."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   2640
      TabIndex        =   35
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code activation data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code paie"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   34
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date d'expiration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code modem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code de validité"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom Etablissement"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   9105
      TabIndex        =   4
      Top             =   2700
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Année scolaire"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Etablissement"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   9135
      TabIndex        =   0
      Top             =   2220
      Width           =   1575
   End
End
Attribute VB_Name = "frm_activator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const anneeScolaire As String = "2024-2025"



Private Sub btn_copy_activation_code_Click()
    Call copy_to_clipboard(1)
End Sub

Private Sub btn_copy_certif_code_Click()
    Call copy_to_clipboard(2)
End Sub
Private Sub btn_copy_Click()
    Call copy_to_clipboard(5)
End Sub

Private Sub btn_copy_reactivation_code_Click()
    Call copy_to_clipboard(3)
End Sub

Private Sub btn_copy_sms_code_Click()
    Call copy_to_clipboard(4)
End Sub

Private Sub btn_extract_etabs_to_json_Click()
    Dim DB As DAO.Database
    Dim rst As DAO.Recordset
    Dim SQL As String
    Dim ojb As New JsonBag
    Dim filePath As String

    Set DB = DBEngine.OpenDatabase("C:\SPIDER\SPIDER-APP.spdb", 0, 0, "MS Access;PWD=")

    SQL = "SELECT tbl_etab.id_etab AS _id, tbl_etab.lib_etab AS name, tbl_etab.Commune_etab AS commune, " & _
          "IIf([tbl_etab].[primaire]=True,1,0) AS primaire, tbl_etab.codedren AS CodeDren, tbl_etab.codeiep AS CodeIep " & vbCrLf & _
          "FROM tbl_etab;"

    Set rst = DB.OpenRecordset(SQL)

    filePath = EnregistrerUnFichier(frm_activator.hWnd, "Enregistrer sous", "etabs.json", CurDir)

    If filePath = "" Then Exit Sub

    On Error GoTo GestErr

    Screen.MousePointer = 11
    With ojb
        .json = ojb.RecordSet2JSON(rst)
        Call .OutPutToTextFile(filePath)
    End With

    MsgBox "Fichier enregistré avec succès vers " & filePath, vbInformation, "spdActivator - Succès"

sortie:
    Screen.MousePointer = 0
    Exit Sub

GestErr:
    MsgBox Err.Description, vbCritical, "spdActivator - Erreur"
    Resume sortie

End Sub

Private Sub btn_fich_activ_Click()

    target_dir = ChoisirUnDossier("VEUILLEZ INDIQUER UN EMPLACEMENT DE SORTIE POUR LE FICHIER D'ACTIVATION", -1)
    If target_dir = "" Then Exit Sub
    target_dir = IIf(Right(target_dir, 1) = "\", target_dir, target_dir & "\")

    target_file = target_dir & "Paramètres d'activation_" & Me.txt_code_etab.Text & "_" & Me.Label_libEtab.Caption & ".ini"

    my_section = Me.zt_anScol

    If Me.opt_data Then
        Call EcrireDansFichierINI(my_section, "CodeEtab", Me.txt_code_etab.Text, target_file)
        Call EcrireDansFichierINI(my_section, "NomEtab", Me.Label_libEtab.Caption, target_file)
        Call EcrireDansFichierINI(my_section, "CodeActivation", Me.zt_activation.Text, target_file)
    End If

    If Me.opt_photo Then Call EcrireDansFichierINI(my_section, "code_cert_etab", Me.zt_photo.Text, target_file)
    If Me.opt_modem Then Call EcrireDansFichierINI(my_section, "code_activation_modem", Me.zt_modem.Text, target_file)

    MsgBox "Le fichier a été créé avec succès!", vbInformation


    Dim myFile As String
    myFile = IIf(Right(App.Path, 1) = "\", App.Path & "DataConfig.ini", App.Path & "\DataConfig.ini")

    Call EcrireDansFichierINI("DATA", Me.txt_code_etab.Text, Me.txt_code_etab.Text & ", " & Me.Label_libEtab.Caption & ", " & Now, myFile)

End Sub

Private Sub btn_gen_all_Click()
    gen_all_codes_to_txt_file
End Sub

Private Sub btn_gen_data_Click()
    Call CreateData
End Sub

Private Sub cbo_dren_Click()

    myText = Me.cbo_dren.Text

    If Len(myText) = 0 Then

    Else
        my_id_dren = Mid(myText, 1, 2)
        Call fill_cbo_dden(my_id_dren)
    End If

    'MsgBox Me.cbo_dren.Text
    'Call fill_cbo_dden("")

End Sub

Private Sub opt_activation_Click()
    generer_code
End Sub

Private Sub opt_cinetpay_Click()
    generer_code
End Sub

Private Sub opt_modem_Click()
    generer_code
End Sub

Private Sub opt_paie_Click()
    generer_code
End Sub

Private Sub opt_photo_Click()
    generer_code
End Sub

Private Sub opt_RhControl_Click()
generer_code
End Sub

Private Sub opt_schoolcontrol_Click()
    generer_code
End Sub

Private Sub opt_validity_Click()
    generer_code
End Sub

Private Sub opt_wave_Click()
generer_code
End Sub

Private Sub opt_web_sms_Click()
    generer_code
End Sub

Private Sub zt_anScol_Change()
    Call generer_code
End Sub


Private Sub txt_code_etab_Change()
    Call generer_code
End Sub


Private Sub Form_Load()
    Dim my_hdd_id As Long

    '    my_hdd_id = GetDiskSerial()
    '
    '    If my_hdd_id <> LOBA_HDD_ID And my_hdd_id <> ROBY_HDD_ID Then
    '        my_new_hdd_id = InputBox("VEUILLEZ SAISIR LE MOT DE PASSE", "spider")
    '        If my_new_hdd_id = "" Or (my_new_hdd_id <> LOBA_HDD_ID And my_new_hdd_id <> ROBY_HDD_ID) Then
    '            End
    '        End If
    '    End If

    'récupérer la dernière année saisie
    Dim iniPath As String
    Dim Fso As New FileSystemObject
    Dim anScol As String
    iniPath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & App.EXEName & ".ini"
    If Fso.FileExists(iniPath) Then
        anScol = LireDansFichierINI("RECENTS", "anScol", iniPath, anneeScolaire)
    Else
        anScol = anneeScolaire
    End If
    Me.zt_anScol.Text = anScol



    'remplissage des drens et ddens
    Call fill_cbo_dren

End Sub
Function copy_to_clipboard(copy_mode As Long)
    Dim clipboard As MSForms.DataObject
    Dim my_text As String

    If Me.zt_activation.Text = "" Then Exit Function

    Select Case copy_mode
        Case 1: my_text = Me.zt_activation.Text
        Case 2: my_text = Me.zt_photo.Text
        Case 3: my_text = Me.zt_validity.Text
        Case 4: my_text = Me.zt_modem.Text

        Case 5: my_text = get_clipboard_text
    End Select

    If my_text = "" Then Exit Function

    Set clipboard = New MSForms.DataObject
    clipboard.SetText my_text
    clipboard.PutInClipboard

End Function

Private Sub zt_cinetpay_Change()
    Call generer_code
End Sub

Private Sub zt_expireDate_Change()
    Call generer_code
End Sub

Private Sub zt_imei_Change()
    Call generer_code
End Sub
Function get_clipboard_text() As String
    X = "Code d'activation: " & Me.zt_activation.Text _
        & vbCrLf & "Code certif. photos: " & Me.zt_photo.Text _
        & vbCrLf & "Code activation sms: " & Me.zt_validity.Text _
        & vbCrLf & "Code activation sms: " & Me.zt_modem.Text
    get_clipboard_text = X
End Function
Sub generer_code()
    Dim DB As DAO.Database
    Dim rst As DAO.Recordset
    'dim oValidity as New cv
    Dim codeEtab As String
    Dim anScol As String
    Dim expireDate As String

    If Len(txt_code_etab.Text) = 6 And Len(Me.zt_anScol.Text) = 9 Then

        codeEtab = txt_code_etab.Text
        anScol = Me.zt_anScol.Text
        expireDate = Me.zt_expireDate.Text

        'Récupérer le nom de l'etablissement
        Set DB = DBEngine.OpenDatabase("C:\SPIDER\SPIDER-APP.spdb", 0, 0, "MS Access;PWD=")
        Set rst = DB.OpenRecordset("SELECT lib_etab FROM tbl_etab WHERE id_etab='" & codeEtab & "'")
        If Not rst.EOF Then
            Me.Label_libEtab.Caption = rst(0)
        Else
            Me.Label_libEtab.Caption = ""
        End If
        rst.Close: Set rst = Nothing
        DB.Close: Set DB = Nothing

        'activation
        If Me.opt_activation.value = Checked Then
            Me.zt_activation.Text = getKeyByType("activation", codeEtab, anScol)
        Else
            Me.zt_activation.Text = ""
        End If

        'photo
        If Me.opt_photo.value = Checked Then
            Me.zt_photo.Text = getKeyByType("photo", codeEtab, anScol)
        Else
            Me.zt_photo.Text = ""
        End If

        'paie
        If Me.opt_paie.value = Checked Then
            Me.zt_paie.Text = getKeyByType("paie", codeEtab, anScol)
        Else
            Me.zt_paie.Text = ""
        End If

        'web sms
        If Me.opt_web_sms.value = Checked Then
            Me.zt_web_sms.Text = getKeyByType("sms", codeEtab, anScol)
        Else
            Me.zt_web_sms.Text = ""
        End If

        'cinetpay
        If Me.opt_cinetpay.value = Checked Then
            Me.zt_cinetpay.Text = getKeyByType("cinetpay", codeEtab, anScol)
        Else
            Me.zt_cinetpay.Text = ""
        End If

        'label_schoolcontrol
        If Me.opt_schoolcontrol.value = Checked Then
            Me.zt_schoolcontrol.Text = getKeyByType("schoolcontrol", codeEtab, anScol)
        Else
            Me.zt_schoolcontrol.Text = ""
        End If

        'Rhcontrol
        If Me.opt_RhControl.value = Checked Then
            Me.zt_RhControl.Text = getKeyByType("rhcontrol", codeEtab, anScol)
        Else
            Me.zt_RhControl.Text = ""
        End If

        'wave
        If Me.opt_wave.value = Checked Then
            Me.zt_wave.Text = getKeyByType("wave", codeEtab, anScol)
        Else
            Me.zt_wave.Text = ""
        End If





        'vKey = getValidityKey_v2(get_request_code(.item("createdDate")), .item("expireDate"))
        If IsDate(expireDate) And Me.opt_validity.value = Checked Then
            Me.zt_validity.Text = getValidityKey_v2(get_request_code(CLng(Date), codeEtab, anScol), expireDate)
        Else
            Me.zt_validity.Text = ""
        End If

        If Len(Me.zt_imei.Text) = 15 Then
            If Me.opt_modem.value = Checked Then
                Me.zt_modem.Text = getKeyByType("modem", codeEtab, anScol, Me.zt_imei.Text)
            Else
                Me.zt_modem.Text = ""
            End If
        End If




        'Mémoriser l'année
        Dim iniPath As String
        iniPath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & App.EXEName & ".ini"
        Call EcrireDansFichierINI("RECENTS", "anScol", Me.zt_anScol.Text, iniPath)

    Else

        Me.Label_libEtab.Caption = ""
        Me.zt_activation.Text = ""
        Me.zt_photo.Text = ""
        Me.zt_validity.Text = ""
        Me.zt_modem.Text = ""

    End If
End Sub

Sub gen_all_codes_to_txt_file()
    Dim Fso As New FileSystemObject
    Dim DB As DAO.Database
    Dim rst As DAO.Recordset
    Dim RetVal As String
    Dim my_id_etab As String
    Dim my_activation_code
    Dim my_cert_code
    Dim my_reactiv_code
    Dim sFilePath As String



    If Len(Me.zt_anScol.Text) = 9 Then

        sFilePath = EnregistrerUnFichier(Me.hWnd, "Enregistrer sous", "codes_" & Me.zt_anScol.Text & ".txt", App.Path)
        If sFilePath = "" Then Exit Sub

        Screen.MousePointer = 11

        Set DB = DBEngine.OpenDatabase("C:\SPIDER\SPIDER-APP.spdb", 0, 0, "MS Access;PWD=")
        Set rst = DB.OpenRecordset("SELECT id_etab, lib_etab FROM tbl_etab ORDER BY lib_etab")

        While Not rst.EOF
            my_id_etab = Right(rst!id_etab, 6)
            my_activation_code = get_activation_code(Me.zt_anScol.Text, my_id_etab)
            my_cert_code = code_cert_etab(Me.zt_anScol.Text, my_id_etab)
            my_reactiv_code = get_reactivation_code(my_id_etab, Me.zt_anScol.Text, 0)

            RetVal = RetVal & rst!lib_etab & " (" & my_id_etab & ")" & vbTab & my_activation_code & vbTab & my_cert_code & vbTab & my_reactiv_code & vbNewLine
            DoEvents
            rst.MoveNext
        Wend

    End If



    'Debug.Print RetVal

    If Right(sFilePath, 4) <> ".txt" Then sFilePath = sFilePath & ".txt"
    Set TxtFile = Fso.CreateTextFile(sFilePath, True)
    TxtFile.Write (RetVal)

    'If Not rst.EOF Then Me.Label_libEtab.Caption = rst(0)
    rst.Close: Set rst = Nothing
    DB.Close: Set DB = Nothing


    MsgBox "terminé", vbInformation, "Activator"

    Screen.MousePointer = 0



End Sub

