VERSION 5.00
Begin VB.Form Configuration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interface Port Série"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ForeColor       =   &H80000008&
   Icon            =   "Configuration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Parametres_Defaut 
      Caption         =   "&Paramètres par défaut"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      ToolTipText     =   "Restaure la configuration par défaut du port série"
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton Sauver_Parametres 
      Caption         =   "&Sauvegarder les paramètres"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      ToolTipText     =   "Enregistre la configuration du port série dans la base de registres"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Analyse_Bas_Niveau 
      Caption         =   "&Analyse du Port COM"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      ToolTipText     =   "Permet le transfert caractère par caractère"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton Quitter 
      Caption         =   "&Quitter"
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      ToolTipText     =   "Au revoir..."
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Frame Frame3 
      Caption         =   "Time Out"
      Height          =   1095
      Left            =   360
      TabIndex        =   22
      Top             =   4080
      Width           =   4575
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Durée du time out fin en seconde :"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label8 
         Caption         =   "Durée du time out début en seconde :"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sélection du port"
      Height          =   735
      Left            =   2880
      TabIndex        =   21
      Top             =   1080
      Width           =   2055
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   360
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Com 2"
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Com 1"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paramètres"
      Height          =   2655
      Left            =   360
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1320
         Width           =   510
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   840
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Configuration.frx":030A
         Left            =   960
         List            =   "Configuration.frx":030C
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Flux :"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Bits d'arrêt :"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Bits de données :"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Parité :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Vitesse :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Image A_Propos_De 
      Height          =   480
      Left            =   3600
      Picture         =   "Configuration.frx":030E
      ToolTipText     =   "A propos de..."
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Caption         =   "Configuration du port série"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Prise_Compte_Paramètres()
    ' Prise en compte des paramètres de configuration du Port Série
    Choix_Vitesse = Combo1.ListIndex
    Choix_Parite = Combo2.ListIndex
    Choix_Bit_Donnee = Combo3.ListIndex
    Choix_Bit_Arret = Combo4.ListIndex
    Choix_Flux = Combo5.ListIndex
    If (Option1.Value = True) Then
        Choix_Port = 1
    Else
        Choix_Port = 2
    End If
    Time_Out_Debut = Int(Val(Text1.Text))
    Time_Out_Fin = Int(Val(Text2.Text))
End Sub

Private Sub A_Propos_De_Click()
    ' Version et auteurs du logiciel Interface
    Dim Reponse
    Dim Chaine As String
    Chaine = "Logiciel Interface :" + Chr(13) + Chr(10) + "-----------------------------" + Chr(13) + Chr(10) + "Communication via le port série." + Chr(13) + Chr(10) + Chr(13) + Chr(10)
    Chaine = Chaine + "Version : 1.1" + Chr(13) + Chr(10)
    Chaine = Chaine + "Date : 10/03/1999." + Chr(13) + Chr(10)
    Chaine = Chaine + "Auteur : sibair"
    Reponse = MsgBox(Chaine, vbInformation + vbOKOnly, "A propos d'Interface")
End Sub

Private Sub Analyse_Bas_Niveau_Click()
    Prise_Compte_Paramètres
    
    Rem Ouverture de la fenetre de transmission
    Me.Visible = False
    frm_Analyse_Bas_Niveau.Initialisation_Affichage_Page
    frm_Analyse_Bas_Niveau.Visible = True
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (UnloadMode <> vbFormCode) Then
        ' l'évenement Unload ne vient pas du code
        Unload frm_Analyse_Bas_Niveau
    End If
End Sub

Private Sub Parametres_Defaut_Click()
    ' Chargement des paramètres par défaut
    Choix_Vitesse = 6
    Choix_Parite = 2
    Choix_Bit_Donnee = 4
    Choix_Bit_Arret = 0
    Choix_Flux = 0
    Choix_Port = 1
    Time_Out_Debut = 5
    Time_Out_Fin = 5
    Affiche_Parametres
End Sub

Private Sub Sauver_Parametres_Click()
    ' Sauver les paramètres du port série
    Prise_Compte_Paramètres
    SaveSetting App.Title, "Settings", "Vitesse", Mid(Str(Choix_Vitesse), 2)
    SaveSetting App.Title, "Settings", "Parite", Mid(Str(Choix_Parite), 2)
    SaveSetting App.Title, "Settings", "Bit_Donnee", Mid(Str(Choix_Bit_Donnee), 2)
    SaveSetting App.Title, "Settings", "Bit_Arret", Mid(Str(Choix_Bit_Arret), 2)
    SaveSetting App.Title, "Settings", "Flux", Mid(Str(Choix_Flux), 2)
    SaveSetting App.Title, "Settings", "Port", Mid(Str(Choix_Port), 2)
    SaveSetting App.Title, "Settings", "Time_Out_Debut", Mid(Str(Time_Out_Debut), 2)
    SaveSetting App.Title, "Settings", "Time_Out_Fin", Mid(Str(Time_Out_Fin), 2)
End Sub

Public Sub Affiche_Parametres()
    ' Paramètres génénaux
    ' La vitesse de la transmission
    Combo1.ListIndex = Choix_Vitesse
    ' La parite de la transmission
    Combo2.ListIndex = Choix_Parite
    ' Le nombre de bits de donnees de la transmission
    Combo3.ListIndex = Choix_Bit_Donnee
    ' Le nombre de bits d'arrêt de la transmission
    Combo4.ListIndex = Choix_Bit_Arret
    ' Le flux de la transmission
    Combo5.ListIndex = Choix_Flux
    ' Choix du port
    If (Choix_Port = 1) Then
        Option1.Value = True
        Option2.Value = False
    Else
        Option1.Value = False
        Option2.Value = True
    End If
    ' Affichage des informations de Time_Out
    Text1.Text = Mid(Str(Time_Out_Debut), 2)
    Text2.Text = Mid(Str(Time_Out_Fin), 2)
End Sub

Private Sub Quitter_Click()
    Unload frm_Analyse_Bas_Niveau
    Unload Me
End Sub

