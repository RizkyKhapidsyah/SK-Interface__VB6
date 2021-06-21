VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frm_Analyse_Bas_Niveau 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interface - Analyse Bas Niveau"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   Icon            =   "frm_Analyse_Bas_Niveau.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Quitter l'application"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Quitter l'application"
      Top             =   7320
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Retour à l'écran de paramètrage"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      ToolTipText     =   "Revenir à la fenêtre précédente afin de configurer le port série"
      Top             =   7320
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Zone de réception (Code ASCII)"
      Height          =   2655
      Left            =   360
      TabIndex        =   12
      Top             =   4200
      Width           =   4455
      Begin VB.TextBox Text2 
         Height          =   2295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9120
      TabIndex        =   3
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Insérer le caractère spécial de code ASCII :"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      ToolTipText     =   "Cette zone permet d'insérer dans la chaîne envoyée des caractères spéciaux"
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Frame Frame4 
      Caption         =   "Zone d'envoi"
      Height          =   2535
      Left            =   5280
      TabIndex        =   10
      Top             =   1080
      Width           =   4575
      Begin VB.TextBox Text1 
         Height          =   1365
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label8 
         Caption         =   "Chaine envoyée :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Envoyer"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      ToolTipText     =   "La chaîne de carctères définie dans la zone d'envoi est envoyée vers le port série."
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Zone de réception"
      Height          =   2655
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   4455
      Begin VB.TextBox Text5 
         Height          =   2295
         HideSelection   =   0   'False
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Effacement de la zone de réception"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Effacement complet de la zone de réception"
      Top             =   3840
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7560
      Top             =   240
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   6480
      TabIndex        =   20
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   17
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   15
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10200
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Analyse Bas Niveau du port série"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frm_Analyse_Bas_Niveau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Initialisation_Affichage_Page()
    Rem Ecriture sur l'écran des paramètres de configuration
    Label2.Caption = "Port Com " + Str(Choix_Port)
    Label3.Caption = "Vitesse : " + Vitesse(Choix_Vitesse) + " Bauds"
    Label4.Caption = "Parité : " + Parite(Choix_Parite)
    Label5.Caption = "Nombre de bits de données : " + Bit_Donnee(Choix_Bit_Donnee)
    Label6.Caption = "Nombre de bits d'arrêt : " + Bit_Arret(Choix_Bit_Arret)
    Label9.Caption = "Contrôle de flux : " + Flux(Choix_Flux)
    
    Rem Effacement des zones d'édition
    Text1.Text = ""
    Text2.Text = ""
    Text5.Text = ""
    Text7.Text = ""
    
    Rem Ouverture du Port
    MSComm1.CommPort = Choix_Port
    MSComm1.Settings = Vitesse(Choix_Vitesse) + "," + Parite(Choix_Parite) + "," + Bit_Donnee(Choix_Bit_Donnee) + "," + Bit_Arret(Choix_Bit_Arret)
    MSComm1.Handshaking = Choix_Flux
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
    Rem Vidage du buffer du Port Série
    Dim Chaine As String
    If (MSComm1.InBufferCount > 0) Then
        Chaine = MSComm1.Input
    End If
    Rem Lancement du Timer de scrutation du port Série
    frm_Analyse_Bas_Niveau.Timer1.Enabled = True
End Sub

Private Sub Command1_Click()
   Rem Effacement de la zone de réception
   Text5.Text = ""
   Text2.Text = ""
End Sub

Private Sub Command2_Click()
    Rem Ecriture du port série
    MSComm1.Output = Text1.Text
End Sub

Private Sub Command3_Click()
    Rem Insérer le code ASCII dans la chaîne à envoyer si valeur est comprise entre 0 et 255
    
    Dim Code As Integer
    Dim PositionCurseur As Integer
    Dim LongueurChaine As Integer
    Dim Chaine As String
    Dim DebutChaine As String
    Dim FinChaine As String
        
    Code = Val(Text7.Text)
    PositionCurseur = Text1.SelStart
    If (Code >= 0 And Code <= 255) Then
        Rem la chaîne est coupée en deux et le caractère est inséré entre les deux morceaux
        Chaine = Text1.Text
        LongueurChaine = Len(Text1.Text)
        DebutChaine = Left(Chaine, PositionCurseur)
        FinChaine = Right(Chaine, LongueurChaine - PositionCurseur)
        Chaine = DebutChaine + Chr(Code) + FinChaine
        Text1.Text = Chaine
        Text1.SetFocus
        Text1.SelStart = PositionCurseur + 1
    Else
        Rem le caractère spécial n'est pas inséré
        Text1.SetFocus
        Text1.SelStart = PositionCurseur
    End If
End Sub

Private Sub Fermeture_Fenetre()
    Rem Fermeture du Port
    MSComm1.PortOpen = False
    Rem Arrêt du Timer de scrutation du port Série
    Timer1.Enabled = False
End Sub

Private Sub Command4_Click()
    'Retour écran parmètrage
    Fermeture_Fenetre
    
    Me.Visible = False
    Configuration.Visible = True
End Sub

Private Sub Command5_Click()
    Rem Quitter l'application
    Fermeture_Fenetre
    
    Unload Configuration
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (UnloadMode <> vbFormCode) Then
        Unload Configuration
    End If
End Sub

Private Sub Text5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Rem Gestion de selection miroir du texte dans la zone de réception (Code ASCII)
    
    Dim PositionSelection As Integer
    Dim LongueurSelection As Integer
    Dim DestinationPositionSelection As Integer
    Dim DestinationLongueurSelection As Integer
    Dim Chaine As String
    Dim Caractere As String
    Dim CodeCaractere As Integer
    Dim CaractereLettre As String
    Dim Compteur As Integer
    
    Chaine = Text5.Text
    PositionSelection = Text5.SelStart + 1
    LongueurSelection = Text5.SelLength
    If (LongueurSelection > 0) Then
        Rem Un ou plusieurs caractères ont été sélectonnés
        DestinationPositionSelection = 0
        DestinationLongueurSelection = 0
        For Compteur = 1 To Len(Chaine)
            Caractere = Mid(Chaine, Compteur, 1)
            CodeCaractere = Asc(Caractere)
            CaractereLettre = Str(CodeCaractere)
            If (Compteur < PositionSelection) Then
                DestinationPositionSelection = DestinationPositionSelection + Len(CaractereLettre) + 1
            End If
            If (Compteur >= PositionSelection And Compteur < PositionSelection + LongueurSelection) Then
                DestinationLongueurSelection = DestinationLongueurSelection + Len(CaractereLettre) + 1
            End If
        Next
        Text2.SetFocus
        Text2.SelStart = DestinationPositionSelection
        Text2.SelLength = DestinationLongueurSelection
    End If
End Sub

Private Sub Timer1_Timer()
    Rem Lit les données du Port de Communication
    Dim Chaine As String
    Dim ChaineASCII As String
    Dim Caractere As String
    Dim CodeCaractere As Integer
    Dim CaractereLettre As String
    Dim Compteur As Integer
    
    If (MSComm1.InBufferCount > 0) Then
        Rem Affiche la zone de reception
        Chaine = Text5.Text
        Chaine = Chaine + MSComm1.Input
        Text5.Text = Chaine
        Text5.SelStart = Len(Chaine)
        
        Rem Affiche la zone de reception (Code ASCII)
        ChaineASCII = ""
        For Compteur = 1 To Len(Chaine)
            Caractere = Mid(Chaine, Compteur, 1)
            CodeCaractere = Asc(Caractere)
            CaractereLettre = Str(CodeCaractere)
            ChaineASCII = ChaineASCII + CaractereLettre + " "
        Next
        Text2.Text = ChaineASCII
        Text2.SelStart = Len(ChaineASCII)
    End If
End Sub
