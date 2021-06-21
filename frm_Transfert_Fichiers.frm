VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frm_Transfert_Fichiers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aérospatiale Méaulte : Interface - Transfert de Fichiers"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frm_Transfert_Fichiers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4320
      Top             =   720
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2640
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   -1  'True
      ParityReplace   =   42
      InputMode       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   720
   End
   Begin VB.CommandButton Bouton_Recevoir_Fichier 
      Caption         =   "Recevoir un Fichier"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      ToolTipText     =   "Recevoir un fichier binaire via le port série"
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reception d'un Fichier"
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   6975
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.CommandButton Bouton_Envoyer_Fichier 
      Caption         =   "Envoyer un Fichier"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      ToolTipText     =   "Envoyer un fichier binaire via le port série"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Envoi d'un Fichier"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   6975
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Retour à l'écran de paramètrage"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Revenir à la fenêtre précédente afin de configurer le port série"
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Quitter l'application"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      ToolTipText     =   "Quitter l'application"
      Top             =   4080
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3480
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Transfert de Fichiers via le port série"
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
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   7455
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7440
      Y1              =   3720
      Y2              =   3720
   End
End
Attribute VB_Name = "frm_Transfert_Fichiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Défnition d'une variable Time_Out_Compteur, Time_Out_Phase et Octet_Recu pour gérer le time out
Dim Time_Out_Compteur As Integer
    'Compteur en seconde
Dim Time_Out_Phase As Integer
    '1 -> début de la phase Time_Out_Debut
    '2 -> fin de la phase Time_Out_Debut et début de la phase Time_Out_Fin
    '3 -> fin de la phase Time_Out_Fin
Dim Time_Out_Caractere_Recu As Boolean
    'Indique si au moins un octet a été recu
Dim Chaine_Essai As String
Private Sub Bouton_Envoyer_Fichier_Click()
    'Lit le fichier sur le disque est l'envoi vers le port série
    Dim Reponse
    Dim Octet As Byte
    dlgCommonDialog.Filter = "Tous les fichiers (*.*)|*.*"
    dlgCommonDialog.filename = ""
    dlgCommonDialog.ShowOpen
    On Error GoTo ErrEnvoi
    If Len(dlgCommonDialog.filename) > 0 Then
        ' Un fichier est sélectionné et le bouton Annuler n'a pas été cliqué
        Reponse = MsgBox("Appuyer sur Ok pour envoyer le fichier", vbOKCancel)
        If Reponse = vbOK Then
            Me.MousePointer = vbHourglass
            Label2.Caption = "Emission en cours..."
            Label2.Refresh
            
            ' Paramètrage du port série
            MSComm1.CommPort = Choix_Port
            MSComm1.Settings = Vitesse(Choix_Vitesse) + "," + Parite(Choix_Parite) + "," + Bit_Donnee(Choix_Bit_Donnee) + "," + Bit_Arret(Choix_Bit_Arret)
            MSComm1.Handshaking = Choix_Flux
            MSComm1.InputLen = 0
            MSComm1.PortOpen = True
            ' Vidage du buffer du Port Série
            Dim Chaine As String
            If (MSComm1.InBufferCount > 0) Then
                Chaine = MSComm1.Input
            End If
            
            ' Lecture du fichier et envoi via le port série
            Open dlgCommonDialog.filename For Binary As #1
            
            Do While Not EOF(1)
                Get #1, , Octet
                MSComm1.Output = Chr(Octet)
            Loop
            Close #1
        
            ' L'envoi s'est correctement exécuté
            Label2.Caption = ""
            Me.MousePointer = vbDefault
            MsgBox ("Le fichier " + dlgCommonDialog.filename + " a été envoyé")
            
            Rem Fermeture du Port
            MSComm1.PortOpen = False
        End If
    End If
    
    Exit Sub
ErrEnvoi:
    MsgBox ("Erreur lors de l'envoi du fichier " + dlgCommonDialog.filename)
End Sub

Private Sub Bouton_Recevoir_Fichier_Click()
    Dim Reponse
    
    'Lit le fichier du port série et l'écrit sur le disque
    Dim Octet As Byte
    dlgCommonDialog.Filter = "Tous les fichiers (*.*)|*.*"
    dlgCommonDialog.filename = ""
    dlgCommonDialog.ShowSave
    On Error GoTo ErrReception
    If Len(dlgCommonDialog.filename) > 0 Then
        ' Un fichier est sélectionné et le bouton Annuler n'a pas été cliqué
        Reponse = MsgBox("Appuyer sur Ok pour réceptionner le fichier", vbOKCancel)
        
        If Reponse = vbOK Then
            Me.MousePointer = vbHourglass
            Label3.Caption = "Reception en cours..."
            Label3.Refresh
            
            ' Paramètrage du port série
            MSComm1.CommPort = Choix_Port
            MSComm1.Settings = Vitesse(Choix_Vitesse) + "," + Parite(Choix_Parite) + "," + Bit_Donnee(Choix_Bit_Donnee) + "," + Bit_Arret(Choix_Bit_Arret)
            MSComm1.Handshaking = Choix_Flux
            MSComm1.InputLen = 0
            MSComm1.PortOpen = True
            ' Vidage du buffer du Port Série
            Dim Chaine As String
            If (MSComm1.InBufferCount > 0) Then
                Chaine = MSComm1.Input
            End If
            
            Rem Ouverture du fichier
            Open dlgCommonDialog.filename For Binary As #1
        
            Rem Lancement du Timer de scrutation du port Série
            Time_Out_Compteur = 0
            Time_Out_Caractere_Recu = False
            Time_Out_Phase = 1
            Timer1.Enabled = True
            Timer2.Enabled = True
            Rem Boucle d'attente
            Do
                If (Time_Out_Phase = 1) Then
                    If (Time_Out_Caractere_Recu = False) Then
                        If (Time_Out_Compteur > Time_Out_Debut) Then
                            'le time out début est dépassé
                            Time_Out_Phase = 3
                        End If
                    Else
                        'On passe à de la phase 1 à la phase 2 (Chaine_Recue <> "")
                        Time_Out_Phase = 2
                    End If
                Else
                    If (Time_Out_Compteur > Time_Out_Fin) Then
                        'le time out fin est dépassé (phase 2 -> phase 3)
                        Time_Out_Phase = 3
                    End If
                End If
       
                'Laisse la main aux timers
                DoEvents
            Loop Until Time_Out_Phase = 3
            
            Rem Arrêt du Timer
            Timer1.Enabled = False
            Timer2.Enabled = False
            
            Rem Fermeture du fichier
            Close #1
            
            Label3.Caption = ""
            Me.MousePointer = vbDefault
            Reponse = MsgBox("La réception du fichier est terminée", vbOKOnly)
                
            Rem Fermeture du Port
            MSComm1.PortOpen = False
        End If
    End If
    
    Exit Sub
ErrReception:
    MsgBox ("Erreur lors de la réception du fichier")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If (UnloadMode <> vbFormCode) Then
        Unload Configuration
        Unload frm_Analyse_Bas_Niveau
        Unload frm_Editeur_Texte
    End If
End Sub

Private Sub Command4_Click()
    'Retour écran parmètrage
    Me.Visible = False
    Configuration.Visible = True
End Sub

Private Sub Command5_Click()
    Rem Quitter l'application
    Unload Configuration
    Unload frm_Analyse_Bas_Niveau
    Unload frm_Editeur_Texte
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Dim b() As Byte
    If (MSComm1.InBufferCount > 0) Then
        Time_Out_Caractere_Recu = True
        b() = MSComm1.Input
        Put #1, , b
        If (Time_Out_Phase = 2) Then
            Time_Out_Compteur = 0
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    ' Actualise l'heure et la date de la barre d'état
    Time_Out_Compteur = Time_Out_Compteur + 1
End Sub

