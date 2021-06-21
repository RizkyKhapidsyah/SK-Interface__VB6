VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frm_Editeur_Texte 
   Caption         =   "Interface"
   ClientHeight    =   6555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "frm_Editeur_Texte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6555
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327680
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Nouveau"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Ouvrir"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Enregistrer"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimer"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Couper"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copier"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Coller"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RecevoirFichier"
            Object.ToolTipText     =   "Recevoir un fichier via le port série"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EnvoyerFichier"
            Object.ToolTipText     =   "Envoyer un fichier via le port série"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Retour"
            Object.ToolTipText     =   "Retour à l'écran de paramètrage"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   840
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DTREnable       =   0   'False
      OutBufferSize   =   1024
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   720
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   840
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.TextBox Text1 
      Height          =   5775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   480
      Width           =   10815
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6285
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327680
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13467
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   4920
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327680
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":065C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":09AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":0D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":1052
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":13A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":16F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":1A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":1D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":207C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":2396
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":26B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Editeur_Texte.frx":29CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nouveau"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Ouvrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Enre&gistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Enregistrer &sous..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Imprimer..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParametrage 
         Caption         =   "&Retour paramètrage"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Editi&on"
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Couper"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Co&pier"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "C&oller"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuCommunication 
      Caption         =   "&Communication"
      Begin VB.Menu mnuEnvoyerFichier 
         Caption         =   "&Envoyer le fichier"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuRecevoirFichier 
         Caption         =   "&Recevoir un fichier"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frm_Editeur_Texte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Déclaration de variable privé au module

' Selection permet de sauvegarder la chaine utilisée dans les copier-coller
Dim Selection As String

' Fichier édité
Dim Fichier As String

' Défnition d'une variable Time_Out_Compteur et Time_Out_Phase pour gérer le time out
Dim Time_Out_Compteur As Integer
    'Compteur en seconde
Dim Time_Out_Phase As Integer
    '1 -> début de la phase Time_Out_Debut
    '2 -> fin de la phase Time_Out_Debut et début de la phase Time_Out_Fin
    '3 -> fin de la phase Time_Out_Fin
    
'Chaine qui stocke la chaîne reçue lors de la reception d'un fichier
Dim Chaine_Recue As String

Private Sub Form_Load()
    'Définition du titre de l'application
    Fichier = App.Path + "\" + "doc.txt"
    Caption = "Aérospatiale Méaulte : Interface - Editeur de Texte : " + Fichier
    Call Affiche_Position_Curseur
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Unload Configuration
        Unload frm_Analyse_Bas_Niveau
        Unload frm_Transfert_Fichiers
    End If
End Sub

Private Sub Form_Resize()
    Refresh
End Sub

Private Sub Form_Paint()
    Text1.Left = 0
    Text1.Width = Me.ScaleWidth
    Text1.Top = tbToolBar.Height
    Text1.Height = Me.ScaleHeight - tbToolBar.Height - sbStatusBar.Height
    Text1.Refresh
End Sub

Private Sub mnuEnvoyerFichier_Click()
    Dim Reponse
    Rem Ouverture du Port
    On Error GoTo ErrPortEcriture
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
    
    Reponse = MsgBox("Appuyer sur Ok pour envoyer le fichier", vbOKCancel)
    
    If Reponse = vbOK Then
        Rem Ecriture du port série
        'MSComm1.Output = Text1.Text
        Dim I As Integer
        For I = 1 To Len(Text1.Text)
            MSComm1.Output = Mid(Text1.Text, I, 1)
        Next I
    End If
    
    Rem Fermeture du Port
    MSComm1.PortOpen = False
    
    Exit Sub
    
ErrPortEcriture:
    MsgBox ("Erreur en écriture sur le port série")
End Sub

Private Sub mnuParametrage_Click()
   Afficher_Ecran_Configuration
End Sub

Private Sub Afficher_Ecran_Configuration()
    'Retour à l'écran de paramètrage
    Timer2.Enabled = False
    Me.Visible = False
    Configuration.Visible = True
End Sub

Private Sub mnuRecevoirFichier_Click()
    Dim Reponse
    Rem Ouverture du Port
    On Error GoTo ErrPortLecture
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
    
    Reponse = MsgBox("Appuyer sur Ok pour réceptionner le fichier", vbOKCancel)
    If Reponse = vbOK Then
        Rem Lancement du Timer de scrutation du port Série
        Time_Out_Compteur = 0
        Time_Out_Phase = 1
        Timer1.Enabled = True
        Chaine_Recue = ""
        Rem Boucle d'attente
        Do
            If (Time_Out_Phase = 1) Then
                If (Chaine_Recue = "") Then
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
        Reponse = MsgBox("La réception du fichier est terminée", vbOKOnly)
        Text1.Text = Text1.Text + Chaine_Recue
    End If
    
    Rem Fermeture du Port
    MSComm1.PortOpen = False
    
    Exit Sub
    
ErrPortLecture:
    MsgBox ("Erreur en lecture sur le port série")
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "RecevoirFichier"
            mnuRecevoirFichier_Click
        Case "EnvoyerFichier"
            mnuEnvoyerFichier_Click
        Case "Retour"
            Afficher_Ecran_Configuration
    End Select
End Sub

Private Sub mnuEditCopy_Click()
Rem Commande Copier
    Selection = Mid(Text1, Text1.SelStart + 1, Text1.SelLength)
End Sub

Private Sub mnuEditCut_Click()
Rem Commande Couper
    Dim ChaineDebut, ChaineFin As String
    Selection = Mid(Text1, Text1.SelStart + 1, Text1.SelLength)
    ChaineDebut = Left(Text1, Text1.SelStart)
    ChaineFin = Mid(Text1, Text1.SelStart + Text1.SelLength + 1)
    Text1 = ChaineDebut + ChaineFin
    Text1.SelStart = Len(ChaineDebut)
End Sub

Private Sub mnuEditPaste_Click()
Rem Commande Coller
    Dim ChaineDebut, ChaineFin As String
    ChaineDebut = Left(Text1, Text1.SelStart)
    ChaineFin = Mid(Text1, Text1.SelStart + 1)
    Text1 = ChaineDebut + Selection + ChaineFin
    Text1.SelStart = Len(ChaineDebut + Selection)
End Sub

Private Sub mnuFileOpen_Click()
    Dim Chaine, Ligne As Variant
    Dim LongueurChaine As Long
    dlgCommonDialog.Filter = "Tous les fichiers (*.*)|*.*"
    dlgCommonDialog.filename = ""
    dlgCommonDialog.ShowOpen
    On Error GoTo ErrLecture
    If Len(dlgCommonDialog.filename) > 0 Then
        'Un fichier est sélectionné et le bouton Annuler n'a pas été cliqué
        Chaine = ""
        Open dlgCommonDialog.filename For Input As #1
        Do While Not EOF(1)
            Line Input #1, Ligne
            Chaine = Chaine + Ligne + Chr(13) + Chr(10)
        Loop
        LongueurChaine = Len(Chaine)
        If (LongueurChaine > 1) Then
            Chaine = Left(Chaine, LongueurChaine - 2)
        End If
        Close #1
        ' La lecture s'est correctement exécutée
        Fichier = dlgCommonDialog.filename
        Caption = "Aérospatiale Méaulte : Interface - Editeur de Texte : " + Fichier
        Text1 = Chaine
        Call Affiche_Position_Curseur
    End If
    Exit Sub
ErrLecture:
    MsgBox ("Impossible d'ouvrir le fichier " + dlgCommonDialog.filename)
End Sub

Private Sub mnuFileSave_Click()
    ' Enregistrement du fichier
    Dim Chaine As Variant
    Chaine = Text1
    On Error GoTo ErrEcriture
    Open Fichier For Output As #1
    Print #1, Chaine
    Close #1
    'Le fichier s'est correctement sauvé
    Exit Sub
ErrEcriture:
    MsgBox ("Impossible d'écrire dans le fichier " + dlgCommonDialog.filename)
End Sub

Private Sub mnuFileSaveAs_Click()
    ' Enregistrement du fichier enregistré sous
    Dim Chaine As Variant
    Chaine = Text1
    dlgCommonDialog.Filter = "Tous les fichiers (*.*)|*.*"
    dlgCommonDialog.filename = Fichier
    dlgCommonDialog.ShowSave
    On Error GoTo ErrEcriture
    If Len(dlgCommonDialog.filename) > 0 Then
        Open dlgCommonDialog.filename For Output As #1
        Print #1, Chaine
        Close #1
        'Le fichier s'est correctement sauvé
        Fichier = dlgCommonDialog.filename
        Caption = "Aérospatiale Méaulte : Interface - Editeur de Texte : " + Fichier
        Text1 = Chaine
    End If
    Exit Sub
ErrEcriture:
    MsgBox ("Impossible d'écrire dans le fichier " + dlgCommonDialog.filename)
End Sub

Private Sub mnuFilePrint_Click()
    Dim Position, NumPage, LongueurChaine, Ligne, Y, NbLignes, NbCaractères, HauteurTexte, LongueurTexte As Long
    Dim Chaine, Header, Footer, ChaineImprimee As String
    ' Impression de la page
    ' Initialisation des variables
    NbLignes = 78
    NbCaractères = 96
    Chaine = ""
    NumPage = 1
    Ligne = 0
    On Error GoTo ErrImpression
    ' Affichage de l'entête de la première page
    Header = "Impression de " + Fichier
    Printer.Print Header
    HauteurTexte = Printer.TextHeight(Header)
    Printer.Line (0, Int(HauteurTexte * 1.5))-(Printer.ScaleWidth, Int(HauteurTexte * 1.5))
    Printer.CurrentX = 0
    Printer.CurrentY = 2 * HauteurTexte
    For Position = 1 To Len(Text1.Text)
        If Ligne >= NbLignes Then
            ' Affichage du bas de page
            Y = Printer.CurrentY
            Footer = "Page" + Str(NumPage)
            HauteurTexte = Printer.TextHeight(Footer)
            LongueurTexte = Printer.TextWidth(Footer)
            Printer.Line (0, Y + Int(HauteurTexte / 2))-(Printer.ScaleWidth, Y + Int(HauteurTexte / 2))
            Printer.CurrentX = Printer.ScaleWidth - LongueurTexte
            Printer.CurrentY = Y + HauteurTexte
            Printer.Print Footer
    
            ' Imprime une nouvelle page
            Printer.NewPage
            NumPage = NumPage + 1
            
            ' Affichage de l'entête
            Header = "Impression de " + Fichier
            Printer.Print Header
            HauteurTexte = Printer.TextHeight(Header)
            Printer.Line (0, Int(HauteurTexte * 1.5))-(Printer.ScaleWidth, Int(HauteurTexte * 1.5))
            Printer.CurrentX = 0
            Printer.CurrentY = 2 * HauteurTexte
            Ligne = 0
        End If
        Chaine = Chaine + Mid(Text1.Text, Position, 1)
        LongueurChaine = Len(Chaine)
        
        If LongueurChaine >= NbCaractères Then
            Chaine = Chaine + Chr(13) + Chr(10)
        End If
        
        If Right(Chaine, 2) = Chr(13) + Chr(10) Then
            Ligne = Ligne + 1
            LongueurChaine = Len(Chaine)
            ChaineImprimee = Left(Chaine, LongueurChaine - 2)
            Printer.Print ChaineImprimee
            Chaine = ""
        End If
    Next
    ' Affichage de la dernière ligne
    If Len(Chaine) > 0 Then
        Printer.Print Chaine
    End If
    ' Affichage du bas de page de la dernière page
    Footer = "Page" + Str(NumPage)
    HauteurTexte = Printer.TextHeight(Footer)
    LongueurTexte = Printer.TextWidth(Footer)
    Y = Printer.ScaleHeight - 2 * HauteurTexte
    Printer.Line (0, Y + Int(HauteurTexte / 2))-(Printer.ScaleWidth, Y + Int(HauteurTexte / 2))
    Printer.CurrentX = Printer.ScaleWidth - LongueurTexte
    Printer.CurrentY = Y + HauteurTexte
    Printer.Print Footer
    ' Fin de l'impression
    Printer.EndDoc
    Exit Sub
    
ErrImpression:
    MsgBox ("Impossible d'imprimer sur l'imprimante " + Printer.DeviceName)
End Sub

Private Sub mnuFileExit_Click()
    Unload Configuration
    Unload frm_Analyse_Bas_Niveau
    Unload frm_Transfert_Fichiers
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    'Nouveau Fichier
    Text1 = ""
    Fichier = App.Path + "\" + "doc.txt"
    Caption = "Aérospatiale Méaulte : Interface " + Fichier
    Call Affiche_Position_Curseur
End Sub

Private Sub Affiche_Position_Curseur()
    Dim PosCurseur, Compteur, Ligne, Colonne As Long
    PosCurseur = Text1.SelStart
    Ligne = 1
    Colonne = 1
    For Compteur = 1 To PosCurseur
        Colonne = Colonne + 1
        If Mid(Text1.Text, Compteur, 1) = Chr(13) Then
            Ligne = Ligne + 1
            Colonne = 0
        End If
    Next
    sbStatusBar.Panels.Item(1).Text = "Col" + Str(Colonne) + ", Ln" + Str(Ligne)
End Sub

Private Sub Text1_Click()
    'Actualise le contenu de la barre d'état
    Call Affiche_Position_Curseur
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    'Actualise le contenu de la barre d'état
    Call Affiche_Position_Curseur
End Sub

Private Sub Timer1_Timer()
    If (MSComm1.InBufferCount > 0) Then
        Chaine_Recue = Chaine_Recue + MSComm1.Input
        If (Time_Out_Phase = 2) Then
            Time_Out_Compteur = 0
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    ' Actualise l'heure et la date de la barre d'état
    sbStatusBar.Panels.Item(2).Text = Date
    sbStatusBar.Panels.Item(3).Text = Time
    Time_Out_Compteur = Time_Out_Compteur + 1
End Sub
