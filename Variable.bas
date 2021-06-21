Attribute VB_Name = "Variable"
'
' Définition des Variables Globales
'

' La vitesse de la transmission
Public Choix_Vitesse As Integer
Public Vitesse(0 To 13) As String

' La parite de la transmission
Public Choix_Parite As Integer
Public Parite(0 To 4) As String

' Le nombre de bits de stop de la transmission
Public Choix_Bit_Donnee As Integer
Public Bit_Donnee(0 To 4) As String

' Le nombre de bits d'arrêt de la transmission
Public Choix_Bit_Arret As Integer
Public Bit_Arret(0 To 2) As String

' Le contrôle du flux de la transmission
Public Choix_Flux As Integer
Public Flux(0 To 3) As String

' Choix du port pour la communication
Public Choix_Port As Integer

' Définition du time_out_debut pour la réception de fichier
Public Time_Out_Debut As Integer

' Définition du time_out_fin pour la réception de fichier
Public Time_Out_Fin As Integer
