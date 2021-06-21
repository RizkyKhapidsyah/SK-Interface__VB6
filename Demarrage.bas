Attribute VB_Name = "Demarrage"
Sub Main()
    'Chargement des paramètres par défaut du port série
    Choix_Vitesse = Val(GetSetting(App.Title, "Settings", "Vitesse", 5))
    Choix_Parite = Val(GetSetting(App.Title, "Settings", "Parite", 2))
    Choix_Bit_Donnee = Val(GetSetting(App.Title, "Settings", "Bit_donnee", 4))
    Choix_Bit_Arret = Val(GetSetting(App.Title, "Settings", "Bit_Arret", 0))
    Choix_Flux = Val(GetSetting(App.Title, "Settings", "Flux", 0))
    Choix_Port = Val(GetSetting(App.Title, "Settings", "Port", 1))
    Time_Out_Debut = Val(GetSetting(App.Title, "Settings", "Time_Out_Debut", 5))
    Time_Out_Fin = Val(GetSetting(App.Title, "Settings", "Time_Out_Fin", 5))
    
    'Chargement et configuration de la fenêtre Configuration
    Load Configuration
    Load frm_Analyse_Bas_Niveau
    
    'La vitesse de la transmission
    Vitesse(0) = "110"
    Configuration.Combo1.AddItem (Vitesse(0))
    Vitesse(1) = "300"
    Configuration.Combo1.AddItem (Vitesse(1))
    Vitesse(2) = "600"
    Configuration.Combo1.AddItem (Vitesse(2))
    Vitesse(3) = "1200"
    Configuration.Combo1.AddItem (Vitesse(3))
    Vitesse(4) = "2400"
    Configuration.Combo1.AddItem (Vitesse(4))
    Vitesse(5) = "4800"
    Configuration.Combo1.AddItem (Vitesse(5))
    Vitesse(6) = "9600"
    Configuration.Combo1.AddItem (Vitesse(6))
    Vitesse(7) = "14400"
    Configuration.Combo1.AddItem (Vitesse(7))
    Vitesse(8) = "19200"
    Configuration.Combo1.AddItem (Vitesse(8))
    Vitesse(9) = "28800"
    Configuration.Combo1.AddItem (Vitesse(9))
    Vitesse(10) = "38400"
    Configuration.Combo1.AddItem (Vitesse(10))
    Vitesse(11) = "56000"
    Configuration.Combo1.AddItem (Vitesse(11))
    Vitesse(12) = "128000"
    Configuration.Combo1.AddItem (Vitesse(12))
    Vitesse(13) = "256000"
    Configuration.Combo1.AddItem (Vitesse(13))
    
    Rem La parite de la transmission
    Parite(0) = "E"
    Configuration.Combo2.AddItem (Parite(0))
    Parite(1) = "M"
    Configuration.Combo2.AddItem (Parite(1))
    Parite(2) = "N"
    Configuration.Combo2.AddItem (Parite(2))
    Parite(3) = "O"
    Configuration.Combo2.AddItem (Parite(3))
    Parite(4) = "S"
    Configuration.Combo2.AddItem (Parite(4))
    
    Rem Le nombre de bits de donnees de la transmission
    Bit_Donnee(0) = "4"
    Configuration.Combo3.AddItem (Bit_Donnee(0))
    Bit_Donnee(1) = "5"
    Configuration.Combo3.AddItem (Bit_Donnee(1))
    Bit_Donnee(2) = "6"
    Configuration.Combo3.AddItem (Bit_Donnee(2))
    Bit_Donnee(3) = "7"
    Configuration.Combo3.AddItem (Bit_Donnee(3))
    Bit_Donnee(4) = "8"
    Configuration.Combo3.AddItem (Bit_Donnee(4))
    
    Rem Le nombre de bits d'arrêt de la transmission
    Bit_Arret(0) = "1"
    Configuration.Combo4.AddItem (Bit_Arret(0))
    Bit_Arret(1) = "1.5"
    Configuration.Combo4.AddItem (Bit_Arret(1))
    Bit_Arret(2) = "2"
    Configuration.Combo4.AddItem (Bit_Arret(2))
    
    Rem Le flux de communication
    Flux(0) = "comNone"
    Configuration.Combo5.AddItem (Flux(0))
    Flux(1) = "comXOnXOff"
    Configuration.Combo5.AddItem (Flux(1))
    Flux(2) = "comRTS"
    Configuration.Combo5.AddItem (Flux(2))
    Flux(3) = "comRTSXOnXOff"
    Configuration.Combo5.AddItem (Flux(3))

    'Lance l'écran de configuration
    Fin_Application = False
    Configuration.Show
    Configuration.Affiche_Parametres
End Sub
