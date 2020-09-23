Attribute VB_Name = "modVariables"
Option Explicit

Private Type typeaDriveList
    DriveName           As String               ' Le nom du disque, exemple C:
    LastPerformance     As DISK_PERFORMANCE     ' Les dernières données pour faire les calculs
    ReadOperations      As Integer              ' Dernière valeur du delta-ReadCount
    WriteOperations     As Integer              ' Dernière valeur du delta-WriteCount
End Type

Public ReadMaxOperations    As Long             ' Valeur la plus forte récupérée (pour mise à l'échelle)
Public WriteMaxOperations   As Long             ' Valeur la plus forte récupérée (pour mise à l'échelle)
Public aDriveList()         As typeaDriveList   ' Collection des drives scannés
Public OffSet               As Integer          ' Echelle d'une barre d'un bargraphe
Public PremierCalculNonNull As Boolean          ' Permet d'afficher l'icone programme en attendant un PremierCalculNonNull
