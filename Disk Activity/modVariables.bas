Attribute VB_Name = "modVariables"
Option Explicit

Private Type typeaDriveList
    DriveName           As String               ' Le nom du disque, exemple C:
    LastPerformance     As DISK_PERFORMANCE     ' Les derni�res donn�es pour faire les calculs
    ReadOperations      As Integer              ' Derni�re valeur du delta-ReadCount
    WriteOperations     As Integer              ' Derni�re valeur du delta-WriteCount
End Type

Public ReadMaxOperations    As Long             ' Valeur la plus forte r�cup�r�e (pour mise � l'�chelle)
Public WriteMaxOperations   As Long             ' Valeur la plus forte r�cup�r�e (pour mise � l'�chelle)
Public aDriveList()         As typeaDriveList   ' Collection des drives scann�s
Public OffSet               As Integer          ' Echelle d'une barre d'un bargraphe
Public PremierCalculNonNull As Boolean          ' Permet d'afficher l'icone programme en attendant un PremierCalculNonNull
