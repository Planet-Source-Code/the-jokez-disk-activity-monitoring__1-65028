Attribute VB_Name = "modDiskPerformance"
Option Explicit

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
                                ByVal lpFileName As String, _
                                ByVal dwDesiredAccess As Long, _
                                ByVal dwShareMode As Long, _
                                lpSecurityAttributes As Any, _
                                ByVal dwCreationDisposition As Long, _
                                ByVal dwFlagsAndAttributes As Long, _
                                ByVal hTemplateFile As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" ( _
                                ByVal hDevice As Long, _
                                ByVal dwIoControlCode As Long, _
                                lpInBuffer As Any, _
                                ByVal nInBufferSize As Long, _
                                lpOutBuffer As Any, _
                                ByVal nOutBufferSize As Long, _
                                lpBytesReturned As Long, _
                                lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" ( _
                                ByVal nBufferLength As Long, _
                                ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const INVALID_HANDLE_VALUE = -1
Private Const OPEN_EXISTING = 3
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const DRIVE_FIXED = 3

Private Const IOCTL_DISK_PERFORMANCE = &H70020
' IOCTL_DISK_PERFORMANCE est un entier de 32 bits décomposés en 4 zones :
' Voir http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcehardware5/html/wce50lrfctlcode.asp
'   et pour les valeurs de constantes --> http://www.allapi.net --> ApiViewer
' ### Méfiez-vous de ne pas trop jouer avec ces paramètres, ça risque de bloquer la machine ###
' Bits 0 et 1 (2) : Method
'                        METHOD_BUFFERED   = 00 (0)
'                        METHOD_IN_DIRECT  = 01 (1)
'                        METHOD_OUT_DIRECT = 10 (2)
'                        METHOD_NEITHER    = 11 (3)
' Bits 2 à 13 (12) : Fonction
'                        Pas eu d'explication. Même dans le fichier WinIOctl.h, la valeur 8 est donnée sans raison
'                        Voir http://www.reactos.org/generated/doxygen/d9/d47/winioctl_8h-source.html
'                        = 0000 0000 1000 (8)
' Bits 14 à 15 (2) : Acces
'                        FILE_ANY_ACCESS   = 00 (0)
'                        FILE_READ_ACCESS  = 01 (1)
'                        FILE_WRITE_ACCESS = 10 (2)
' Bits 16 à 31 (16) : DeviceType
'                        Voir longue liste. Nous, c'est FILE_DEVICE_DISK = 7
'                        Devant, les 8 derniers bits sont à 1 --> 1111 1111 0000 0111 (7)
' Résultat (de 31 à 0) : 0000 0000 0000 0111 0000 0000 0010 0000 = &h70020 = 458784
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2

Public Type DISK_PERFORMANCE   ' Currency = LARGE_INTEGER en C (8 bytes)
    BytesRead           As Currency ' The number of bytes read.
    BytesWritten        As Currency ' The number of bytes written.
    ReadTime            As Currency ' The time it takes to complete a read.
    WriteTime           As Currency ' The time it takes to complete a write.
    IdleTime            As Currency ' The idle time.
    ReadCount           As Long     ' The number of read operations.
    WriteCount          As Long     ' The number of write operations.
    QueueDepth          As Long     ' The depth of the queue.
    SplitCount          As Long     ' The cumulative count of I/Os that are associated I/Os.
                                    ' An associated I/O is a fragmented I/O, where multiple I/Os
                                    '   to a disk are required to fulfill the original logical I/O request.
                                    ' The most common example of this scenario is a file that is fragmented
                                    '   on a disk. The multiple I/Os are counted as split I/O counts.
    QueryTime           As Currency ' The system time stamp when a query for this structure is returned.
                                    ' Use this member to synchronize between the file system driver and a caller.
    StorageDeviceNumber As Long     ' The unique number for a device that identifies it to the storage manager
                                    '   that is indicated in the StorageManagerName member.
    StorageManagerName  As String * 8 ' The name of the storage manager that controls this device.
                                      ' Examples of storage managers are "PhysDisk," "FTDISK," and "DMIO".
    ReadCount2          As Long ' \
    WriteCount2         As Long ' | Non documentés correctement. Sur MSDN, ces trois valeurs manquent
    QueueDepth2         As Long ' /
End Type
'
Public Sub ListAllDrives()
    
    Dim sAllDrives As String
    Dim sTemp As String
    Dim lRet As Long
    Dim iDriveIndex As Integer
    
    iDriveIndex = -1
    ' Récupère la liste des drives, séparés par des vbNull
    sAllDrives = fGetDrives
    Do While sAllDrives <> ""
        ' Isole le drive à gauche
        sTemp = Mid$(sAllDrives, 1, InStr(sAllDrives, vbNullChar) - 1)
        ' Le supprime de la liste
        sAllDrives = Mid$(sAllDrives, InStr(sAllDrives, vbNullChar) + 1)
        ' Récupère le type de drive
        lRet = GetDriveType(sTemp)
        ' On ne le prend en compte que si c'est un disque local
        If lRet = DRIVE_FIXED Then
            ' Incrémente le nombre de drive trouvé
            iDriveIndex = iDriveIndex + 1
            ' Redimensionne le tableau
            ReDim Preserve aDriveList(iDriveIndex)
            ' Mémorise la lettre du drive
            aDriveList(iDriveIndex).DriveName = Left$(sTemp, 1) & ":"
        End If
    Loop

End Sub

Public Sub ScanDrives(ByVal iDriveIndex As Integer)

    ' Lance la demande d'infos et fait les calculs
    
    Dim Valeurs      As DISK_PERFORMANCE    ' Valeurs collectées
    Dim cElapseTime  As Currency            ' Temps écoulé depuis dernière lecture
    Dim cBytesRead   As Currency            ' Nombre de Bytes échangés
    Dim cBytesWrite  As Currency            ' Nombre de Bytes échangés
    Dim iPourcentage As Integer             ' Calcul du % de charge
    
    ' Récupère les données
    Valeurs = LecturePerformance(aDriveList(iDriveIndex).DriveName)
    ' Pas de calcul si on n'a pas déjà fait un cycle
    If aDriveList(iDriveIndex).LastPerformance.StorageDeviceNumber = 0 Then
        ' Mémorise les infos pour le prochain cycle
        aDriveList(iDriveIndex).LastPerformance = Valeurs
        ' Ressort
        Exit Sub
    End If
'    ' Je vous ai laissé la liste pour faire mumuse
'    With Valeurs
'        Debug.Print "------------------------------------------"
'        Debug.Print "BytesRead              "; .BytesRead
'        Debug.Print "BytesWritten           "; .BytesWritten
'        Debug.Print "ReadTime               "; .ReadTime
'        Debug.Print "WriteTime              "; .WriteTime
'        Debug.Print "IdleTime               "; .IdleTime
'        Debug.Print "ReadCount              "; .ReadCount
'        Debug.Print "WriteCount             "; .WriteCount
'        Debug.Print "QueueDepth             "; .QueueDepth
'        Debug.Print "SplitCount             "; .SplitCount
'        Debug.Print "QueryTime              "; .QueryTime
'        Debug.Print "StorageDeviceNumber    "; .StorageDeviceNumber
'        Debug.Print "StorageManagerName      "; .StorageManagerName
'        Debug.Print "ReadCount2             "; .ReadCount2
'        Debug.Print "WriteCount2            "; .WriteCount2
'        Debug.Print "QueueDepth2            "; .QueueDepth2
'    End With
    
    ' Calcule le temps écoulé depuis la dernière lecture
    ' QueryTime nous fourni un TimeStamp qui s'incrémente de 100 nanoSecondes depuis
    '   le 1er Janvier 1601 (drôle de référence !) mais n'est rafraichit que toutes
    '   les 10 milliSecondes
    ' Donc, en faisant la différence entre Dernière et AvantDernière valeur, on récupère un
    '   chiffre à virgule en milliSecondes
    cElapseTime = Valeurs.QueryTime - aDriveList(iDriveIndex).LastPerformance.QueryTime
    
    ' Calcule le nombre de Bytes d'échanges en Lecture
'    cBytesRead = Valeurs.BytesRead - aDriveList(iDriveIndex).LastPerformance.BytesRead
    cBytesRead = Valeurs.ReadCount - aDriveList(iDriveIndex).LastPerformance.ReadCount
    ' Puis calcule ce nombre par secondes
    cBytesRead = cBytesRead / cElapseTime * 1000@
    ' Mémorise si le record du nombre de Bytes échangés en Lecture est atteint
    '    Cette mémoire permet de faire la mise à l'échelle
    If cBytesRead > ReadMaxOperations Then ReadMaxOperations = cBytesRead
    ' En se basant sur le record maxi de Lecture, on va calculer le pourcentage (en entier, ça suffit)
    If ReadMaxOperations > 0 Then
        iPourcentage = CInt(100@ / ReadMaxOperations * cBytesRead)
        If iPourcentage > 100 Then iPourcentage = 100
        aDriveList(iDriveIndex).ReadOperations = iPourcentage
    End If
    
    ' Calcule le nombre de Bytes d'échanges en Ecriture
'    cBytesWrite = Valeurs.BytesWritten - aDriveList(iDriveIndex).LastPerformance.BytesWritten
    cBytesWrite = Valeurs.WriteCount - aDriveList(iDriveIndex).LastPerformance.WriteCount
    ' Puis calcule ce nombre par secondes
    cBytesWrite = cBytesWrite / cElapseTime * 1000@
    ' Mémorise si le record du nombre de Bytes échangés en Ecriture est atteint
    '    Cette mémoire permet de faire la mise à l'échelle
    If cBytesWrite > WriteMaxOperations Then WriteMaxOperations = cBytesWrite
    ' En se basant sur le record maxi de Lecture, on va calculer le pourcentage (en entier, ça suffit)
    If WriteMaxOperations > 0 Then
        iPourcentage = CInt(100@ / WriteMaxOperations * cBytesWrite)
        If iPourcentage > 100 Then iPourcentage = 100
        aDriveList(iDriveIndex).WriteOperations = iPourcentage
    End If
    
    ' Mémorise les infos pour le prochain cycle
    aDriveList(iDriveIndex).LastPerformance = Valeurs

 'Debug.Print Time, aDriveList(iDriveIndex).DriveName; "("; aDriveList(iDriveIndex).LastPerformance.StorageDeviceNumber; ")", _
             cElapseTime; "mS", cBytesRead; "(r/s)", cBytesWrite; "(w/s)", _
             aDriveList(iDriveIndex).ReadOperations; "%r", aDriveList(iDriveIndex).WriteOperations; "%w", _
             ReadMaxOperations; "(MaxR)", WriteMaxOperations; "(MaxW)"
    
End Sub

Private Function LecturePerformance(ByVal theDrive As String) As DISK_PERFORMANCE
    
    ' theDrive doit être une lettre représentant un disque physique local
    ' Si l'unité n'est pas accessible, la variable QueryTime renvoie -1
    
    Dim hDrive As Long
    Dim DummyReturnedBytes As Long
    Dim Resultat As DISK_PERFORMANCE
    Dim Ret As Long
    
    theDrive = UCase(theDrive)
    ' On se connecte au drive et récupère son handle
    hDrive = CreateFile("\\.\" & theDrive, _
                        GENERIC_READ Or GENERIC_WRITE, _
                        FILE_SHARE_READ + FILE_SHARE_WRITE, _
                        ByVal 0, _
                        OPEN_EXISTING, _
                        0, _
                        0)
    ' Si on l'a bien trouvé, demande les infos performance
    If hDrive <> INVALID_HANDLE_VALUE Then
        Ret = DeviceIoControl(hDrive, _
                              IOCTL_DISK_PERFORMANCE, _
                              ByVal 0, _
                              0, _
                              Resultat, _
                              Len(Resultat), _
                              DummyReturnedBytes, _
                              ByVal 0)
        Call CloseHandle(hDrive)  ' Déconnecte
        If Ret = 0 Then
            ' La fonction a échoué : Signale le problème
            Resultat.QueryTime = -1
        End If
    Else
        ' Le lecteur n'est pas accessible : Signale le problème
        Resultat.QueryTime = -1
    End If
    LecturePerformance = Resultat   ' Transmet les valeurs

End Function

Private Function fGetDrives() As String
    ' Renvoie toutes les unités de disque mappées
    ' Toutes les lettres de disques sont séparées par des vbNull
    
    Dim lRet As Long
    Dim lTemp As Long
    Dim sDrives As String * 255
    
    lTemp = Len(sDrives)
    lRet = GetLogicalDriveStrings(lTemp, sDrives)
    fGetDrives = Left(sDrives, lRet)

End Function
