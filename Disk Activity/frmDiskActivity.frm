VERSION 5.00
Begin VB.Form frmDiskActivity 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   630
   ClientLeft      =   180
   ClientTop       =   12300
   ClientWidth     =   4725
   ControlBox      =   0   'False
   Icon            =   "frmDiskActivity.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   42
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   315
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2160
      Picture         =   "frmDiskActivity.frx":0CCA
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picTravail 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   2
      Enabled         =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   3360
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picVide 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF00FF&
      ForeColor       =   &H00FF00FF&
      Height          =   480
      Left            =   2760
      Picture         =   "frmDiskActivity.frx":1994
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picBase 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   1560
      Picture         =   "frmDiskActivity.frx":26DE
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer timerM�J 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   1080
      Top             =   0
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   143
      Width           =   135
   End
   Begin VB.Image imgDA 
      Height          =   240
      Index           =   0
      Left            =   600
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Menu mnuPrincipal 
      Caption         =   "Codes-Sources"
      Visible         =   0   'False
      Begin VB.Menu mnuMasquer 
         Caption         =   "&Masquer les d�tails"
      End
      Begin VB.Menu mnuRunAtStartUp 
         Caption         =   "Lancer DiskActivity au d�marrage de Windows"
      End
      Begin VB.Menu mnuS�parateur0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "frmDiskActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================================================
' Titre  : DiskActivity
' Auteur : Jack
' Source : http://www.vbfrance.com/code.aspx?ID=37086
'=======================================================================================


Option Explicit

' ## D�clarations pour assurer le d�placement de la forme sans caption � la souris
' Voir http://www.vbfrance.com/codes/DEPLACER-FEUILLE-SANS-CAPTION_16982.aspx
' D�finition de coordonn�es d'un objet
Private Type POINTAPI
    x As Long
    y As Long
End Type
' D�finition de position et taille d'un objet
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Pour r�cup�rer la position souris (en coordonn�e �cran)
Private Declare Function GetCursorPos Lib "user32" ( _
                    lpPoint As POINTAPI) As Long
' Pour d�placer la feuille (en coordonn�e �cran)
Private Declare Function MoveWindow Lib "user32" ( _
                    ByVal hWnd As Long, _
                    ByVal x As Long, _
                    ByVal y As Long, _
                    ByVal nWidth As Long, _
                    ByVal nHeight As Long, _
                    ByVal bRepaint As Long) As Long
' Pour conna�tre la position de la feuille (en coordonn�e �cran)
Private Declare Function GetWindowRect Lib "user32" ( _
                    ByVal hWnd As Long, _
                    lpRect As RECT) As Long

' Nos variables de m�moire de position
Private D�placementEnCours As Boolean
Private Coord              As POINTAPI
Private TailleFeuille      As RECT
'

Private Sub Form_Load()
    
    Dim Temp As String, bRet As Boolean
    
    ' Initialisation
    Me.ScaleMode = vbPixels ' facilite la gestion des Images
    Call SetTop(Me, True)   ' Notre forme sera toujours visible
    mnuMasquer.Tag = 0
    mnuRunAtStartUp.Checked = IIf(WillRunAtStartup(App.EXEName) = True, vbChecked, vbUnchecked)
    
    OffSet = 100 / 32 ' D�calage de chaque barre du bargraphe sur une base 100% et image de 32 pixels
    picTravail.BackColor = vbMagenta ' D�fini le fond transparent
    ' Initialisation du tableau des caract�ristiques
    ReDim aDriveList(0)
    ReadMaxOperations = 400    ' valeurs de base pour ne pa que l'affichage
    WriteMaxOperations = 400   '   s'affole les premi�res minutes
    
    ' R�cup�re les param�tres enregistr�s dans la base de registres
    Temp = GetSetting("Codes-Sources", App.EXEName, "Position fen�tre", CStr(Screen.Width / 2) & ";" & CStr(Screen.Height / 2))
    Me.Move Split(Temp, ";")(0), Split(Temp, ";")(1)
    bRet = GetSetting("Codes-Sources", App.EXEName, "D�tails masqu�s ?", False)
    If bRet Then mnuMasquer_Click   ' car par d�faut, pas coch�e
    
    ' Cr�e une ic�ne dans le SysTray
    PremierCalculNonNull = False
    With TrayIcon
        .cbSize = Len(TrayIcon)             ' make the tray icon
        .hWnd = Me.hWnd                     ' Handle of the window used to handle messages
        .uId = vbNull                       ' ID code of the icon
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE ' Flags
        .ucallbackMessage = WM_MOUSEMOVE    ' ID of the call back message
        .hIcon = picLogo.Picture            ' The start icon
        .szTip = "DiskActivity par Jack - Codes Sources" & Chr$(0) ' The Tooltip for the icon
    End With
    ' Add icon to the tray
    Call Shell_NotifyIcon(NIM_ADD, TrayIcon)
    
    ' D�marre la proc�dure de hooking pour notre forme pour le Magn�tisme des formes
    ' ### Si vous voulez faire du debuggage, mettez cette ligne en commentaire
    '     car le hooking emp�che d'acc�der au feuilles de code
    ' Ici, on ne lance pas le hook si on est en mode IDE
    ' ???????????? Je viens de m'apercevoir que ce Magn�tisme ne fonctionne pas si
    '              la forme n'a pas de Caption - Dommage
    '              En fait, une forme sans Caption ne g�n�re pas d'�v�nement WM_ENTERSIZEMOVE
    '              Si vous trouvez une astuce ...
    'DockingStart Me, [Aimantable � toutes les formes du bureau]
    
    ' Recherche tous les disques durs locaux
    Call ListAllDrives
    ' Cr�� autant de composant que de disque
    Call CreateComposants

    ' On peut lancer la surveillance
    timerM�J.Enabled = True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Info : Echelle de la forme en Pixels (pas en Twips)
    If (Button And vbRightButton) Then
        PopupMenu mnuPrincipal, vbPopupMenuRightAlign, , , mnuQuitter
    
    ' Si on appuie sur le bouton gauche et
    ' qu'on n'est pas d�j� en cours de d�placement
    ElseIf (Button And vbLeftButton) And Not D�placementEnCours Then
        ' On est en train de faire un Double-Click --> Pas de recherche de la position de la forme
        If (x <> WM_LBUTTONDBLCLK) And (Not mnuPrincipal.Visible) Then
            ' On m�morise
            D�placementEnCours = True
            ' On r�cup�re la position initiale de la souris
            Call GetCursorPos(Coord)
            ' et les positions et dimensions initiales de notre feuille
            Call GetWindowRect(Me.hWnd, TailleFeuille)
        End If
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Info : Echelle de la forme en Pixels (pas en Twips)
    Static Occup� As Boolean
    
    ' Si on est en cours de d�placement avec le bouton gauche
    If (Button And vbLeftButton) And D�placementEnCours Then
        ' Dimensionne notre variable souris
        Dim NewCoord As POINTAPI
        ' R�cup�re nouvelle position de la souris
        Call GetCursorPos(NewCoord)
        ' D�place notre feuille � la nouvelle position
        Call MoveWindow(Me.hWnd, _
                        TailleFeuille.Left + NewCoord.x - Coord.x, _
                        TailleFeuille.Top + NewCoord.y - Coord.y, _
                        TailleFeuille.Right - TailleFeuille.Left, _
                        TailleFeuille.Bottom - TailleFeuille.Top, _
                        True)
        ' Laisse le temps � Windows de g�rer le graphisme
        DoEvents
        Exit Sub
    End If
        
    ' On fait un Click sur la forme ?
    If Occup� = False Then
        Occup� = True
        Select Case x
            Case WM_LBUTTONDBLCLK   ' Double-Click gauche
                mnuMasquer.Tag = -1
                Call mnuMasquer_Click
            Case WM_RBUTTONUP       ' Click-Droit
                PopupMenu mnuPrincipal, vbPopupMenuRightAlign, , , mnuMasquer
        End Select
        Occup� = False
    End If
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Si on relache la souris, on remet � z�ro notre m�moire
    If (Button And vbLeftButton) And D�placementEnCours Then
            D�placementEnCours = False
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim r As Integer
    
    ' Stoppe les scrutations
    timerM�J.Enabled = False
    
    ' M�morise l'emplacement de la fen�tre pour le prochain red�marrage
    ' Les donn�es sont stock�es dans la base de registres � cet endroit :
    '   HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Codes-Sources\DiskActivity
    SaveSetting "Codes-Sources", App.EXEName, "Position fen�tre", CStr(Me.Left) & ";" & CStr(Me.Top)
    SaveSetting "Codes-Sources", App.EXEName, "D�tails masqu�s ?", Str(mnuMasquer.Tag)
    
    ' Demande de stopper le hooking de notre forme
    DockingTerminate Me
    
    ' D�truit les composants charg�s (sauf l'original)
    For r = lblDrive.Count To 2 Step -1
        Unload lblDrive(r - 1)
    Next r
    For r = imgDA.Count To 2 Step -1
        Unload imgDA(r - 1)
    Next r
    
    ' D�monte l'ic�ne du Systray
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hWnd = Me.hWnd
    TrayIcon.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, TrayIcon)
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmDiskActivity = Nothing
    
End Sub

' D�placements quand on clique sur une des Images
Private Sub imgDA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Renvoie � la feuille les evenements du (seul) controle
    Call Form_MouseDown(Button, Shift, x, y)
End Sub
Private Sub imgDA_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Idem pour les Move
    Call Form_MouseMove(Button, Shift, x, y)
End Sub
Private Sub imgDA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Idem pour le Up
    Call Form_MouseUp(Button, Shift, x, y)
End Sub

' D�placements quand on clique sur un des Labels
Private Sub lblDrive_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Renvoie � la feuille les evenements du (seul) controle
    Call Form_MouseDown(Button, Shift, x, y)
End Sub
Private Sub lblDrive_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Idem pour les Move
    Call Form_MouseMove(Button, Shift, x, y)
End Sub
Private Sub lblDrive_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Idem pour le Up
    Call Form_MouseUp(Button, Shift, x, y)
End Sub

Private Sub mnuMasquer_Click()

    mnuMasquer.Tag = Not mnuMasquer.Tag
    If mnuMasquer.Tag Then
        Me.Hide
        mnuMasquer.Caption = "&Voir les d�tails"
    Else
        Me.Show
        Me.WindowState = vbNormal
        mnuMasquer.Caption = "&Masquer les d�tails"
    End If
        
End Sub

Private Sub mnuQuitter_Click()

    Unload Me
    
End Sub

Private Sub mnuRunAtStartUp_Click()

    ' D�marrera l'application au d�marrage de la session Windows si le menu est coch�
    mnuRunAtStartUp.Checked = Not mnuRunAtStartUp.Checked
    If mnuRunAtStartUp.Checked Then
        If Not WillRunAtStartup(App.EXEName) Then
            Call SetRunAtStartup(App.EXEName, App.Path, True)
        End If
    Else
        If WillRunAtStartup(App.EXEName) Then
            Call SetRunAtStartup(App.EXEName, App.Path, False)
        End If
    End If
    
End Sub

Private Sub timerM�J_Timer()
    
    Dim r As Integer
    
'    Debug.Print "-----------------------------"
    ' Recherche les infos d'activit� des disques
    For r = 0 To UBound(aDriveList)
        Call ScanDrives(r)
    Next r
    
    ' Fabrique les images
    ' Les disques sont g�r�s en sens inverse afin de d�tecter le 0 = le dernier
    ' Au dernier passage, on cr�e une icone de plus pour le Systray
    For r = UBound(aDriveList) To 0 Step -1
        Call CreateImages(r)
    Next r
    DoEvents
    
End Sub


Private Sub CreateComposants()
    ' Charge les couples Label-Image pour chaque disque
    ' Rappel : la forme est dimentionn�e en Pixels, pas en Twips (variables � virgule)
    
    Dim r As Integer
    Dim LargeurCouple As Single
    
    LargeurCouple = lblDrive(0).Width + imgDA(0).Width + 16
    
    ' 1er couple : Composants de base
    lblDrive(0).Caption = Left$(aDriveList(0).DriveName, 1)
    lblDrive(0).Move 2, 5
    imgDA(0).Move lblDrive(0).Left + lblDrive(0).Width + 2, 3
    
    ' les couples suivants
    For r = 1 To UBound(aDriveList)
        ' Si le Label n'existe pas, on le cr�� et on le positionne
        If lblDrive.UBound < (r + 1) Then Load lblDrive(r)
        lblDrive(r).Caption = Left$(aDriveList(r).DriveName, 1)
        lblDrive(r).Move lblDrive(r - 1).Left + LargeurCouple, lblDrive(0).Top
        ' Si l'Image n'existe pas, on la cr�� et on la positionne
        If imgDA.Count < (r + 1) Then Load imgDA(r)
        imgDA(r).Move lblDrive(r).Left + lblDrive(r).Width + 2, imgDA(0).Top
        ' Rend les deux composants visibles
        lblDrive(r).Visible = True
        imgDA(r).Visible = True
    Next r
    
    ' D�finition de la taille de la forme
    Me.Width = (imgDA(imgDA.UBound).Left + imgDA(imgDA.UBound).Width + 8) * Screen.TwipsPerPixelX
    Me.Height = (imgDA(0).Top + imgDA(0).Height + 5) * Screen.TwipsPerPixelY
    Me.Refresh

End Sub

Private Sub CreateImages(ByVal iDriveIndex As Integer)
    ' G�n�re une icone dont les bargraphes rep�sentent l'activit� du disque
    '   Voir http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=64964&lngWId=1
    ' iDrive d�termine le drive � traiter et g�n�re l'image dans imgDA() associ�
    
    ' Les disques sont scann�s du dernier vers le premier :
    '   Durant les appels des disques, on a m�moris� lequel lit et �crit le plus
    '   Une fois qu'on sera � l'index 0 (le dernier), on mettra � jour l'ic�ne du Systray
    '   pour qu'elle soit l'image de l'activit� globale de tous les disques
    
    Static iMaxRead As Integer, iMaxWrite As Integer
    
    Dim hIcon As Long
    Dim IconPic As StdPicture
    
    picTravail.Picture = picBase.Picture ' Image de base
    
    ' Donn�es c�t� Lecture
    Call BitBlt(picTravail.hDC, 0, 0, 16, 32 - (aDriveList(iDriveIndex).ReadOperations / OffSet), picVide.hDC, 0, 0, vbSrcCopy)
    ' Donn�es c�t� Ecriture
    Call BitBlt(picTravail.hDC, 16, 0, 16, 32 - (aDriveList(iDriveIndex).WriteOperations / OffSet), picVide.hDC, 0, 0, vbSrcCopy)
    ' Redessine l'icone
    picTravail.Picture = picTravail.Image
    
    ' Transforme le Magenta en transparence
    hIcon = BitmapToIcon(picTravail.Picture.handle, vbMagenta)
    Set IconPic = GDIToPicture(hIcon)
    If (IconPic Is Nothing) Then
        ' Lib�re le handle si la cr�ation a �chou� (resources)
        Call DestroyIcon(hIcon)
    Else ' Attribue notre image au composant index� final
        Set imgDA(iDriveIndex) = GDIToPicture(hIcon)
        imgDA(iDriveIndex).ToolTipText = aDriveList(iDriveIndex).DriveName & "  " & _
                                         "Lecture " & CStr(aDriveList(iDriveIndex).ReadOperations) & "%, " & _
                                         "Ecriture " & CStr(aDriveList(iDriveIndex).WriteOperations) & "%"
    End If

    ' M�morise les Max
    If aDriveList(iDriveIndex).ReadOperations > iMaxRead Then iMaxRead = aDriveList(iDriveIndex).ReadOperations
    If aDriveList(iDriveIndex).WriteOperations > iMaxWrite Then iMaxWrite = aDriveList(iDriveIndex).WriteOperations
    
    ' S'agit-il du dernier disque ?
    If iDriveIndex = 0 Then
        Call CreateSystrayIcon(iMaxRead, iMaxWrite)
'Debug.Print iMaxRead, iMaxWrite, , ReadMaxOperations, WriteMaxOperations
        ' Remet � z�ro les compteurs
        iMaxRead = 0
        iMaxWrite = 0
        DoEvents
    End If

End Sub

Private Sub CreateSystrayIcon(ByVal ReadVal As Integer, _
                              ByVal WriteVal As Integer)

    ' A peu de chose pr�s, la m�me proc�dure que dans CreateImages
    
    Static Compteur As Integer
    
    Dim hIcon As Long
    Dim IconPic As StdPicture
    
    picTravail.Picture = picBase.Picture ' Image de base
    
    If ReadVal = 0 And WriteVal = 0 Then
        ' pas d'activit� : Incr�mente le compteur
        Compteur = Compteur + 1
    Else
        ' Sinon, RaZ du compteur
        Compteur = 0
        ' Ca y est, on a des donn�es � afficher pour la 2ere fois
        PremierCalculNonNull = True
    End If
    If Not PremierCalculNonNull Or Compteur > 5 Then
        ' Plusieurs cycle qu'on n'a pas d'activit� --> Affiche le logo
        picTravail.Picture = picLogo.Picture
    Else
        ' Donn�es c�t� Lecture
        Call BitBlt(picTravail.hDC, 0, 0, 16, 32 - (ReadVal / OffSet), picVide.hDC, 0, 0, vbSrcCopy)
        ' Donn�es c�t� Ecriture
        Call BitBlt(picTravail.hDC, 16, 0, 16, 32 - (WriteVal / OffSet), picVide.hDC, 0, 0, vbSrcCopy)
    End If
    ' Redessine l'icone
    picTravail.Picture = picTravail.Image
    
    ' Transforme le Magenta en transparence
    hIcon = BitmapToIcon(picTravail.Picture.handle, vbMagenta)
    Set IconPic = GDIToPicture(hIcon)
    If (IconPic Is Nothing) Then
        ' Lib�re le handle si la cr�ation a �chou� (resources)
        Call DestroyIcon(hIcon)
    Else ' Attribue notre image � l'ic�ne du SysTray
        TrayIcon.hIcon = IconPic.handle
        Shell_NotifyIcon NIM_MODIFY, TrayIcon
    End If
    
End Sub

' Renvoie -1 (True) si on est en mode IDE, ou renvoie 0 (False) sur on est en mode Compil�
Private Function InIDE() As Long
    Static Value As Long
    If Value = 0 Then
        Value = 1
        Debug.Assert (True Or InIDE())  ' Cette ligne n'existe pas en mode Compil�
        InIDE = Value - 1
    End If
    Value = 0
End Function

