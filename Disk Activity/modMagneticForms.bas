Attribute VB_Name = "modMagneticForm"
' Magnetische Forms Modul.
' Copyright (C) 2001 Benjamin Wilger
' Benjamin@ActiveVB.de

'----------------------------------------------------------
' Traduit et aménagé par Jack, www.vbfrance.com
' Source : http://www.vbfrance.com/code.aspx?ID=27194
' Honnetement, je n'ai pas détaillé le code en profondeur.
' J'ai mis des commentaires ... mais manque de temps ...
'----------------------------------------------------------
' J'ai aussi ajouté un gadget qui permet de nous aimanter
' à toutes les formes ouvertes, pas seulement celles de
' notre application (paramètre supplémentaire dans la
' variable DockingLog, individualisable par forme, et ajout
' de la procédure EnumToutesFenêtres)
'----------------------------------------------------------

' Utilisation :
' =============
' Ajoutez ce module à votre application
' Dans votre forme, dans le Form_Load par exemple, démarrer l'aimantation avec :
'         DockingStart Me, [Aimantable à toutes les formes du bureau]
'     ou  DockingStart Me, [Aimantable aux formes de mon Appli]
' Pour stopper l'aimantation ou avant de refermer votre forme, dans Form_Unload) :
'         DockingTerminate Me

' ===========
' Important :
' ===========
' N'ARRÊTEZ SURTOUT PAS L'APPLI AVEC LE BOUTON "STOP" DE l'EDITEUR VB
'   ---> Vous auriez un beau crash de l'éditeur !!

' Car, avec le Hooking, on créé un nouveau process et on dit à Windows
' d'envoyer les évenements sur ce nouveau process. Comme on redonne ces
' évènements au programme d'origine (et des évènements, il s'en produit
' trois ou quatre à chaque fois que vous touchez à la souris !), si ce
' programme est fermé sans dire au process hooké qu'il se ferme, ce dernier
' enverra les évènements à un process qui n'existe plus, et ça, Windows
' il aime pas du tout !
' Le problème est le même en mode compilé : vous aurez une alerte système.
'
' Pour connaître la bonne manière de démonter le hook, regardez la
' procédure Form_QueryUnload
'

Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function SystemParametersInfo_Rect Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

' Api pour la recherche des formes présentes sur le bureau
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Const WM_MOVING         As Long = &H216
Private Const WM_SIZING         As Long = &H214
Private Const WM_ENTERSIZEMOVE  As Long = &H231

Private Const GWL_WNDPROC       As Long = (-4)

Private Const SPI_GETWORKAREA   As Long = 48

Private Const WMSZ_LEFT         As Long = 1
Private Const WMSZ_TOPLEFT      As Long = 4
Private Const WMSZ_BOTTOMLEFT   As Long = 7
Private Const WMSZ_TOP          As Long = 3
Private Const WMSZ_TOPRIGHT     As Long = 5

' Les modes des formes
'----------------------
Private Enum SnapFormMode
    Moving = 1
    Sizing = 2
End Enum

Public Enum AimantationType
    [Aimantable aux formes de mon Appli] = 0
    [Aimantable à toutes les formes du bureau] = 1
End Enum

' Les caractéristiques des formes qui demandent à être aimantées
Private Type DockingLog
    hWnd        As Long             ' le handle
    oldProc     As Long             ' la référence du process
    Aimantation As AimantationType  ' Voir au dessus
End Type

Private Logs() As DockingLog, LogCount As Integer, MaxLogs As Integer

Private MouseX As Long, MouseY As Long
Private SnappedX As Boolean, SnappedY As Boolean
Private Rects() As RECT ' Collection des formes auxquelles on va essayer de s'aimanter
Private hWnd_à_exclure As Long

' Distance mini avant aimantation des formes
Private Const SnapWidth = 15 ' pixels

' Puisque le Subclassing rend généralement le Debugging impossible,
'   il peut être mis hors circuit ici
'   (dans le cas où vous utiliseriez ce module dans un autre projet)
Private Const DoSubClass As Boolean = True
'

' Active le Hook sur cette forme
Public Sub DockingStart(f As Form, _
                        ByVal ModeAimantation As AimantationType)
    
    Dim H As Long, t As Integer
    
    If Not DoSubClass Then Exit Sub
    
    ' Redimensionne le tableau du stock s'il devient trop petit
    If LogCount + 10 > MaxLogs Then
        MaxLogs = LogCount + 10
        ReDim Preserve Logs(MaxLogs)
    End If
    
    ' Vérifie que cette forme n'est pas déjà hookée
    For t = 0 To LogCount - 1
        If Logs(t).hWnd = f.hWnd Then
            ' Elle existe déja. On ressort
            Exit Sub
        End If
    Next t

    ' Démarre le hooking de la forme
    H = f.hWnd
    Logs(LogCount).hWnd = H
    Logs(LogCount).oldProc = SetWindowLong(H, GWL_WNDPROC, AddressOf WindowProc)
    Logs(LogCount).Aimantation = ModeAimantation
    ' Incrémente le compteur de fenêtres gérées
    LogCount = LogCount + 1
    
End Sub

' On supprime le Hook (appelé avant de fermer la forme,
'   sinon, risque de crash)
Public Sub DockingTerminate(f As Form)
    
    Dim t As Integer, H As Long
    
    H = f.hWnd
    ' Recherche notre forme dans le stock
    For t = 0 To LogCount - 1
        If Logs(t).hWnd = H Then
            ' On l'a trouvée
            ' Remet l'ancienne référence du process
            SetWindowLong H, GWL_WNDPROC, Logs(t).oldProc
            ' Décale le stock qui suit pour garder
            '   une liste sans 'blanc'
            For H = t To LogCount - 2
                Logs(H) = Logs(H + 1)
            Next H
            ' Décrémente le nombre de fenêtres gérées
            LogCount = LogCount - 1
            Exit For
        End If
    Next t
    
End Sub

' Ici, tous les évènements de toutes les formes qui ont été hookées sont interceptées.
' Si ce ne sont pas des informations utiles qui sont transmises, on les renverra à
'   l'ancien process
Public Function WindowProc(ByVal hWnd As Long, _
                           ByVal wMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    
    Dim t As Integer
    Dim oldProc As Long     ' Adresse du process d'origine
    Dim ModeAimantation As Integer
    Dim r As RECT, p As POINTAPI
    Dim runProc As Boolean  ' Sera lu à la fin pour savoir si on a utilisé les données
    
    ' Par défaut, données non utilisées
    runProc = True
    
    ' On recherche la forme associée au handle fourni
    For t = 0 To LogCount - 1
        If Logs(t).hWnd = hWnd Then
            ' pour récupérer le numéro du process original
            oldProc = Logs(t).oldProc
            ' et le mode d'aimantation choisi pour cette forme
            ModeAimantation = Logs(t).Aimantation
            Exit For
        End If
    Next t

    If oldProc = 0 Then Exit Function ' On ressort si on ne l'a pas trouvé (bizarre)
    
    If wMsg = WM_ENTERSIZEMOVE Then ' Windows nous informe d'un Resize ou un Move vient
        GetWindowRect hWnd, r       ' de commencer
        GetCursorPos p
        MouseX = p.X - r.Left
        MouseY = p.Y - r.Top
        ' Récupère les coordonnées des autres formes (position et taille)
        GetFrmRects hWnd, ModeAimantation

    ElseIf wMsg = WM_SIZING Or _
           wMsg = WM_MOVING Then ' Un Resize ou Move est en cours
        CopyMemory r, ByVal lParam, Len(r) ' Récupère les données grace au pointeur
        ' Aimantation de la forme si besoin
        If wMsg = WM_SIZING Then
            DockFormRect hWnd, Sizing, r, wParam
        Else
            DockFormRect hWnd, Moving, r, wParam, MouseX, MouseY
        End If
        ' On réécrit les données vers le pointeur
        CopyMemory ByVal lParam, r, Len(r)
        
        WindowProc = 1  ' Fonction Ok
        runProc = False ' On a utilisé les données. On ne les renverra pas à la forme
    End If
    
    ' Si les infos n'ont pas été utilisées, on les renvoie à la forme
    If runProc Then WindowProc = CallWindowProc(oldProc, hWnd, wMsg, wParam, lParam)
    
End Function

' Fabrique la liste des coordonnées des formes auxquelles on veut s'aimanter
Private Function GetFrmRects(ByVal hWnd As Long, _
                             ByVal ModeAimantation As AimantationType)

    Dim frm As Form, i As Integer
    
    ' Réinitialise la liste des coordonnées
    ReDim Rects(0 To 0)
    ' Récupère l'espace de travail
    SystemParametersInfo_Rect SPI_GETWORKAREA, vbNull, Rects(0), 0
    
    Select Case ModeAimantation
        Case [Aimantable aux formes de mon Appli]
            ' On ne va faire les recherches que par rapports aux formes de notre Appli
            i = 1
            For Each frm In Forms   ' Pour chaque fenêtre de notre projet
                ' Si une forme est visible (et que ce n'est pas la notre)
                If frm.Visible And Not frm.hWnd = hWnd Then
                    ReDim Preserve Rects(0 To i)
                    ' On mémorise la position/taille de la fenêtre
                    GetWindowRect frm.hWnd, Rects(i)
                    i = i + 1   ' Incrémente le nombre de fenêtres
                End If
            Next frm
    
        Case [Aimantable à toutes les formes du bureau]
            ' Désigne notre handle pour ne pas en tenir compte dans la liste
            hWnd_à_exclure = hWnd
            ' On va rechercher les positions de toutes les fenêtres du bureau
            EnumWindows AddressOf EnumToutesFenêtres, ByVal 0&
    
    End Select
    
End Function

' Recherche les positions de toutes les fenêtres du bureau
' (CallBack : Windows appellera cette function tant qu'il y aura des données)
Private Function EnumToutesFenêtres(ByVal hWnd As Long, _
                                    ByVal lParam As Long) As Boolean
    
    Dim Nb As Long, REC As RECT

    ' Pas de mémorisation si cette forme est la notre
    '   ou si la forme est en icone
    '   ou n'est pas accessible
    '   ou n'est pas visible
    If hWnd = hWnd_à_exclure Or _
       IsIconic(hWnd) <> 0 Or _
       IsWindowEnabled(hWnd) = 0 Or _
       IsWindowVisible(hWnd) = 0 Then GoTo Fin
    
    ' Récupère les coordonnées de la forme
    GetWindowRect hWnd, REC
    ' On ressort si les coordonnées sont incohérentes
    If REC.Top = REC.Bottom Or _
       REC.Left = REC.Right Then GoTo Fin
    
    ' Incrémente le nombre de coordonnées dans la liste
    Nb = UBound(Rects) + 1
    ReDim Preserve Rects(0 To Nb)
    ' Ajoute ces coordonnées à la liste
    Rects(Nb) = REC
    
Fin:
    ' Continue l'énumération
    EnumToutesFenêtres = True
End Function

' Le coeur du programme : les tests pour savoir s'il faut aimanter
Private Sub DockFormRect(ByVal hWnd As Long, _
                         ByVal Mode As SnapFormMode, _
                         GivenRect As RECT, _
                         Optional SizingEdge As Long, _
                         Optional MouseX As Long, _
                         Optional MouseY As Long)
    
    Dim p As POINTAPI
    Dim i As Long, W As Long, H As Long
    Dim tmpRect As RECT, frmRect As RECT
    Dim diffX As Long, diffY As Long
    Dim XPos As Long, YPos As Long
    Dim tmpXPos As Long, tmpYPos As Long
    Dim tmpMouseX As Long, tmpMouseY As Long
    Dim FoundX As Boolean, FoundY As Boolean
    
    diffX = SnapWidth
    diffY = SnapWidth
    
    ' Par défaut, les futures coordonnées sont celles de notre forme
    tmpRect = GivenRect
    frmRect = GivenRect
    
    '
    If Mode = Moving Then
        GetCursorPos p
        If SnappedX Then
            tmpMouseX = p.X - tmpRect.Left
            OffsetRect tmpRect, tmpMouseX - MouseX, 0
            OffsetRect GivenRect, tmpMouseX - MouseX, 0
        Else
            MouseX = p.X - tmpRect.Left
        End If
        If SnappedY Then
            tmpMouseY = p.Y - tmpRect.Top
            OffsetRect tmpRect, 0, tmpMouseY - MouseY
            OffsetRect GivenRect, 0, tmpMouseY - MouseY
        Else
            MouseY = p.Y - tmpRect.Top
        End If
    End If
    
    ' Précalcule la largeur et hauteur de la fenêtre
    '   (+ facile dans les équations qui suivent)
    W = tmpRect.Right - tmpRect.Left
    H = tmpRect.Bottom - tmpRect.Top
    
    ' Et maintenant, la partie difficile à lire (lol)
    ' ----- 1er cas : Si la fenêtre se déplace
    If Mode = Moving Then
        For i = 0 To UBound(Rects)
            '----- Déplacements horizontaux :
            If (tmpRect.Left >= (Rects(i).Left - SnapWidth) And _
                tmpRect.Left <= (Rects(i).Left + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpRect.Left - Rects(i).Left) < diffX Then
                    GivenRect.Left = Rects(i).Left
                    GivenRect.Right = GivenRect.Left + W
                    diffX = Abs(tmpRect.Left - Rects(i).Left)
                    FoundX = True
                
            ElseIf i > 0 And (tmpRect.Left >= (Rects(i).Right - SnapWidth) And _
                tmpRect.Left <= (Rects(i).Right + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpRect.Left - Rects(i).Right) < diffX Then
                    GivenRect.Left = Rects(i).Right
                    GivenRect.Right = GivenRect.Left + W
                    diffX = Abs(tmpRect.Left - Rects(i).Right)
                    FoundX = True
                
            ElseIf i > 0 And (tmpRect.Right >= (Rects(i).Left - SnapWidth) And _
                tmpRect.Right <= (Rects(i).Left + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpRect.Right - Rects(i).Left) < diffX Then
                    GivenRect.Right = Rects(i).Left
                    GivenRect.Left = GivenRect.Right - W
                    diffX = Abs(tmpRect.Right - Rects(i).Left)
                    FoundX = True
                
            ElseIf (tmpRect.Right >= (Rects(i).Right - SnapWidth) And _
                tmpRect.Right <= (Rects(i).Right + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpRect.Right - Rects(i).Right) < diffX Then
                    GivenRect.Right = Rects(i).Right
                    GivenRect.Left = GivenRect.Right - W
                    diffX = Abs(tmpRect.Right - Rects(i).Right)
                    FoundX = True
            End If
            
            '----- Déplacements verticaux :
            If (tmpRect.Top >= (Rects(i).Top - SnapWidth) And _
                tmpRect.Top <= (Rects(i).Top + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpRect.Top - Rects(i).Top) < diffY Then
                    GivenRect.Top = Rects(i).Top
                    GivenRect.Bottom = GivenRect.Top + H
                    diffY = Abs(tmpRect.Top - Rects(i).Top)
                    FoundY = True
                
            ElseIf i > 0 And (tmpRect.Top >= (Rects(i).Bottom - SnapWidth) And _
                tmpRect.Top <= (Rects(i).Bottom + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpRect.Top - Rects(i).Bottom) < diffY Then
                    GivenRect.Top = Rects(i).Bottom
                    GivenRect.Bottom = GivenRect.Top + H
                    diffY = Abs(tmpRect.Top - Rects(i).Bottom)
                    FoundY = True
                
            ElseIf i > 0 And (tmpRect.Bottom >= (Rects(i).Top - SnapWidth) And _
                tmpRect.Bottom <= (Rects(i).Top + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpRect.Bottom - Rects(i).Top) < diffY Then
                    GivenRect.Bottom = Rects(i).Top
                    GivenRect.Top = GivenRect.Bottom - H
                    diffY = Abs(tmpRect.Bottom - Rects(i).Top)
                    FoundY = True
                
            ElseIf (tmpRect.Bottom >= (Rects(i).Bottom - SnapWidth) And _
                tmpRect.Bottom <= (Rects(i).Bottom + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpRect.Bottom - Rects(i).Bottom) < diffY Then
                    GivenRect.Bottom = Rects(i).Bottom
                    GivenRect.Top = GivenRect.Bottom - H
                    diffY = Abs(tmpRect.Bottom - Rects(i).Bottom)
                    FoundY = True
            End If
        Next i
        
        ' Mémorise si on doit faire un déplacement horizontal et/ou vertical
        SnappedX = FoundX
        SnappedY = FoundY
        
    ' ----- 1er cas : Si la fenêtre est redimensionnée
    ElseIf Mode = Sizing Then
        ' Si on manipule la fenêtre par la gauche
        '   ou par un des coins haut ou bas
        If SizingEdge = WMSZ_LEFT Or _
           SizingEdge = WMSZ_TOPLEFT Or _
           SizingEdge = WMSZ_BOTTOMLEFT Then
                XPos = GivenRect.Left
        Else
                XPos = GivenRect.Right
        End If
        
        ' Si on manipule la fenêtre par le sommet
        '   ou par un des coins gauche ou droit
        If SizingEdge = WMSZ_TOP Or _
           SizingEdge = WMSZ_TOPLEFT Or _
           SizingEdge = WMSZ_TOPRIGHT Then
                YPos = GivenRect.Top
        Else
                YPos = GivenRect.Bottom
        End If

        tmpXPos = XPos
        tmpYPos = YPos

        For i = 0 To UBound(Rects)
            '----- Dimensionnements horizontaux :
            If ((tmpXPos >= (Rects(i).Left - SnapWidth) And _
                tmpXPos <= (Rects(i).Left + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpXPos - Rects(i).Left) < diffX) Then
                    XPos = Rects(i).Left
                    diffX = Abs(tmpXPos - Rects(i).Left)
                
            ElseIf (tmpXPos >= (Rects(i).Right - SnapWidth) And _
                tmpXPos <= (Rects(i).Right + SnapWidth)) And _
                ((tmpRect.Top - SnapWidth) < Rects(i).Bottom And _
                (tmpRect.Bottom + SnapWidth) > Rects(i).Top) And _
                Abs(tmpXPos - Rects(i).Right) < diffX Then
                    XPos = Rects(i).Right
                    diffX = Abs(tmpXPos - Rects(i).Right)
            End If
            
            '----- Dimensionnements verticaux :
            If (tmpYPos >= (Rects(i).Top - SnapWidth) And _
                tmpYPos <= (Rects(i).Top + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpYPos - Rects(i).Top) < diffY Then
                    YPos = Rects(i).Top
                    diffY = Abs(tmpYPos - Rects(i).Top)
                
            ElseIf (tmpYPos >= (Rects(i).Bottom - SnapWidth) And _
                tmpYPos <= (Rects(i).Bottom + SnapWidth)) And _
                ((tmpRect.Left - SnapWidth) < Rects(i).Right And _
                (tmpRect.Right + SnapWidth) > Rects(i).Left) And _
                Abs(tmpYPos - Rects(i).Bottom) < diffY Then
                    YPos = Rects(i).Bottom
                    diffY = Abs(tmpYPos - Rects(i).Bottom)
            End If
        Next i

        ' Si on manipule la fenêtre par la gauche
        '   ou par un des coins haut ou bas
        If SizingEdge = WMSZ_LEFT Or _
           SizingEdge = WMSZ_TOPLEFT Or _
           SizingEdge = WMSZ_BOTTOMLEFT Then
                GivenRect.Left = XPos
        Else
                GivenRect.Right = XPos
        End If
    
        ' Si on manipule la fenêtre par le sommet
        '   ou par un des coins gauche ou droit
        If SizingEdge = WMSZ_TOP Or _
           SizingEdge = WMSZ_TOPLEFT Or _
           SizingEdge = WMSZ_TOPRIGHT Then
                GivenRect.Top = YPos
        Else
                GivenRect.Bottom = YPos
        End If
    End If
    
End Sub
