VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================================
' Name:         TButton Class Creates a Title Bar Button
'               cTbutton.cls
' Author:       Nitin Kohli (pulsatingstar20@yahoo.com)
' Date:         20 March 2009
'
' Description:  Class Draws a button Image on a Forms Title Bar
'               and raises Click ,DblClick Events when user clicks
'               on button.
'
'                Set properties
'                IconFilename    : 24x24 Icon Path&filename to draw button
'                IconFilenameBG  : 24x24 Icon Path&filename to draw Background
'                                  button Image to lighten , Hover ,Selected effects
'                Edge            : Distance of button from forms right edge
'       Finally  Hwnd            : Forms handle to Draw button & subclass form
'
'                ResourceID      : Instead of IconFilename/IconFilenameBG
'                                : ID,ID+1 will work in executable only

' Depedencies   vbAccelerator Image List
'               http://www.vbaccelerator.com/home/VB/Code/Controls/ImageList/vbAccelerator_Image_List_Control/article.asp
'               cVBALImageList.cls

'               vbAccelerator SSubTmr6
'               http://www.vbaccelerator.com/home/VB/Code/Libraries/Subclassing/SSubTimer_ASM_Version/VB6_ASM_SSubTmr6_Binary.asp
'               Download Dll File & Resister


'Further Suggestions

'           Multiple Instances can be used to draw multiple buttons
'           Can be basis for multi button toolbar

'           Systemwide Hook can be used to place this on all windows
'           to perform action for ur application
'           'Music toolbar etc
'
'********** Votes will be encorage more postings *****************
' ======================================================================================


Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    left As Long
    tOp As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long



Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Private Declare Function TRACKMOUSEEVENT Lib "user32" Alias "TrackMouseEvent" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long
    Private Const TME_NONCLIENT = &H10
    Private Const WM_NCMOUSELEAVE = &H2A2
    Private Const WM_NCPAINT = &H85
    Private Const WM_NCACTIVATE = &H86
    Private Const WM_ACTIVATE = &H6
    Private Const WM_NCMOUSEMOVE = &HA0
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const WM_NCLBUTTONUP = &HA2
    Private Const WM_NCLBUTTONDBLCLK = &HA3
    Private Const WM_NCHITTEST = &H84
    Private Const WM_DESTROY = &H2




Dim mIml                As cVBALImageList

Dim mIconFileName       As String
Dim mIconFileNameBG     As String
Dim mResourceID         As Long
Dim mHwnd               As Long
Dim mEdge               As Long
Dim mbTracking          As Boolean


Public Event Click()
Public Event DblClick()

Private m_emr As EMsgResponse



Implements ISubclass

Public Property Let IconFilename(RHS As String)
    mIconFileName = RHS
End Property

Public Property Get IconFilename() As String
    IconFilename = mIconFileName
End Property
Public Property Let IconFilenameBG(RHS As String)
    mIconFileNameBG = RHS
End Property

Public Property Get IconFilenameBG() As String
    IconFilenameBG = mIconFileNameBG
End Property
Public Property Let ResourceID(RHS As Long)
    mResourceID = RHS
End Property
Public Property Get ResourceID() As Long
    ResourceID = mResourceID
End Property
Public Property Let Edge(RHS As Long)
    mEdge = RHS
End Property
Public Property Get Edge() As Long
    Edge = mEdge
End Property


Public Property Let hwnd(RHS As Long)
    mHwnd = RHS
    
    With mIml
        If ResourceID <> 0 Then
            .AddFromResourceID mResourceID, App.hInstance, IMAGE_ICON, 1
            .AddFromResourceID mResourceID + 1, App.hInstance, IMAGE_ICON, 2
        Else
            .AddFromFile IconFilename, IMAGE_ICON, 1
            .AddFromFile IconFilenameBG, IMAGE_ICON, 2
        End If
    
        If .ImageCount >= 2 Then
        
            AttachMessage Me, hwnd, WM_ACTIVATE
            AttachMessage Me, hwnd, WM_NCPAINT
            AttachMessage Me, hwnd, WM_NCACTIVATE
            AttachMessage Me, hwnd, WM_NCMOUSEMOVE
            AttachMessage Me, hwnd, WM_NCMOUSELEAVE
            AttachMessage Me, hwnd, WM_NCLBUTTONDOWN
            AttachMessage Me, hwnd, WM_NCLBUTTONUP
            AttachMessage Me, hwnd, WM_NCLBUTTONDBLCLK
            AttachMessage Me, hwnd, WM_NCHITTEST
            AttachMessage Me, hwnd, WM_DESTROY
            
        Else
            Debug.Print "Insufficient Icon Information to draw Tbutton"
        End If
    End With
    
    
End Property
Public Property Get hwnd() As Long
    hwnd = mHwnd
End Property

Private Sub Class_Initialize()
    Set mIml = New cVBALImageList
    With mIml
        .IconSizeX = 24
        .IconSizeY = 24
        .ColourDepth = ILC_COLOR32
        .Create
    End With
   m_emr = emrPostProcess
End Sub



Private Sub Class_Terminate()
    mIml.Destroy
    Set mIml = Nothing
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
    m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
        ISubclass_MsgResponse = m_emr
        
End Property


Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Static bUserMode As Boolean
    Static bButtonHover        As Boolean 'Mouse Move over Button
    Static bButtonDown         As Boolean 'Mouse Down over Button
    
    
    
    Select Case iMsg
    
    Case WM_NCACTIVATE, WM_NCPAINT, WM_ACTIVATE
    
        'When UserMode Flag Set
        'Let original WndProc handle messages first
        'We Never know what Class User Might do in Click event
        
        If bUserMode Then
            m_emr = emrPostProcess
        Else
            m_emr = emrPreProcess
        End If
        DrawButton
        
    Case WM_NCMOUSEMOVE
        m_emr = emrPostProcess
        
        If Not mbTracking Then TrackMouse
        'Redraw Button to mark mouser over
        If IsOverButton(lParam) Then
            If Not bButtonHover Then
                DrawButton False, True
                bButtonHover = True
            End If
        Else
            If bButtonHover Or bButtonDown Then
                DrawButton
                bButtonHover = False
                bButtonDown = False
            End If
        End If

        
    Case WM_NCMOUSELEAVE
        'Catch Leave to restore buttons Normal image
        mbTracking = False
        If bButtonHover Then
            DrawButton
            bButtonHover = False
        End If
        
    Case WM_NCLBUTTONDOWN
        'Mouse Down on button
        If IsOverButton(lParam) Then
            DrawButton True, False
            
            bButtonHover = False
            bButtonDown = True
        End If
        
        
    Case WM_NCLBUTTONUP
            'Never Recived Unless Window is maximise
            'Bcoz of that COSTLY WM_NCHITTEST needs to be checked for button up
            'Somebody has a better idea suggest or mail me
            'pulsatingstar20@yahoo.com
            m_emr = emrPostProcess
    Case WM_NCHITTEST
        m_emr = emrPostProcess
        
        If IsOverButton(lParam) And bButtonDown Then
            bButtonDown = False
            bUserMode = True
            RaiseEvent Click
            bUserMode = False
            DrawButton
        End If
        
    Case WM_NCLBUTTONDBLCLK
    
        If IsOverButton(lParam) Then
            bUserMode = True
            RaiseEvent DblClick
            bUserMode = False
            m_emr = emrConsume
        Else
            m_emr = emrPostProcess
        End If
    Case WM_DESTROY
        Call Detach
    End Select
    
    

End Function


Public Sub DrawButton(Optional bSelected As Boolean, Optional bCut As Boolean)
    Dim lDc         As Long
    Dim rec         As RECT
    
    lDc = GetWindowDC(hwnd)
    
    GetWindowRect hwnd, rec
    
    With rec
        mIml.DrawImage 2, lDc, (.Right - .left) - Edge, 4, bSelected, bCut
        mIml.DrawImage 1, lDc, (.Right - .left) - Edge, 4, bSelected, bCut
    End With
    
    ReleaseDC hwnd, lDc
    
End Sub


Private Function IsOverButton(ByVal lPos As Long)
    ' Determine if the specified Coords are within our custom button

    Dim xPos As Long, ypos As Long
    Dim ActiveRec As RECT
    
    GetWindowRect hwnd, ActiveRec
    
    With ActiveRec
    
        xPos = LoWord(lPos)
        ypos = HiWord(lPos)
        
        .left = .Right - Edge
        .Right = .left + 24
        .Bottom = .tOp + 24
        IsOverButton = xPos > .left And xPos < .Right And ypos > .tOp And ypos < .Bottom
    End With
    
    
    
    
End Function

Private Function TrackMouse()

    'Track needs to be activated to receive WM_NCMOUSELEAVE
    '
    
    Dim trk As tagTRACKMOUSEEVENT
     
    mbTracking = True
    
    With trk
        .cbSize = 16
        .dwFlags = TME_NONCLIENT
        .hwndTrack = hwnd
    End With

    TRACKMOUSEEVENT trk
    
    
End Function

Private Sub Detach()

    DetachMessage Me, hwnd, WM_ACTIVATE
    DetachMessage Me, hwnd, WM_NCACTIVATE
    DetachMessage Me, hwnd, WM_NCPAINT
    DetachMessage Me, hwnd, WM_NCMOUSEMOVE
    DetachMessage Me, hwnd, WM_NCMOUSELEAVE
    DetachMessage Me, hwnd, WM_NCLBUTTONDOWN
    DetachMessage Me, hwnd, WM_NCLBUTTONUP
    DetachMessage Me, hwnd, WM_NCLBUTTONDBLCLK
    DetachMessage Me, hwnd, WM_NCHITTEST
    DetachMessage Me, hwnd, WM_DESTROY
        

End Sub
Private Property Get HiWord(ByRef lThis As Long) As Long
   If (lThis And &H80000000) = &H80000000 Then
      HiWord = ((lThis And &H7FFF0000) \ &H10000) Or &H8000&
   Else
      HiWord = (lThis And &HFFFF0000) \ &H10000
   End If
End Property

Private Property Get LoWord(ByRef lThis As Long) As Long
   LoWord = (lThis And &HFFFF&)
End Property

