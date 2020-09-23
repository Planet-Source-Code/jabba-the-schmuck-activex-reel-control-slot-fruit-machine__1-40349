VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucReel 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   ScaleHeight     =   152
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   Begin MSComctlLib.ImageList iDefault 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":18A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":3D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":49EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":563E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":6290
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":6EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":7B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucMultiReel.ctx":8786
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   1680
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   2640
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picReel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   0
      Left            =   0
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   0
      Width           =   510
      Begin VB.Shape shHold 
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00808080&
         FillStyle       =   7  'Diagonal Cross
         Height          =   975
         Index           =   0
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "ucReel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'Ensure all variables are declared.  (Improves performance if they are.)

'Declare global variables, arrays, and events.
Dim myReels As Long
Dim myPicsPerReel As Long
Dim myNaturalFinish As Boolean
Dim myHorizontalSpin As Boolean
Dim myEnableHold As Boolean

Dim stopNow As Boolean
Dim doNotAlign As Boolean

Dim sourceNo() As Long
Dim myDirectionUp() As Boolean
Dim myStepDistance() As Long
Dim mySpeed() As Long
Dim myMinCyclesForSpin() As Long
Dim myReelFixedStop() As Long
Dim myCount1() As Long
Dim myCount2() As Long
Dim destX() As Long
Dim destY() As Long
Dim cycles() As Long
Dim stepNo() As Double
Dim resetAfterNaturalFinish() As Boolean

Public Event OnReelStart(ByVal whatReel As Integer)
Public Event OnReelStop(ByVal whatReel As Integer)
Public Event OnEachFullCycle(ByVal whatReel As Integer)
Public Event OnReelClick(ByVal whatReel As Integer)

Public Sub Spin(Optional ByVal whatReel As Long = 0, _
                Optional ByVal spinType As Long = 0, _
                Optional ByVal reelFixedStop As Long = 0, _
                Optional ByVal changeDirection, _
                Optional ByVal minCyclesForSpin, _
                Optional ByVal spinSpeed)

    'Set these two variables to False to allow the reel(s) to cycle more
    'than once.
    stopNow = False
    doNotAlign = False
    
    'If an invalid reel number is passed to this routine, whatReel will
    'be set to zero, which means all reels will spin.
    If whatReel > Reels Or whatReel < 0 Then whatReel = 0
    'If an invalid spinType is passed to this routine, a random one will
    'be chosen.
    If spinType > 6 Or spinType < 1 Then spinType = Int(Rnd * 5) + 1

    If whatReel = 0 Then
        'Loop through all reels if whatReel is zero.
        For whatReel = 0 To myReels - 1
            'Set the speed, direction, and minimum cycles for the spin.
            Call SetSpinType(spinType, whatReel)
            'Set the reels to stop on whatever position has been passed.  If zero,
            'or invalid then a random stop is set.
            If reelFixedStop < 1 Or reelFixedStop > myPicsPerReel + 1 Then
                myReelFixedStop(whatReel) = Int(Rnd * (myPicsPerReel + 1))
            Else
                myReelFixedStop(whatReel) = reelFixedStop
            End If
            If myReelFixedStop(whatReel) = myPicsPerReel + 1 Then myReelFixedStop(whatReel) = 0
            'Change the variables previously set in the SetSpinType routine, IF SPECIFIED.
            If IsMissing(changeDirection) = False Then myDirectionUp(whatReel) = changeDirection
            If IsMissing(minCyclesForSpin) = False Then myMinCyclesForSpin(whatReel) = minCyclesForSpin
            If IsMissing(spinSpeed) = False Then mySpeed(whatReel) = spinSpeed
            'The reels will soon begin spinning, so raise the OnReelStart event for each
            'reel.
            RaiseEvent OnReelStart(whatReel)
        Next whatReel
        'Spin.
        Call MySpinRoutine(0, myReels - 1)
    Else
        '_____________________________________________________________
        'Do the same as above only with ONE reel.
        If reelFixedStop < 1 Or reelFixedStop > myPicsPerReel + 1 Then
            myReelFixedStop(whatReel - 1) = Int(Rnd * (myPicsPerReel + 1))
        Else
            myReelFixedStop(whatReel - 1) = reelFixedStop
        End If
        If spinType = 6 Then Call StopSpinning
        If myReelFixedStop(whatReel - 1) = myPicsPerReel + 1 Then myReelFixedStop(whatReel - 1) = 0
        If IsMissing(changeDirection) = False Then myDirectionUp(whatReel - 1) = changeDirection
        If IsMissing(minCyclesForSpin) = False Then myMinCyclesForSpin(whatReel - 1) = minCyclesForSpin
        If IsMissing(spinSpeed) = False Then mySpeed(whatReel - 1) = spinSpeed
        RaiseEvent OnReelStart(whatReel - 1)
        Call MySpinRoutine(whatReel - 1, whatReel - 1)
    End If

End Sub

Private Sub MySpinRoutine(ByVal whatStart As Long, ByVal whatEnd As Long)

'Declare local variable.
Dim whatReel As Long

    'Begin loop.
    Do
        'Far quicker than simple DoEvents, as it will check if there is anything
        'to 'do' first.
        If GetInputState <> 0 Then DoEvents
        'Begin loop to cycle through each reel.
        For whatReel = whatStart To whatEnd
            'Don't do anything if the reel is held.
            If shHold(whatReel).Visible = False Then
                'If the reel isn't ready to stop.
                If GetEndOfSpin(whatReel) = False Or cycles(whatReel) = 0 Then
                    'Get the first 'time'.
                    myCount2(whatReel) = GetTickCount
                    'If the reel is aligning, set it to be much slower for better effect.
                    If resetAfterNaturalFinish(whatReel) = True Then mySpeed(whatReel) = 90
                    'If enough time has elapsed since the last cycle of the loop.
                    If myCount2(whatReel) - myCount1(whatReel) > mySpeed(whatReel) Then
                        'Get the second 'time' and draw the next step of the reel.
                        myCount1(whatReel) = GetTickCount
                        Call Draw(whatReel)
                    End If
                Else 'The reel is ready to stop...
                    'If the reel needs to align...
                    If myNaturalFinish = True Then
                        'Initialize variables to incorperate the aligning.
                        If resetAfterNaturalFinish(whatReel) = False Then
                            resetAfterNaturalFinish(whatReel) = True
                            myStepDistance(whatReel) = 1
                            myDirectionUp(whatReel) = Not myDirectionUp(whatReel)
                            Call Draw(whatReel)
                        End If
                    End If
                End If
            End If
        Next whatReel

        'Loop through each spinning reel and loop if it hasn't finished spinning.
        For whatReel = whatStart To whatEnd
            If shHold(whatReel).Visible = False Then
                If GetEndOfSpin(whatReel) = False Or stepNo(whatReel) <> 4 Then GoTo DoNotEndYet
            End If
        Next whatReel
        'Exit the loop; end the spin.
        Exit Do
DoNotEndYet:
    Loop

'Reset variables for next spin.
For whatReel = whatStart To whatEnd
    If resetAfterNaturalFinish(whatReel) = True Then
        resetAfterNaturalFinish(whatReel) = False
        myStepDistance(whatReel) = 4
    End If
    cycles(whatReel) = 0
    mySpeed(whatReel) = 28
    'The reel(s) have finished spinning so raise the OnReelStop event for each
    'relevant reel.
    RaiseEvent OnReelStop(whatReel)
Next whatReel

End Sub

Private Function GetEndOfSpin(ByVal whatReel As Long) As Boolean

'If the StopSpinning routine has been called, this will stop the spin,
'and align if requested.
If stopNow And doNotAlign Then GetEndOfSpin = True: Exit Function

    'Return False if a smooth finish is needed and the reel isn't in the
    'stopping position (step 4).
    If myNaturalFinish = False And stepNo(whatReel) <> 4 Then
        GetEndOfSpin = False
        Exit Function
    ElseIf myNaturalFinish = True Then
        'A natural finish is needed, so return False if the reel isn't one step
        'above or below the stopping position - depending on the direction of
        'the spin.
        If resetAfterNaturalFinish(whatReel) = False Then
            If myDirectionUp(whatReel) = True And stepNo(whatReel) <> 3 Then
                GetEndOfSpin = False
                Exit Function
            ElseIf myDirectionUp(whatReel) = False And stepNo(whatReel) <> 5 Then
                GetEndOfSpin = False
                Exit Function
            End If
        Else
            'Unless the reel is aligning, then we want to stop in the normal place.
            If stepNo(whatReel) = 4 Then
                GetEndOfSpin = True
            Else
                GetEndOfSpin = False
            End If
            Exit Function
        End If
    End If
    
    'If none of the above conditions have caused an Exit Function, then it's possible
    'that the spin has come to an end.  Just need to check that the number of cycles
    'made is more than the minimum cycles for the spin, and that the reel position is
    'correct.  If any of these aren't True, GetEndOfSpin will return False.
    If Abs(cycles(whatReel)) > myMinCyclesForSpin(whatReel) Then
        If stopNow = True Or GetSource(whatReel) = myReelFixedStop(whatReel) Then
            GetEndOfSpin = True
        Else
            GetEndOfSpin = False
        End If
    End If

End Function

Private Sub SetSpinType(spinType As Long, ByVal whatReel As Long)

'Random direction if the spinType is 1 or 2.
If spinType < 3 Then
    myDirectionUp(whatReel) = Int(Rnd * 2)
Else
    'If between 3 and 6 then each reel will spin in a different direction to that next to it.
    If whatReel <> 0 Then myDirectionUp(whatReel) = Not myDirectionUp(whatReel - 1)
End If

    Select Case spinType
        Case 1
            'The reels (should) stop from left to right.
            If whatReel > 0 Then myMinCyclesForSpin(whatReel) = myMinCyclesForSpin(whatReel - 1) + myReels
        Case 2
            'The reels (should) stop from right to left.
            If whatReel = 0 Then myMinCyclesForSpin(whatReel) = myReels + myReels * 4 Else myMinCyclesForSpin(whatReel) = myMinCyclesForSpin(whatReel - 1) - 5
        Case 3
            'Short spin.
            myMinCyclesForSpin(whatReel) = 0
        Case 4
            'Average spin, with the outside reels spinning faster than the inner ones.
            myMinCyclesForSpin(whatReel) = myPicsPerReel
            If whatReel = 0 Or whatReel = myReels - 1 Then mySpeed(whatReel) = 15 Else mySpeed(whatReel) = 30
        Case 5
            'Every other reel (ie: 1st, 3rd, 5th, etc) will spin longer than the others,
            'and quicker.
            If whatReel Mod 2 = 0 Then myMinCyclesForSpin(whatReel) = 0 Else myMinCyclesForSpin(whatReel) = 14
            If whatReel Mod 2 = 0 Then mySpeed(whatReel) = 30 Else mySpeed(whatReel) = 15
        Case 6
            'Shuffle; each reel will make one cycle in the opposite direction to that next
            'to it.
            Call StopSpinning
    End Select

End Sub

Public Sub Nudge(Optional ByVal whatReel As Long, Optional changeDirection As Boolean = False)
    
    'Move the reel(s) one cycle in the given direction.
    Call StopSpinning
    Call Spin(whatReel, 6, , changeDirection)

End Sub

Private Sub Draw(whatReel As Long)

'Increment/decrement (depending on the direction of the spin) the StepNo variable.
If myDirectionUp(whatReel) = False Then
    stepNo(whatReel) = stepNo(whatReel) + myStepDistance(whatReel) / 4
    'Adjust the destination point to draw from, DestY if spinning vertically, DestX if
    'spinning horizontally.
    If myHorizontalSpin = False Then destY(whatReel) = destY(whatReel) + myStepDistance(whatReel) Else destX(whatReel) = destX(whatReel) + myStepDistance(whatReel)
Else
    stepNo(whatReel) = stepNo(whatReel) - myStepDistance(whatReel) / 4
    If myHorizontalSpin = False Then destY(whatReel) = destY(whatReel) - myStepDistance(whatReel) Else destX(whatReel) = destX(whatReel) - myStepDistance(whatReel)
End If

'Reset the stepNo() variable, when each full cycle has passed. (1 cycle = 8 Steps)
If myDirectionUp(whatReel) = False Then
    If stepNo(whatReel) = 9 Then
        stepNo(whatReel) = 1
        'As one full cycle has passed, raise the OnEachFullCycle event.
        RaiseEvent OnEachFullCycle(whatReel)
        cycles(whatReel) = cycles(whatReel) + 1
        'Reset where to draw from (DestX()/DestY()), and decrement the picture source (sourceNo()).
        If myHorizontalSpin = False Then destY(whatReel) = 0 - 32 + 32 / 8 Else destX(whatReel) = 0 - 32 + 32 / 8
        Incr sourceNo(whatReel), -1
        If sourceNo(whatReel) < 0 Then sourceNo(whatReel) = myPicsPerReel
    End If
Else
    'The same as above, but if the reel is spinning the other way.
    If stepNo(whatReel) = 0 Then
        stepNo(whatReel) = 8
        RaiseEvent OnEachFullCycle(whatReel)
        Incr cycles(whatReel), -1
        If myHorizontalSpin = False Then destY(whatReel) = 0 Else destX(whatReel) = 0
        Incr sourceNo(whatReel)
        If sourceNo(whatReel) > myPicsPerReel Then sourceNo(whatReel) = 0
    End If
End If

'Clear the PictureBox ready to draw...
picReel(whatReel).Cls

Dim i As Long
    '3 cycles in this loop, to draw the top, middle, and bottom picture.  (Or left, middle,
    'and right picture!)
    For i = 0 To 2
        'If you're not sure about BitBlt search planet-source-code.com for a tutorial, plenty
        'out there.
        If myHorizontalSpin = False Then
            Call BitBlt(picReel(whatReel).hDC, 0, destY(whatReel) + i * 32, 32, 32, picMask(GetSource(whatReel, i)).hDC, 0, 0, SRCAND)
            Call BitBlt(picReel(whatReel).hDC, 0, destY(whatReel) + i * 32, 32, 32, picSprite(GetSource(whatReel, i)).hDC, 0, 0, SRCINVERT)
        'Top one for vertical spin, bottom for horizontal.  Using DestY & DestX respectively.
        Else
            Call BitBlt(picReel(whatReel).hDC, destX(whatReel) + i * 32, 0, 32, 32, picMask(GetSource(whatReel, i)).hDC, 0, 0, SRCAND)
            Call BitBlt(picReel(whatReel).hDC, destX(whatReel) + i * 32, 0, 32, 32, picSprite(GetSource(whatReel, i)).hDC, 0, 0, SRCINVERT)
        End If
    Next i
'Refresh the PictureBox after each draw (or step).
picReel(whatReel).Refresh

End Sub

Private Function GetSource(ByVal whatReel As Long, Optional ByVal whatPos As Long = 2) As Long
'Used to get the relevant sourceNo().  This tells the program what picture to copy.
GetSource = sourceNo(whatReel) + whatPos
If GetSource > myPicsPerReel Then GetSource = GetSource - (myPicsPerReel + 1)

End Function

Private Sub picReel_Click(Index As Integer)

    'As the reel has been clicked, raise the OnReelClick event.
    RaiseEvent OnReelClick(Index)
    
    'Show the hold shape, if enabled. (To enable, set the EnableHold property to True)
    If myEnableHold = True Then
        shHold(Index).Visible = Not shHold(Index).Visible
    End If

End Sub

Private Sub UserControl_Initialize()

    'Initialize VB's randomizing routine.
    Randomize
    
    'Initialize properties.  If anyone knows how to do this better please post a reply
    'or e-mail peter_oakey@hotmail.com :)
    If Reels = 0 Then Reels = 5
    If PicsPerReel = 0 Then PicsPerReel = 5
    
    'Do one step of the reel, on initialization.
    Dim whatReel As Long
        For whatReel = 0 To myReels - 1
            Call Draw(whatReel)
        Next whatReel
    
End Sub

Private Sub InitializeVariables()

    Dim i As Integer
        'As suggested by the name of this routine, initialize variables.
        For i = 0 To myReels - 1
            destY(i) = 0
            destX(i) = 0
            sourceNo(i) = Int(Rnd * myPicsPerReel + 1)
            myStepDistance(i) = 32 / 8
            stepNo(i) = 4
            mySpeed(i) = 28
            If myHorizontalSpin = False Then destY(i) = 0 - 32 / 2 - 32 / 8 Else destX(i) = 0 - 32 / 2 - 32 / 8
        Next i

End Sub

Private Sub ArrangePics()
        
    Dim i As Integer
        'Set the mask and sprite pictures into the relevant PictureBoxes.
        For i = 0 To picMask.Count - 1
                picMask(i).Picture = iDefault.ListImages(i * 2 + 1).Picture
                picSprite(i).Picture = iDefault.ListImages(i * 2 + 2).Picture
        Next i

End Sub

Private Sub LoadUnloadControls(ByVal picsChanged As Boolean)

    'Declare local variable. (To be used in loops)
    Dim i As Integer

    If picsChanged = False Then

        ReDim destY(myReels - 1) As Long
        ReDim destX(myReels - 1) As Long
        ReDim sourceNo(myReels - 1) As Long
        ReDim myDirectionUp(myReels - 1) As Boolean
        ReDim myStepDistance(myReels - 1) As Long
        ReDim stepNo(myReels - 1) As Double
        ReDim cycles(myReels - 1) As Long
        ReDim mySpeed(myReels - 1) As Long
        ReDim myMinCyclesForSpin(myReels - 1) As Long
        ReDim myReelFixedStop(myReels - 1) As Long
        ReDim myCount1(myReels - 1) As Long
        ReDim myCount2(myReels - 1) As Long
        ReDim resetAfterNaturalFinish(myReels - 1) As Boolean
            
        'Unload unwanted PictureBoxes and Shapes, if too many are loaded already.
        If picReel.Count > Reels Then
            For i = Reels To picReel.Count - 1
                Unload shHold(i)
                Unload picReel(i)
            Next i
        End If
        
        'Load needed PictureBoxes and Shapes if there aren't enough already.
        If picReel.Count < Reels Then
            For i = picReel.Count To Reels - 1
                Load picReel(i)
                Load shHold(i)
            Next i
        End If
    Else
        'Unload unwanted PictureBoxes if there are too many already loaded.
        If picMask.Count > myPicsPerReel + 1 Then
            For i = myPicsPerReel + 1 To picMask.Count - 1
                Unload picMask(i)
                Unload picSprite(i)
            Next i
        End If
        'Load needed PictureBoxes if there aren't enough already loaded.
        If picMask.Count < myPicsPerReel Then
            For i = picMask.Count To myPicsPerReel
                Load picMask(i)
                Load picSprite(i)
            Next i
        End If
    End If
       
End Sub

Private Sub SetReelPosition()

'Error handling.
On Error GoTo ErrorHandler

    Dim i As Integer
        'Clear, adjust the size/position, and show the PictureBoxes (reels).
        For i = 0 To Reels - 1
            With picReel(i)
                .Cls
                If HorizontalSpin = False Then
                    .Height = 32 * 2
                    .Width = 32
                    .Left = 0 + i * picReel(0).Width
                    .Top = 0
                Else
                    .Height = 32
                    .Width = 32 * 2
                    .Left = 0
                    .Top = 0 + i * picReel(0).Height
                End If
                .Visible = True
                'Set the size and position of the hold Shapes, and place them in
                'the relevant PictureBox container.
                With shHold(i)
                    .Top = 0
                    .Left = 0
                    .Height = picReel(i).Height
                    .Width = picReel(i).Width
                End With
            End With
            Set shHold(i).Container = picReel(i)
            'If the pictures have been placed in the mask/sprite PictureBoxes then
            'draw one step to show them.
            If picMask(myPicsPerReel).Picture <> Empty Then Call Draw(CLng(i))
        Next i

'Exit so that the Error Handling routine is only done when called.
Exit Sub

'Error handling routine.
ErrorHandler:
    If Err.Number = 340 Then '340 = Missing a control.
        Call LoadUnloadControls(False)  '|
        Call InitializeVariables        '| - Should fix the problem.
        Resume Next                     '|
    Else
        Dim myResponse As Byte
            'If not, give as much info as possible.
            myResponse = MsgBox(Err.Description & " from " & Err.Source & Chr(10) _
            & Chr(10) & "Continue anyway?", vbYesNo, "Error number " & Err.Number & _
            " Has Occured")
            'Hopefully no a critical error, so give the option to continue...
            If myResponse = vbYes Then Resume Next Else
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    'Set my variables to hold what's in the property bags.
    myReels = PropBag.ReadProperty("Reels", "Reels")
    myPicsPerReel = PropBag.ReadProperty("PicsPerReel", "PicsPerReel")
    myNaturalFinish = PropBag.ReadProperty("NaturalFinish", "NaturalFinish")
    myHorizontalSpin = PropBag.ReadProperty("HorizontalSpin", "HorizontalSpin")
    myEnableHold = PropBag.ReadProperty("EnableHold", "EnableHold")
        
    Call SetReelPosition
    
End Sub

Private Sub UserControl_Resize()
       
    'Adjust the size of the control to fit the reels.
    If HorizontalSpin = False Then
        Width = (Reels * 32) * 15
        Height = (32 * 2) * 15
    Else
        Width = (32 * 2) * 15
        Height = (Reels * 32) * 15
    End If
    
End Sub

Public Property Get Reels() As Variant
Attribute Reels.VB_Description = "Returns/Sets the number of reels for the control."
    Reels = myReels
End Property

Public Property Let Reels(ByVal vNewValue As Variant)

If vNewValue < 1 Or vNewValue > 20 Then vNewValue = myReels

    myReels = vNewValue
    PropertyChanged "Reels"

    Call LoadUnloadControls(False)
    Call LoadUnloadControls(True)
    Call InitializeVariables
    Call ArrangePics
    Call SetReelPosition
    Call UserControl_Resize
        
End Property

Public Property Get PicsPerReel() As Variant
Attribute PicsPerReel.VB_Description = "Returns/Sets how many pictures are used for each reel.  NB: Includes zero."
    PicsPerReel = myPicsPerReel
End Property

Public Property Let PicsPerReel(ByVal vNewValue As Variant)

If vNewValue < 2 Or vNewValue > 5 Then vNewValue = myPicsPerReel

    myPicsPerReel = vNewValue
    PropertyChanged "PicsPerReel"

    Call LoadUnloadControls(True)
    Call InitializeVariables
    Call ArrangePics
    
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "Reels", myReels
    PropBag.WriteProperty "PicsPerReel", myPicsPerReel
    PropBag.WriteProperty "Size", 32
    PropBag.WriteProperty "NaturalFinish", myNaturalFinish
    PropBag.WriteProperty "HorizontalSpin", myHorizontalSpin
    PropBag.WriteProperty "EnableHold", myEnableHold
    
End Sub

Public Property Get NaturalFinish() As Boolean
Attribute NaturalFinish.VB_Description = "Returns/Sets whether each reel will finish smoothly or naturally."
    NaturalFinish = myNaturalFinish
End Property

Public Property Let NaturalFinish(ByVal vNewValue As Boolean)

If vNewValue = 1 Or vNewValue = True Then vNewValue = True Else vNewValue = False

    myNaturalFinish = vNewValue
    PropertyChanged "NaturalFinish"

End Property

Public Property Get HorizontalSpin() As Boolean
Attribute HorizontalSpin.VB_Description = "Returns/Sets whether the reel point and spin vertically, or horizontally.  Also re-shapes the reel."
    HorizontalSpin = myHorizontalSpin
End Property

Public Property Let HorizontalSpin(ByVal vNewValue As Boolean)

If vNewValue = 1 Or vNewValue = True Then vNewValue = True Else vNewValue = False

    myHorizontalSpin = vNewValue
    PropertyChanged "HorizontalSpin"

    Call InitializeVariables
    Call SetReelPosition
    Call UserControl_Resize
    
End Property

Public Function IsHeld(ByVal whatReel As Integer)
    
    IsHeld = shHold(whatReel).Visible

End Function

Public Property Get EnableHold() As Boolean
Attribute EnableHold.VB_Description = "Returnes/Sets whether clicking on a reel will prevent it from spinning and show the hold Shape."
    EnableHold = myEnableHold
End Property

Public Property Let EnableHold(ByVal vNewValue As Boolean)

If vNewValue = 1 Or vNewValue = True Then vNewValue = True Else vNewValue = False

    myEnableHold = vNewValue
    PropertyChanged "EnableHold"
    
    If myEnableHold = False Then
        Dim i As Integer
            For i = 0 To shHold.Count - 1
                If shHold(i).Visible = True Then shHold(i).Visible = False
            Next i
    End If

End Property

Public Sub StopSpinning(Optional ByVal alignReels As Boolean = True)

    'Will stop the reels immediately, and align unless told otherwise.
    stopNow = True
    doNotAlign = Not alignReels
    
    Dim whatReel As Long
        For whatReel = 0 To myReels - 1
            myMinCyclesForSpin(whatReel) = 0
        Next whatReel
    
End Sub

Public Function ReelPosition(ByVal whatReel As Long)
    'Returns the current reel position, from 1 to how ever many pictures there are
    'per reel.
    ReelPosition = GetSource(whatReel)
    If ReelPosition = 0 Then ReelPosition = myPicsPerReel + 1

End Function

