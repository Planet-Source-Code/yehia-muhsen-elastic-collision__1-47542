VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elastic Collision"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmBG 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton CmdStop 
         Caption         =   "&Stop"
         Height          =   495
         Left            =   2640
         TabIndex        =   23
         Top             =   5160
         Width           =   2295
      End
      Begin VB.CommandButton CmdExit 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   495
         Left            =   5400
         TabIndex        =   22
         Top             =   5160
         Width           =   2295
      End
      Begin VB.CommandButton Obj2 
         BackColor       =   &H008080FF&
         Caption         =   "2"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.CheckBox ChkSound 
         Caption         =   "Enable Sound"
         Height          =   255
         Left            =   4200
         TabIndex        =   20
         Top             =   4800
         Width           =   3495
      End
      Begin VB.CheckBox ChkStop 
         Caption         =   "Stop movements after collision"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   4800
         Width           =   3495
      End
      Begin VB.CommandButton Obj1 
         BackColor       =   &H00FF8080&
         Caption         =   "1"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   840
         Width           =   615
      End
      Begin VB.VScrollBar VScale 
         Height          =   1335
         LargeChange     =   5
         Left            =   7440
         Max             =   200
         Min             =   1
         TabIndex        =   15
         Top             =   1890
         Value           =   5
         Width           =   255
      End
      Begin VB.PictureBox ObjGraph 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1785
         ScaleWidth      =   7260
         TabIndex        =   14
         Top             =   1680
         Width           =   7290
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "&Go"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Caption         =   "Object # 2"
         Height          =   1095
         Left            =   3960
         TabIndex        =   2
         Top             =   3600
         Width           =   3735
         Begin VB.TextBox Obj2Speed 
            Height          =   285
            Left            =   960
            TabIndex        =   11
            Text            =   "-4.000"
            Top             =   600
            Width           =   2055
         End
         Begin VB.HScrollBar Obj2Size 
            Height          =   255
            LargeChange     =   2
            Left            =   840
            Max             =   10
            TabIndex        =   6
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label6 
            Caption         =   "m/s"
            Height          =   255
            Left            =   3240
            TabIndex        =   12
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label4 
            Caption         =   "Velocity :"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Size      :"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Object # 1"
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   3600
         Width           =   3735
         Begin VB.TextBox Obj1Speed 
            Height          =   285
            Left            =   960
            TabIndex        =   8
            Text            =   "4.000"
            Top             =   600
            Width           =   2055
         End
         Begin VB.HScrollBar Obj1Size 
            Height          =   255
            LargeChange     =   2
            Left            =   840
            Max             =   10
            TabIndex        =   4
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label5 
            Caption         =   "m/s"
            Height          =   255
            Left            =   3240
            TabIndex        =   9
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Velocity :"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Size      :"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label8 
         Caption         =   " +"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7440
         TabIndex        =   17
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "  -"
         ForeColor       =   &H000000C0&
         Height          =   135
         Left            =   7440
         TabIndex        =   16
         Top             =   1680
         Width           =   255
      End
      Begin VB.Shape BG 
         BorderColor     =   &H00808080&
         FillColor       =   &H00E0E0E0&
         FillStyle       =   2  'Horizontal Line
         Height          =   1360
         Left            =   120
         Top             =   220
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed By: Yehia Muhsen
'Date : 8-6-2003
'Description: This program is an application for the elastic collision
'             between two objects. In elastic collision, the total momentum
'             is always conserved, and the final velocities get distributed
'             according to the famous equation of elastic collision, which is
'             based on the masses and initial velocities.
'             This program is a good example of how to use math to move objects,
'             detect collisions, draw graphs, and change objects sizes and positions.
'             In this program, I made sure that all control objects get disabled during
'             the movement except the Stop button, and I used the idea of global variables
'             to exit the loop of movement.
'
'             ***I hope the code is easy to understand***
'
'Eamil: If you have any question plase email me (yehia_sm@hotmail.com)

Option Explicit
Dim StopAction As Boolean

Private Sub CmdExit_Click()
Unload Form1
End Sub

Private Sub CmdGo_Click()
'variables
Dim Spd1 As Single, Spd2 As Single, SpdTemp As Single
Dim OldSpeed1 As Single, OldSpeed2 As Single
Dim Distance1 As Single, Distance2 As Single
Dim Sz1 As Long, Sz2 As Long
Dim nScale As Byte
Dim I As Single
Dim Re As Byte, CollisionOccured As Boolean

'Disable some elements
Obj1Size.Enabled = False
Obj1Speed.Enabled = False
Obj2Size.Enabled = False
Obj2Speed.Enabled = False
CmdGo.Enabled = False
CmdExit.Enabled = False
CmdStop.SetFocus

'Get data
Spd1 = Val(Obj1Speed)
OldSpeed1 = Spd1
Spd2 = Val(Obj2Speed)
OldSpeed2 = Spd2
Sz1 = Obj1.Width * Obj1.Height
Sz2 = Obj2.Width * Obj2.Height

'Pre-Action
Distance1 = Obj1.Left
Distance2 = Obj2.Left
StopAction = False
CollisionOccured = False
ObjGraph.Cls
'Action

Do
    
    
    'When collision occures
    If (Obj1.Left + Obj1.Width) >= Obj2.Left Then
        SpdTemp = Spd1
        Spd1 = ((Sz1 - Sz2) / (Sz1 + Sz2)) * Spd1 + 2 * Sz2 * Spd2 / (Sz1 + Sz2)
        Spd2 = ((Sz2 - Sz1) / (Sz1 + Sz2)) * Spd2 + 2 * Sz1 * SpdTemp / (Sz1 + Sz2)
        If ChkSound = 1 Then Beep
        
        CollisionOccured = True
        
        'Prevent one object goes inside another
        'Obj2.Left = Obj1.Left + Obj1.Width
    End If
    
    'When hitting the wall
    If Obj1.Left <= BG.Left Then
        
        'Change velocity to the opposite direction
        Spd1 = -Spd1
        
        'Prevent the object from being stuck
        'Obj1.Left = BG.Left

        If ChkSound = 1 Then Beep
        
    End If
    
    If (Obj2.Left + Obj2.Width) >= (BG.Left + BG.Width) Then
        
        'Change velociy to the opposite direction
        Spd2 = -Spd2
        
        'Prevent the object from being stuck
        'Obj2.Left = BG.Left + BG.Width - Obj2.Width

        '
        If ChkSound = 1 Then Beep
    End If
    
    'Move objects
    'Using Distance1 and Distance2 is very helpful to keep decimal parts
    Distance1 = Distance1 + Spd1
    Distance2 = Distance2 + Spd2
    
    Obj1.Left = Distance1
    Obj2.Left = Distance2
 
    Obj1Speed = Format(Spd1, "0.000")
    Obj2Speed = Format(Spd2, "0.000")

    'Show the graph
    
    'When changing the scale, start the graph over
    If Not nScale = VScale.Value Then
        I = 0
        ObjGraph.Cls
    End If
    nScale = VScale.Value
    I = I + nScale / 20
    
    'Clear graph when the screen is full
    If I >= ObjGraph.ScaleWidth Then I = 0: ObjGraph.Cls
    ObjGraph.PSet (I, (Obj1.Left + Obj1.Width - BG.Left) * (ObjGraph.ScaleHeight / BG.Width)), vbBlue
    ObjGraph.PSet (I, (Obj2.Left - BG.Left) * (ObjGraph.ScaleHeight / BG.Width)), vbRed
    
    'Stop Movement if required
    If ChkStop = 1 And CollisionOccured Then
        Re = MsgBox("Do you want to continue", vbQuestion + vbYesNo, "Collisioin occured")
        If Re = vbNo Then CmdStop_Click
    End If
    
    '
    CollisionOccured = False
    
    '
    DoEvents
Loop Until StopAction

'Re-enable some elements
Obj1Speed = OldSpeed1
Obj2Speed = OldSpeed2
Obj1Size.Enabled = True
Obj1Speed.Enabled = True
Obj2Size.Enabled = True
Obj2Speed.Enabled = True
CmdGo.Enabled = True
CmdExit.Enabled = True
CmdGo.SetFocus
End Sub

Private Sub CmdStop_Click()
Form_Load
End Sub

Private Sub Form_Load()
'
StopAction = True
'
Obj1Size_Change
Obj2Size_Change

End Sub


Private Sub Obj1Size_Change()
Dim SizeChange As Long

SizeChange = Obj1Size.Value * 70

Obj1.Width = 615 + SizeChange
Obj1.Height = 615 + SizeChange
Obj1.Left = 940 - SizeChange

Obj1.Top = BG.Top + (BG.Height / 2) - (Obj1.Height / 2)
End Sub

Private Sub Obj2Size_Change()
Dim SizeChange As Long

SizeChange = Obj2Size.Value * 70

Obj2.Width = 615 + SizeChange
Obj2.Height = 615 + SizeChange
Obj2.Left = 6250

Obj2.Top = BG.Top + (BG.Height / 2) - (Obj2.Height / 2)
End Sub
