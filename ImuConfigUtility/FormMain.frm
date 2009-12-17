VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormMain 
   Caption         =   "IMU Configuration"
   ClientHeight    =   8385
   ClientLeft      =   5505
   ClientTop       =   4095
   ClientWidth     =   12450
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   559
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameTab 
      BorderStyle     =   0  'None
      Height          =   4035
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   12315
      Begin VB.Frame frame 
         Caption         =   "Info"
         Height          =   3915
         Left            =   0
         TabIndex        =   93
         Top             =   0
         Width           =   2535
         Begin VB.TextBox txtRefreshInt 
            Height          =   315
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   99
            Text            =   "5"
            Top             =   3000
            Width           =   975
         End
         Begin VB.TextBox txtCaptureCount 
            Height          =   285
            Left            =   1200
            MaxLength       =   5
            TabIndex        =   97
            Text            =   "1000"
            Top             =   2640
            Width           =   975
         End
         Begin VB.CheckBox chkCapture 
            Caption         =   "Start Capture"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   3420
            Width           =   2055
         End
         Begin VB.Label Label31 
            Caption         =   "Sample Rate:"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   3060
            Width           =   1035
         End
         Begin VB.Label Label28 
            Caption         =   "Max Samples:"
            Height          =   195
            Left            =   120
            TabIndex        =   96
            Top             =   2700
            Width           =   1035
         End
         Begin VB.Label lblVersion 
            AutoSize        =   -1  'True
            Caption         =   "Firmware Version"
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
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   2295
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Buttons"
         Height          =   3915
         Left            =   11460
         TabIndex        =   14
         Top             =   0
         Width           =   795
         Begin VB.Shape switchShape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   450
            Index           =   1
            Left            =   120
            Shape           =   3  'Circle
            Top             =   180
            Width           =   270
         End
         Begin VB.Label switch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   22
            Top             =   300
            Width           =   255
         End
         Begin VB.Shape switchShape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   450
            Index           =   2
            Left            =   120
            Shape           =   3  'Circle
            Top             =   600
            Width           =   270
         End
         Begin VB.Shape switchShape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   450
            Index           =   3
            Left            =   120
            Shape           =   3  'Circle
            Top             =   1020
            Width           =   270
         End
         Begin VB.Shape switchShape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   450
            Index           =   4
            Left            =   120
            Shape           =   3  'Circle
            Top             =   1500
            Width           =   270
         End
         Begin VB.Shape switchShape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   450
            Index           =   5
            Left            =   120
            Shape           =   3  'Circle
            Top             =   1980
            Width           =   270
         End
         Begin VB.Shape switchShape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   450
            Index           =   6
            Left            =   120
            Shape           =   3  'Circle
            Top             =   2400
            Width           =   270
         End
         Begin VB.Shape switchShape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   450
            Index           =   7
            Left            =   120
            Shape           =   3  'Circle
            Top             =   2820
            Width           =   270
         End
         Begin VB.Shape switchShape 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   450
            Index           =   8
            Left            =   120
            Shape           =   3  'Circle
            Top             =   3300
            Width           =   270
         End
         Begin VB.Label switch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   21
            Top             =   720
            Width           =   255
         End
         Begin VB.Label switch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   420
            TabIndex        =   20
            Top             =   1080
            Width           =   255
         End
         Begin VB.Label switch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   420
            TabIndex        =   19
            Top             =   1620
            Width           =   255
         End
         Begin VB.Label switch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   420
            TabIndex        =   18
            Top             =   2100
            Width           =   255
         End
         Begin VB.Label switch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   420
            TabIndex        =   17
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label switch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   420
            TabIndex        =   16
            Top             =   2940
            Width           =   255
         End
         Begin VB.Label switch 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   420
            TabIndex        =   15
            Top             =   3420
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Data"
         Height          =   3915
         Left            =   2580
         TabIndex        =   2
         Top             =   0
         Width           =   4935
         Begin VB.CheckBox chkChart 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   4200
            TabIndex        =   31
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox chkChart 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   4200
            TabIndex        =   30
            Top             =   2340
            Width           =   255
         End
         Begin VB.CheckBox chkChart 
            Caption         =   "Check1"
            Height          =   195
            Index           =   5
            Left            =   4200
            TabIndex        =   29
            Top             =   2640
            Width           =   255
         End
         Begin VB.CheckBox chkChart 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   4200
            TabIndex        =   25
            Top             =   1080
            Width           =   255
         End
         Begin VB.CheckBox chkChart 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   4200
            TabIndex        =   24
            Top             =   1380
            Width           =   255
         End
         Begin VB.CheckBox chkChart 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   4200
            TabIndex        =   23
            Top             =   1680
            Width           =   255
         End
         Begin VB.CheckBox chkChart 
            Caption         =   "Check1"
            Height          =   195
            Index           =   6
            Left            =   4200
            TabIndex        =   8
            Top             =   480
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.CheckBox chkChart 
            Caption         =   "Check1"
            Height          =   195
            Index           =   7
            Left            =   4200
            TabIndex        =   7
            Top             =   780
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.Label lblAdcGyroMid 
            Caption         =   "N/A"
            Height          =   195
            Index           =   0
            Left            =   1260
            TabIndex        =   46
            Top             =   2940
            Width           =   2775
         End
         Begin VB.Label lblAdcGyroMid 
            Caption         =   "N/A"
            Height          =   195
            Index           =   1
            Left            =   1260
            TabIndex        =   45
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label lblAdcGyroMid 
            Caption         =   "N/A"
            Height          =   195
            Index           =   2
            Left            =   1260
            TabIndex        =   44
            Top             =   3540
            Width           =   2775
         End
         Begin VB.Line Line32 
            X1              =   120
            X2              =   4800
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Label Label33 
            Caption         =   "Z-Zero Gyro"
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
            Left            =   120
            TabIndex        =   43
            Top             =   3540
            Width           =   1035
         End
         Begin VB.Label Label32 
            Caption         =   "Y-Zero Gyro"
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
            Left            =   120
            TabIndex        =   42
            Top             =   3240
            Width           =   1035
         End
         Begin VB.Line Line31 
            X1              =   120
            X2              =   4800
            Y1              =   3180
            Y2              =   3180
         End
         Begin VB.Label Label30 
            Caption         =   "X-Zero Gyro"
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
            Left            =   120
            TabIndex        =   41
            Top             =   2940
            Width           =   1035
         End
         Begin VB.Label Label29 
            Caption         =   "X Gyro"
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
            Left            =   120
            TabIndex        =   40
            Top             =   2040
            Width           =   1035
         End
         Begin VB.Label Label27 
            Caption         =   "Y Gyro"
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
            Left            =   120
            TabIndex        =   39
            Top             =   2340
            Width           =   1035
         End
         Begin VB.Label Label26 
            Caption         =   "Z Gyro"
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
            Left            =   120
            TabIndex        =   38
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Label Label25 
            Caption         =   "Z Accel."
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
            Left            =   120
            TabIndex        =   37
            Top             =   1680
            Width           =   1035
         End
         Begin VB.Label Label20 
            Caption         =   "Y Accel."
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
            Left            =   120
            TabIndex        =   36
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label Label19 
            Caption         =   "X Accel."
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
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label lblInpReading 
            Caption         =   "0V 0 deg/s"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   3
            Left            =   1260
            TabIndex        =   34
            Top             =   2100
            Width           =   2775
         End
         Begin VB.Label lblInpReading 
            Caption         =   "0V 0 deg/s"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   4
            Left            =   1260
            TabIndex        =   33
            Top             =   2400
            Width           =   2775
         End
         Begin VB.Shape shapeColor 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   3
            Left            =   4620
            Top             =   2040
            Width           =   150
         End
         Begin VB.Shape shapeColor 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   4
            Left            =   4620
            Top             =   2340
            Width           =   150
         End
         Begin VB.Shape shapeColor 
            BackColor       =   &H0018D3D3&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   5
            Left            =   4620
            Top             =   2640
            Width           =   150
         End
         Begin VB.Label lblInpReading 
            Caption         =   "0V 0 deg/s"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   5
            Left            =   1260
            TabIndex        =   32
            Top             =   2700
            Width           =   2775
         End
         Begin VB.Line Line30 
            X1              =   120
            X2              =   4800
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line29 
            X1              =   120
            X2              =   4800
            Y1              =   2580
            Y2              =   2580
         End
         Begin VB.Line Line28 
            X1              =   120
            X2              =   4800
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line23 
            X1              =   120
            X2              =   4800
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Label lblInpReading 
            Caption         =   "0V 0g"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   0
            Left            =   1260
            TabIndex        =   28
            Top             =   1140
            Width           =   2775
         End
         Begin VB.Label lblInpReading 
            Caption         =   "0V 0g"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   1
            Left            =   1260
            TabIndex        =   27
            Top             =   1440
            Width           =   2775
         End
         Begin VB.Label lblInpReading 
            Caption         =   "0V 0g"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   2
            Left            =   1260
            TabIndex        =   26
            Top             =   1740
            Width           =   2775
         End
         Begin VB.Shape shapeColor 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   0
            Left            =   4620
            Top             =   1080
            Width           =   150
         End
         Begin VB.Shape shapeColor 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   1
            Left            =   4620
            Top             =   1380
            Width           =   150
         End
         Begin VB.Shape shapeColor 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   2
            Left            =   4620
            Top             =   1680
            Width           =   150
         End
         Begin VB.Line Line27 
            X1              =   120
            X2              =   4800
            Y1              =   1620
            Y2              =   1620
         End
         Begin VB.Line Line24 
            X1              =   120
            X2              =   4800
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line8 
            X1              =   120
            X2              =   4800
            Y1              =   1020
            Y2              =   1020
         End
         Begin VB.Label Label5 
            Caption         =   "X Output"
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
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   915
         End
         Begin VB.Label lblOut 
            Caption         =   "255"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   0
            Left            =   1260
            TabIndex        =   10
            Top             =   540
            Width           =   2775
         End
         Begin VB.Label lblOut 
            Caption         =   "255"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1260
            TabIndex        =   9
            Top             =   840
            Width           =   2775
         End
         Begin VB.Shape shapeColor 
            BackColor       =   &H00800080&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   7
            Left            =   4620
            Top             =   780
            Width           =   150
         End
         Begin VB.Shape shapeColor 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   150
            Index           =   6
            Left            =   4620
            Top             =   480
            Width           =   150
         End
         Begin VB.Line Line26 
            X1              =   1200
            X2              =   1200
            Y1              =   180
            Y2              =   3780
         End
         Begin VB.Line Line25 
            X1              =   4140
            X2              =   4140
            Y1              =   180
            Y2              =   3780
         End
         Begin VB.Line Line19 
            X1              =   120
            X2              =   4800
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label Label24 
            Caption         =   "Chart"
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
            Left            =   4200
            TabIndex        =   6
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label23 
            Caption         =   "Last Value"
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
            Left            =   1260
            TabIndex        =   5
            Top             =   180
            Width           =   1695
         End
         Begin VB.Label Label18 
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
            Left            =   180
            TabIndex        =   4
            Top             =   180
            Width           =   915
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   4800
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label4 
            Caption         =   "Y Output"
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
            Left            =   120
            TabIndex        =   3
            Top             =   780
            Width           =   975
         End
      End
      Begin VB.Shape pointer 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         Height          =   150
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   1740
         Width           =   150
      End
      Begin VB.Shape pad 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   3825
         Left            =   7560
         Top             =   60
         Width           =   3825
      End
   End
   Begin VB.Frame frameTab 
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   360
      Width           =   12315
      Begin VB.Frame Frame4 
         Caption         =   "Output"
         Height          =   1815
         Left            =   8580
         TabIndex        =   132
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtOutScale 
            Height          =   285
            Index           =   0
            Left            =   720
            MaxLength       =   7
            TabIndex        =   136
            Text            =   "0000000"
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtOutScale 
            Height          =   285
            Index           =   1
            Left            =   720
            MaxLength       =   7
            TabIndex        =   135
            Text            =   "0000000"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtOutSmoothing 
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   7
            TabIndex        =   134
            Text            =   "000"
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtOutSmoothing 
            Height          =   285
            Index           =   1
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   133
            Text            =   "000"
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label17 
            Caption         =   "Y"
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
            Left            =   180
            TabIndex        =   141
            Top             =   960
            Width           =   255
         End
         Begin VB.Label Label16 
            Caption         =   "X"
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
            Left            =   180
            TabIndex        =   140
            Top             =   600
            Width           =   255
         End
         Begin VB.Line Line11 
            X1              =   60
            X2              =   3240
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label9 
            Caption         =   "Axis"
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
            Left            =   120
            TabIndex        =   139
            Top             =   240
            Width           =   435
         End
         Begin VB.Line Line12 
            X1              =   600
            X2              =   600
            Y1              =   240
            Y2              =   1380
         End
         Begin VB.Line Line18 
            X1              =   60
            X2              =   3240
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Label Label21 
            Caption         =   "Scale"
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
            Left            =   720
            TabIndex        =   138
            Top             =   240
            Width           =   795
         End
         Begin VB.Line Line20 
            X1              =   1560
            X2              =   1560
            Y1              =   240
            Y2              =   1380
         End
         Begin VB.Label Label2 
            Caption         =   "Smoothing"
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
            Left            =   1680
            TabIndex        =   137
            Top             =   240
            Width           =   915
         End
         Begin VB.Line Line3 
            X1              =   2640
            X2              =   2640
            Y1              =   240
            Y2              =   1380
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Configuration"
         Height          =   1755
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   3435
         Begin VB.CommandButton btnConfigWrite 
            Caption         =   "Write Device"
            Height          =   315
            Left            =   120
            TabIndex        =   102
            Top             =   660
            Width           =   1155
         End
         Begin VB.CommandButton btnConfigRead 
            Caption         =   "Read Device"
            Height          =   315
            Left            =   120
            TabIndex        =   101
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label14 
            Caption         =   "Device must be written before any change will take effect."
            Height          =   435
            Left            =   120
            TabIndex        =   142
            Top             =   1080
            Width           =   3075
         End
         Begin VB.Label lblLastWrite 
            Caption         =   "Last Write Time"
            Height          =   195
            Left            =   1320
            TabIndex        =   104
            Top             =   720
            Width           =   1995
         End
         Begin VB.Label lblLastRead 
            Caption         =   "Last Read Time"
            Height          =   195
            Left            =   1320
            TabIndex        =   103
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Accelerometer"
         Height          =   1815
         Left            =   3480
         TabIndex        =   67
         Top             =   0
         Width           =   5055
         Begin VB.TextBox txtZeroLevel 
            Height          =   285
            Index           =   0
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   79
            Text            =   "0.000"
            Top             =   540
            Width           =   555
         End
         Begin VB.TextBox txtInpSens 
            Height          =   285
            Index           =   0
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   78
            Text            =   "0000000"
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtZeroLevel 
            Height          =   285
            Index           =   1
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   77
            Text            =   "0.000"
            Top             =   960
            Width           =   555
         End
         Begin VB.TextBox txtInpSens 
            Height          =   285
            Index           =   1
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   76
            Text            =   "0000000"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtZeroLevel 
            Height          =   285
            Index           =   2
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   75
            Text            =   "0.000"
            Top             =   1380
            Width           =   555
         End
         Begin VB.TextBox txtInpSens 
            Height          =   285
            Index           =   2
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   74
            Text            =   "0000000"
            Top             =   1380
            Width           =   735
         End
         Begin VB.CheckBox chkInpInvert 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   1320
            TabIndex        =   73
            Top             =   600
            Width           =   255
         End
         Begin VB.CheckBox chkInpInvert 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   72
            Top             =   1020
            Width           =   255
         End
         Begin VB.CheckBox chkInpInvert 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   1320
            TabIndex        =   71
            Top             =   1440
            Width           =   255
         End
         Begin VB.ComboBox cbInpAnNum 
            Height          =   315
            Index           =   1
            ItemData        =   "FormMain.frx":014A
            Left            =   300
            List            =   "FormMain.frx":0163
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   960
            Width           =   795
         End
         Begin VB.ComboBox cbInpAnNum 
            Height          =   315
            Index           =   2
            ItemData        =   "FormMain.frx":018A
            Left            =   300
            List            =   "FormMain.frx":01A5
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   1380
            Width           =   795
         End
         Begin VB.ComboBox cbInpAnNum 
            Height          =   315
            Index           =   0
            ItemData        =   "FormMain.frx":01CC
            Left            =   300
            List            =   "FormMain.frx":01E7
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   540
            Width           =   795
         End
         Begin VB.Line Line6 
            X1              =   60
            X2              =   4980
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Line Line2 
            X1              =   120
            X2              =   4980
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Port"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   92
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblAn 
            Caption         =   "X"
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
            Left            =   120
            TabIndex        =   91
            Top             =   600
            Width           =   255
         End
         Begin VB.Line Line1 
            X1              =   1140
            X2              =   1140
            Y1              =   240
            Y2              =   1740
         End
         Begin VB.Label Label3 
            Caption         =   "Zero Level @ 0g"
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
            Left            =   3240
            TabIndex        =   90
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label label 
            Caption         =   "V (actual)"
            Height          =   195
            Index           =   0
            Left            =   3840
            TabIndex        =   89
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Sensitivity"
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
            Left            =   1800
            TabIndex        =   88
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label labelSensmVG 
            Caption         =   "mV/g"
            Height          =   195
            Index           =   0
            Left            =   2580
            TabIndex        =   87
            Top             =   600
            Width           =   435
         End
         Begin VB.Line Line7 
            X1              =   60
            X2              =   4980
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lblAn 
            Caption         =   "Y"
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
            Index           =   1
            Left            =   120
            TabIndex        =   86
            Top             =   1020
            Width           =   255
         End
         Begin VB.Label label 
            Caption         =   "V (actual)"
            Height          =   195
            Index           =   4
            Left            =   3840
            TabIndex        =   85
            Top             =   1020
            Width           =   1155
         End
         Begin VB.Label labelSensmVG 
            Caption         =   "mV/g"
            Height          =   195
            Index           =   1
            Left            =   2580
            TabIndex        =   84
            Top             =   1020
            Width           =   495
         End
         Begin VB.Label lblAn 
            Caption         =   "Z"
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
            Index           =   2
            Left            =   120
            TabIndex        =   83
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label label 
            Caption         =   "V (actual)"
            Height          =   195
            Index           =   2
            Left            =   3840
            TabIndex        =   82
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label labelSensmVG 
            Caption         =   "mV/g"
            Height          =   195
            Index           =   2
            Left            =   2580
            TabIndex        =   81
            Top             =   1440
            Width           =   495
         End
         Begin VB.Line Line9 
            X1              =   3180
            X2              =   3180
            Y1              =   240
            Y2              =   1740
         End
         Begin VB.Label Label7 
            Caption         =   "Invert"
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
            Left            =   1200
            TabIndex        =   80
            Top             =   240
            Width           =   495
         End
         Begin VB.Line Line10 
            Index           =   0
            X1              =   1740
            X2              =   1740
            Y1              =   240
            Y2              =   1740
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Gyroscope"
         Height          =   1815
         Left            =   3480
         TabIndex        =   47
         Top             =   1800
         Width           =   8715
         Begin VB.TextBox txtGyroNoise 
            Height          =   285
            Index           =   2
            Left            =   7080
            MaxLength       =   5
            TabIndex        =   125
            Text            =   "0.000"
            Top             =   1380
            Width           =   555
         End
         Begin VB.TextBox txtGyroNoise 
            Height          =   285
            Index           =   1
            Left            =   7080
            MaxLength       =   5
            TabIndex        =   123
            Text            =   "0.000"
            Top             =   960
            Width           =   555
         End
         Begin VB.TextBox txtGyroNoise 
            Height          =   285
            Index           =   0
            Left            =   7080
            MaxLength       =   5
            TabIndex        =   121
            Text            =   "0.000"
            Top             =   540
            Width           =   555
         End
         Begin VB.TextBox txtGyroDrift 
            Height          =   285
            Index           =   2
            Left            =   6000
            MaxLength       =   5
            TabIndex        =   119
            Text            =   "0.000"
            Top             =   1380
            Width           =   555
         End
         Begin VB.TextBox txtGyroDrift 
            Height          =   285
            Index           =   1
            Left            =   6000
            MaxLength       =   5
            TabIndex        =   117
            Text            =   "0.000"
            Top             =   960
            Width           =   555
         End
         Begin VB.TextBox txtGyroDrift 
            Height          =   285
            Index           =   0
            Left            =   6000
            MaxLength       =   5
            TabIndex        =   115
            Text            =   "0.000"
            Top             =   540
            Width           =   555
         End
         Begin VB.TextBox txtZeroLevel 
            Height          =   285
            Index           =   5
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   111
            Text            =   "0.000"
            Top             =   1380
            Width           =   555
         End
         Begin VB.TextBox txtZeroLevel 
            Height          =   285
            Index           =   4
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   110
            Text            =   "0.000"
            Top             =   960
            Width           =   555
         End
         Begin VB.TextBox txtZeroLevel 
            Height          =   285
            Index           =   3
            Left            =   3240
            MaxLength       =   5
            TabIndex        =   109
            Text            =   "0.000"
            Top             =   540
            Width           =   555
         End
         Begin VB.CheckBox chkGyroAutoZero 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   5280
            TabIndex        =   107
            Top             =   1440
            Width           =   255
         End
         Begin VB.CheckBox chkGyroAutoZero 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   5280
            TabIndex        =   106
            Top             =   1020
            Width           =   255
         End
         Begin VB.CheckBox chkGyroAutoZero 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   5280
            TabIndex        =   105
            Top             =   600
            Width           =   255
         End
         Begin VB.CheckBox chkInpInvert 
            Caption         =   "Check1"
            Height          =   195
            Index           =   4
            Left            =   1320
            TabIndex        =   56
            Top             =   1020
            Width           =   255
         End
         Begin VB.CheckBox chkInpInvert 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   1320
            TabIndex        =   55
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtInpSens 
            Height          =   285
            Index           =   4
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   54
            Text            =   "0000000"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtInpSens 
            Height          =   285
            Index           =   3
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   53
            Text            =   "0000000"
            Top             =   540
            Width           =   735
         End
         Begin VB.ComboBox cbInpAnNum 
            Height          =   315
            Index           =   3
            ItemData        =   "FormMain.frx":020E
            Left            =   300
            List            =   "FormMain.frx":0229
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   540
            Width           =   795
         End
         Begin VB.ComboBox cbInpAnNum 
            Height          =   315
            Index           =   4
            ItemData        =   "FormMain.frx":0250
            Left            =   300
            List            =   "FormMain.frx":026B
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   960
            Width           =   795
         End
         Begin VB.ComboBox cbInpAnNum 
            Height          =   315
            Index           =   5
            ItemData        =   "FormMain.frx":0292
            Left            =   300
            List            =   "FormMain.frx":02AD
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   1380
            Width           =   795
         End
         Begin VB.TextBox txtInpSens 
            Height          =   285
            Index           =   5
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   49
            Text            =   "0000000"
            Top             =   1380
            Width           =   735
         End
         Begin VB.CheckBox chkInpInvert 
            Caption         =   "Check1"
            Height          =   195
            Index           =   5
            Left            =   1320
            TabIndex        =   48
            Top             =   1440
            Width           =   255
         End
         Begin VB.Label Label8 
            Caption         =   "Max Noise"
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
            Index           =   2
            Left            =   7080
            TabIndex        =   128
            Top             =   240
            Width           =   1095
         End
         Begin VB.Line Line15 
            Index           =   2
            X1              =   7020
            X2              =   7020
            Y1              =   240
            Y2              =   1740
         End
         Begin VB.Label Label8 
            Caption         =   "Max Drift"
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
            Index           =   1
            Left            =   6000
            TabIndex        =   127
            Top             =   240
            Width           =   795
         End
         Begin VB.Label labelAnMid 
            Caption         =   "V"
            Height          =   195
            Index           =   9
            Left            =   6600
            TabIndex        =   126
            Top             =   1440
            Width           =   195
         End
         Begin VB.Label labelAnMid 
            Caption         =   "V"
            Height          =   195
            Index           =   8
            Left            =   6600
            TabIndex        =   124
            Top             =   1020
            Width           =   195
         End
         Begin VB.Label labelAnMid 
            Caption         =   "V"
            Height          =   195
            Index           =   7
            Left            =   6600
            TabIndex        =   122
            Top             =   600
            Width           =   195
         End
         Begin VB.Line Line15 
            Index           =   1
            X1              =   5940
            X2              =   5940
            Y1              =   240
            Y2              =   1740
         End
         Begin VB.Label labelAnMid 
            Caption         =   "V"
            Height          =   195
            Index           =   6
            Left            =   7680
            TabIndex        =   120
            Top             =   1440
            Width           =   195
         End
         Begin VB.Label labelAnMid 
            Caption         =   "V"
            Height          =   195
            Index           =   5
            Left            =   7680
            TabIndex        =   118
            Top             =   1020
            Width           =   195
         End
         Begin VB.Label labelAnMid 
            Caption         =   "V"
            Height          =   195
            Index           =   4
            Left            =   7680
            TabIndex        =   116
            Top             =   600
            Width           =   195
         End
         Begin VB.Label labelGyroZeroLevel 
            Caption         =   "V (per spec)"
            Height          =   195
            Index           =   2
            Left            =   3840
            TabIndex        =   114
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label labelGyroZeroLevel 
            Caption         =   "V (per spec)"
            Height          =   195
            Index           =   1
            Left            =   3840
            TabIndex        =   113
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label labelGyroZeroLevel 
            Caption         =   "V (per spec)"
            Height          =   195
            Index           =   0
            Left            =   3840
            TabIndex        =   112
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Auto Zero"
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
            Left            =   5040
            TabIndex        =   108
            Top             =   240
            Width           =   915
         End
         Begin VB.Line Line13 
            X1              =   1740
            X2              =   1740
            Y1              =   240
            Y2              =   1740
         End
         Begin VB.Label Label10 
            Caption         =   "Invert"
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
            Left            =   1200
            TabIndex        =   66
            Top             =   240
            Width           =   495
         End
         Begin VB.Line Line14 
            X1              =   3180
            X2              =   3180
            Y1              =   240
            Y2              =   1740
         End
         Begin VB.Label labelSensmVG 
            Caption         =   "mV/°/s"
            Height          =   195
            Index           =   4
            Left            =   2580
            TabIndex        =   65
            Top             =   1020
            Width           =   555
         End
         Begin VB.Label lblAn 
            Caption         =   "Y"
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
            Index           =   4
            Left            =   120
            TabIndex        =   64
            Top             =   1020
            Width           =   135
         End
         Begin VB.Line Line16 
            X1              =   60
            X2              =   8520
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Line Line17 
            X1              =   60
            X2              =   8520
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Label labelSensmVG 
            Caption         =   "mV/°/s"
            Height          =   195
            Index           =   5
            Left            =   2580
            TabIndex        =   63
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label11 
            Caption         =   "Sensitivity"
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
            Left            =   1800
            TabIndex        =   62
            Top             =   240
            Width           =   1275
         End
         Begin VB.Line Line21 
            X1              =   120
            X2              =   8520
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line Line22 
            X1              =   1140
            X2              =   1140
            Y1              =   240
            Y2              =   1740
         End
         Begin VB.Label lblAn 
            Caption         =   "X"
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
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label15 
            Caption         =   "Port"
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
            Left            =   300
            TabIndex        =   60
            Top             =   240
            Width           =   615
         End
         Begin VB.Line Line15 
            Index           =   0
            X1              =   4980
            X2              =   4980
            Y1              =   240
            Y2              =   1740
         End
         Begin VB.Label lblAn 
            Caption         =   "Z"
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
            Index           =   5
            Left            =   120
            TabIndex        =   59
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label labelSensmVG 
            Caption         =   "mV/°/s"
            Height          =   195
            Index           =   3
            Left            =   2580
            TabIndex        =   58
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label Label13 
            Caption         =   "Zero Level @ 0°/s "
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
            Left            =   3240
            TabIndex        =   57
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "General"
         Height          =   1815
         Left            =   0
         TabIndex        =   13
         Top             =   1800
         Width           =   3435
         Begin VB.CheckBox chkSwapButtons 
            Caption         =   "Swap buttons [1..4] with [5..8]"
            Height          =   375
            Left            =   120
            TabIndex        =   144
            Top             =   720
            Width           =   2535
         End
         Begin VB.ComboBox cbVdd 
            Height          =   315
            ItemData        =   "FormMain.frx":02D4
            Left            =   780
            List            =   "FormMain.frx":02DE
            TabIndex        =   130
            Text            =   "0.000"
            Top             =   240
            Width           =   855
         End
         Begin VB.Label label 
            Caption         =   "V"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   131
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "VDD (Supply Voltage):"
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
            Left            =   120
            TabIndex        =   129
            Top             =   300
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox chart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3825
      Left            =   0
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   829
      TabIndex        =   143
      Top             =   4500
      Width           =   12435
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   11820
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TabStrip tabStrip 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Monitor"
            Object.Tag             =   "info"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Configuration"
            Object.Tag             =   "config"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' vendor and product IDs
Private Const VendorID = 121
Private Const ProductID = 23823

' read and write buffers
Private Const VENDOR_DATA_POS = 4 ' this is  VENDOR_DATA_OFFSET (as defined in firmware) + 1
Private Const BufferInSize = 57
Private Const BufferOutSize = 54
Dim BufferIn(0 To BufferInSize) As Byte
Dim BufferOut(0 To BufferOutSize) As Byte
Dim init As Boolean
Dim refreshCount As Integer
Dim config As ConfigType
Dim info As InfoType
Dim configLoaded As Boolean
Dim lastChartV()
Dim captureStr As String
Dim captureCount As Long
Dim bufferPos As Integer
Dim strDeviceName As String


' ****************************************************************
' when the form loads, connect to the HID controller - pass
' the form window handle so that you can receive notification
' events...
'*****************************************************************
Private Sub Form_Load()
    Dim i, j
    lblVersion = "Device Unplugged"
    For i = 0 To INPUT_COUNT - 1
        cbInpAnNum(i).Clear
        For j = 0 To 12
            cbInpAnNum(i).AddItem "AN" & j
            cbInpAnNum(i).ItemData(cbInpAnNum(i).ListCount - 1) = j
        Next
        cbInpAnNum(i).AddItem "N/A"
        cbInpAnNum(i).ItemData(cbInpAnNum(i).ListCount - 1) = 255
        cbInpAnNum(i).ListIndex = cbInpAnNum(i).ListCount - 1
    Next
    ReDim lastChartV(0 To INPUT_COUNT + 1)
    For i = 0 To INPUT_COUNT + 1
        lastChartV(i) = -1
    Next
    
    Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    tabStrip_Click
    
    btnConfigWrite.Enabled = False
    btnConfigRead.Enabled = False
    
    ConnectToHID (Me.hWnd)
End Sub

'*****************************************************************
' disconnect from the HID controller...
'*****************************************************************
Private Sub Form_Unload(Cancel As Integer)
   DisconnectFromHID
End Sub

'*****************************************************************
' a HID device has been plugged in...
'*****************************************************************
Public Sub OnPlugged(ByVal pHandle As Long)
    Dim vid As Long, pid As Long
    vid = hidGetVendorID(pHandle)
    pid = hidGetProductID(pHandle)
    Debug.Print "OnPlugged VID: " & vid & " hex:" & Hex(vid) & " PID: " & pid & " hex:" & Hex(pid)
    If vid = VendorID And pid = ProductID Then
        Debug.Print "Plugged InputReportLength:" & hidGetInputReportLength(pHandle) & " OutputReportLength:" & hidGetOutputReportLength(pHandle)
        strDeviceName = String(255, " ")
        hidGetProductName pHandle, strDeviceName, 255
        strDeviceName = Left(strDeviceName, InStr(1, strDeviceName, Chr(0), vbBinaryCompare) - 1)
        lblVersion = "Device Plugged ..." & vbNewLine & strDeviceName
        configLoaded = False
        SetTxMode TX_MODE_CONFIG
    End If
End Sub

'*****************************************************************
' a HID device has been unplugged...
'*****************************************************************
Public Sub OnUnplugged(ByVal pHandle As Long)
   If hidGetVendorID(pHandle) = VendorID And hidGetProductID(pHandle) = ProductID Then
      Debug.Print "Unplugged"
      lblVersion = "Device Unplugged"
      btnConfigWrite.Enabled = False
      btnConfigRead.Enabled = False
   End If
End Sub

'*****************************************************************
' controller changed notification - called
' after ALL HID devices are plugged or unplugged
'*****************************************************************
Public Sub OnChanged()
   Dim DeviceHandle As Long
   
   ' get the handle of the device we are interested in, then set
   ' its read notify flag to true - this ensures you get a read
   ' notification message when there is some data to read...
   DeviceHandle = hidGetHandle(VendorID, ProductID)
   hidSetReadNotify DeviceHandle, True
End Sub

'*****************************************************************
' on read event...
'*****************************************************************
Public Sub OnRead(ByVal pHandle As Long)
    Dim i, strLine As String
    
    ' read the data (don't forget, pass the whole array)...
    If hidRead(pHandle, BufferIn(0)) Then
        ' IMPORTANT: first byte is the report ID, e.g. BufferIn(0)
        ' the other bytes are the data from the microcontrolller
    
       
       
        'handle veondor-defined data
        strLine = ""
        For i = 1 To VENDOR_DATA_POS
            strLine = strLine & BufferIn(i) & "," 'x,y,buttons output and tx mode
        Next i
        
        If TX_MODE_INFO = BufferIn(VENDOR_DATA_POS) Then
                info.sequence = BufferIn(VENDOR_DATA_POS + 1)
                info.interval = LbHb(BufferIn(VENDOR_DATA_POS + 2), BufferIn(VENDOR_DATA_POS + 3))
                strLine = strLine & info.sequence & "," & info.interval & ","      'sequence and interval in us

                
                For i = 0 To INPUT_COUNT - 1
                    info.adc(i) = LbHb(BufferIn(VENDOR_DATA_POS + 4 + i * 2), BufferIn(VENDOR_DATA_POS + 4 + i * 2 + 1))
                    strLine = strLine & info.adc(i) & ","  'adc outputs
                Next
                
                For i = 0 To 2
                    info.adcGyroMid(i) = LbHb(BufferIn(VENDOR_DATA_POS + 4 + INPUT_COUNT * 2 + i * 2), BufferIn(VENDOR_DATA_POS + 4 + INPUT_COUNT * 2 + i * 2 + 1))
                    strLine = strLine & info.adcGyroMid(i) & IIf(i < 2, ",", "") 'adcGyroMid outputs
                Next
                
                
                If (chkCapture.Value = vbChecked) Then
                    captureStr = captureStr & strLine & vbNewLine
                    captureCount = captureCount + 1
                    If captureCount >= StrToLong(txtCaptureCount, 1) Then
                        chkCapture.Value = vbUnchecked
                    End If
                End If
                    
        End If
   
  
        
        If TX_MODE_CONFIG = BufferIn(VENDOR_DATA_POS) Then
            For i = VENDOR_DATA_POS + 1 To BufferInSize
                strLine = strLine & BufferIn(i) & ","
            Next i
        
            If Not configLoaded Then
                ConfigLoad
                configLoaded = True
                lblLastRead = "Last Read Time: " & Format(Now, "hh:mm:ss")
                SetTxMode TX_MODE_INFO
            End If
            
            Debug.Print strLine
            
        End If
        
        
        
       
        
        
        'buttons
        For i = 1 To 8
          switchShape(i).BackColor = IIf((BufferIn(3) And (2 ^ (i - 1))) = 0, &HFFFFFF, &HFFC0C0)
        Next i
       
        
        refreshCount = refreshCount + 1
        Dim refreshInterval
        refreshInterval = StrToInt(txtRefreshInt, 1)
        If refreshCount >= refreshInterval And refreshInterval <> 0 Then
        
             
            'adcGyroMid
            For i = 0 To 2
                lblAdcGyroMid(i).Caption = IIf(info.adcGyroMid(i) = 65535, "N/A", Format(info.adcGyroMid(i) * (config.vdd / 1000) / 1023, "0.00V"))
            Next
        
        
            'last readings
            Dim v, s, m, d
            For i = 0 To INPUT_COUNT - 1
                v = info.adc(i) * (config.vdd / 1000) / 1023    'voltage in V
                s = StrToDouble(txtInpSens(i), 0)               'sensitivity mV/G , mV/deg/s
                m = config.zeroLevel(i) / 1000                  'zero-level (middle) value V
                d = v - m                                       'voltage relative to zero-level
                If (chkInpInvert(i).Value = vbChecked) Then d = -d 'invert input ?
                lblInpReading(i) = "AN" & IIf(config.inpAnNum(i) <> 255, config.inpAnNum(i), "?") & Format(info.adc(i), "=0000") & " " & Format(v, "0.000V ")
                If s > 0 Then
                    If i < 3 Then
                        lblInpReading(i) = lblInpReading(i) & Format(d * 1000 / s, "+0.00;-0.00") & "g"
                    Else
                        lblInpReading(i) = lblInpReading(i) & Format(d * 1000 / s, "+0.00;-0.00") & "deg/s"
                    End If
                End If
            Next
        
            'pad update
            pointer.Left = pad.Left + BufferIn(1) * Screen.TwipsPerPixelX - pointer.Width / 2
            pointer.Top = pad.Top + BufferIn(2) * Screen.TwipsPerPixelY - pointer.Height / 2
            
            'last readings
            lblOut(0).Caption = BufferIn(1)
            lblOut(1).Caption = BufferIn(2)
                    
            'chart
            Call BitBlt(chart.hDC, 0, 0, chart.Width - 1, chart.Height, chart.hDC, 1, 0, SRCCOPY)
            For i = 0 To chart.Height - 1
                Call SetPixelV(chart.hDC, chart.Width - 1, i, vbWhite)
            Next i
            
            Dim j
            For i = 0 To INPUT_COUNT + 1
                v = 0
                If i < INPUT_COUNT Then
                    If j < INPUT_COUNT Then v = Round(info.adc(i) * 255 / 1023)
                Else
                    v = BufferIn(1 + i - INPUT_COUNT)  ' output readings X, Y
                End If
                
                If (chkChart(i)) Then
                    If lastChartV(i) < 0 Then
                        chart.PSet (chart.Width - 1, 255 - v), shapeColor(i).BackColor
                    Else
                        chart.Line (chart.Width - 2, 255 - lastChartV(i))-(chart.Width - 1, 255 - v), shapeColor(i).BackColor
                    End If
                    lastChartV(i) = v
                Else
                    lastChartV(i) = -1
                End If
            Next
            
            chart.Refresh
            
            
            refreshCount = 0
        End If

      
      
    End If
End Sub



'*****************************************************************
' this is how you write some data...
'*****************************************************************

 
'converts  Low Byte and High Byte to word
Public Function LbHb(ByVal byte_low As Long, ByVal byte_high As Long) As Long
    LbHb = byte_low + byte_high * 256
End Function

Public Sub bufferPutByte(b As Byte)
    BufferOut(bufferPos) = b
    bufferPos = bufferPos + 1
End Sub

Public Sub bufferPutWord(w As Long)
    BufferOut(bufferPos) = w Mod 256    'low byte
    BufferOut(bufferPos + 1) = w \ 256  'high byte
    bufferPos = bufferPos + 2
End Sub

Public Function ConfigSave()
    Dim i
    Debug.Print "ConfigSave"
    bufferPos = 0
    
    bufferPutByte 0                  ' first byte is always the report ID
    bufferPutByte RX_CMD_CONFIG       ' command type
    
    'version
    bufferPutByte config.version
    
    'input invert
    config.inpInvert = 0
    For i = 0 To INPUT_COUNT - 1
        If chkInpInvert(i).Value = vbChecked Then config.inpInvert = config.inpInvert + 2 ^ i
    Next
    'swap buttons
    If chkSwapButtons.Value = vbChecked Then config.inpInvert = config.inpInvert + 2 ^ 6
    
    bufferPutByte config.inpInvert
    
    'analog ports assignments
    For i = 0 To INPUT_COUNT - 1
        config.inpAnNum(i) = cbInpAnNum(i).ItemData(cbInpAnNum(i).ListIndex)
        bufferPutByte config.inpAnNum(i)
    Next
           
    'acc zero level
    For i = 0 To 5
        config.zeroLevel(i) = Round(StrToDouble(txtZeroLevel(i).Text, 0) * 1000)
        bufferPutWord config.zeroLevel(i)
    Next
    
    'input sensibility
    For i = 0 To INPUT_COUNT - 1
        If i < 3 Then
            config.inpSens(i) = Round(StrToDouble(txtInpSens(i).Text, 0)) ' accelerometer mV/g
        Else
            config.inpSens(i) = Round(StrToDouble(txtInpSens(i).Text, 0) * 1000) ' gyro convert mV/deg/s to uV/deg/s
        End If
        bufferPutWord config.inpSens(i)
    Next
    
    'output scale
    For i = 0 To 1
        config.outScale(i) = Round(StrToDouble(txtOutScale(i).Text, 0) * 1000)
        bufferPutWord config.outScale(i)
    Next
    
    'gyro weight
    For i = 0 To 1
        config.outSmoothing(i) = Round(StrToDouble(txtOutSmoothing(i).Text, 0))
        bufferPutByte config.outSmoothing(i)
    Next i
    
    
    
    'VDD
    config.vdd = Round(StrToDouble(cbVdd.Text, 0) * 1000)
    bufferPutWord config.vdd
    
          
    'gyroNoise
    For i = 0 To 2
        config.gyroNoise(i) = Round(StrToDouble(txtGyroNoise(i).Text, 0) * 1000)
        bufferPutWord config.gyroNoise(i)
    Next
    
    
    'gyroDrift
    For i = 0 To 2
        config.gyroDrift(i) = Round(StrToDouble(txtGyroDrift(i).Text, 0) * 1000)
        bufferPutWord config.gyroDrift(i)
    Next
    
    'autoZero
    config.gyroAutoZero = 0
    For i = 0 To 2
        If chkGyroAutoZero(i).Value = vbChecked Then config.gyroAutoZero = config.gyroAutoZero + 2 ^ i
    Next
    bufferPutByte config.gyroAutoZero
        
    
    'debug print
    For i = 0 To BufferOutSize
        Debug.Print BufferOut(i);
    Next
    Debug.Print
    
    'after writing config, device will switch automatically to TX_MODE_CONFIG (and send us the new config)
    configLoaded = False 'prepare to load config after write
    lblLastWrite = "Last Write Time: " & Format(Now, "hh:mm:ss")
    hidWriteEx VendorID, ProductID, BufferOut(0)
    
End Function

Public Function bufferGetByte() As Byte
    bufferGetByte = BufferIn(bufferPos)
    bufferPos = bufferPos + 1
    Debug.Print "*b*"; bufferPos
End Function

Public Function bufferGetWord() As Long
    bufferGetWord = LbHb(BufferIn(bufferPos), BufferIn(bufferPos + 1))
    bufferPos = bufferPos + 2
    Debug.Print "*w*"; bufferPos
End Function

Public Function CheckConfigVersion()
    CheckConfigVersion = True
    If (config.version <> CONFIG_VERSION) Then
        MsgBox "Device has wrong configuration version " & config.version & vbNewLine & _
        "This software was build for configuration version " & CONFIG_VERSION, vbCritical
        CheckConfigVersion = False
    End If
End Function

Public Function ConfigLoad()
    Dim i, j
    
    btnConfigWrite.Enabled = False
    btnConfigRead.Enabled = False
    
    'version
    bufferPos = VENDOR_DATA_POS + 1
    config.version = bufferGetByte
    lblVersion.Caption = "Device Plugged." & vbNewLine & _
        strDeviceName & vbNewLine & _
        "Configuration Version " & (config.version)
    
    If (Not CheckConfigVersion) Then
        ConfigLoad = False
        btnConfigRead.Enabled = True
        Exit Function
    End If
    
    'input invert
    config.inpInvert = bufferGetByte
    For i = 0 To INPUT_COUNT - 1
        chkInpInvert(i).Value = IIf(0 <> (config.inpInvert And 2 ^ i), vbChecked, vbUnchecked)
    Next
    'swap buttons
    chkSwapButtons.Value = IIf(0 <> (config.inpInvert And 2 ^ 6), vbChecked, vbUnchecked)
     
    'analog ports assignments
    For i = 0 To INPUT_COUNT - 1
        config.inpAnNum(i) = bufferGetByte
        cbInpAnNum(i).ListIndex = cbInpAnNum(i).ListCount - 1
        For j = 0 To cbInpAnNum(i).ListCount - 1
           If cbInpAnNum(i).ItemData(j) = config.inpAnNum(i) Then cbInpAnNum(i).ListIndex = j
        Next
    Next
        
    'acc zero level
    For i = 0 To 5
         config.zeroLevel(i) = bufferGetWord
         txtZeroLevel(i).Text = config.zeroLevel(i) / 1000
    Next
        
    'input sensibility
    For i = 0 To INPUT_COUNT - 1
        config.inpSens(i) = bufferGetWord
        If i < 3 Then
            txtInpSens(i).Text = config.inpSens(i) ' accelerometer mV/g
        Else
            txtInpSens(i).Text = config.inpSens(i) / 1000 ' gyro convert uV/deg/s to mV/deg/s
        End If
    Next
    
    'output scale
    For i = 0 To 1
        config.outScale(i) = bufferGetWord
        txtOutScale(i).Text = config.outScale(i) / 1000
    Next
    
    'output smoothing
    For i = 0 To 1
        config.outSmoothing(i) = bufferGetByte
        txtOutSmoothing(i).Text = config.outSmoothing(i)
    Next
    
    'VDD
    config.vdd = bufferGetWord
    cbVdd.Text = config.vdd / 1000
   

    'gyroNoise
    For i = 0 To 2
         config.gyroNoise(i) = bufferGetWord
         txtGyroNoise(i).Text = config.gyroNoise(i) / 1000
    Next
    
    
    'gyroDrift
    For i = 0 To 2
         config.gyroDrift(i) = bufferGetWord
         txtGyroDrift(i).Text = config.gyroDrift(i) / 1000
    Next
    
    'autoZero
    config.gyroAutoZero = bufferGetByte
    For i = 0 To 2
        chkGyroAutoZero(i).Value = IIf(0 <> (config.gyroAutoZero And 2 ^ i), vbChecked, vbUnchecked)
    Next
     
    btnConfigWrite.Enabled = True
    btnConfigRead.Enabled = True
    ConfigLoad = True
     
End Function

Private Sub btnConfigRead_Click()
    configLoaded = False
    SetTxMode TX_MODE_CONFIG
End Sub


Private Sub SetTxMode(txMode As Byte)
    Debug.Print "SetTxMode:" & txMode
    BufferOut(0) = 0   ' first by is always the report ID
    BufferOut(1) = RX_CMD_SET_TX_MODE
    BufferOut(2) = txMode
    hidWriteEx VendorID, ProductID, BufferOut(0)
End Sub


Private Sub btnConfigWrite_Click()
    ConfigSave
End Sub

Private Sub chkCapture_Click()
    
    If chkCapture.Value = vbChecked Then
        chkCapture.Caption = "Stop Capture"
        captureStr = "OutX,OutY,ButtonsBits,Cmd,Sequence,IntervalUs,AN0,AN1,AN2,AN3,AN4,AN5,adcGyroMidX,adcGyroMidY,adcGyroMidZ" & vbNewLine
        captureCount = 0
    Else
        chkCapture.Caption = "Start Capture"
        cd.DialogTitle = "Save " & captureCount & " Captured Records As"
        cd.DefaultExt = ".csv"
        cd.Filter = "Comma Separated Values (*.csv)|*.csv"
        cd.CancelError = True
        On Error Resume Next
save_lbl:
        cd.ShowSave
        
        If FileExists(cd.FileName) Then
            If vbNo = MsgBox("File " & cd.FileName & " already exists, overwrite ?", vbYesNo + vbExclamation) Then
                GoTo save_lbl
            End If
        End If
        
        If Err Then
            Exit Sub
        End If
        
        Str2File captureStr, cd.FileName, True
        If Err Then
            MsgBox "Failed to save file " & vbNewLine & cd.FileName & vbNewLine & "Error:" & Err.Description, vbCritical
            GoTo save_lbl
        End If
    End If
    
    
End Sub



'VALIDATION
Private Sub lblHelp_Click()
    ShellExecute hWnd, "open", "http://www.starlino.com/", vbNullString, vbNullString, Empty
End Sub


Private Sub tabStrip_Click()
    Dim i
    For i = 1 To tabStrip.Tabs.Count
        frameTab(i).Visible = (i = tabStrip.SelectedItem.Index)
    Next
End Sub

Private Sub txtOutScale_Validate(Index As Integer, Cancel As Boolean)
    Dim v
    v = StrToDouble(txtOutScale(Index), -1)
    If v < 0 Or v > 65 Then
        MsgBox "Output scale must be in the range of 0 to 65", vbCritical
        Cancel = True
    End If
End Sub


Private Sub txtZeroLevel_Validate(Index As Integer, Cancel As Boolean)
    Dim v
    v = StrToDouble(txtZeroLevel(Index), -1)
    If v < 0 Or v > 5 Then
        MsgBox "Zero-Level (0g) voltage must be in the range of 0 to 5", vbCritical
        Cancel = True
    End If
End Sub


Private Sub txtInpSens_Validate(Index As Integer, Cancel As Boolean)
    Dim v
    v = StrToDouble(txtInpSens(Index), -1)
    
    If Index < 3 Then
        If v < 0 Or v > 65535 Then
            MsgBox "Accelerometer sensibility must be in the range of 0 to 65535", vbCritical
            Cancel = True
        End If
    Else
        If v < 0 Or v > 65 Then
            MsgBox "Gyro sensibility must be in the range of 0 to 65", vbCritical
            Cancel = True
        End If
    End If
    
End Sub

Private Sub txtOutSmoothing_Validate(Index As Integer, Cancel As Boolean)
    Dim v
    v = StrToDouble(txtOutSmoothing(Index), -1)
    If v < 0 Or v > 255 Then
        MsgBox "Smoothing must be in the range of 0 to 255", vbCritical
        Cancel = True
    End If
End Sub


Private Sub txtCaptureCount_Validate(Cancel As Boolean)
     Dim v
    v = StrToInt(txtCaptureCount, -1)
    If v < 0 Or v > 65535 Then
        MsgBox "Max number of samples to capture must be in the range of 0 to 65535", vbCritical
        Cancel = True
    End If
End Sub



