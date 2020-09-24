VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7FDF243A-2E06-4F93-989D-6C9CC526FFC5}#10.0#0"; "HoverButton.ocx"
Begin VB.Form Options 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4530
   ClientLeft      =   7545
   ClientTop       =   2385
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOptions.frx":1042
   ScaleHeight     =   4530
   ScaleWidth      =   5745
   Begin VB.PictureBox picWinIface 
      BorderStyle     =   0  'None
      Height          =   3870
      Left            =   2070
      ScaleHeight     =   3870
      ScaleWidth      =   3600
      TabIndex        =   30
      Top             =   90
      Visible         =   0   'False
      Width           =   3600
      Begin VB.TextBox txtButtonWidth 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2490
         MaxLength       =   4
         TabIndex        =   59
         Text            =   "100"
         Top             =   345
         Width           =   540
      End
      Begin VB.CheckBox chkFlatButtons 
         Caption         =   " &Flat Taskbar Buttons "
         Height          =   240
         Left            =   60
         TabIndex        =   49
         Top             =   675
         Width           =   2745
      End
      Begin VB.Frame Frame1 
         Caption         =   " Windows 2000 Options "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CC9F7C&
         Height          =   660
         Left            =   60
         TabIndex        =   45
         Top             =   1620
         Width           =   3495
         Begin VB.TextBox txtTrans 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   2250
            MaxLength       =   3
            TabIndex        =   47
            Text            =   "0"
            Top             =   255
            Width           =   405
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2730
            TabIndex        =   48
            Top             =   300
            Width           =   135
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Window Translucency &:"
            Height          =   195
            Left            =   405
            TabIndex        =   46
            Top             =   285
            Width           =   1710
         End
      End
      Begin VB.CheckBox chkAlwaysShowST 
         Caption         =   " &Always show in System Tray"
         Height          =   225
         Left            =   60
         TabIndex        =   40
         Top             =   1275
         Width           =   3150
      End
      Begin VB.CheckBox chkMinToSystray 
         Caption         =   " &Minimize to System Tray "
         Height          =   225
         Left            =   60
         TabIndex        =   39
         Top             =   975
         Width           =   3150
      End
      Begin VB.CheckBox chkStretch 
         Caption         =   " &Stretch taskbar buttons to fit taskbar "
         Height          =   225
         Left            =   60
         TabIndex        =   31
         Top             =   60
         Width           =   3150
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Button Width &:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1215
         TabIndex        =   58
         Top             =   375
         Width           =   1140
      End
   End
   Begin VB.PictureBox picIRC 
      BorderStyle     =   0  'None
      Height          =   3930
      Left            =   2055
      ScaleHeight     =   3930
      ScaleWidth      =   3585
      TabIndex        =   32
      Top             =   60
      Width           =   3585
      Begin VB.CheckBox chkShowLag 
         Caption         =   " Show Lag when joining channels "
         Height          =   255
         Left            =   75
         TabIndex        =   57
         Top             =   1710
         Width           =   3360
      End
      Begin RichTextLib.RichTextBox rtfQuitMsg 
         Height          =   345
         Left            =   195
         TabIndex        =   53
         Top             =   1260
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   609
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"frmOptions.frx":1D0C
      End
      Begin VB.CheckBox chkHidePing 
         Caption         =   " Hide Ping? Pong! "
         Height          =   255
         Left            =   90
         TabIndex        =   38
         Top             =   690
         Width           =   2700
      End
      Begin VB.CheckBox chkRejoin 
         Caption         =   " Rejoin Channels when kicked "
         Height          =   255
         Left            =   90
         TabIndex        =   37
         Top             =   375
         Width           =   2700
      End
      Begin VB.CheckBox chkWhoIsQ 
         Caption         =   " Do WhoIs in Queries"
         Height          =   255
         Left            =   90
         TabIndex        =   33
         Top             =   75
         Width           =   3405
      End
      Begin VB.Label Label17 
         Caption         =   "Default Quit Message &:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   105
         TabIndex        =   52
         Top             =   1005
         Width           =   2820
      End
   End
   Begin VB.PictureBox picDisplay 
      BorderStyle     =   0  'None
      Height          =   3960
      Left            =   2100
      ScaleHeight     =   3960
      ScaleWidth      =   3675
      TabIndex        =   21
      Top             =   45
      Width           =   3675
      Begin HoverButton.Button cmdBGDef 
         Height          =   330
         Left            =   1740
         TabIndex        =   63
         Top             =   75
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "Def"
         CaptionDown     =   "&Def"
         CaptionOver     =   "&Def"
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button Button1 
         Height          =   300
         Left            =   3225
         TabIndex        =   60
         Top             =   2535
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   529
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "..."
         CaptionDown     =   "..."
         CaptionOver     =   "&..."
         ShowFocusRect   =   -1  'True
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin VB.CheckBox chkTileImg 
         Alignment       =   1  'Right Justify
         Caption         =   " &Tile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2865
         TabIndex        =   56
         Top             =   2250
         Width           =   675
      End
      Begin VB.CheckBox chkBGImage 
         Caption         =   " &Client background-image"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   51
         Top             =   2265
         Width           =   2565
      End
      Begin VB.TextBox txtBGImage 
         Height          =   315
         Left            =   390
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2535
         Width           =   2760
      End
      Begin VB.PictureBox picClientBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2280
         ScaleHeight     =   135
         ScaleWidth      =   255
         TabIndex        =   43
         Top             =   1545
         Width           =   285
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   315
         Left            =   2790
         TabIndex        =   36
         Text            =   "8.25"
         Top             =   1875
         Width           =   765
      End
      Begin VB.ComboBox cmbFontName 
         Height          =   315
         Left            =   780
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1875
         Width           =   1965
      End
      Begin VB.PictureBox picRightColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2280
         ScaleHeight     =   135
         ScaleWidth      =   255
         TabIndex        =   28
         Top             =   1200
         Width           =   285
      End
      Begin VB.PictureBox picLeftColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2280
         ScaleHeight     =   135
         ScaleWidth      =   255
         TabIndex        =   26
         Top             =   855
         Width           =   285
      End
      Begin VB.PictureBox picForeColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2280
         ScaleHeight     =   135
         ScaleWidth      =   255
         TabIndex        =   22
         Top             =   510
         Width           =   285
      End
      Begin VB.PictureBox picBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   2280
         ScaleHeight     =   135
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   165
         Width           =   285
      End
      Begin HoverButton.Button cmdFGDEf 
         Height          =   330
         Left            =   1740
         TabIndex        =   64
         Top             =   420
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "Def"
         CaptionDown     =   "&Def"
         CaptionOver     =   "&Def"
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdLBDef 
         Height          =   330
         Left            =   1740
         TabIndex        =   65
         Top             =   765
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "Def"
         CaptionDown     =   "&Def"
         CaptionOver     =   "&Def"
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdRCDef 
         Height          =   330
         Left            =   1740
         TabIndex        =   66
         Top             =   1110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "Def"
         CaptionDown     =   "&Def"
         CaptionOver     =   "&Def"
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdCBDef 
         Height          =   330
         Left            =   1740
         TabIndex        =   67
         Top             =   1455
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "Def"
         CaptionDown     =   "&Def"
         CaptionOver     =   "&Def"
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdBGColor 
         Height          =   330
         Left            =   2190
         TabIndex        =   68
         Top             =   75
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "       Change..."
         CaptionDown     =   "       &Change..."
         CaptionOver     =   "       &Change..."
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdForeColor 
         Height          =   330
         Left            =   2190
         TabIndex        =   69
         Top             =   420
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "       Change..."
         CaptionDown     =   "       &Change..."
         CaptionOver     =   "       &Change..."
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdLeftColor 
         Height          =   330
         Left            =   2190
         TabIndex        =   70
         Top             =   765
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "       Change..."
         CaptionDown     =   "       &Change..."
         CaptionOver     =   "       &Change..."
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdRightColor 
         Height          =   330
         Left            =   2190
         TabIndex        =   71
         Top             =   1110
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "       Change..."
         CaptionDown     =   "       &Change..."
         CaptionOver     =   "       &Change..."
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin HoverButton.Button cmdClientBack 
         Height          =   330
         Left            =   2190
         TabIndex        =   72
         Top             =   1455
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BackColor       =   -2147483633
         HoverBackColor  =   14268829
         Border          =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontOverX {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   10649190
         HilightColor    =   -2147483633
         ShadowColor     =   -2147483633
         HoverHilightColor=   14928823
         HoverShadowColor=   13410172
         ForeColor       =   -2147483630
         HoverForeColor  =   4210752
         Caption         =   "       Change..."
         CaptionDown     =   "       &Change..."
         CaptionOver     =   "       &Change..."
         ShowFocusRect   =   0   'False
         Sink            =   -1  'True
         Style           =   0
         PictureLocation =   0
         ButtonStyleX    =   0
         State           =   0
         IconHeight      =   0
         IconWidth       =   0
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Client Back Color:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   44
         Top             =   1515
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Font:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   34
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Right Back Color:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   29
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Left Back Color:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   27
         Top             =   825
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "BackGround Color:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   25
         Top             =   135
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "ForeGround Color:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   24
         Top             =   480
         Width           =   1530
      End
   End
   Begin VB.PictureBox picConnecting 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   2100
      ScaleHeight     =   3900
      ScaleWidth      =   3630
      TabIndex        =   12
      Top             =   45
      Width           =   3630
      Begin VB.TextBox txtAutoJoin 
         Height          =   300
         Left            =   195
         TabIndex        =   42
         Text            =   "#projectIRC"
         Top             =   3570
         Width           =   3345
      End
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   2955
         MaxLength       =   5
         TabIndex        =   20
         Text            =   "6667"
         Top             =   2955
         Width           =   600
      End
      Begin VB.TextBox txtRetry 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1650
         MaxLength       =   2
         TabIndex        =   18
         Text            =   "99"
         Top             =   2430
         Width           =   225
      End
      Begin VB.CheckBox chkRetry 
         Caption         =   " &Retry Connect        times "
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   2415
         Width           =   3240
      End
      Begin VB.TextBox txtIdent 
         Height          =   315
         Left            =   1530
         TabIndex        =   4
         Text            =   "~IDENT"
         Top             =   1185
         Width           =   2025
      End
      Begin VB.CheckBox chkInvisible 
         Caption         =   " Invisible Mode "
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   2130
         Width           =   3240
      End
      Begin VB.CheckBox chkReconnect 
         Caption         =   " Reconnect to server on disconnect "
         Height          =   255
         Left            =   210
         TabIndex        =   7
         Top             =   1845
         Width           =   3240
      End
      Begin VB.CheckBox chkStartUp 
         Caption         =   " Connect to server on Client load "
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Top             =   1560
         Width           =   3240
      End
      Begin VB.ComboBox cbServer 
         Height          =   315
         Left            =   210
         Sorted          =   -1  'True
         TabIndex        =   5
         Text            =   "irc.otherside.com"
         Top             =   2955
         Width           =   2685
      End
      Begin VB.TextBox txtFullName 
         Height          =   315
         Left            =   1530
         TabIndex        =   3
         Text            =   "projectIRC User"
         Top             =   810
         Width           =   2025
      End
      Begin VB.TextBox txtOtherNick 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Text            =   "OtherNick"
         Top             =   435
         Width           =   2025
      End
      Begin VB.TextBox txtNick 
         Height          =   315
         Left            =   1530
         TabIndex        =   1
         Text            =   "pIRCu"
         Top             =   60
         Width           =   2025
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AutoJoin Channels &:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   41
         Top             =   3315
         Width           =   1650
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2790
         TabIndex        =   19
         Top             =   2700
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IDENT:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   17
         Top             =   1215
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   16
         Top             =   2700
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   15
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alternate Nick:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   14
         Top             =   465
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nick:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   60
         TabIndex        =   13
         Top             =   90
         Width           =   390
      End
   End
   Begin HoverButton.Button btnOk 
      Height          =   375
      Left            =   3825
      TabIndex        =   61
      Top             =   4095
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "Ok"
      CaptionDown     =   "Ok"
      CaptionOver     =   "&Ok"
      ShowFocusRect   =   -1  'True
      Sink            =   -1  'True
      Picture         =   "frmOptions.frx":1D8B
      PictureDown     =   "frmOptions.frx":2DDD
      Pictureover     =   "frmOptions.frx":3E2F
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   3000
      TabIndex        =   55
      Top             =   4830
      Width           =   1665
   End
   Begin MSComctlLib.TreeView tvOptions 
      Height          =   4425
      Left            =   45
      TabIndex        =   54
      Top             =   45
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   7805
      _Version        =   393217
      Style           =   6
      Appearance      =   1
   End
   Begin HoverButton.Button btnCancel 
      Height          =   375
      Left            =   4740
      TabIndex        =   62
      Top             =   4095
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      BackColor       =   -2147483633
      HoverBackColor  =   14268829
      Border          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   10649190
      HilightColor    =   -2147483633
      ShadowColor     =   -2147483633
      HoverHilightColor=   14928823
      HoverShadowColor=   13410172
      ForeColor       =   -2147483630
      HoverForeColor  =   -2147483635
      Caption         =   "Cancel"
      CaptionDown     =   "Cancel"
      CaptionOver     =   "&Cancel"
      ShowFocusRect   =   -1  'True
      Sink            =   -1  'True
      Picture         =   "frmOptions.frx":4E81
      Style           =   0
      PictureLocation =   0
      ButtonStyleX    =   0
      State           =   0
      IconHeight      =   0
      IconWidth       =   0
   End
   Begin VB.Shape shpCurve2 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   1935
      Shape           =   2  'Oval
      Top             =   0
      Width           =   360
   End
   Begin VB.Label lblConnecting 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Not Added yet."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1950
      TabIndex        =   11
      Top             =   1875
      Width           =   3255
   End
   Begin VB.Label lblWhich 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting..."
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   1950
      TabIndex        =   0
      Top             =   -270
      Width           =   3285
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Options..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   -270
      Width           =   1800
   End
   Begin VB.Shape shpTop 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   360
      Left            =   -210
      Top             =   -360
      Width           =   5955
   End
   Begin VB.Shape shpBlue 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   0
      Top             =   -345
      Width           =   1950
   End
   Begin VB.Shape shpCorner 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   165
      Left            =   705
      Shape           =   4  'Rounded Rectangle
      Top             =   -300
      Width           =   2175
   End
   Begin VB.Shape shpCurve 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   180
      Left            =   1845
      Top             =   -240
      Width           =   240
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bC As Boolean
Dim bControl As Boolean
Sub AddServer(strName As String, bAppend As Boolean)
    If strName = "" Then Exit Sub
    Dim lngRet As Long
    lngRet = SendMessage(cbServer.hWnd, CB_FINDSTRINGEXACT, -1, ByVal strName)
    
    If lngRet <> -1 Then Exit Sub
    cbServer.AddItem strName
    If Not bAppend Then Exit Sub
    Open path & "serverlist.data" For Append As #1
        Print #1, strName
    Close #1
End Sub


Public Function GetImage(strItem As String) As Integer
    Select Case strItem
        Case "Connecting"
            GetImage = 1
        Case "IRC"
            GetImage = 2
        Case "Display"
            GetImage = 3
        Case Else
            GetImage = -1
    End Select
End Function

Sub HideAll()
    picConnecting.Visible = False
    picDisplay.Visible = False
    picWinIface.Visible = False
    picIRC.Visible = False
End Sub


Sub LoadFonts()
    Dim i As Integer
    For i = 0 To Screen.FontCount
        cmbFontName.AddItem Screen.Fonts(i)
    Next i
    For i = 0 To cmbFontName.ListCount - 1
        If LCase(cmbFontName.List(i)) = LCase(strFontName) Then
            cmbFontName.ListIndex = i
            Exit For
        End If
    Next i
    cmbFontSize.Text = intFontSize
End Sub

Sub LoadOptions()

    'this procedure WAS used to load buddies from the ini to treeview.
    'taken from Chad Cox's AIM example, thanks Chad.
    
    Dim strBuffer As String * 600, lngSize As Long, arrBuddies() As String, lngDo As Long
    Dim nod() As Node, intGroup As Integer, strOptions As String
    
    
    strOptions = "g Connecting" & chr(1) & _
                 "b Options" & chr(1) & _
                 "b Firewall" & chr(1) & _
                 "g IRC" & chr(1) & _
                 "b on Connect" & chr(1) & _
                 "b Ignore" & chr(1) & _
                 "b Logging" & chr(1) & _
                 "g Display" & chr(1) & _
                 "b Client" & chr(1) & _
                 "b Interface" & chr(1) & _
                 "b Output" & chr(1)
                 

    With tvOptions
        arrBuddies$ = Split(strOptions, chr(1))
        .Nodes.Clear
        For lngDo& = LBound(arrBuddies$) To UBound(arrBuddies$)
            ReDim Preserve nod(1 To .Nodes.Count + 1)
            If arrBuddies$(lngDo&) <> "" Then
                If left(arrBuddies$(lngDo&), 1) = "g" Then
                    Set nod(.Nodes.Count) = .Nodes.Add(, , , Right(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2))
                    intGroup% = .Nodes.Count
                Else
                    If .Nodes.Count > 0 Then
                        Set nod(.Nodes.Count) = .Nodes.Add(nod(intGroup%), tvwChild, , Right(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2))
                        nod(.Nodes.Count).EnsureVisible
                    End If
                End If
            End If
        Next
    End With

End Sub

Sub SaveConnect()
    'this saves all the Connection settings
    
    strMyNick = txtNick         'nick
    WriteINI "connect", "nick", strMyNick
    strOtherNick = txtOtherNick 'altnick
    WriteINI "connect", "altnick", strOtherNick
    strFullName = txtFullName   'full name
    WriteINI "connect", "fullname", strFullName
    strMyIdent = txtIdent       'ident
    WriteINI "connect", "ident", strMyIdent
    strServer = cbServer.Text   'server
    WriteINI "connect", "server", strServer
    lngPort = CLng(txtPort)     'port
    WriteINI "connect", "port", CStr(lngPort)
    bConOnLoad = chkStartUp     'connect on load
    WriteINI "connect", "connonload", CStr(bConOnLoad)
    bReconnect = chkReconnect   'reconnect on disconnect
    WriteINI "connect", "reconnect", CStr(bReconnect)
    bInvisible = chkInvisible   'invisible mode
    WriteINI "connect", "invisible", CStr(bInvisible)
    bRetry = chkRetry           'retry connect
    WriteINI "connect", "retry", CStr(bRetry)
    intRetry = CInt(txtRetry)   'number of retries
    WriteINI "connect", "retrynum", CStr(intRetry)
    strAutoJoin = txtAutoJoin   'autojoin channels
    WriteINI "connect", "autojoin", strAutoJoin
End Sub
Sub SaveDisplay()
    lngBackColor = picBackColor.BackColor
    WriteINI "display", "backcolor", CStr(lngBackColor)
    lngForeColor = picForeColor.BackColor
    WriteINI "display", "forecolor", CStr(lngForeColor)
    lngLeftColor = picLeftColor.BackColor
    WriteINI "display", "leftcolor", CStr(lngLeftColor)
    lngRightColor = picRightColor.BackColor
    WriteINI "display", "rightcolor", CStr(lngRightColor)

    intFontSize = CInt(cmbFontSize.Text)
    WriteINI "display", "fontsize", CStr(intFontSize)
    lngClientBack = picClientBack.BackColor
    WriteINI "display", "clientbg", CStr(lngClientBack)
    strBGImage = txtBGImage.Text
    WriteINI "display", "bgimage", strBGImage
    bBGImage = chkBGImage
    WriteINI "display", "bgimageon", CStr(bBGImage)
    bTileImg = chkTileImg
    WriteINI "display", "tileimg", CStr(bTileImg)
    
    
    strFontName = cmbFontName.Text
    WriteINI "display", "fontname", strFontName
    BuddyList.lstNicks.FontName = strFontName
    BuddyList.lstSetup.FontName = strFontName
    BuddyList.lstSetup.FontSize = intFontSize
    BuddyList.lstNicks.FontSize = intFontSize
    Dim i As Integer
    For i = 1 To intChannels
        If Channels(i).strName <> "" Then
            On Error Resume Next
            Channels(i).lstNicks.FontName = strFontName
            Channels(i).lstNicks.FontSize = intFontSize
            Channels(i).rtbTopic.Font.Name = strFontName
            Channels(i).rtbTopic.Font.Size = strFontSize
        End If
    Next i
End Sub



Sub SaveIRC()
    bWhoisInQuery = chkWhoIsQ   'do whois in query
    WriteINI "irc", "whoisquery", CStr(bWhoisInQuery)
    bRejoinOnKick = chkRejoin   'rejoin channel on kick
    WriteINI "irc", "rejoinonkick", CStr(bRejoinOnKick)
    bHidePing = chkHidePing     'hide ping? pong!
    WriteINI "irc", "hideping", CStr(bHidePing)
    strQuitMsg = ANSICode(rtfQuitMsg)   'quit message
    WriteINI "irc", "quitmsg", strQuitMsg
    bShowLag = chkShowLag        'show lag in channel
    WriteINI "irc", "showlag", CStr(bShowLag)
End Sub

Sub SaveWinIface()
    bFlatButtons = chkFlatButtons    'flat buttons in taskbar
    WriteINI "windows", "flatbuttons", CStr(bFlatButtons)
    bStretchButtons = chkStretch     'retry connect
    WriteINI "windows", "stretch", CStr(bStretchButtons)
    bMinToSystray = chkMinToSystray
    WriteINI "windows", "mintosystray", CStr(bMinToSystray)
    bAlwaysShowST = chkAlwaysShowST
    WriteINI "windows", "alwaysshowst", CStr(bAlwaysShowST)
    intTranslucency = CInt(txtTrans)
    WriteINI "windows", "translucency", CStr(intTranslucency)
    intButtonWidth = CInt(txtButtonWidth)   'width of non-stretched buttons
    WriteINI "windows", "buttonwidth", CStr(intButtonWidth)
    
    If bAlwaysShowST Then Client.SysTray.ShowIcon
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    Me.MousePointer = 11
    SaveConnect
    SaveDisplay
    SaveWinIface
    SaveIRC
    Me.MousePointer = 0
    Unload Me
    Client.DrawToolbar

End Sub

Private Sub Button1_Click()
    Dim strFile As SelectedFile
    strFile = ShowOpen(Me.hWnd)
    If strFile.bCanceled Then Exit Sub
    txtBGImage = strFile.sLastDirectory & strFile.sFiles(1)

End Sub

Private Sub cbServer_GotFocus()
    With cbServer
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub cbServer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        AddServer cbServer.Text, True
        txtPort.SetFocus
        KeyAscii = 0
        Exit Sub
    End If
    Dim lngFind As Long, intPos As Integer, intLength As Integer
    With cbServer
        If KeyAscii = 8 Then
            If .SelStart = 0 Then Exit Sub
            .SelStart = .SelStart - 1
            .SelLength = 32000
            .SelText = ""
        Else
            .SelText = chr(KeyAscii)
        End If
        KeyAscii = 0
        lngFind = SendMessage(.hWnd, CB_FINDSTRING, 0, ByVal .Text)
        If lngFind = -1 Then Exit Sub
        intPos = .SelStart
        intLength = Len(.List(lngFind)) - Len(.Text)
        .SelText = .SelText & Right(.List(lngFind), intLength)
        .SelStart = intPos
        .SelLength = intLength
    End With
End Sub


Private Sub cmdBGColor_Click()
    Dim lngColor As SelectedColor
    lngColor = ShowColor(Me.hWnd, True)
    If lngColor.bCanceled Then Exit Sub
    picBackColor.BackColor = lngColor.oSelectedColor
End Sub

Private Sub cmdBGDef_Click()
    picBackColor.BackColor = RGB(255, 255, 255)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCBDef_Click()
    picClientBack.BackColor = &H8000000C
End Sub

Private Sub cmdClientBack_Click()
    Dim lngColor As SelectedColor
    lngColor = ShowColor(Me.hWnd, True)
    If lngColor.bCanceled Then Exit Sub
    picClientBack.BackColor = lngColor.oSelectedColor

End Sub

Private Sub cmdFGDef_Click()
    picForeColor.BackColor = 0
End Sub

Private Sub cmdForeColor_Click()
    Dim lngColor As SelectedColor
    lngColor = ShowColor(Me.hWnd, True)
    If lngColor.bCanceled Then Exit Sub
    picForeColor.BackColor = lngColor.oSelectedColor

End Sub

Private Sub cmdLBDef_Click()
    picLeftColor.BackColor = &H800000
End Sub

Private Sub cmdLeftColor_Click()
    Dim lngColor As SelectedColor
    lngColor = ShowColor(Me.hWnd, True)
    If lngColor.bCanceled Then Exit Sub
    picLeftColor.BackColor = lngColor.oSelectedColor

End Sub


Private Sub cmdOK_Click()
    Me.MousePointer = 11
    SaveConnect
    SaveDisplay
    SaveWinIface
    SaveIRC
    Me.MousePointer = 0
    Unload Me
    Client.DrawToolbar
End Sub

Private Sub cmdRCDef_Click()
    picRightColor.BackColor = &H8000000F
End Sub

Private Sub cmdRightColor_Click()
    Dim lngColor As SelectedColor
    lngColor = ShowColor(Me.hWnd, True)
    If lngColor.bCanceled Then Exit Sub
    picRightColor.BackColor = lngColor.oSelectedColor

End Sub

Private Sub Command2_Click()
    Dim strFile As SelectedFile
    strFile = ShowOpen(Me.hWnd)
    If strFile.bCanceled Then Exit Sub
    txtBGImage = strFile.sLastDirectory & strFile.sFiles(1)

End Sub

Private Sub Command3_Click()
    'this procedure is used to load our buddies from the ini to our treeview.
    Dim strBuffer As String * 600, lngSize As Long, arrBuddies() As String, lngDo As Long
    Dim nod() As Node, intGroup As Integer, strOptions As String
    
    
    strOptions = "g Connecting" & chr(1) & _
                 "b Options" & chr(1) & _
                 "b Firewall" & chr(1) & _
                 "g IRC" & chr(1) & _
                 "b on Connect" & chr(1) & _
                 "b Ignore" & chr(1) & _
                 "b Logging" & chr(1) & _
                 "b General" & chr(1)

    With tvOptions
        arrBuddies$ = Split(strOptions, chr(1))
        .Nodes.Clear
        For lngDo& = LBound(arrBuddies$) To UBound(arrBuddies$)
            ReDim Preserve nod(1 To .Nodes.Count + 1)
            If arrBuddies$(lngDo&) <> "" Then
                If left(arrBuddies$(lngDo&), 1) = "g" Then
                    Set nod(.Nodes.Count) = .Nodes.Add(, , , Right(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2), 0)
                    intGroup% = .Nodes.Count
                Else
                    If .Nodes.Count > 0 Then
                        Set nod(.Nodes.Count) = .Nodes.Add(nod(intGroup%), tvwChild, , Right(arrBuddies$(lngDo&), Len(arrBuddies$(lngDo&)) - 2), 0)
                        nod(.Nodes.Count).EnsureVisible
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture("")
    HideAll
    picConnecting.Visible = True
        LoadOptions
    
    '* Connection settings
    txtNick = strMyNick
    txtOtherNick = strOtherNick
    txtFullName = strFullName
    txtIdent = strMyIdent
    txtPort = CStr(lngPort)
    cbServer.Text = strServer
    chkStartUp = TF(bConOnLoad)
    chkReconnect.Value = TF(bReconnect)
    chkInvisible.Value = TF(bInvisible)
    chkRetry.Value = TF(bRetry)
    txtRetry = CInt(intRetry)
    txtAutoJoin = strAutoJoin
    
    '* Display settings
    chkBGImage = TF(bBGImage)
    chkTileImg = TF(bTileImg)
    txtBGImage = strBGImage
    picClientBack.BackColor = CLng(lngClientBack)
    shpTop.BackColor = CLng(lngLeftColor)
    picBackColor.BackColor = CLng(lngBackColor)
    picForeColor.BackColor = CLng(lngForeColor)
    shpBlue.BackColor = CLng(lngLeftColor)
    picLeftColor.BackColor = CLng(lngLeftColor)
    picRightColor.BackColor = CLng(lngRightColor)
    shpCurve.BackColor = CLng(lngLeftColor)
    shpCurve2.BackColor = CLng(lngRightColor)
    shpCorner.FillColor = CLng(lngLeftColor)
    lbl1.ForeColor = CLng(lngBackColor)
    
    '* Windows/Interface
    chkFlatButtons.Value = TF(bFlatButtons)
    chkStretch.Value = TF(bStretchButtons)
    chkMinToSystray.Value = TF(bMinToSystray)
    chkAlwaysShowST.Value = TF(bAlwaysShowST)
    txtTrans.Text = CStr(intTranslucency)
    txtButtonWidth.Text = CStr(intButtonWidth)
    
    '* IRC
    chkWhoIsQ.Value = TF(bWhoisInQuery)
    chkRejoin.Value = TF(bRejoinOnKick)
    chkHidePing.Value = TF(bHidePing)
    chkShowLag.Value = TF(bShowLag)
    rtfQuitMsg = ""
    PutData rtfQuitMsg, strQuitMsg
    
    '* Blah, servers
    Dim srvlst As String, strData As String, strList() As String
    srvlst = path & "serverlist.data"
    If Not FileExists(srvlst) Then Open srvlst For Output As #1: Print #1, "": Close #1
    
    Open srvlst For Binary As #1
        strData = String(LOF(1), 0)
        Get #1, 1, strData
    Close #1
    strList = Split(strData, vbCrLf)
    strData = ""
    For i = LBound(strList) To UBound(strList)
        AddServer strList(i), False
    Next i
    
    LoadFonts
    
End Sub

Private Sub rtfQuitMsg_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = True   'control
End Sub


Private Sub rtfQuitMsg_KeyPress(KeyAscii As Integer)

    If KeyAscii = 11 Then
        ColorPicker.Move Client.left + Me.left, Client.top + Me.top + Me.Height - 100
        ColorPicker.Show
    ElseIf IsNumeric(chr(KeyAscii)) Then
    Else
        ColorPicker.Hide
    End If
    
    Dim strText As String, strNick As String, i As Integer, bFound As Boolean
    Dim strData As String, strTemp As String
    
    bFound = False
    On Error Resume Next
    
    If bControl Then
        If KeyAscii = 11 Then
            rtfQuitMsg.SelText = strColor & rtfQuitMsg.SelText
        ElseIf KeyAscii = 2 Then
            'rtfQuitMsgSelText = strBold
            rtfQuitMsg.SelBold = Not rtfQuitMsg.SelBold
        ElseIf KeyAscii = 21 Then
            'rtfQuitMsgSelText = strUnderline
            rtfQuitMsg.SelUnderline = Not rtfQuitMsg.SelUnderline
        ElseIf KeyAscii = 18 Then
            'rtfQuitMsgSelText = strReverse
            rtfQuitMsg.SelStrikeThru = Not rtfQuitMsg.SelStrikeThru
        End If
    End If
    

End Sub


Private Sub rtfQuitMsg_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then bControl = False   'control
End Sub


Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem.Caption
        Case "Connecting"  'connecting
            HideAll
            picConnecting.Visible = True
            lblWhich.Caption = "Connection Settings"
        Case "IRC"
            HideAll
            picIRC.Visible = True
            lblWhich.Caption = "IRC Settings"
        
        
        Case "Display"  'display
            HideAll
            picDisplay.Visible = True
            lblWhich.Caption = "Display Settings"
        Case "Windows"  'windows/iface
            HideAll
            picWinIface.Visible = True
            lblWhich.Caption = "Windows/Interface Settings"
    End Select

End Sub

Private Sub tvOptions_Click()
    
    HideAll
    Select Case tvOptions.SelectedItem.Text
        Case "Options"
            picConnecting.Visible = True
        Case "IRC"
            picIRC.Visible = True
        Case "Client"
            picDisplay.Visible = True
        Case "Interface"
            picWinIface.Visible = True
        Case Else

    End Select
End Sub


Private Sub txtAutoJoin_GotFocus()
    With txtAutoJoin
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub


Private Sub txtAutoJoin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Call cmdOK_Click
        KeyAscii = 0
    End If
End Sub


Private Sub txtButtonWidth_KeyPress(KeyAscii As Integer)
    If IsNumeric(chr(KeyAscii)) = False And _
        KeyAscii <> 8 Then KeyAscii = 0
End Sub


Private Sub txtFullName_GotFocus()
    With txtFullName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtFullName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtIdent.SetFocus: KeyAscii = 0
End Sub


Private Sub txtIdent_GotFocus()
    With txtIdent
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtIdent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cbServer.SetFocus: KeyAscii = 0
End Sub


Private Sub txtNick_GotFocus()
    With txtNick
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtNick_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOtherNick.SetFocus: KeyAscii = 0
    If KeyAscii = Asc(" ") Then KeyAscii = 0
End Sub


Private Sub txtOtherNick_GotFocus()
    With txtOtherNick
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtOtherNick_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFullName.SetFocus: KeyAscii = 0
    If KeyAscii = Asc(" ") Then KeyAscii = 0
End Sub


Private Sub txtPort_GotFocus()
    With txtPort
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAutoJoin.SetFocus
        KeyAscii = 0
    End If
End Sub


Private Sub txtRetry_KeyPress(KeyAscii As Integer)
    If IsNumeric(chr(KeyAscii)) Then Else KeyAscii = 0
End Sub


