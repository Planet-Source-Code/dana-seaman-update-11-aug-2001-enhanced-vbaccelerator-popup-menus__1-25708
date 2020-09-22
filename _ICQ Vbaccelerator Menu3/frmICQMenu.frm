VERSION 5.00
Begin VB.Form frmICQMenu 
   BackColor       =   &H00404040&
   Caption         =   "Enhanced VbAccelerator Menus (Uses standard Vb menus)"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8115
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1725
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmICQMenu.frx":0000
         Top             =   360
         Width           =   5565
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmICQMenu.frx":0008
         Top             =   4800
         Width           =   5565
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmICQMenu.frx":0010
         Top             =   3360
         Width           =   5565
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmICQMenu.frx":0018
         Top             =   2400
         Width           =   5565
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "frmICQMenu.frx":0020
         Top             =   6960
         Width           =   5565
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   " New "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   " New "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1860
         TabIndex        =   13
         Top             =   4560
         Width           =   525
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   " New "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1770
         TabIndex        =   12
         Top             =   3120
         Width           =   525
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   " Menus "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   150
         TabIndex        =   11
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   " Vertical Bar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   150
         TabIndex        =   10
         Top             =   4440
         Width           =   1605
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   " Separators "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   150
         TabIndex        =   9
         Top             =   3000
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   " Highlights "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   150
         TabIndex        =   8
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   " New "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2220
         TabIndex        =   7
         Top             =   6720
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   " AlphaBlending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   150
         TabIndex        =   6
         Top             =   6600
         Width           =   1965
      End
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   5880
      Picture         =   "frmICQMenu.frx":0028
      Top             =   0
      Visible         =   0   'False
      Width           =   24000
   End
   Begin VB.Image Image1 
      Height          =   1920
      Index           =   0
      Left            =   5880
      Picture         =   "frmICQMenu.frx":2A81
      Top             =   240
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   1920
      Index           =   2
      Left            =   5880
      Picture         =   "frmICQMenu.frx":4804
      Top             =   2160
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   1920
      Index           =   3
      Left            =   5880
      Picture         =   "frmICQMenu.frx":7E82
      Top             =   4080
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "File"
      Tag             =   "1605"
      Begin VB.Menu mnuFile 
         Caption         =   "-Group"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Filter"
         Index           =   1
         Tag             =   "1421"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Open with ... Assoc"
         Index           =   2
         Tag             =   "1422"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Clone"
         Index           =   3
         Tag             =   "1135"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Date Stamp"
         Index           =   4
         Tag             =   "1142"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Attr Stamp"
         Index           =   5
         Tag             =   "1144"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-Blowfish"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Enc"
         Index           =   7
         Tag             =   "1640"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Dec"
         Index           =   8
         Tag             =   "1641"
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-UUENCODE"
         Index           =   9
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&UUEncode"
         Index           =   10
      End
      Begin VB.Menu mnuFile 
         Caption         =   "UUDec&ode"
         Index           =   11
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-Misc"
         Index           =   12
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&NotePad ... WordPad"
         Index           =   13
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Menu Options"
         Index           =   14
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "-Highlight Style"
            Index           =   0
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Highlight Standard"
            Index           =   1
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Highlight Gradient"
            Index           =   2
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Highlight Background (wallpaper)"
            Index           =   3
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "-Vertical Style"
            Index           =   4
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Vertical Standard"
            Index           =   5
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Vertical Gradient"
            Index           =   6
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Vertical Background (wallpaper)"
            Index           =   7
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "-Boolean"
            Index           =   8
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Vertical Bar"
            Index           =   9
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Vertical Bar Align Right"
            Index           =   10
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Vertical Text 270°"
            Index           =   11
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Transparent text"
            Index           =   12
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Embossed Separator"
            Index           =   13
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "Translucency"
            Index           =   14
         End
         Begin VB.Menu mnuRadioCheck 
            Caption         =   "#Radio Check#Reserved 02"
            Index           =   15
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "#Group Functions#E&xit"
         Index           =   15
      End
   End
   Begin VB.Menu mnuEditTOP 
      Caption         =   "&Clipboard"
      Begin VB.Menu mnuEdit 
         Caption         =   "-Long Filenames"
         Index           =   0
         Tag             =   "1400"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Nam"
         Index           =   1
         Tag             =   "1401"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Pth"
         Index           =   2
         Tag             =   "1402"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Nam + Pth"
         Index           =   3
         Tag             =   "1403"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-Alias"
         Index           =   4
         Tag             =   "1404"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   ".Nam"
         Index           =   5
         Tag             =   "1405"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   ".Pth"
         Index           =   6
         Tag             =   "1406"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   ".Nam + .Pth"
         Index           =   7
         Tag             =   "1407"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-Selected"
         Index           =   8
         Tag             =   "1408"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Tab"
         Index           =   9
         Tag             =   "1409"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Comma"
         Index           =   10
         Tag             =   "1410"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-ALL"
         Index           =   11
         Tag             =   "1411"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "*Tab"
         Index           =   12
         Tag             =   "1412"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "#Clipboard (Grid)#*Comma"
         Index           =   13
         Tag             =   "1413"
      End
   End
   Begin VB.Menu mnuViewTOP 
      Caption         =   "View"
      Tag             =   "1014"
      Begin VB.Menu mnuView 
         Caption         =   ""
         Index           =   0
         Tag             =   "1419"
      End
      Begin VB.Menu mnuView 
         Caption         =   ""
         Index           =   1
         Tag             =   "1420"
      End
      Begin VB.Menu mnuView 
         Caption         =   ""
         Index           =   2
         Tag             =   "1424"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&Kilobytes"
         Index           =   3
      End
      Begin VB.Menu mnuView 
         Caption         =   "Tips"
         Index           =   4
         Tag             =   "1327"
      End
      Begin VB.Menu mnuView 
         Caption         =   "#Sounds"
         Index           =   5
         Tag             =   "1426"
      End
   End
   Begin VB.Menu mnuToolsTOP 
      Caption         =   "Tools"
      Tag             =   "1606"
      Begin VB.Menu mnuTools 
         Caption         =   "Dos"
         Index           =   0
         Tag             =   "1005"
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Fmt Disk"
         Index           =   1
         Tag             =   "1322"
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Sync Atom"
         Index           =   2
         Tag             =   "1323 "
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Build FS Db"
         Index           =   3
         Tag             =   "1324"
      End
      Begin VB.Menu mnuTools 
         Caption         =   "FF"
         Index           =   4
         Tag             =   "1325 "
      End
      Begin VB.Menu mnuTools 
         Caption         =   "DC"
         Index           =   5
         Tag             =   "1326"
      End
   End
   Begin VB.Menu mnuCtrlTOP 
      Caption         =   "Ctrl"
      Tag             =   "1607"
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   "&Modem"
         Index           =   8
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu mnuCtrl 
         Caption         =   ""
         Index           =   15
      End
   End
   Begin VB.Menu mnuDriveTOP 
      Caption         =   "Drv"
      Tag             =   "1608"
      Begin VB.Menu mnuDrive 
         Caption         =   "-Volume"
         Index           =   0
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "VolLbl"
         Index           =   1
         Tag             =   "1328"
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "-Network"
         Index           =   2
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "MapNetDrv"
         Index           =   3
         Tag             =   "1329"
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "UnMapNetDrv"
         Index           =   4
         Tag             =   "1330"
      End
   End
   Begin VB.Menu mnuLangTOP 
      Caption         =   "Lang"
      Tag             =   "1146"
      Begin VB.Menu mnuLang 
         Caption         =   "&Deutsch"
         Index           =   0
      End
      Begin VB.Menu mnuLang 
         Caption         =   "&English"
         Index           =   1
      End
      Begin VB.Menu mnuLang 
         Caption         =   "E&spanhõl"
         Index           =   2
      End
      Begin VB.Menu mnuLang 
         Caption         =   "&Français"
         Index           =   3
      End
      Begin VB.Menu mnuLang 
         Caption         =   "&Italiano"
         Index           =   4
      End
      Begin VB.Menu mnuLang 
         Caption         =   "#Language#&Português (brasileiro)"
         Index           =   5
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "Hlp"
      Tag             =   "1100"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Con"
         Index           =   0
         Shortcut        =   {F1}
         Tag             =   "1250"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Srch"
         Index           =   1
         Tag             =   "1251"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Idx"
         Index           =   2
         Tag             =   "1252"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&Abt"
         Index           =   4
         Tag             =   "1002"
      End
   End
   Begin VB.Menu mnuCopyTOP 
      Caption         =   "Copy"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Cop"
         Index           =   0
         Shortcut        =   {F7}
         Tag             =   "1041"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Smart"
         Index           =   1
         Shortcut        =   {F8}
         Tag             =   "1042"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Mov"
         Index           =   2
         Shortcut        =   {F9}
         Tag             =   "1108"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "FTP &Upload"
         Index           =   4
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "FTP &Download"
         Index           =   5
      End
   End
   Begin VB.Menu mnuSortTOP 
      Caption         =   "Sort"
      Visible         =   0   'False
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   1
         Tag             =   "1521"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   2
         Tag             =   "1522"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   3
         Tag             =   "1523"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   4
         Tag             =   "1525"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   5
         Tag             =   "1524"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   6
         Tag             =   "1526"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   7
         Tag             =   "1527"
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu mnuSort 
         Caption         =   ""
         Index           =   10
         Tag             =   "1528"
      End
   End
   Begin VB.Menu mnuSortZipTOP 
      Caption         =   "SortZip"
      Visible         =   0   'False
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   0
         Tag             =   "1520"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   1
         Tag             =   "1521"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   2
         Tag             =   "1522"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   3
         Tag             =   "1523"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   4
         Tag             =   "1525"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   5
         Tag             =   "1524"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   6
         Tag             =   "1712"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   7
         Tag             =   "1713"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   8
         Tag             =   "1526"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   9
         Tag             =   "1527"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   10
         Tag             =   "1714"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   11
         Tag             =   "1715"
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   "CRC"
         Index           =   12
      End
      Begin VB.Menu mnuSortZip 
         Caption         =   ""
         Index           =   13
         Tag             =   "1716"
      End
   End
   Begin VB.Menu mnuZipTOP 
      Caption         =   "Zip"
      Visible         =   0   'False
      Begin VB.Menu mnuZip 
         Caption         =   ""
         Index           =   0
         Tag             =   "1550"
         Begin VB.Menu mnuZipAdd 
            Caption         =   "Ovr"
            Index           =   0
            Tag             =   "1239"
         End
         Begin VB.Menu mnuZipAdd 
            Caption         =   "Adv"
            Index           =   1
            Tag             =   "1240"
         End
      End
      Begin VB.Menu mnuZip 
         Caption         =   ""
         Index           =   1
         Tag             =   "1551"
      End
      Begin VB.Menu mnuZip 
         Caption         =   ""
         Index           =   2
         Tag             =   "1552"
      End
   End
   Begin VB.Menu mnuSelectTOP 
      Caption         =   "Select"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   0
         Tag             =   "1031"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   1
         Tag             =   "1032"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   2
         Tag             =   "1033"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   4
         Tag             =   "1035"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   ""
         Index           =   5
         Tag             =   "1036"
      End
   End
   Begin VB.Menu mnuPopTOP 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu mnuPop 
         Caption         =   "C"
         Index           =   0
         Tag             =   "1041"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "CI"
         Index           =   1
         Tag             =   "1042"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "M"
         Index           =   2
         Tag             =   "1108"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "R"
         Index           =   3
         Tag             =   "1111"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "D"
         Index           =   4
         Tag             =   "1107"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "S"
         Index           =   5
         Tag             =   "1118"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "N"
         Index           =   6
         Tag             =   "1109"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "P"
         Index           =   7
         Tag             =   "1012"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "D"
         Index           =   8
         Tag             =   "1142"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "A"
         Index           =   9
         Tag             =   "1144"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Z"
         Index           =   10
         Tag             =   "1110"
      End
      Begin VB.Menu mnuPop 
         Caption         =   "#Popup Menu#Create Shortcut"
         Index           =   11
      End
   End
End
Attribute VB_Name = "frmICQMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private ControlArray    As Variant
Private sCreated        As String
Private sAccessed       As String
Private sNone           As String

'--------------------------------------------
Private cIM As New cIconMenu
Private m_cIL16          As New cVBALImageList
Private Sub PrepareImageLists()

   Dim pic As StdPicture
  
   Set pic = Image1(1) 'MenuIcons 16x16
   With m_cIL16
      .ColourDepth = &H8
      .IconSizeX = 16
      .IconSizeY = 16
      .Create
      .AddFromHandle pic.Handle, IMAGE_BITMAP, , &H4080C0
   End With
   
End Sub


Private Sub Form_Load()
   
   SoundOn = True
   ControlArray = Split("sysdm.cpl @1|appwiz.cpl @1|timedate.cpl|desk.cpl|main.cpl @3|inetcpl.cpl|joy.cpl|main.cpl @1|modem.cpl|main.cpl|mmsys.cpl|netcpl.cpl|password.cpl|main.cpl @2|intl.cpl|sysdm.cpl", "|")

'---- Rip resource strings from Windows Dll's ----

   sCreated = GetResourceString(8996)
   sAccessed = GetResourceString(8997)
   sNone = GetResourceString(9808)

   PrepareImageLists 'Load picture strip into imagelist class
   With cIM
      .Attach Me.hwnd
      .Font = frmICQMenu.Font
      Set .BackgroundPicture = Image1(0).Picture
      .ImageList = m_cIL16.hIml 'was vbalImageList1
      .HighlightStyle = ECPStyleGradient: Check (2)
      '--------------------------------------------
      'Following properties are NOT in original control
      'If there is no BackgroundPictureVertical set
      'then SeparatorBackColor is used for vertical bar
      .VerticalStyle = ECPStyleGradient: Check (6)
      Set .BackgroundPictureVertical = Image1(2).Picture
      Set .BackgroundPictureHighlight = Image1(3).Picture
      .GradientStartColor = vbRed '&H80FF& 'Orange
      .GradientEndColor = vbYellow
      .SeparatorBackColor = &H98CCD0    'Gold
      .SeparatorForeColor = vbBlack
      .SeparatorTextEmboss = True: Check (13)
      .SeparatorTextTransparent = True: Check (12)
      .HasVerticalCaptionBar = True: Check (9)
      .VerticalFont "Arial", 16, FW700_BOLD
     ' .VerticalPlaceRight = True: Check (10)
     ' .VerticalEscapement270 = True: Check (11)
     ' .HasTranslucency = True: Check (14)
     ' .TranslucencyPercentage = 20
   End With

   '** Defaults to English (1000) if no registry entry
   Lang = GetSetting(App.Title, "Settings", "Lang", 1000)
   
   mnuLang(Lang \ c1000).Checked = True
   UpdateLanguage
   
   Text1.Text = "1. Uses standard Vb menus. No need to build complicated menus at run-time." & vbCrLf & _
                "2. Does Popup menus as well." & vbCrLf & _
                "3. Item Icons." & vbCrLf & _
                "4. Background image (wallpaper)." & vbCrLf & _
                "5. Easily compiled into Active-X DLL." & vbCrLf & _
                "6. Note: If you are using the menu tag property and a resource file to support multiple languages be sure to include the " & Chr$(34) & "#" & Chr$(34) & " sign(s) and vertical captions in ALL last menu items in your resource file."
   
   Text2.Text = "1. Standard, Gradient, or Background image (wallpaper)."

   Text3.Text = "1. Transparent caption." & vbCrLf & _
                "2. Opaque caption. Background Color, Foreground Color, Emboss properties." & vbCrLf & _
                "3. Separator Icon." & vbCrLf & _
                "4. Caption Example " & Chr$(34) & "-MyCaption" & Chr$(34) & "."

   Text4.Text = "1. Built into cIconMenu class by writing direct to DC's." & vbCrLf & _
                "2. Left or right bar placement." & vbCrLf & _
                "3. 90° or 270° text rotation." & vbCrLf & _
                "4. Standard, Gradient, or Background image (wallpaper)." & vbCrLf & _
                "5. Vertical Icon." & vbCrLf & _
                "6. Vertical Font properties. Bar width auto-sizes to font size." & vbCrLf & _
                "7. Transparent caption." & vbCrLf & _
                "8. Opaque caption. Background Color, Foreground Color properties."
   Text5.Text = "1. True translucency with user specified percentage. NOT WORKING CORRECTLY. CAN ANYONE HELP WITH THIS. SEE cIconMenu class."

End Sub

Private Sub Check(Index As Long)
   mnuRadioCheck(Index).Checked = True
End Sub
Private Sub UpdateLanguage()
   On Error GoTo ProcedureError
   Dim j As Long, L4 As Long
   '-- first update toolbar and misc controls

SetControlCaptionStrings Me    ' indexed by language
   
'------------------------------
'To specify an IconIndex for the vertical caption bar
'simply combine 2 integer indexes into a long
'using function MakeLong.
'cIconMenu(CPopenu.cls) will split the indexes.
With cIM
   .ClearIcons

   For L4 = 0 To 1
      .IconIndex(mnuZipAdd(L4).Caption) = 27
   Next
   '----------Radio Buttons and Check Boxes
   .IconIndex(mnuRadioCheck(0).Caption) = 78
   For L4 = 1 To 7
      If L4 <> 4 Then 'Ignore item 4 (separator)
         .IconIndex(mnuRadioCheck(L4).Caption) = IIf(mnuRadioCheck(L4).Checked, 78, 81)
      End If
   Next
   .IconIndex(mnuRadioCheck(4).Caption) = 78
   .IconIndex(mnuRadioCheck(8).Caption) = 77
   For L4 = 9 To 15
      .IconIndex(mnuRadioCheck(L4).Caption) = IIf(mnuRadioCheck(L4).Checked, 77, 76)
   Next
   .IconIndex(mnuRadioCheck(15).Caption) = MakeLong(.IconIndex(mnuRadioCheck(11).Caption), 77)

   '----------
   For L4 = 0 To 2
      .IconIndex(mnuZip(L4).Caption) = 27 + L4
   Next
    
   For L4 = 0 To 4
      .IconIndex(mnuDrive(L4).Caption) = Choose(L4 + 1, -1, 75, -1, 57, MakeLong(58, 72))
   Next

   For L4 = 0 To 4
      .IconIndex(mnuHelp(L4).Caption) = Choose(L4 + 1, 49, 55, 56, -1, MakeLong(42, 49))
   Next

   For L4 = 0 To 5
      .IconIndex(mnuView(L4).Caption) = Choose(L4 + 1, 32, 26, 46, 63, 42, MakeLong(23, 26))
      .IconIndex(mnuLang(L4).Caption) = Choose(L4 + 1, 36, 37, 38, 39, 40, MakeLong(41, 37))
      .IconIndex(mnuCopy(L4).Caption) = Choose(L4 + 1, 25, 42, 25, -1, 66, 65)
      .IconIndex(mnuSelect(L4).Caption) = Choose(L4 + 1, 31, 50, 52, -1, 31, 50)
      .IconIndex(mnuTools(L4).Caption) = Choose(L4 + 1, 47, 53, 10, 63, 55, MakeLong(53, 13))
   Next
   
   mnuSort(0).Caption = sNone
   mnuSort(8).Caption = sCreated
   mnuSort(9).Caption = sAccessed
   For L4 = 0 To 10
      .IconIndex(mnuSort(L4).Caption) = Choose(L4 + 1, 47, 31, 78, 76, 63, 77, 32, 33, 10, 32, 75)
   Next

   For L4 = 0 To 11
      .IconIndex(mnuPop(L4).Caption) = Choose(L4 + 1, 25, 42, 25, 50, 28, 64, 75, 75, 10, 75, 35, MakeLong(84, 42))
   Next

   For L4 = 0 To 13
      .IconIndex(mnuEdit(L4).Caption) = Choose(L4 + 1, 44, 44, 44, 44, 47, 47, 47, 47, 31, 31, 31, 79, 79, MakeLong(79, 44))
   Next

   For L4 = 0 To 13
      .IconIndex(mnuSortZip(L4).Caption) = Choose(L4 + 1, 47, 31, 78, 76, 63, 77, 35, 34, 32, 33, 35, 20, -1, 67)
   Next
   For L4 = 0 To 15
      .IconIndex(mnuFile(L4).Caption) = Choose(L4 + 1, 59, 45, 43, 25, 10, 75, 2, 2, 1, 62, 62, 61, 60, 60, 78, MakeLong(28, 42))
   Next
   
   'Some *.cpl files are 16-bit (even in Win ME)
   'so we can't rip those resources without Thunking
   mnuCtrl(0).Caption = GetResourceString(1610)   'Add new hardware
   ExtractMenuCaption 1, 2001                'Add/Remove Programs
   ExtractMenuCaption 2, 300                 'Time/Date
   ExtractMenuCaption 3, 100                 'Display
   ExtractMenuCaption 4, 106                 'Fonts
   ExtractMenuCaption 5, 4312                'Internet
   ExtractMenuCaption 6, 1076                'Game
   ExtractMenuCaption 7, 102                 'Kybd
   mnuCtrl(8).Caption = "Modems"             'Modems
   ExtractMenuCaption 9, 100                 'Mouse
   ExtractMenuCaption 10, 4867               'Sounds & MM
   mnuCtrl(11).Caption = GetResourceString(1621)           'Network
   ExtractMenuCaption 12, 2002               'Password
   ExtractMenuCaption 13, 104                'Printer
   ExtractMenuCaption 14, 1                  'Regional Settings
   mnuCtrl(15).Caption = GetResourceString(1625)            'System

   For L4 = 0 To 15
      j = Choose(L4 + 1, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 23, 19, 20, 21, 22, MakeLong(24, 22))
      .IconIndex(mnuCtrl(L4).Caption) = j
   Next

End With

   SaveSetting App.Title, "Settings", "Language", Lang

ProcedureExit:
  Exit Sub
ProcedureError:
  Resume Next
  If ErrMsgBox(Me.Name & ".UpdateLanguagee") = vbRetry Then Resume Next

End Sub
Private Function IndicesToLong(Vert As Integer, Norm As Integer) As Long
   IndicesToLong = Vert * &HFFFF + Norm
End Function

Private Sub ExtractMenuCaption(Index As Long, Key As Long)
   Dim CPL As String, L4 As Long
   CPL = ControlArray(Index)
   L4 = InStr(CPL, ".")
   If L4 Then
      CPL = left$(CPL, L4 + 3)
      mnuCtrl(Index).Caption = "&" & GetResourceStringFromFile(CPL, Key)
   End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then 'Mouse right
      Me.PopupMenu mnuPopTOP
   End If
End Sub

Private Sub Form_Resize()

   On Error GoTo ProcedureError
 
   If Me.WindowState <> vbMinimized Then
  '    If Me.Width < 3000 Then
  '       Me.Width = 3000
  '    End If
  '    Picture1.Move Me.ScaleWidth - (Picture1.Width), 0, Picture1.Width, Picture1.Height
  '    Picture2.Move 0, 0, Picture1.left, Picture1.Height
   End If
 
ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".Form_Resize") = vbRetry Then Resume Next
End Sub

Private Sub mnuCtrl_Click(Index As Integer)
   Dim CPL As String
   On Error GoTo ProcedureError
   
   CPL = ControlArray(Index)
   Shell "rundll32.exe shell32.dll,Control_RunDLL " & CPL, vbNormalFocus
ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".mnuCtrl_Click") = vbRetry Then Resume Next
End Sub

Private Sub mnuFile_Click(Index As Integer)
   If Index = mnuFile.UBound Then
      Unload Me
   End If
End Sub

Private Sub mnuLang_Click(Index As Integer)
   On Error GoTo ProcedureError
   Dim temp As Long, LangIndex As Long

   'uncheck old Lang
   mnuLang(Lang \ c1000).Checked = False
   'check new Lang
   mnuLang(Index).Checked = True

   Lang = Index * c1000

   UpdateLanguage ' fix remaining stuff

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".mnuLang_Click") = vbRetry Then Resume Next

End Sub

Private Sub mnuRadioCheck_Click(Index As Integer)
   Dim L4 As Long

With cIM
   Select Case Index
      Case 1 To 3 'Radio
         cIM.HighlightStyle = Index - 1
         For L4 = 1 To 3
            mnuRadioCheck(L4).Checked = IIf(L4 = Index, True, False)
            .IconIndex(mnuRadioCheck(L4).Caption) = IIf(L4 = Index, 78, 81)
         Next
      Case 5 To 7 'Radio
         cIM.VerticalStyle = Index - 5
         For L4 = 5 To 7
            mnuRadioCheck(L4).Checked = IIf(L4 = Index, True, False)
            .IconIndex(mnuRadioCheck(L4).Caption) = IIf(L4 = Index, 78, 81)
         Next
      Case 9 To 15 'Check
         mnuRadioCheck(Index).Checked = Not mnuRadioCheck(Index).Checked
         If mnuRadioCheck(Index).Checked Then
            .IconIndex(mnuRadioCheck(Index).Caption) = 77
         Else
            .IconIndex(mnuRadioCheck(Index).Caption) = 76
         End If
         Select Case Index
            Case 9
               .HasVerticalCaptionBar = mnuRadioCheck(Index).Checked
            Case 10
               .VerticalPlaceRight = mnuRadioCheck(Index).Checked
            Case 11
               .VerticalEscapement270 = mnuRadioCheck(Index).Checked
            Case 12
               .SeparatorTextTransparent = mnuRadioCheck(Index).Checked
            Case 13
               .SeparatorTextEmboss = mnuRadioCheck(Index).Checked
            Case 14
               .HasTranslucency = mnuRadioCheck(Index).Checked
            Case 15
         End Select
    End Select
End With
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then 'Mouse right
      Me.PopupMenu mnuPopTOP
   End If
End Sub

