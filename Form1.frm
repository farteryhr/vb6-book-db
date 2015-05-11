VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "book db"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   13665
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox picAll 
      Height          =   255
      Left            =   6720
      ScaleHeight     =   195
      ScaleWidth      =   1755
      TabIndex        =   77
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdUser 
      Caption         =   "user"
      Height          =   375
      Left            =   5160
      TabIndex        =   76
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdBook 
      Caption         =   "book"
      Height          =   375
      Left            =   3720
      TabIndex        =   75
      Top             =   5520
      Width           =   1335
   End
   Begin VB.PictureBox picUser 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   6720
      ScaleHeight     =   4935
      ScaleWidth      =   6615
      TabIndex        =   55
      Top             =   480
      Width           =   6615
      Begin VB.TextBox txtAge 
         Height          =   270
         Left            =   1200
         TabIndex        =   79
         Top             =   1560
         Width           =   1815
      End
      Begin VB.PictureBox picSuper 
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2835
         ScaleWidth      =   6315
         TabIndex        =   71
         Top             =   1920
         Width           =   6375
         Begin VB.ListBox lstUser 
            Height          =   1320
            Left            =   120
            TabIndex        =   74
            Top             =   120
            Width           =   6135
         End
         Begin VB.CommandButton cmdUserDelete 
            Caption         =   "delete user!"
            Height          =   375
            Left            =   4440
            TabIndex        =   73
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CommandButton cmdNewUser 
            Caption         =   "new user!"
            Height          =   375
            Left            =   2520
            TabIndex        =   72
            Top             =   1560
            Width           =   1815
         End
      End
      Begin VB.TextBox txtGender 
         Height          =   270
         Left            =   1200
         TabIndex        =   70
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtRealName 
         Height          =   270
         Left            =   1200
         TabIndex        =   67
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdUserSearch 
         Caption         =   "?"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   66
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdUserSearch 
         Caption         =   "?"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   65
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdUserUpdate 
         Caption         =   "update user!"
         Height          =   375
         Left            =   4680
         TabIndex        =   64
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtInPasswordR 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   4920
         PasswordChar    =   "#"
         TabIndex        =   63
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtInPassword 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   4920
         PasswordChar    =   "#"
         TabIndex        =   62
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtInUsername 
         Height          =   270
         Left            =   1200
         TabIndex        =   59
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtWorkId 
         Height          =   270
         Left            =   1200
         TabIndex        =   57
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "age"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "gander"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "real name"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "again"
         Height          =   255
         Left            =   3720
         TabIndex        =   61
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "password"
         Height          =   255
         Left            =   3720
         TabIndex        =   60
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "username"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "workid"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picRecord 
      Height          =   1815
      Left            =   10200
      ScaleHeight     =   1755
      ScaleWidth      =   3315
      TabIndex        =   45
      Top             =   5520
      Width           =   3375
      Begin VB.CommandButton cmdQuery 
         Caption         =   "query!"
         Height          =   375
         Left            =   1920
         TabIndex        =   52
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtDateEnd 
         Height          =   270
         Left            =   1320
         TabIndex        =   49
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtDateStart 
         Height          =   270
         Left            =   1320
         TabIndex        =   46
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblQuery 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "result"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "date end"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "date start"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picPayCancel 
      Height          =   1815
      Left            =   6720
      ScaleHeight     =   1755
      ScaleWidth      =   3315
      TabIndex        =   32
      Top             =   5520
      Width           =   3375
      Begin VB.CommandButton cmdInArrival 
         Caption         =   "arrival!"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   80
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdInCancel 
         Caption         =   "cancel!"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   44
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdInPay 
         Caption         =   "pay!"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblStatus 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   81
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblImportId 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   41
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblTotalPay 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblInPrice 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2520
         TabIndex        =   39
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblQuantity 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "order id"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "total"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "price"
         Height          =   255
         Left            =   1680
         TabIndex        =   34
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "quantity"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox picBook 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   6615
      TabIndex        =   5
      Top             =   480
      Width           =   6615
      Begin VB.CommandButton cmdImportProcess 
         Caption         =   "< import process"
         Height          =   375
         Left            =   5160
         TabIndex        =   31
         Top             =   4440
         Width           =   1335
      End
      Begin VB.PictureBox picTradeImport 
         Height          =   1815
         Left            =   1680
         ScaleHeight     =   1755
         ScaleWidth      =   3315
         TabIndex        =   26
         Top             =   3000
         Width           =   3375
         Begin VB.CommandButton cmdInImport 
            Caption         =   "import!"
            Height          =   375
            Left            =   1920
            TabIndex        =   54
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdInTrade 
            Caption         =   "trade!"
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtInPrice 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   2520
            TabIndex        =   30
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox txtQuantity 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   1080
            TabIndex        =   27
            Text            =   "1"
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1080
            TabIndex        =   42
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "total"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "price"
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "quantity"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "< new import"
         Height          =   375
         Left            =   5160
         TabIndex        =   25
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdTrade 
         Caption         =   "< trade"
         Height          =   855
         Left            =   5160
         TabIndex        =   24
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdRecord 
         Caption         =   "record >"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   4440
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "modify book!"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "new book!"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtPrice 
         Height          =   270
         Left            =   1080
         TabIndex        =   19
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "?"
         Height          =   255
         Index           =   3
         Left            =   6000
         TabIndex        =   18
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "?"
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   17
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "?"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   16
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "?"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
      Begin VB.ListBox lstBook 
         Height          =   1320
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   6375
      End
      Begin VB.TextBox txtPress 
         Height          =   270
         Left            =   1080
         TabIndex        =   12
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txtAuthor 
         Height          =   270
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txtIsbn 
         Height          =   270
         Left            =   3240
         TabIndex        =   7
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txtTitle 
         Height          =   270
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "price"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "press"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "author"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "isbn"
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "title"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtUsername 
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "password"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "username"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dbconn As New ADODB.Connection
Private dbrs As New ADODB.Recordset

Private bLoggedIn As Boolean
Private bImportProcess As Boolean
Private userLevel As Long

Private Sub cmdBook_Click()
    picBook.Visible = True
    picUser.Visible = False
End Sub

Private Sub cmdImport_Click()
    picRecord.Visible = False
    picTradeImport.Visible = True
    picPayCancel.Visible = False
    cmdInTrade.Visible = False
    cmdInImport.Visible = True
    If bImportProcess = True Then
        lstBook.Clear
    End If
    bImportProcess = False
End Sub

Private Sub cmdImportProcess_Click()
    picRecord.Visible = False
    picTradeImport.Visible = False
    picPayCancel.Visible = True
    If bImportProcess = False Then
        lstBook.Clear
    End If
    bImportProcess = True
End Sub

Private Sub cmdInArrival_Click()
    On Error GoTo die
    Dim isbn As String
    dbconn.Execute "update import set status = 2 where importid = " & lblImportId.Caption
    dbrs.Open "select top 1 * from import where importid = " & lblImportId.Caption, dbconn
    If dbrs.EOF Then
        MsgBox "order id not found", vbExclamation
        dbrs.Close
        Exit Sub
    End If
    isbn = dbrs.Fields("isbn")
    dbrs.Close
    dbconn.Execute "update book set quantity = quantity + " & lblQuantity.Caption & " where isbn=" & SqlQuote(isbn)
    lstBook_Click 'refresh this view
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdInCancel_Click()
    On Error GoTo die
    Dim isbn As String
    dbconn.Execute "update import set status = 3 where importid = " & lblImportId.Caption
    lstBook_Click 'refresh this view
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdInImport_Click()
    On Error GoTo die
    If txtQuantity.Text <> "" And txtInPrice.Text <> "" And _
        validateNumber(txtQuantity.Text, 10000, True, True) And validateNumber(txtInPrice.Text, 100000, False) And txtIsbn.BackColor = RGB(192, 255, 192) Then
        dbrs.Open "select top 1 * from book where isbn=" & SqlQuote(txtIsbn.Text), dbconn
        If dbrs.EOF Then
            MsgBox "invalid isbn", vbExclamation
            Exit Sub
        End If
        dbrs.Close
        dbconn.Execute "insert into import(isbn,price,quantity,status) values(" & _
            SqlQuote(txtIsbn.Text) & "," & txtInPrice.Text & "," & txtQuantity.Text & ", 0)" ' auto generates import id
        MsgBox "import order created!", vbInformation
    Else
        MsgBox "wrong data / isbn", vbExclamation
    End If
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdInPay_Click()
    On Error GoTo die
    Dim isbn As String
    dbconn.Execute "update import set status = 1 where importid = " & lblImportId.Caption
    dbrs.Open "select top 1 * from import where importid = " & lblImportId.Caption, dbconn
    If dbrs.EOF Then
        MsgBox "order id not found", vbExclamation
        dbrs.Close
        Exit Sub
    End If
    isbn = dbrs.Fields("isbn")
    dbrs.Close
    dbconn.Execute "insert into trade(isbn,price,quantity,tradetime) values (" & _
        SqlQuote(isbn) & "," & lblInPrice.Caption & "," & " -" & lblQuantity.Caption & _
        ", '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "')" 'negative: money outgoing
    lstBook_Click 'refresh this view
    Exit Sub
die:
    dbrs.Close
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdInTrade_Click()
    On Error GoTo die
    If txtQuantity.Text <> "" And txtInPrice.Text <> "" And _
    validateNumber(txtQuantity.Text, 10000, True, True) And validateNumber(txtInPrice.Text, 100000, False) And txtIsbn.BackColor = RGB(192, 255, 192) Then
        dbrs.Open "select top 1 * from book where isbn=" & SqlQuote(txtIsbn.Text), dbconn
        If dbrs.EOF Then
            MsgBox "invalid isbn", vbExclamation
            dbrs.Close
            Exit Sub
        ElseIf dbrs.Fields("quantity") - (Val(txtQuantity.Text)) < 0 Then
            MsgBox "no enough quantity in storage!" & vbCrLf & "current quantity: " & dbrs.Fields("quantity"), vbExclamation
            dbrs.Close
            Exit Sub
        End If
        dbconn.Execute "insert into trade(isbn,price,quantity,tradetime) values(" & _
            SqlQuote(txtIsbn.Text) & "," & txtInPrice.Text & "," & txtQuantity.Text & _
            ", '" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "')" 'positive: money incoming
        dbconn.Execute "update book set quantity=quantity-" & txtQuantity.Text & " where isbn=" & SqlQuote(txtIsbn.Text)
        MsgBox "trade success! " & (dbrs.Fields("quantity")) & " left.", vbInformation ' this is a reference!
        dbrs.Close
        
    Else
        MsgBox "wrong data / isbn", vbExclamation
    End If
    Exit Sub
die:
    dbrs.Close
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdLogin_Click()
    If dbconn.State <> 0 Then dbconn.Close
    If bLoggedIn And txtUsername.Text = "" Then
        GoTo logout
    End If
    On Error GoTo diedb
    dbconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db.mdb;Persist Security Info=False"
    dbrs.Open "select * from users where username=" & SqlQuote(txtUsername.Text), dbconn
    'dbrs.Open "select * from users", dbconn
    'On Error GoTo dieuser
    If dbrs.EOF Then GoTo dieuser
    MD5String txtPassword.Text
    If UCase(dbrs.Fields("md5password")) = UCase(GetMD5Text()) Then
        userLevel = dbrs.Fields("userlevel")
        txtWorkId.Text = dbrs.Fields("workid")
        txtInUsername.Text = dbrs.Fields("username")
        txtRealName.Text = CNStr(dbrs.Fields("realname"))
        txtGender.Text = CNStr(dbrs.Fields("gender"))
        picAll.Visible = True
        Dim bSuper As Boolean
        bSuper = (userLevel >= 1)
        txtInUsername.Locked = Not bSuper
        txtWorkId.Locked = Not bSuper
        picSuper.Visible = bSuper
    Else
        GoTo dieuser
    End If
    Me.Caption = "book db - " & txtUsername.Text
    bLoggedIn = True
    txtPassword.Text = ""
    txtUsername.Text = ""
    dbrs.Close
    Exit Sub
diedb:
    MsgBox "db error" & vbCrLf & Err.Description, vbCritical
    GoTo fail
dieuser:
    MsgBox "username / password wrong", vbExclamation
    GoTo fail
fail:
    dbrs.Close
    dbconn.Close
logout:
    picAll.Visible = False
    Me.Caption = "book db"
    bLoggedIn = False
    cmdLogin.Caption = "login"
End Sub

Private Sub cmdnew_Click()
    On Error GoTo die
    If txtIsbn.Text = "" Then
        MsgBox "isbn can't be empty.", vbExclamation
        Exit Sub
    End If
    dbconn.Execute "insert into book (isbn,title,author,press,price,quantity) values (" & _
        SqlQuote(txtIsbn.Text) & "," & SqlQuote(txtTitle.Text) & "," & SqlQuote(txtAuthor.Text) & "," & SqlQuote(txtPress.Text) & "," & _
        CNumNStr(txtPrice.Text) & "," & "0 )" 'new book quantity=0
    MsgBox "book created.", vbInformation
    cmdSearch_Click 0 'refresh the view
    lstBook.ListIndex = 0
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdNewUser_Click()
On Error GoTo die
    Dim pwmd5 As String
    If txtInPassword.Text <> txtInPasswordR.Text Then
        MsgBox "two password fields mismatch; user is not created."
        Exit Sub
    Else
        MD5String txtInPassword.Text
        pwmd5 = GetMD5Text()
        txtInPassword.Text = ""
        txtInPasswordR.Text = ""
    End If
    dbconn.Execute "insert into users (workid,username,realname,gender,age,md5password,userlevel) values (" & _
        SqlQuote(txtWorkId.Text) & "," & SqlQuote(txtInUsername.Text) & "," & SqlQuote(txtRealName.Text) & "," & _
        SqlQuote(txtGender.Text) & "," & CNumNStr(txtAge.Text) & "," & SqlQuote(pwmd5) & "," & "0 )"  'new user can only have level 0
    MsgBox "new user created.", vbInformation
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdQuery_Click()
    On Error GoTo die
    Dim dateStart As Date, dateEnd As Date
    dateStart = CDate(txtDateStart.Text)
    dateEnd = CDate(txtDateEnd.Text)
    If dateEnd = dateStart Then dateEnd = dateStart + 1
    
    On Error GoTo diehard
    dbrs.Open "select sum(quantity*price) from trade where (" & _
        "tradetime between " & "#" & Format(dateStart, "yyyy-MM-dd HH:mm:ss") & "#" & " and " & _
        "#" & Format(dateEnd, "yyyy-MM-dd HH:mm:ss") & "# )"
    If dbrs.EOF Then
        lblQuery.Caption = "empty"
    ElseIf IsNull(dbrs.Fields(0)) Then
        lblQuery.Caption = "empty"
    Else
        lblQuery.Caption = dbrs.Fields(0)
    End If
    dbrs.Close
    Exit Sub
die:
    MsgBox "wrong date. use yyyy-mm-dd [hh-mm-ss].", vbExclamation
    Exit Sub
diehard:
    dbrs.Close
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdRecord_Click()
    picRecord.Visible = True
    picTradeImport.Visible = False
    picPayCancel.Visible = False
    txtDateEnd.Text = Format(Now, "yyyy-MM-dd")
    If txtDateStart.Text = "" Then txtDateStart.Text = Format(Now, "yyyy-MM-dd")
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    On Error GoTo die
    lstBook.Clear
    Select Case Index
    Case 0
        If Not bImportProcess Then dbrs.Open "select top 100 * from book where isbn like " & SqlQuote(txtIsbn.Text, True), dbconn _
        Else dbrs.Open "select * from book inner join import on book.isbn=import.isbn where book.isbn like " & SqlQuote(txtIsbn.Text, True) & " order by status", dbconn
        If dbrs.EOF Then txtIsbn.BackColor = RGB(255, 192, 192)
    Case 1
        If Not bImportProcess Then dbrs.Open "select top 100 * from book where title like " & SqlQuote(txtTitle.Text, True), dbconn _
        Else dbrs.Open "select * from book inner join import on book.isbn=import.isbn where title like " & SqlQuote(txtTitle.Text, True) & " order by status", dbconn
        If dbrs.EOF Then txtTitle.BackColor = RGB(255, 192, 192)
    Case 2
        If Not bImportProcess Then dbrs.Open "select top 100 * from book where author like " & SqlQuote(txtAuthor.Text, True), dbconn _
        Else dbrs.Open "select * from book inner join import on book.isbn=import.isbn where author like " & SqlQuote(txtAuthor.Text, True) & " order by status", dbconn
        If dbrs.EOF Then txtAuthor.BackColor = RGB(255, 192, 192)
    Case 3
        If Not bImportProcess Then dbrs.Open "select top 100 * from book where press like " & SqlQuote(txtPress.Text, True), dbconn _
        Else dbrs.Open "select * from book inner join import on book.isbn=import.isbn where press like " & SqlQuote(txtPress.Text, True) & " order by status", dbconn
        If dbrs.EOF Then txtPress.BackColor = RGB(255, 192, 192)
    End Select
    
    Do Until dbrs.EOF
        If bImportProcess Then
            Dim status As Long, strstatus As String
            status = dbrs.Fields("status")
            strstatus = Array("Ordered", "Paid", "Arrived", "Canceled")(status)
            Select Case Index
                Case 0: lstBook.AddItem dbrs.Fields("importid") & vbTab & dbrs.Fields("book.isbn") & vbTab & strstatus & vbTab & dbrs.Fields("title") & vbTab & dbrs.Fields("press") & vbTab & dbrs.Fields("author")
                Case 1: lstBook.AddItem dbrs.Fields("importid") & vbTab & dbrs.Fields("book.isbn") & vbTab & strstatus & vbTab & dbrs.Fields("title") & vbTab & dbrs.Fields("press") & vbTab & dbrs.Fields("author")
                Case 2: lstBook.AddItem dbrs.Fields("importid") & vbTab & dbrs.Fields("book.isbn") & vbTab & strstatus & vbTab & dbrs.Fields("author") & vbTab & dbrs.Fields("title") & vbTab & dbrs.Fields("press")
                Case 3: lstBook.AddItem dbrs.Fields("importid") & vbTab & dbrs.Fields("book.isbn") & vbTab & strstatus & vbTab & dbrs.Fields("press") & vbTab & dbrs.Fields("title") & vbTab & dbrs.Fields("author")
            End Select
        Else
            Select Case Index
                Case 0: lstBook.AddItem dbrs.Fields("quantity") & vbTab & dbrs.Fields("isbn") & vbTab & dbrs.Fields("title") & vbTab & dbrs.Fields("press") & vbTab & dbrs.Fields("author")
                Case 1: lstBook.AddItem dbrs.Fields("quantity") & vbTab & dbrs.Fields("isbn") & vbTab & dbrs.Fields("title") & vbTab & dbrs.Fields("press") & vbTab & dbrs.Fields("author")
                Case 2: lstBook.AddItem dbrs.Fields("quantity") & vbTab & dbrs.Fields("isbn") & vbTab & dbrs.Fields("author") & vbTab & dbrs.Fields("title") & vbTab & dbrs.Fields("press")
                Case 3: lstBook.AddItem dbrs.Fields("quantity") & vbTab & dbrs.Fields("isbn") & vbTab & dbrs.Fields("press") & vbTab & dbrs.Fields("title") & vbTab & dbrs.Fields("author")
            End Select
        End If
        dbrs.MoveNext
    Loop
    dbrs.Close
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdtrade_Click()
    picRecord.Visible = False
    picTradeImport.Visible = True
    picPayCancel.Visible = False
    cmdInTrade.Visible = True
    cmdInImport.Visible = False
    If bImportProcess = True Then
        lstBook.Clear
    End If
    bImportProcess = False
    txtInPrice.Text = txtPrice.Text
    txtQuantity.Text = "1"
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo die
    dbconn.Execute "update book set " & _
        "title=" & SqlQuote(txtTitle.Text) & "," & _
        "author=" & SqlQuote(txtAuthor.Text) & "," & _
        "press=" & SqlQuote(txtPress.Text) & _
        "where isbn = " & SqlQuote(txtIsbn.Text)
        ' doesn't affect quantity
    lstBook_Click ' refresh the view
    MsgBox "book modified.", vbInformation
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdUser_Click()
    picUser.Visible = True
    picBook.Visible = False
End Sub

Private Sub cmdUserDelete_Click()
    On Error GoTo die
    dbrs.Open "select top 1 * from users where username=" & SqlQuote(txtInUsername.Text)
    
    If dbrs.EOF Then
        MsgBox "user '" & txtInUsername.Text & "' not found", vbExclamation
        dbrs.Close
        Exit Sub
    End If
    If dbrs.Fields("userlevel") = 1 Then
        MsgBox "can't delete super user", vbExclamation
        dbrs.Close
        Exit Sub
    End If
    Dim sure As Long
    sure = MsgBox("sure to delete user '" & txtInUsername.Text & "'?", vbQuestion Or vbYesNo Or vbDefaultButton2)
    If sure = vbYes Then
        dbconn.Execute "delete from users where username=" & SqlQuote(txtInUsername.Text)
        MsgBox "user deleted.", vbInformation
    End If
    dbrs.Close
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
    dbrs.Close
End Sub

Private Sub cmdUserSearch_Click(Index As Integer)
    On Error GoTo die
    If userLevel = 0 Then
        MsgBox "you are not permitted to use user search.", vbExclamation
        Exit Sub
    End If
    lstUser.Clear
    Select Case Index
    Case 0
        dbrs.Open "select top 100 * from users where workid like " & SqlQuote(txtWorkId.Text, True), dbconn
        If dbrs.EOF Then txtIsbn.BackColor = RGB(255, 192, 192)
    Case 1
        dbrs.Open "select top 100 * from users where username like " & SqlQuote(txtInUsername.Text, True), dbconn
        If dbrs.EOF Then txtTitle.BackColor = RGB(255, 192, 192)
    End Select
    
    Do Until dbrs.EOF
        lstUser.AddItem dbrs.Fields("username") & vbTab & dbrs.Fields("workid") & vbTab & dbrs.Fields("realname") & vbTab & dbrs.Fields("gender") & vbTab & dbrs.Fields("age")
        dbrs.MoveNext
    Loop
    dbrs.Close
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
    dbrs.Close
End Sub

Private Sub cmdUserUpdate_Click()
    On Error GoTo die
    Dim pwmd5 As String
    If txtInPassword.Text <> "" Then
        If txtInPassword.Text <> txtInPasswordR.Text Then
            MsgBox "two password fields mismatch; password is not updated."
        Else
            MD5String txtInPassword.Text
            pwmd5 = GetMD5Text()
            txtInPassword.Text = ""
            txtInPasswordR.Text = ""
        End If
    End If
    dbconn.Execute "update users set " & _
        IIf(pwmd5 = "", "", "md5password=" & SqlQuote(pwmd5) & ",") & _
        "workid=" & SqlQuote(txtWorkId.Text) & "," & _
        "realname=" & SqlQuote(txtRealName.Text) & "," & _
        "gender=" & SqlQuote(txtGender.Text) & _
        "where username=" & SqlQuote(txtInUsername.Text)
        ' doesn't affect quantity
    MsgBox "user info updated.", vbInformation
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
End Sub



Private Sub Form_Load()
    picAll.Move picBook.Left, picBook.Top, picBook.Width, (cmdUser.Top + cmdUser.Height - picBook.Top + 120)
    picAll.Visible = False
    
    SetParent picBook.hWnd, picAll.hWnd
    SetParent picUser.hWnd, picAll.hWnd
    SetParent cmdBook.hWnd, picAll.hWnd
    SetParent cmdUser.hWnd, picAll.hWnd
    
    picUser.Move 0, 0
    picUser.Visible = False
    picBook.Move 0, 0
    picBook.Visible = True
    
    cmdBook.Move cmdBook.Left - picAll.Left, cmdBook.Top - picAll.Top
    cmdUser.Move cmdUser.Left - picAll.Left, cmdUser.Top - picAll.Top
    
    SetParent picPayCancel.hWnd, picBook.hWnd
    picPayCancel.Move picTradeImport.Left, picTradeImport.Top
    
    SetParent picRecord.hWnd, picBook.hWnd
    picRecord.Move picTradeImport.Left, picTradeImport.Top
    
    picTradeImport.Visible = False
    picPayCancel.Visible = False
    
    cmdInTrade.Move cmdInImport.Left, cmdInImport.Top
    cmdInTrade.Visible = False
    cmdInImport.Visible = False
    
    Me.Width = picAll.Width + (Me.Width - Me.ScaleWidth)
    Me.Height = picAll.Height + picAll.Top + (Me.Height - Me.ScaleHeight)
    'txtUsername.Text = "fart"
    'txtPassword.Text = "fart"
    
End Sub

Private Sub lstBook_Click()
    On Error GoTo die
    
    If lstBook.ListIndex = -1 Or lstBook.Text = "" Then Exit Sub
    If bImportProcess Then
        txtIsbn.Text = Split(lstBook.Text, vbTab)(1)
        lblImportId = Split(lstBook.Text, vbTab)(0)
        dbrs.Open "select top 1 * from book where isbn = " & SqlQuote(txtIsbn.Text, False), dbconn
    Else
        txtIsbn.Text = Split(lstBook.Text, vbTab)(1)
        dbrs.Open "select top 1 * from book where isbn = " & SqlQuote(txtIsbn.Text, False), dbconn
    End If
    If dbrs.EOF Then
        txtIsbn.BackColor = RGB(255, 192, 192)
    Else
        txtTitle.Text = dbrs.Fields("title")
        txtAuthor.Text = dbrs.Fields("author")
        txtPress.Text = dbrs.Fields("press")
        txtPrice.Text = dbrs.Fields("price")
        
        txtIsbn.BackColor = RGB(192, 255, 192)
        txtTitle.BackColor = RGB(192, 255, 192)
        txtAuthor.BackColor = RGB(192, 255, 192)
        txtPress.BackColor = RGB(192, 255, 192)
        txtPrice.BackColor = RGB(192, 255, 192)
    End If
    If bImportProcess Then
        dbrs.Close
        dbrs.Open "select * from import where importid = " & lblImportId.Caption, dbconn
        If dbrs.EOF Then
            lblImportId.BackColor = RGB(255, 192, 192)
        Else
            lblImportId.BackColor = &H8000000F
            lblInPrice.Caption = dbrs.Fields("price")
            lblQuantity.Caption = dbrs.Fields("quantity")
            lblTotalPay.Caption = CStr(dbrs.Fields("price") * dbrs.Fields("quantity"))
            Dim status As Long
            status = dbrs.Fields("status")
            lblStatus.Caption = Array("Ordered", "Paid", "Arrived", "Canceled")(status)
            
            cmdInPay.Enabled = (status = 0)
            cmdInCancel.Enabled = (status = 0)
            cmdInArrival.Enabled = (status = 1)
        End If
    End If
    dbrs.Close
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
    dbrs.Close
End Sub

Private Sub lstUser_Click()
    On Error GoTo die
    
    If lstUser.ListIndex = -1 Or lstUser.Text = "" Then Exit Sub
    txtInUsername.Text = Split(lstUser.Text, vbTab)(0)
    dbrs.Open "select top 1 * from users where username = " & SqlQuote(txtInUsername.Text, False), dbconn
    If dbrs.EOF Then
        txtInUsername.BackColor = RGB(255, 192, 192)
    Else
        txtInUsername.Text = CNStr(dbrs.Fields("username"))
        txtWorkId.Text = CNStr(dbrs.Fields("workid"))
        txtRealName.Text = CNStr(dbrs.Fields("realname"))
        txtGender.Text = CNStr(dbrs.Fields("gender"))
        txtAge.Text = CNStr(dbrs.Fields("age"))
        
        txtInUsername.BackColor = RGB(192, 255, 192)
    End If
    
    dbrs.Close
    Exit Sub
die:
    MsgBox Err.Description, vbCritical
    dbrs.Close

End Sub

Private Sub txtAuthor_Change()
    txtAuthor.BackColor = &H80000005
End Sub

Private Sub txtAuthor_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSearch_Click 2
End Sub

Private Sub txtInPrice_Change()
    If Not validateNumber(txtInPrice.Text, 100000, False) Then
        txtInPrice.BackColor = RGB(255, 192, 192)
    Else
        txtInPrice.BackColor = &H80000005
        If validateNumber(txtQuantity.Text, 10000, True) Then lblTotal.Caption = Val(txtInPrice.Text) * Val(txtQuantity.Text)
    End If
End Sub

Private Sub txtInUsername_Change()
    txtInUsername.BackColor = &H80000005
End Sub

Private Sub txtInUsername_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdUserSearch_Click 1
End Sub

Private Sub txtIsbn_Change()
    txtIsbn.BackColor = &H80000005
    txtTitle.BackColor = &H80000005
    txtAuthor.BackColor = &H80000005
    txtPress.BackColor = &H80000005
    txtPrice.BackColor = &H80000005
End Sub

Private Sub txtIsbn_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSearch_Click 0
End Sub

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdLogin_Click
End Sub

Private Sub txtPress_Change()
    txtPress.BackColor = &H80000005
End Sub

Private Sub txtPress_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSearch_Click 3
End Sub

Private Sub txtPrice_Change()
    If Not validateNumber(txtPrice.Text, 100000, False) Then
        txtPrice.BackColor = RGB(255, 192, 192)
    Else
        txtPrice.BackColor = &H80000005
    End If
End Sub

Private Sub txtQuantity_Change()
    If Not validateNumber(txtQuantity.Text, 10000, True) Then
        txtQuantity.BackColor = RGB(255, 192, 192)
    Else
        txtQuantity.BackColor = &H80000005
        If validateNumber(txtInPrice.Text, 100000, False) Then lblTotal.Caption = Val(txtInPrice.Text) * Val(txtQuantity.Text)
    End If
End Sub

Private Sub txtTitle_Change()
    txtTitle.BackColor = &H80000005
End Sub

Private Function validateNumber(ByRef s As String, range As Double, Optional ByVal bInteger As Boolean = False, Optional nozero As Boolean = False) As Boolean
    On Error GoTo die
    validateNumber = True
    If s = "" Then Exit Function
    If IsNumeric(s) Then
        Dim num As Double, cs As String
        num = Val(s)
        If num >= 0 And Not (num = 0 And nozero) And num < range And ((Not bInteger) Or Int(num) = num) Then
            cs = CStr(num)
            If s Like "*.*" And (Not cs Like "*.*") Then cs = cs + "."
            If InStr(1, s, cs) <> 0 And isEveryChar(Right(s, Len(s) - InStr(1, s, cs) + 1 - Len(cs)), "0") Then Exit Function
        End If
    End If
die:
    validateNumber = False
End Function

Private Function isEveryChar(ByVal s As String, ByVal char As String) As Boolean
    Dim xhl As Long
    isEveryChar = True
    For xhl = 0 To Len(s) - 1
        If Mid(s, xhl + 1, 1) <> char Then
            isEveryChar = False
            Exit Function
        End If
    Next xhl
End Function


Private Sub txtTitle_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSearch_Click 1
End Sub

Private Sub txtUsername_Change()
    If txtUsername.Text = "" And bLoggedIn Then cmdLogin.Caption = "logout" Else cmdLogin.Caption = "login"
End Sub

Private Sub txtWorkId_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdUserSearch_Click 0
End Sub
