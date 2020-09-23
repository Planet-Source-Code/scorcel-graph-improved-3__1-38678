VERSION 5.00
Begin VB.Form GRAPH 
   Caption         =   "GRAPH"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13620
   Icon            =   "Graph.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra_Cont 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "YOU CAN RESIZE FORM AND SEE!!!"
      Top             =   3240
      Width           =   13335
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   8520
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   6720
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Frame fra_ChartType 
         Caption         =   "Option"
         Height          =   1185
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   3045
         Begin VB.OptionButton OptVal 
            Caption         =   "Line"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptVal 
            Caption         =   "3D Column"
            Height          =   345
            Index           =   5
            Left            =   1800
            TabIndex        =   10
            Top             =   180
            Width           =   1125
         End
         Begin VB.OptionButton OptVal 
            Caption         =   "Column"
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   9
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton OptVal 
            Caption         =   "3D Bar"
            Height          =   195
            Index           =   4
            Left            =   1800
            TabIndex        =   8
            Top             =   840
            Width           =   945
         End
         Begin VB.OptionButton OptVal 
            Caption         =   "3D Area"
            Height          =   195
            Index           =   3
            Left            =   1800
            TabIndex        =   7
            Top             =   600
            Width           =   945
         End
         Begin VB.OptionButton OptVal 
            Caption         =   "Bar"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   810
            Width           =   975
         End
         Begin VB.OptionButton OptVal 
            Caption         =   "Area"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   570
            Width           =   945
         End
      End
      Begin VB.ListBox lstFields 
         Height          =   1035
         Left            =   6720
         MultiSelect     =   1  'Simple
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.ListBox lst_Tables 
         Height          =   1035
         ItemData        =   "Graph.frx":08CA
         Left            =   3720
         List            =   "Graph.frx":08D7
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Fields to Graph"
         Height          =   255
         Left            =   6720
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Source Table"
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   $"Graph.frx":0908
         Height          =   855
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   5775
      End
   End
   Begin VB.OLE Ole_Graph 
      AutoActivate    =   0  'Manual
      AutoVerbMenu    =   0   'False
      Class           =   "MSGraph.Chart.8"
      Height          =   3255
      Left            =   0
      OleObjectBlob   =   "Graph.frx":0A17
      TabIndex        =   0
      Top             =   0
      Width           =   13635
   End
End
Attribute VB_Name = "GRAPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I just couldn't locate codes for smooth lines graph so I created my own
'Ive been looking for this code for long
'I hope this helps other programmers out there!
'For comments and bugs please e-mail me
'j_jirehp@yahoo.com
'you may use my code freely but please leave some credits to me
'please do leave some comments and I believe this deserves your vote.

'Update Sept 7, 2002
'Added feature to Dynamically select what fields to Graph
'Added feature to Load available fields for the Source Table
'Thanks for your Votes!!!!

'Update Sept 13, 2002
'Added feature to select different table sources with the same graph to reuse
'shows the power of the graph function
'Thanks for your vote

'Update Sept 27, 2002
'Added resize of graph

'Sept 28 2002
'Include Graph Printing
'I take no credit in the Orientation Module I just downloaded it also in PSC
'Though I just made few changes to fit to my requirement


Option Explicit
Dim strTbl As String 'Declare table name
Public db As Database

Private Sub cmdPrint_Click()
'prepares your printer to a landscape orientation
ChngPrinterOrientationLandscape Me
Me.PrintForm
End Sub

Private Sub cmdRefresh_Click()
'Ignore refresh if no Fields are selected
If lstFields.SelCount < 1 Then Exit Sub

'Do Graph and select fields to Fill the ListBox
DoGraph GetSelect & " FROM " & strTbl & " ORDER BY DateCollected "
End Sub
Private Function GetRecordSet(Str As String) As Recordset
'Dynamic Recordset Collection
Set db = OpenDatabase(App.Path & "\TestDB.mdb")
Set GetRecordSet = db.OpenRecordset(Str)
End Function
Private Function Graphing(Str As String) As String
Dim RST As Recordset
Dim X As Integer
Dim TB, NL
Graphing = ""
'tab character
TB = Chr(9)
'new line character
NL = Chr(10)
'blank out the trend data
Set RST = GetRecordSet(Str)

'This sets header
For X = 1 To RST.Fields.Count - 1
    Graphing = Graphing + TB
    Graphing = Graphing & RST(X).Name
Next

'This sets Data here..
While Not RST.EOF
    Graphing = Graphing + NL
    For X = 0 To RST.Fields.Count - 1
        Graphing = Graphing & RST(X)
        If X <> RST.Fields.Count Then
            Graphing = Graphing + TB
        End If
    Next
RST.MoveNext
Wend
Set RST = Nothing
End Function
Private Sub Form_Load()
Dim Str As String
'Sets default Graph type
OptVal_Click 4
OptVal_Click 3

strTbl = "tbl_Collections" 'Set initial table source
'Sets default data to graph
Str = "Select DateCollected AS [Shipped Date], Amount AS [Actual Amount], 700 AS Target FROM tbl_Collections ORDER BY DateCollected "
DoGraph Str 'Execute Function giving the initial select statement
LoadFields 'Load available fields for the initial table source
End Sub
Private Sub Form_Resize()
On Error Resume Next
Ole_Graph.Width = Me.Width - 150
Ole_Graph.Height = Me.Height - 2800
fra_Cont.Top = Me.Height - 2800
End Sub
Private Sub Form_Unload(Cancel As Integer)
db.Close 'End database connection
Set db = Nothing 'Free Up memory
End Sub
Private Sub lst_Tables_Click()
strTbl = lst_Tables.Text 'Change source of data
LoadFields 'Load Available fields for the new table
End Sub
Private Sub OptVal_Click(Index As Integer)
Dim OptionValue As Integer
Select Case Index
    Case 0
        OptionValue = 1 'set Graph to Area
    Case 1
        OptionValue = 2 'set Graph to Bar
    Case 2
        OptionValue = 3 'set Graph to Column
    Case 3
        OptionValue = -4098 'set Graph to 3D Area
    Case 4
        OptionValue = -4099 'set Graph to 3D Bar
    Case 5
        OptionValue = -4100 'set Graph to 3D Column
    Case 6
        OptionValue = 4 'set Graph to Line
End Select
    Ole_Graph.Object.Application.Chart.Type = OptionValue
    
End Sub
Private Sub DoGraph(Str As String)
Dim DataToTrend As String
With Ole_Graph
   'set the file format to text
   .Format = "CF_TEXT"
   'stretch the graph object to fit the ole container
   .SizeMode = 1
End With
'Please remember that first column of your select must be date or any identifier for your data
'the rest of the colums will be your data to be graphed

DataToTrend = Graphing(Str)
'OLE.Object.Application.Chart.Type = 4
With Ole_Graph
   'activate MSGRAPH as hidden
   .DoVerb -3
   If .AppIsRunning Then
      'send the data to trend
      .DataText = DataToTrend
      'update the object
      .Update
   Else
      MsgBox "Graph isn't active", , "JINGPOLS/SCORCEL Graph"
   End If
End With

End Sub
Private Sub LoadFields() ' This collects all fields available in the table
Dim Str As String
Dim RST As Recordset
Dim I As Integer

'Clears the ListBox
lstFields.Clear

'Querries table for Fields
Str = "Select * from " & strTbl & ""
Set RST = GetRecordSet(Str)
If RST.EOF Then Exit Sub 'Exit when no fields is found
For I = 1 To RST.Fields.Count - 1
    lstFields.AddItem (RST(I).Name) 'Add them
Next
Set RST = Nothing
End Sub
Function GetSelect() As String 'Collect Fields to Trend
Dim I As Integer
GetSelect = ""
If lstFields.SelCount > 0 Then
    For I = 0 To lstFields.ListCount - 1 'Get number of Fields to Collect
        If lstFields.Selected(I) = True Then
            lstFields.ListIndex = I
            If GetSelect <> "" Then
                GetSelect = GetSelect & ", " & lstFields
            Else
                GetSelect = "Select DateCollected, " & lstFields.Text 'Initial Value Only
            End If
        End If
    Next
End If
End Function
