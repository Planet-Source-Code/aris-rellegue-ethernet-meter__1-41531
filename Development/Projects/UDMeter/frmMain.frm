VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Upload/Download Meter"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timNICStat 
      Interval        =   1000
      Left            =   2505
      Top             =   2250
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   420
      Left            =   3600
      TabIndex        =   0
      Top             =   2340
      Width           =   1560
   End
   Begin VB.Label lblNIC 
      Alignment       =   2  'Center
      Caption         =   "NIC:"
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   45
      Width           =   5100
   End
   Begin VB.Label lblMTU 
      Height          =   315
      Left            =   2535
      TabIndex        =   10
      Top             =   1845
      Width           =   2400
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "MTU:"
      Height          =   270
      Left            =   690
      TabIndex        =   9
      Top             =   1845
      Width           =   1650
   End
   Begin VB.Label lblUpTroughput 
      Caption         =   "0"
      Height          =   285
      Left            =   2535
      TabIndex        =   8
      Top             =   1500
      Width           =   2430
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Upload Thoughput:"
      Height          =   225
      Left            =   675
      TabIndex        =   7
      Top             =   1500
      Width           =   1665
   End
   Begin VB.Label lblDownTroughput 
      Caption         =   "0"
      Height          =   285
      Left            =   2535
      TabIndex        =   6
      Top             =   1170
      Width           =   2430
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Download Throughput:"
      Height          =   240
      Left            =   675
      TabIndex        =   5
      Top             =   1170
      Width           =   1650
   End
   Begin VB.Label lblSent 
      Caption         =   "0"
      Height          =   285
      Left            =   2535
      TabIndex        =   4
      Top             =   840
      Width           =   2430
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Bytes Sent (KB):"
      Height          =   225
      Left            =   315
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblReceived 
      Caption         =   "0"
      Height          =   285
      Left            =   2535
      TabIndex        =   2
      Top             =   510
      Width           =   2430
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Bytes Received (KB):"
      Height          =   270
      Left            =   315
      TabIndex        =   1
      Top             =   510
      Width           =   1980
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngOldReceived As Long
Dim lngOldSent As Long


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
GetEthernetInfo
End Sub

Private Sub timNICStat_Timer()

GetEthernetInfo

End Sub


Private Sub GetEthernetInfo()

   Dim IPInterfaceRow As MIB_IFROW
   Dim buff() As Byte
   Dim cbRequired As Long
   Dim nStructSize As Long
   Dim nRows As Long
   Dim cnt As Long
   Dim n As Long

   Dim tmp As String
   
   Call GetIfTable(ByVal 0&, cbRequired, 1)

   If cbRequired > 0 Then
    
      ReDim buff(0 To cbRequired - 1) As Byte
      
      If GetIfTable(buff(0), cbRequired, 1) = ERROR_SUCCESS Then
      
        'saves using LenB in the CopyMemory calls below
         nStructSize = LenB(IPInterfaceRow)
   
        'first 4 bytes is a long indicating the
        'number of entries in the table
         CopyMemory nRows, buff(0), 4
      
         For cnt = 1 To nRows
         
           'moving past the four bytes obtained
           'above, get one chunk of data and cast
           'into an IPInterfaceRow type
            CopyMemory IPInterfaceRow, buff(4 + (cnt - 1) * nStructSize), nStructSize
            
               
             'Only get the first ethernet interface found
             If IPInterfaceRow.dwType = MIB_IF_TYPE_ETHERNET Then
               
               lblNIC = GetName(IPInterfaceRow.bDescr)
               
               lngOldReceived = CLng(lblReceived)
               
               'lblReceived = FormatNumber(((IPInterfaceRow.dwInUcastPkts + IPInterfaceRow.dwInNUcastPkts) * IPInterfaceRow.dwMtu) / 1024, 0)
               lblReceived = FormatNumber(((IPInterfaceRow.dwInUcastPkts) * IPInterfaceRow.dwMtu) / 1024, 0)
               
               lngOldSent = CLng(lblSent)
               
               lblSent = FormatNumber(((IPInterfaceRow.dwOutUcastPkts + IPInterfaceRow.dwOutNUcastPkts) * IPInterfaceRow.dwMtu) / 1024, 0)
               
               lblDownTroughput = FormatNumber(((CLng(lblReceived) - lngOldReceived)), 0) & " KB/sec."
               
               lblUpTroughput = FormatNumber(((CLng(lblSent) - lngOldSent)) / 1024, 0) & " KB/sec."
               
               lblMTU = FormatNumber(IPInterfaceRow.dwMtu, 0)
               
               Exit For
                
             End If
               
          Next cnt
          
      End If  'If GetIfTable( ...
      
   End If  'If cbRequired > 0

End Sub



Private Function GetName(ByRef arr) As String
Dim i As Integer
Dim str As String

For i = 0 To UBound(arr)
    str = str + Chr(arr(i))
Next

GetName = str

End Function
