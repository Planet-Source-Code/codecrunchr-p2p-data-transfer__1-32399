VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeerB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' User defined type for storing detail data
Private Type DetailData
    FirstName As String * 30
    LastName As String * 30
    Department As String * 30
End Type

' User Defined type for storeing information into flat file.
Private Type SummaryData
    Buffer As String * 180
End Type

Function StoreData()
    ' vars for udt's
    Dim udtDetail As DetailData
    Dim udtSummary As SummaryData
    
    ' store values into detail udt
    With udtDetail
        .FirstName = frmPeerB.txtFirstName.Text
        .LastName = frmPeerB.txtLastName.Text
        .Department = frmPeerB.txtDepartment.Text
    End With
    
    ' store detail into flat file udt
    LSet udtSummary = udtDetail
    
    ' store flat file into function var.
    StoreData = udtSummary.Buffer
End Function

