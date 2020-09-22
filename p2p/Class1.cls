VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPeerA"
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

Public Sub ReturnData(Buffer As String)
    ' vars for udt's
    Dim udtDetail As DetailData
    Dim udtSummary As SummaryData
    
    ' var for textbox object
    Dim oTextBox As Object

    Set oTextBox = frmPeerA.txtOutput
    
    ' store the buffer value into the summary udt
    udtSummary.Buffer = Buffer
    
    ' move flat file to detail udt
    LSet udtDetail = udtSummary
    
    ' send the data to textbox object
    With udtDetail
        oTextBox.Text = oTextBox.Text & "First Name: " & Trim(.FirstName) & vbCrLf
        oTextBox.Text = oTextBox.Text & "Last Name: " & Trim(.LastName) & vbCrLf
        oTextBox.Text = oTextBox.Text & "Department: " & Trim(.Department) & vbCrLf
    End With
    
    ' At this point you could actualy send the data to a database or whatever
    ' creating a central data store server. with a little more work you could
    ' also retrieve the data and send it back.
End Sub

