VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   " ""D:\Program Files\Microsoft Office\Office10\Samples\Northwind.mdb"""
      Top             =   1440
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start "
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'http://support.microsoft.com/support/kb/articles/Q246/3/35.ASP

'The information in this example applies to:
'
'Microsoft Excel 2002
'Microsoft Excel 2000
'Microsoft Excel 97 for Windows
'Microsoft Visual Basic Professional Edition for Windows, versions 5.0, 6.0
'Microsoft Visual Basic Enterprise Edition for Windows, versions 5.0, 6.0
'ActiveX Data Objects (ADO), versions 2.0, 2.1, 2.5
'
'--------------------------------------------------------------------------------
'
'
'SUMMARY
'
'You can transfer the contents of an ADO recordset to a Microsoft Excel worksheet
'by automating Excel. The approach that you can use depends on the version of Excel
'you are automating. Excel 97, Excel 2000, and Excel 2002 have a CopyFromRecordset
'method that you can use to transfer a recordset to a range. CopyFromRecordset in
'Excel 2000 and 2002 can be used to copy either a DAO or an ADO recordset.
'However, CopyFromRecordset in Excel 97 supports only DAO recordsets.
'To transfer an ADO recordset to Excel 97, you can create an array from the recordset
'and then populate a range with the contents of that array.
'
'This article discusses both approaches.
'The sample code presented illustrates how you can transfer an ADO recordset to Excel 97,
'Excel 2000, or Excel 2002.
'
'
'
'MORE Information
'
'The code sample provided below shows how to copy an ADO recordset to a Microsoft Excel
'worksheet using automation from Microsoft Visual Basic. The code first checks the version
'of Excel. If Excel 2000 or 2002 is detected, the CopyFromRecordset method is used because
'it is efficient and requires less code. However, if Excel 97 or earlier is detected,
'the recordset is first copied to an array using the GetRows method of the ADO recordset
'object. The array is then transposed so that records are in the first dimension (in rows),
'and fields are in the second dimension (in columns).
'Then, the array is copied to an Excel worksheet through assigning the array to a range of
'cells. (The array is copied in one step rather than looping through each cell in the
'worksheet.)
'
'The code sample uses the Northwind sample database that is included with Microsoft Office.
'If you selected the default folder when you installed Microsoft Office, the database
'is located in:
'
'\Program Files\Microsoft Office\Office\Samples\Northwind.mdb
'
'If the Northwind database is located in a different folder on your computer,
'you need to edit the path of the database in the code provided below.
'
'If you do not have the Northwind database installed on your system,
'you can use the Add/Remove option for Microsoft Office setup to install the
'sample databases.
'
'Steps to Create Sample
'Start Visual Basic and create a new Standard EXE project. Form1 is created by default.
'
'
'Add a CommandButton to Form1.
'
'
'Click References from the Project menu. Add a reference to the
'Microsoft ActiveX Data Objects 2.1 Library.
'
'
'Paste the following code into the code section of Form1:


Private Sub Command1_Click()
    Dim cnt As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    
    Dim recArray As Variant
    
    Dim strDB As String
    Dim fldCount As Integer
    Dim recCount As Long
    Dim iCol As Integer
    Dim iRow As Integer
    
    ' Set the string to the path of your Northwind database
    strDB = Text1.Text
  
    ' Open connection to the database
    cnt.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & strDB & ";"
        
    ' Open recordset based on Orders table
    rst.Open "Select * From Orders", cnt
    
    ' Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1) '("Sheet1")
  
    ' Display Excel and give user control of Excel's lifetime
    xlApp.Visible = True
    xlApp.UserControl = True
    
    ' Copy field names to the first row of the worksheet
    fldCount = rst.Fields.Count
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = rst.Fields(iCol - 1).Name
    Next
        
    ' Check version of Excel
    If Val(Mid(xlApp.Version, 1, InStr(1, xlApp.Version, ".") - 1)) > 8 Then
        'EXCEL 2000 or 2002: Use CopyFromRecordset
         
        ' Copy the recordset to the worksheet, starting in cell A2
        xlWs.Cells(2, 1).CopyFromRecordset rst
        'Note: CopyFromRecordset will fail if the recordset
        'contains an OLE object field or array data such
        'as hierarchical recordsets
        
    Else
        'EXCEL 97 or earlier: Use GetRows then copy array to Excel
    
        ' Copy recordset to an array
        recArray = rst.GetRows
        'Note: GetRows returns a 0-based array where the first
        'dimension contains fields and the second dimension
        'contains records. We will transpose this array so that
        'the first dimension contains records, allowing the
        'data to appears properly when copied to Excel
        
        ' Determine number of records

        recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array
        

        ' Check the array for contents that are not valid when
        ' copying the array to an Excel worksheet
        For iCol = 0 To fldCount - 1
            For iRow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
        Next iCol 'next field
            
        ' Transpose and Copy the array to the worksheet,
        ' starting in cell A2
        xlWs.Cells(2, 1).Resize(recCount, fldCount).Value = _
            TransposeDim(recArray)
    End If

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit

    ' Close ADO objects
    rst.Close
    cnt.Close
    Set rst = Nothing
    Set cnt = Nothing
    
    ' Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing

    Set xlApp = Nothing

End Sub


Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X
    
    TransposeDim = tempArray

End Function
 

