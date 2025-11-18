Option Compare Database
Option Explicit

Public Sub ExportObjectsToExcel()
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    
    Dim rowNum As Long
    Dim connDict As Object ' Scripting.Dictionary for distinct connections
    Dim connStr As String
    Dim key As Variant

    ' Use the currently opened database
    Set db = CurrentDb
    
    ' Create Excel instance (late binding – no reference needed)
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True ' show Excel
    
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets(1)
    
    ' Headers
    xlWs.Cells(1, 1).Value = "ObjectName"
    xlWs.Cells(1, 2).Value = "ObjectType"
    xlWs.Cells(1, 3).Value = "Details / SQL / Connection"
    
    rowNum = 2
    
    ' Dictionary for distinct connection strings
    Set connDict = CreateObject("Scripting.Dictionary")
    
    ' --- 1. Tables ---
    For Each tdf In db.TableDefs
        
        ' Skip system & hidden tables (MSys*, etc.)
        If Left$(tdf.Name, 4) <> "MSys" Then
            xlWs.Cells(rowNum, 1).Value = tdf.Name
            xlWs.Cells(rowNum, 2).Value = "Table"
            
            ' If it's a linked table, tdf.Connect will hold the connection string
            If Len(tdf.Connect & "") > 0 Then
                xlWs.Cells(rowNum, 3).Value = tdf.Connect
                
                ' Collect distinct connections
                connStr = tdf.Connect
                If Not connDict.Exists(connStr) Then
                    connDict.Add connStr, True
                End If
            End If
            
            rowNum = rowNum + 1
        End If
    Next tdf
    
    ' --- 2. Queries ---
    For Each qdf In db.QueryDefs
        ' Skip internal / temporary queries (~)
        If Left$(qdf.Name, 1) <> "~" Then
            xlWs.Cells(rowNum, 1).Value = qdf.Name
            xlWs.Cells(rowNum, 2).Value = "Query"
            xlWs.Cells(rowNum, 3).Value = qdf.SQL
            
            rowNum = rowNum + 1
        End If
    Next qdf
    
    ' --- 3. Connections (distinct) ---
    If connDict.Count > 0 Then
        ' Optional: blank row as separator
        rowNum = rowNum + 1
        
        For Each key In connDict.Keys
            xlWs.Cells(rowNum, 1).Value = key
            xlWs.Cells(rowNum, 2).Value = "Connection"
            ' Column 3 left empty or you can repeat key there as well
            rowNum = rowNum + 1
        Next key
    End If
    
    ' Format sheet
    With xlWs
        .Columns("A:C").EntireColumn.AutoFit
        .Rows("1:1").Font.Bold = True
        .Name = "AccessObjects"
    End With
    
    MsgBox "Export completed. Please review and save the Excel file.", vbInformation

Exit_Handler:
    On Error Resume Next
    Set qdf = Nothing
    Set tdf = Nothing
    Set db = Nothing
    ' Do not quit Excel – user may want to save manually
    Exit Sub

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "ExportObjectsToExcel"
    Resume Exit_Handler
End Sub
