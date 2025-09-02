' Module: book-1.vba

Option Explicit

' ========================
' Main Function: Get File List
' ========================
Sub ListFilesInFolderAndSubfolders()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim row As Long
    Dim lastRow As Long
    
    Dim fso As Object
    Dim folder As Object
    
    Set ws = Sheets("Sheet1")
    folderPath = "C:\test101\"   ' <<< เปลี่ยนเป็นโฟลเดอร์ที่ต้องการ
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    row = 4   ' <<< เริ่มแสดงข้อมูลตั้งแต่แถว 4
    
    ' Clear old data (เฉพาะตาราง ไม่ลบ SearchBox)
    ws.Range("A3:G" & ws.Rows.Count).Clear
    
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 12
    
    ' ====== Search Box ======
    ws.Cells(1, 1).Value = "Search File:"
    ws.Cells(1, 2).ClearContents
    ws.Cells(1, 2).Interior.Color = RGB(255, 255, 200)
    ws.Cells(1, 2).Font.Color = RGB(0, 0, 0)
    ws.Cells(1, 2).Name = "SearchBox"   ' ตั้งชื่อ B1
    
    ' ====== Add Refresh Button next to SearchBox ======
    Dim btn As Button
    ' ลบปุ่มเดิมก่อน (ถ้ามี)
    On Error Resume Next
    ws.Buttons("btnRefresh").Delete
    On Error GoTo 0
    
    ' สร้างปุ่มใหม่ที่ C1
    Set btn = ws.Buttons.Add(ws.Cells(1, 3).Left, ws.Cells(1, 3).Top, ws.Cells(1, 3).Width, ws.Cells(1, 3).Height)
    With btn
        .OnAction = "ListFilesInFolderAndSubfolders"   ' กดแล้วรีเฟรช
        .Caption = "Refresh"
        .Name = "btnRefresh"
    End With
    
    ' ====== Table Header ======
    ws.Cells(3, 1).Value = "File Name"
    ws.Cells(3, 2).Value = "File Extension"
    ws.Cells(3, 3).Value = "File Size"
    ws.Cells(3, 4).Value = "Date Created"
    ws.Cells(3, 5).Value = "Last Modified"
    ws.Cells(3, 6).Value = "Folder Path"
    ws.Cells(3, 7).Value = "Open File"
    
    ' Header formatting
    With ws.Range("A3:G3")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 112, 192)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Get all files
    Call ListFilesInFolderRecursive(folder, ws, row)
    
    ' ===== Format Table =====
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Sort by Last Modified
    If lastRow > 4 Then
        ws.Range("A4:G" & lastRow).Sort Key1:=ws.Range("E4"), Order1:=xlDescending, Header:=xlNo
    End If
    
    ' Apply borders
    If lastRow > 3 Then
        With ws.Range("A3:G" & lastRow).Borders
            .LineStyle = xlContinuous
            .Color = RGB(200, 200, 200)
            .Weight = xlThin
        End With
    End If
    
    ' Column widths
    ws.Columns("A").ColumnWidth = 40
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 50
    ws.Columns("G").ColumnWidth = 12
    
    ' Wrap text in file name
    ws.Columns("A").WrapText = True
    
    ' AutoFit row height
    ws.Rows("4:" & lastRow).AutoFit
    
    ' Alternate row color
    Dim i As Long
    For i = 4 To lastRow
        If i Mod 2 = 0 Then
            ws.Range("A" & i & ":G" & i).Interior.Color = RGB(235, 241, 222) ' เขียวอ่อน
        Else
            ws.Range("A" & i & ":G" & i).Interior.Color = RGB(242, 242, 242) ' เทาอ่อน
        End If
    Next i
    
    ' Show total files
    ws.Cells(lastRow + 2, 1).Value = "Total Files:"
    ws.Cells(lastRow + 2, 1).Font.Bold = True
    ws.Cells(lastRow + 2, 2).Value = row - 4
    
    ' Enable AutoFilter
    ws.Range("A3:G" & lastRow).AutoFilter
    
    MsgBox "File list updated successfully!", vbInformation, "Completed"
End Sub


' ========================
' Recursive Function: Read Files
' ========================
Sub ListFilesInFolderRecursive(ByVal folder As Object, ByRef ws As Worksheet, ByRef row As Long)
    Dim file As Object
    Dim subfolder As Object
    
    For Each file In folder.Files
        ws.Cells(row, 1).Value = file.Name
        ws.Cells(row, 2).Value = Mid(file.Name, InStrRev(file.Name, ".") + 1)
        ws.Cells(row, 3).Value = Format(file.Size, "#,##0") & " Bytes"
        ws.Cells(row, 4).Value = file.DateCreated
        ws.Cells(row, 5).Value = file.DateLastModified
        ws.Cells(row, 6).Value = folder.Path
        
        ' Hyperlink to open file
        ws.Hyperlinks.Add Anchor:=ws.Cells(row, 7), _
            Address:=folder.Path & "\" & file.Name, _
            TextToDisplay:="Open"
        
        ' Align center
        ws.Cells(row, 3).HorizontalAlignment = xlCenter
        ws.Cells(row, 4).HorizontalAlignment = xlCenter
        ws.Cells(row, 5).HorizontalAlignment = xlCenter
        ws.Cells(row, 7).HorizontalAlignment = xlCenter
        
        row = row + 1
    Next file
    
    For Each subfolder In folder.SubFolders
        Call ListFilesInFolderRecursive(subfolder, ws, row)
    Next subfolder
End Sub


' ========================
' Search Function: AutoFilter by B1
' ========================
Sub SearchFileName()
    Dim ws As Worksheet
    Dim searchText As String
    Dim lastRow As Long
    
    Set ws = Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    searchText = Trim(ws.Range("B1").Value)   ' B1 is Search Box
    
    ' If search box is empty, clear filter
    If searchText = "" Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Apply AutoFilter on Column A (File Name)
    ws.Range("A3:G" & lastRow).AutoFilter Field:=1, Criteria1:="*" & searchText & "*"
End Sub
