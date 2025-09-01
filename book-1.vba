Sub ListFilesInFolderAndSubfolders()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim row As Long
    Dim lastRow As Long
    
    ' ใช้ FileSystemObject สำหรับการจัดการไฟล์และโฟลเดอร์
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
    
    Set ws = Sheets("Sheet1")
    folderPath = "C:\test101\"   ' <<< เปลี่ยนเส้นทางไปยังโฟลเดอร์ที่ต้องการ
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    row = 2
    
    ' ล้างข้อมูลเก่า
    ws.Range("A1:G1000").Clear
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 12
    
    ' สร้างส่วนหัวตาราง
    ws.Cells(1, 1).Value = "File Name"
    ws.Cells(1, 2).Value = "File Extension"
    ws.Cells(1, 3).Value = "File Size (Bytes)"
    ws.Cells(1, 4).Value = "Date Created"
    ws.Cells(1, 5).Value = "Last Modified"
    ws.Cells(1, 6).Value = "Folder Path"
    ws.Cells(1, 7).Value = "Open File"
    
    ' รูปแบบส่วนหัว
    With ws.Range("A1:G1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 112, 192)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' ฟังก์ชันช่วยเพิ่มไฟล์
    Call ListFilesInFolderRecursive(folder, ws, row)
    
    ' ค้นหาแถวสุดท้ายที่มีข้อมูล
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' เรียงข้อมูลตาม Last Modified (คอลัมน์ E) แบบล่าสุดก่อน
    If lastRow > 2 Then
        ws.Range("A2:G" & lastRow).Sort Key1:=ws.Range("E2"), Order1:=xlDescending, Header:=xlNo
    End If
    
    ' ใช้ขอบเฉพาะช่วงที่ใช้เท่านั้น
    If lastRow > 1 Then
        ws.Range("A1:G" & lastRow).Borders.LineStyle = xlContinuous
    End If
    
    ' ตั้งค่าความกว้างของคอลัมน์
    ws.Columns("A").ColumnWidth = 50
    ws.Columns("B").ColumnWidth = 15
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 50
    ws.Columns("G").ColumnWidth = 12
    
    ' ปรับแถวให้พอดีอัตโนมัติ
    ws.Rows.AutoFit
    
    ' แสดงจำนวนไฟล์รวมท้ายตาราง
    ws.Cells(lastRow + 2, 1).Value = "Total Files:"
    ws.Cells(lastRow + 2, 2).Value = row - 2
    
    MsgBox "File list updated successfully!", vbInformation, "Completed"
End Sub

' ฟังก์ชัน recursive สำหรับรวมไฟล์ทุกโฟลเดอร์ย่อย
Sub ListFilesInFolderRecursive(ByVal folder As Object, ByRef ws As Worksheet, ByRef row As Long)
    Dim file As Object
    Dim subfolder As Object
    
    ' รายการไฟล์ในโฟลเดอร์ปัจจุบัน
    For Each file In folder.Files
        ws.Cells(row, 1).Value = file.Name
        ws.Cells(row, 2).Value = Mid(file.Name, InStrRev(file.Name, ".") + 1)
        ws.Cells(row, 3).Value = Format(file.Size, "#,##0") & " Bytes"
        ws.Cells(row, 4).Value = file.DateCreated
        ws.Cells(row, 5).Value = file.DateLastModified
        ws.Cells(row, 6).Value = folder.Path
        
        ' สร้างไฮเปอร์ลิงก์เพื่อเปิดไฟล์
        ws.Hyperlinks.Add Anchor:=ws.Cells(row, 7), _
            Address:=folder.Path & "\" & file.Name, _
            TextToDisplay:="Open"
        
        ' จัดกึ่งกลางสำหรับคอลัมน์ C, D, E, G
        ws.Cells(row, 3).HorizontalAlignment = xlCenter
        ws.Cells(row, 3).VerticalAlignment = xlCenter
        ws.Cells(row, 4).HorizontalAlignment = xlCenter
        ws.Cells(row, 4).VerticalAlignment = xlCenter
        ws.Cells(row, 5).HorizontalAlignment = xlCenter
        ws.Cells(row, 5).VerticalAlignment = xlCenter
        ws.Cells(row, 7).HorizontalAlignment = xlCenter
        ws.Cells(row, 7).VerticalAlignment = xlCenter
        
        ' สีแถวสลับ (สไตล์ม้าลาย)
        If row Mod 2 = 0 Then
            ws.Range("A" & row & ":G" & row).Interior.Color = RGB(235, 241, 222)
        Else
            ws.Range("A" & row & ":G" & row).Interior.Color = RGB(242, 242, 242)
        End If
        
        row = row + 1
    Next file
    
    ' Loop โฟลเดอร์ย่อย
    For Each subfolder In folder.SubFolders
        Call ListFilesInFolderRecursive(subfolder, ws, row)
    Next subfolder
End Sub
