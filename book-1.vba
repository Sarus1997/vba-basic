Sub ListFilesInFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim row As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = Sheets("Sheet1")
    folderPath = "C:\test101\"   ' <<< เปลี่ยนเส้นทางไปยังโฟลเดอร์ที่ต้องการ
    fileName = Dir(folderPath & "*.*")
    row = 2

    ' ล้างข้อมูลเก่า
    ws.Range("A1:E1000").Clear
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 12
    
    ' สร้างส่วนหัวตาราง
    ws.Cells(1, 1).Value = "File Name"
    ws.Cells(1, 2).Value = "File Size (Bytes)"
    ws.Cells(1, 3).Value = "Last Modified"
    ws.Cells(1, 4).Value = "Open File"
    
    ' รูปแบบส่วนหัว
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 112, 192)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' รายการไฟล์
    Do While fileName <> ""
        ws.Cells(row, 1).Value = fileName
        ws.Cells(row, 2).Value = Format(FileLen(folderPath & fileName), "#,##0") & " Bytes"
        ws.Cells(row, 3).Value = FileDateTime(folderPath & fileName)
        
        ' สร้างไฮเปอร์ลิงก์เพื่อเปิดไฟล์
        ws.Hyperlinks.Add Anchor:=ws.Cells(row, 4), _
            Address:=folderPath & fileName, _
            TextToDisplay:="Open"
        
        ' จัดกึ่งกลางสำหรับคอลัมน์ B, C, D
        ws.Cells(row, 2).HorizontalAlignment = xlCenter
        ws.Cells(row, 2).VerticalAlignment = xlCenter
        ws.Cells(row, 3).HorizontalAlignment = xlCenter
        ws.Cells(row, 3).VerticalAlignment = xlCenter
        ws.Cells(row, 4).HorizontalAlignment = xlCenter
        ws.Cells(row, 4).VerticalAlignment = xlCenter
        
        ' สีแถวสลับ (สไตล์ม้าลาย)
        If row Mod 2 = 0 Then
            ws.Range("A" & row & ":D" & row).Interior.Color = RGB(235, 241, 222)
        Else
            ws.Range("A" & row & ":D" & row).Interior.Color = RGB(242, 242, 242)
        End If
        
        row = row + 1
        fileName = Dir
    Loop
    
    ' ค้นหาแถวสุดท้ายที่มีข้อมูล
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' ใช้ขอบเฉพาะช่วงที่ใช้เท่านั้น
    If lastRow > 1 Then
        ws.Range("A1:D" & lastRow).Borders.LineStyle = xlContinuous
    End If
    
    ' ตั้งค่าความกว้างของคอลัมน์
    ws.Columns("A").ColumnWidth = 50
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("D").ColumnWidth = 12
    
    ' Autofit rows
    ws.Rows.AutoFit
    
    MsgBox "? File list updated successfully!", vbInformation, "Completed"
End Sub


