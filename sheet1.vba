' ฟังชันค้นหาไฟล์ในโฟลเดอร์ที่ระบุ
' กด Alt + F11 เพื่อเปิดหน้าต่าง VBA แล้วกด CTRL + R และวางโค้ดนี้ใน Sheet1(Sheet1)

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("B1")) Is Nothing Then
        Call SearchFileName
    End If
End Sub

