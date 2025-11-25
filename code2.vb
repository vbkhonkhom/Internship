Public Class Form2

    ' [ปุ่ม Cancel] ยกเลิกการตั้งค่า
    Private Sub Button2_Click(ByVal sender As System.Object, Byval e As System.EventArgs) Handles Button2.Click
        ' DialogResult: ตัวแปรระบบที่บอกว่าฟอร์มนี้ปิดลงด้วยสถานะอะไร (ในที่นี้คือ Cancel)
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

    ' [ปุ่ม OK] ยืนยันการบันทึกค่า
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' TextBox1.Text: รับค่า Connection String ของ Server RIST
        If TextBox1.Text = "" Then
            MsgBox("Enter the RIST server connection string.")
            Exit Sub
        End If
        
        ' TextBox2.Text: รับค่า Connection String ของ Server Yukuhashi
        If TextBox2.Text = "" Then
            MsgBox("Enter the Yukuhashi server connection string.")
        End If

        ' ส่งค่ากลับไปเก็บที่ตัวแปร Global ใน Form1
        Form1.StrServerConnection_RIST = TextBox1.Text ' เก็บ IP/Config ของ RIST
        Form1.StrServerConnection_Yuku = TextBox2.Text ' เก็บ IP/Config ของ Yukuhashi
        
        ' DateTimePicker1.Value: เก็บวันที่เริ่มต้น (Start Date) ที่ User เลือก
        ' Format(): แปลงวันที่เป็น String รูปแบบมาตรฐาน Database (yyyy-MM-dd ...)
        Form1.StrStartDate = Format(DateTimePicker1.Value, "yyyy-MM-dd 00:00:00.000")
        
        ' เรียกฟังก์ชันบันทึกลงไฟล์ CSV
        Form1.SaveConfigData()
        
        ' แจ้งว่าจบการทำงานด้วยสถานะ OK
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    ' เมื่อโหลดฟอร์ม ให้ดึงค่าเดิมมาแสดง
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' ดึงค่าจากตัวแปรใน Form1 มาใส่ในกล่องข้อความเพื่อให้ User แก้ไข
        TextBox1.Text = Form1.StrServerConnection_RIST
        TextBox2.Text = Form1.StrServerConnection_Yuku
        DateTimePicker1.Value = Form1.StrStartDate
    End Sub

    ' Event Handler อื่นๆ ที่ไม่ได้ใช้งาน (ว่างไว้)
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
    End Sub
    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
    End Sub
    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, Byval e As System.EventArgs) Handles TextBox2.TextChanged
    End Sub
    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
    End Sub
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
    End Sub
End Class