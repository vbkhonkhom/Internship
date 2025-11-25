Public Class Form3

    ' เมื่อฟอร์มถูกปิด (User กดปิดเอง หรือโปรแกรมสั่งปิด)
    Private Sub Form3_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    End Sub

    ' ก่อนฟอร์มจะปิด
    Private Sub Form3_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosing
        ' StopFlag: ตัวแปร Boolean ใน Form1 เพื่อบอกลูปทำงานหลักว่า "ให้หยุดทำงานได้แล้ว"
        Form1.StopFlag = True 
        
        ' สั่งให้ Form1 (หน้าจอหลัก) กลับมาแสดงอีกครั้ง
        Form1.Show()
    End Sub

    Private Sub Form3_load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub

    ' [ปุ่ม Stop] ปุ่มสำหรับสั่งหยุดการทำงาน
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' สั่ง ReportProgress ไปที่ BackgroundWorker เพื่อแจ้งสถานะว่า "กำลังหยุด" (Stop_Convert)
        Form1.BackgroundWorker1.ReportProgress(Form1.Stop_Convert)
        
        ' สั่ง CancelAsync เพื่อยกเลิกเธรดที่ทำงานอยู่เบื้องหลัง
        Form1.BackgroundWorker1.CancelAsync()
    End Sub
End Class