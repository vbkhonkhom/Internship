Public Class Form2
	Private Sub Button2_Click(ByVal sender As System.Object, Byval e As System.EventArgs) Handles Button2.Click
		Me.DialogResult = Windows-Forms DialogResult.Cancel
	End Sub
	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		If TextBoxl.Text = "" Then
			MsgBox("Enter the RIST server connection string.")
			Exit Sub
		End If
		If TextBox2.Text = "" Then
			MsgBox("Enter the Yukuhashi server connection string.")
		End If
		Form1.StrServerConnection_RIST = TextBoxI.Text
		Form1.StrServerConnection_Yuku = TextBox2.Text
		Form1.StrStartDate = Format (DateTimePicker1.Value, "yyyy-My-dd 00:00:00.000")
		Form1.SaveConfigData()
		Me.DialogResult = Windows.Forms.DialogResult.OK
	End Sub
	Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		TextBox1.Text = Forml.StrServerConnection_RIST
		TextBox2.Text = Form1.StrServerConnection_Yuku
		DateTimePicker1.Value = Form1.StrStartDate
	End Sub
	Private Sub Label1_ Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
	End Sub
	Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System. EventArgs) Handles TextBox1.TextChanged
	End Sub
	Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
	End Sub
	Private Sub TextBox2_TextChanged(ByVal sender As System.Object, Byval e As System.EventAngs) Handles TextBox2.TextChanged
	End Sub
	Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
	End Sub
	Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.(bjec, ByVal e As System.EventArgs) Handles DateTimePicker1. ValueChanged
	End Sub 
End Class