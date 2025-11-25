Public Class Form3
	Private Sub Form3_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs)Handles Me.FormClosed
	End Sub 
	Private Sub Form3_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEvevtArgs) Handles Me.FormClosing
		Form1.StopFlag = True 
		Form1.Show()
	End Sub 
	Private Sub Form3_load(ByVal sender As System.Object,ByVal e As System.EventArgs) Handles MyBase.Load
	End Sub 
	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		Form1.BackgroundWorker1.ReportProgress(Form1.Stop_Coonvert)
		Form1.BackgroundWorker1.CancelAsync()
	End Sub 
End Class 
