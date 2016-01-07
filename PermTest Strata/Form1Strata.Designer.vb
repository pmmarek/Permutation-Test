<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPermTest
   Inherits System.Windows.Forms.Form

   'Form overrides dispose to clean up the component list.
   <System.Diagnostics.DebuggerNonUserCode()> _
   Protected Overrides Sub Dispose(ByVal disposing As Boolean)
      Try
         If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
         End If
      Finally
         MyBase.Dispose(disposing)
      End Try
   End Sub

   'Required by the Windows Form Designer
   Private components As System.ComponentModel.IContainer

   'NOTE: The following procedure is required by the Windows Form Designer
   'It can be modified using the Windows Form Designer.  
   'Do not modify it using the code editor.
   <System.Diagnostics.DebuggerStepThrough()> _
   Private Sub InitializeComponent()
      Me.dlgOpen = New System.Windows.Forms.OpenFileDialog()
      Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
      Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem()
      Me.mnuFileOpen = New System.Windows.Forms.ToolStripMenuItem()
      Me.mnuFileOpenData = New System.Windows.Forms.ToolStripMenuItem()
      Me.mnuFileOpenResults = New System.Windows.Forms.ToolStripMenuItem()
      Me.mnuFilePrint = New System.Windows.Forms.ToolStripMenuItem()
      Me.mnuRun = New System.Windows.Forms.ToolStripMenuItem()
      Me.mnuExit = New System.Windows.Forms.ToolStripMenuItem()
      Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem()
      Me.txtProgress = New System.Windows.Forms.TextBox()
      Me.MenuStrip1.SuspendLayout()
      Me.SuspendLayout()
      '
      'MenuStrip1
      '
      Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile, Me.mnuRun, Me.mnuExit, Me.mnuAbout})
      Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
      Me.MenuStrip1.Name = "MenuStrip1"
      Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(4, 2, 0, 2)
      Me.MenuStrip1.Size = New System.Drawing.Size(470, 24)
      Me.MenuStrip1.TabIndex = 0
      Me.MenuStrip1.Text = "MenuStrip1"
      '
      'mnuFile
      '
      Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileOpen, Me.mnuFilePrint})
      Me.mnuFile.Name = "mnuFile"
      Me.mnuFile.Size = New System.Drawing.Size(37, 20)
      Me.mnuFile.Text = "&File"
      '
      'mnuFileOpen
      '
      Me.mnuFileOpen.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileOpenData, Me.mnuFileOpenResults})
      Me.mnuFileOpen.Name = "mnuFileOpen"
      Me.mnuFileOpen.Size = New System.Drawing.Size(103, 22)
      Me.mnuFileOpen.Text = "&Open"
      '
      'mnuFileOpenData
      '
      Me.mnuFileOpenData.Name = "mnuFileOpenData"
      Me.mnuFileOpenData.Size = New System.Drawing.Size(111, 22)
      Me.mnuFileOpenData.Text = "&Data"
      '
      'mnuFileOpenResults
      '
      Me.mnuFileOpenResults.Name = "mnuFileOpenResults"
      Me.mnuFileOpenResults.Size = New System.Drawing.Size(111, 22)
      Me.mnuFileOpenResults.Text = "R&esults"
      '
      'mnuFilePrint
      '
      Me.mnuFilePrint.Name = "mnuFilePrint"
      Me.mnuFilePrint.Size = New System.Drawing.Size(103, 22)
      Me.mnuFilePrint.Text = "&Print"
      '
      'mnuRun
      '
      Me.mnuRun.Name = "mnuRun"
      Me.mnuRun.Size = New System.Drawing.Size(40, 20)
      Me.mnuRun.Text = "&Run"
      '
      'mnuExit
      '
      Me.mnuExit.Name = "mnuExit"
      Me.mnuExit.Size = New System.Drawing.Size(37, 20)
      Me.mnuExit.Text = "Exit"
      '
      'mnuAbout
      '
      Me.mnuAbout.Name = "mnuAbout"
      Me.mnuAbout.Size = New System.Drawing.Size(52, 20)
      Me.mnuAbout.Text = "About"
      '
      'txtProgress
      '
      Me.txtProgress.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.txtProgress.Font = New System.Drawing.Font("Arial", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txtProgress.Location = New System.Drawing.Point(31, 32)
      Me.txtProgress.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
      Me.txtProgress.Multiline = True
      Me.txtProgress.Name = "txtProgress"
      Me.txtProgress.ReadOnly = True
      Me.txtProgress.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
      Me.txtProgress.Size = New System.Drawing.Size(398, 156)
      Me.txtProgress.TabIndex = 1
      '
      'frmPermTest
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(470, 208)
      Me.Controls.Add(Me.txtProgress)
      Me.Controls.Add(Me.MenuStrip1)
      Me.MainMenuStrip = Me.MenuStrip1
      Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
      Me.Name = "frmPermTest"
      Me.Text = "HSYAS Permutation Test with Strata"
      Me.MenuStrip1.ResumeLayout(False)
      Me.MenuStrip1.PerformLayout()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents dlgOpen As System.Windows.Forms.OpenFileDialog
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents mnuFileOpen As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents mnuFilePrint As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents mnuFileOpenData As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents mnuFileOpenResults As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents mnuRun As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents txtProgress As System.Windows.Forms.TextBox
   Friend WithEvents mnuExit As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem

End Class
