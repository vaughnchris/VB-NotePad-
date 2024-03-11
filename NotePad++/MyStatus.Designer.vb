<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MyStatus
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssFontFace = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssFontSize = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssWordWrap = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssWordCount = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel7 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssSentanceCount = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.tssFontFace, Me.tssFontSize, Me.ToolStripStatusLabel2, Me.tssWordWrap, Me.ToolStripStatusLabel3, Me.tssWordCount, Me.ToolStripStatusLabel7, Me.tssSentanceCount})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(2, 0, 28, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(1006, 42)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(67, 32)
        Me.ToolStripStatusLabel1.Text = "Font:"
        '
        'tssFontFace
        '
        Me.tssFontFace.AutoSize = False
        Me.tssFontFace.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.tssFontFace.IsLink = True
        Me.tssFontFace.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline
        Me.tssFontFace.Name = "tssFontFace"
        Me.tssFontFace.Size = New System.Drawing.Size(150, 32)
        Me.tssFontFace.Text = "Fn: LBS"
        Me.tssFontFace.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.tssFontFace.ToolTipText = "Select Font"
        '
        'tssFontSize
        '
        Me.tssFontSize.AutoSize = False
        Me.tssFontSize.BorderSides = CType(((System.Windows.Forms.ToolStripStatusLabelBorderSides.Top Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.tssFontSize.Margin = New System.Windows.Forms.Padding(-3, 3, 0, 2)
        Me.tssFontSize.Name = "tssFontSize"
        Me.tssFontSize.Size = New System.Drawing.Size(65, 37)
        Me.tssFontSize.ToolTipText = "Font Size"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Margin = New System.Windows.Forms.Padding(50, 3, 0, 2)
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(139, 37)
        Me.ToolStripStatusLabel2.Text = "Word Wrap:"
        '
        'tssWordWrap
        '
        Me.tssWordWrap.AutoSize = False
        Me.tssWordWrap.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.tssWordWrap.Name = "tssWordWrap"
        Me.tssWordWrap.Size = New System.Drawing.Size(45, 32)
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(150, 32)
        Me.ToolStripStatusLabel3.Spring = True
        Me.ToolStripStatusLabel3.Text = "Word Count:"
        Me.ToolStripStatusLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'tssWordCount
        '
        Me.tssWordCount.AutoSize = False
        Me.tssWordCount.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.tssWordCount.Name = "tssWordCount"
        Me.tssWordCount.Size = New System.Drawing.Size(55, 32)
        '
        'ToolStripStatusLabel7
        '
        Me.ToolStripStatusLabel7.Margin = New System.Windows.Forms.Padding(14, 3, 0, 2)
        Me.ToolStripStatusLabel7.Name = "ToolStripStatusLabel7"
        Me.ToolStripStatusLabel7.Size = New System.Drawing.Size(127, 37)
        Me.ToolStripStatusLabel7.Text = "Sentances:"
        '
        'tssSentanceCount
        '
        Me.tssSentanceCount.AutoSize = False
        Me.tssSentanceCount.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
            Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.tssSentanceCount.Name = "tssSentanceCount"
        Me.tssSentanceCount.Size = New System.Drawing.Size(55, 32)
        '
        'MyStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 25.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Controls.Add(Me.StatusStrip1)
        Me.Name = "MyStatus"
        Me.Size = New System.Drawing.Size(1006, 42)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents tssFontSize As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
    Friend WithEvents tssWordWrap As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel3 As ToolStripStatusLabel
    Friend WithEvents tssWordCount As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel7 As ToolStripStatusLabel
    Friend WithEvents tssSentanceCount As ToolStripStatusLabel
    Public WithEvents tssFontFace As ToolStripStatusLabel
End Class
