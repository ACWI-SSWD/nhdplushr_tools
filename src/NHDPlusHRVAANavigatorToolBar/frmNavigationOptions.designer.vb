<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmNavigationOptions
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents btnNavigate As System.Windows.Forms.Button
    Public WithEvents txtMaxDistance As System.Windows.Forms.TextBox
    Public WithEvents lblDistanceUnits As System.Windows.Forms.Label
    Public WithEvents lblNavigationStopMessage As System.Windows.Forms.Label
    Public WithEvents lblNavigationStopDistance As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNavigationOptions))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.tbStartMeasure = New System.Windows.Forms.TextBox()
        Me.radWholeReach = New System.Windows.Forms.RadioButton()
        Me.btnNavigate = New System.Windows.Forms.Button()
        Me.txtMaxDistance = New System.Windows.Forms.TextBox()
        Me.lblDistanceUnits = New System.Windows.Forms.Label()
        Me.lblNavigationStopMessage = New System.Windows.Forms.Label()
        Me.lblNavigationStopDistance = New System.Windows.Forms.Label()
        Me.gbNavigationStopDistance = New System.Windows.Forms.GroupBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.gbNavigationStart = New System.Windows.Forms.GroupBox()
        Me.radStartMeasure = New System.Windows.Forms.RadioButton()
        Me.gbNavigationStopAttributes = New System.Windows.Forms.GroupBox()
        Me.lblStopAttributesMessage = New System.Windows.Forms.Label()
        Me.tbAttrValue = New System.Windows.Forms.TextBox()
        Me.cbOperator = New System.Windows.Forms.ComboBox()
        Me.cbAttrName = New System.Windows.Forms.ComboBox()
        Me.lblAttrComparisonValue = New System.Windows.Forms.Label()
        Me.lblAttrOperator = New System.Windows.Forms.Label()
        Me.lblAttrName = New System.Windows.Forms.Label()
        Me.gbNavigationInformation = New System.Windows.Forms.GroupBox()
        Me.btnNavDBPathBrowse = New System.Windows.Forms.Button()
        Me.lblNavDBPath = New System.Windows.Forms.Label()
        Me.txtNavDBPath = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.gbNavigationStopDistance.SuspendLayout()
        Me.gbNavigationStart.SuspendLayout()
        Me.gbNavigationStopAttributes.SuspendLayout()
        Me.gbNavigationInformation.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbStartMeasure
        '
        Me.tbStartMeasure.Location = New System.Drawing.Point(214, 52)
        Me.tbStartMeasure.Name = "tbStartMeasure"
        Me.tbStartMeasure.Size = New System.Drawing.Size(123, 20)
        Me.tbStartMeasure.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.tbStartMeasure, "Starting measure ")
        '
        'radWholeReach
        '
        Me.radWholeReach.AutoSize = True
        Me.radWholeReach.Location = New System.Drawing.Point(26, 29)
        Me.radWholeReach.Name = "radWholeReach"
        Me.radWholeReach.Size = New System.Drawing.Size(279, 18)
        Me.radWholeReach.TabIndex = 2
        Me.radWholeReach.TabStop = True
        Me.radWholeReach.Text = "Start at top or bottom of ""clicked"" NHDFlowline"
        Me.radWholeReach.UseVisualStyleBackColor = True
        '
        'btnNavigate
        '
        Me.btnNavigate.BackColor = System.Drawing.SystemColors.Control
        Me.btnNavigate.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnNavigate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNavigate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnNavigate.Location = New System.Drawing.Point(205, 499)
        Me.btnNavigate.Name = "btnNavigate"
        Me.btnNavigate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnNavigate.Size = New System.Drawing.Size(110, 49)
        Me.btnNavigate.TabIndex = 9
        Me.btnNavigate.Text = "Navigate"
        Me.btnNavigate.UseVisualStyleBackColor = False
        '
        'txtMaxDistance
        '
        Me.txtMaxDistance.AcceptsReturn = True
        Me.txtMaxDistance.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaxDistance.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaxDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxDistance.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaxDistance.Location = New System.Drawing.Point(195, 16)
        Me.txtMaxDistance.MaxLength = 0
        Me.txtMaxDistance.Name = "txtMaxDistance"
        Me.txtMaxDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaxDistance.Size = New System.Drawing.Size(73, 20)
        Me.txtMaxDistance.TabIndex = 5
        Me.txtMaxDistance.Text = "0"
        '
        'lblDistanceUnits
        '
        Me.lblDistanceUnits.BackColor = System.Drawing.SystemColors.Control
        Me.lblDistanceUnits.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDistanceUnits.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDistanceUnits.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDistanceUnits.Location = New System.Drawing.Point(275, 18)
        Me.lblDistanceUnits.Name = "lblDistanceUnits"
        Me.lblDistanceUnits.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDistanceUnits.Size = New System.Drawing.Size(49, 17)
        Me.lblDistanceUnits.TabIndex = 4
        Me.lblDistanceUnits.Text = "KM"
        '
        'lblNavigationStopMessage
        '
        Me.lblNavigationStopMessage.BackColor = System.Drawing.SystemColors.Control
        Me.lblNavigationStopMessage.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNavigationStopMessage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNavigationStopMessage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNavigationStopMessage.Location = New System.Drawing.Point(11, 44)
        Me.lblNavigationStopMessage.Name = "lblNavigationStopMessage"
        Me.lblNavigationStopMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNavigationStopMessage.Size = New System.Drawing.Size(414, 35)
        Me.lblNavigationStopMessage.TabIndex = 3
        Me.lblNavigationStopMessage.Text = "If a non-zero value is provided, navigation will stop when it reaches that distan" & _
    "ce from the navigation starting point..  "
        '
        'lblNavigationStopDistance
        '
        Me.lblNavigationStopDistance.BackColor = System.Drawing.SystemColors.Control
        Me.lblNavigationStopDistance.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblNavigationStopDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNavigationStopDistance.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblNavigationStopDistance.Location = New System.Drawing.Point(11, 16)
        Me.lblNavigationStopDistance.Name = "lblNavigationStopDistance"
        Me.lblNavigationStopDistance.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNavigationStopDistance.Size = New System.Drawing.Size(169, 25)
        Me.lblNavigationStopDistance.TabIndex = 2
        Me.lblNavigationStopDistance.Text = "Navigation Stop Distance:"
        '
        'gbNavigationStopDistance
        '
        Me.gbNavigationStopDistance.Controls.Add(Me.txtMaxDistance)
        Me.gbNavigationStopDistance.Controls.Add(Me.lblDistanceUnits)
        Me.gbNavigationStopDistance.Controls.Add(Me.lblNavigationStopMessage)
        Me.gbNavigationStopDistance.Controls.Add(Me.lblNavigationStopDistance)
        Me.gbNavigationStopDistance.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbNavigationStopDistance.Location = New System.Drawing.Point(13, 182)
        Me.gbNavigationStopDistance.Name = "gbNavigationStopDistance"
        Me.gbNavigationStopDistance.Size = New System.Drawing.Size(430, 86)
        Me.gbNavigationStopDistance.TabIndex = 5
        Me.gbNavigationStopDistance.TabStop = False
        Me.gbNavigationStopDistance.Text = "Stop Navigation based on Distance Traveled"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(335, 500)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 48)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'gbNavigationStart
        '
        Me.gbNavigationStart.Controls.Add(Me.tbStartMeasure)
        Me.gbNavigationStart.Controls.Add(Me.radStartMeasure)
        Me.gbNavigationStart.Controls.Add(Me.radWholeReach)
        Me.gbNavigationStart.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbNavigationStart.Location = New System.Drawing.Point(14, 87)
        Me.gbNavigationStart.Name = "gbNavigationStart"
        Me.gbNavigationStart.Size = New System.Drawing.Size(429, 89)
        Me.gbNavigationStart.TabIndex = 7
        Me.gbNavigationStart.TabStop = False
        Me.gbNavigationStart.Text = "Navigation Start Options"
        '
        'radStartMeasure
        '
        Me.radStartMeasure.AutoSize = True
        Me.radStartMeasure.Location = New System.Drawing.Point(26, 52)
        Me.radStartMeasure.Name = "radStartMeasure"
        Me.radStartMeasure.Size = New System.Drawing.Size(183, 18)
        Me.radStartMeasure.TabIndex = 3
        Me.radStartMeasure.TabStop = True
        Me.radStartMeasure.Text = "Start at Reachcode measure:"
        Me.radStartMeasure.UseVisualStyleBackColor = True
        '
        'gbNavigationStopAttributes
        '
        Me.gbNavigationStopAttributes.Controls.Add(Me.lblStopAttributesMessage)
        Me.gbNavigationStopAttributes.Controls.Add(Me.tbAttrValue)
        Me.gbNavigationStopAttributes.Controls.Add(Me.cbOperator)
        Me.gbNavigationStopAttributes.Controls.Add(Me.cbAttrName)
        Me.gbNavigationStopAttributes.Controls.Add(Me.lblAttrComparisonValue)
        Me.gbNavigationStopAttributes.Controls.Add(Me.lblAttrOperator)
        Me.gbNavigationStopAttributes.Controls.Add(Me.lblAttrName)
        Me.gbNavigationStopAttributes.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbNavigationStopAttributes.Location = New System.Drawing.Point(13, 274)
        Me.gbNavigationStopAttributes.Name = "gbNavigationStopAttributes"
        Me.gbNavigationStopAttributes.Size = New System.Drawing.Size(430, 206)
        Me.gbNavigationStopAttributes.TabIndex = 8
        Me.gbNavigationStopAttributes.TabStop = False
        Me.gbNavigationStopAttributes.Text = "Filter Navigation Results"
        '
        'lblStopAttributesMessage
        '
        Me.lblStopAttributesMessage.AutoSize = True
        Me.lblStopAttributesMessage.Location = New System.Drawing.Point(17, 27)
        Me.lblStopAttributesMessage.Name = "lblStopAttributesMessage"
        Me.lblStopAttributesMessage.Size = New System.Drawing.Size(0, 14)
        Me.lblStopAttributesMessage.TabIndex = 6
        '
        'tbAttrValue
        '
        Me.tbAttrValue.Location = New System.Drawing.Point(204, 171)
        Me.tbAttrValue.Name = "tbAttrValue"
        Me.tbAttrValue.Size = New System.Drawing.Size(159, 20)
        Me.tbAttrValue.TabIndex = 8

        '  lblStopAttributesMessage
        Me.lblStopAttributesMessage.Text = "Navigation starts based on the mouse click and " & vbCrLf & "start options selected above.  Only NHDFlowline features" & vbCrLf & "that satisfy the filter condition specified below" & vbCrLf & "will be included in the navigation results."
        '
        'cbOperator
        '
        Me.cbOperator.FormattingEnabled = True
        Me.cbOperator.Items.AddRange(New Object() {"<", "<=", ">=", ">"})
        Me.cbOperator.Location = New System.Drawing.Point(145, 134)
        Me.cbOperator.Name = "cbOperator"
        Me.cbOperator.Size = New System.Drawing.Size(112, 22)
        Me.cbOperator.TabIndex = 7
        '
        'cbAttrName
        '
        Me.cbAttrName.FormattingEnabled = True
        Me.cbAttrName.Items.AddRange(New Object() {"PATHLENGTH", "ARBOLATESU", "TOTDASQKM", "DIVDASQKM"})
        Me.cbAttrName.Location = New System.Drawing.Point(145, 101)
        Me.cbAttrName.Name = "cbAttrName"
        Me.cbAttrName.Size = New System.Drawing.Size(112, 22)
        Me.cbAttrName.TabIndex = 6
        '
        'lblAttrComparisonValue
        '
        Me.lblAttrComparisonValue.AutoSize = True
        Me.lblAttrComparisonValue.Location = New System.Drawing.Point(15, 174)
        Me.lblAttrComparisonValue.Name = "lblAttrComparisonValue"
        Me.lblAttrComparisonValue.Size = New System.Drawing.Size(163, 14)
        Me.lblAttrComparisonValue.TabIndex = 2
        Me.lblAttrComparisonValue.Text = "Attribute Comparison Value:"
        '
        'lblAttrOperator
        '
        Me.lblAttrOperator.AutoSize = True
        Me.lblAttrOperator.Location = New System.Drawing.Point(15, 137)
        Me.lblAttrOperator.Name = "lblAttrOperator"
        Me.lblAttrOperator.Size = New System.Drawing.Size(59, 14)
        Me.lblAttrOperator.TabIndex = 1
        Me.lblAttrOperator.Text = "Operator:"
        '
        'lblAttrName
        '
        Me.lblAttrName.AutoSize = True
        Me.lblAttrName.Location = New System.Drawing.Point(15, 104)
        Me.lblAttrName.Name = "lblAttrName"
        Me.lblAttrName.Size = New System.Drawing.Size(93, 14)
        Me.lblAttrName.TabIndex = 0
        Me.lblAttrName.Text = "Attribute Name:"
        '
        'gbNavigationInformation
        '
        Me.gbNavigationInformation.Controls.Add(Me.btnNavDBPathBrowse)
        Me.gbNavigationInformation.Controls.Add(Me.lblNavDBPath)
        Me.gbNavigationInformation.Controls.Add(Me.txtNavDBPath)
        Me.gbNavigationInformation.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbNavigationInformation.Location = New System.Drawing.Point(13, 12)
        Me.gbNavigationInformation.Name = "gbNavigationInformation"
        Me.gbNavigationInformation.Size = New System.Drawing.Size(429, 69)
        Me.gbNavigationInformation.TabIndex = 9
        Me.gbNavigationInformation.TabStop = False
        Me.gbNavigationInformation.Text = "Navigation Information"
        '
        'btnNavDBPathBrowse
        '
        Me.btnNavDBPathBrowse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNavDBPathBrowse.Image = CType(resources.GetObject("btnNavDBPathBrowse.Image"), System.Drawing.Image)
        Me.btnNavDBPathBrowse.Location = New System.Drawing.Point(367, 8)
        Me.btnNavDBPathBrowse.Name = "btnNavDBPathBrowse"
        Me.btnNavDBPathBrowse.Size = New System.Drawing.Size(56, 50)
        Me.btnNavDBPathBrowse.TabIndex = 1
        Me.btnNavDBPathBrowse.Text = "Browse"
        Me.btnNavDBPathBrowse.UseVisualStyleBackColor = True
        '
        'lblNavDBPath
        '
        Me.lblNavDBPath.AutoSize = True
        Me.lblNavDBPath.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNavDBPath.Location = New System.Drawing.Point(14, 26)
        Me.lblNavDBPath.Name = "lblNavDBPath"
        Me.lblNavDBPath.Size = New System.Drawing.Size(129, 14)
        Me.lblNavDBPath.TabIndex = 1
        Me.lblNavDBPath.Text = "Navigator Database Path:"
        '
        'txtNavDBPath
        '
        Me.txtNavDBPath.Location = New System.Drawing.Point(148, 20)
        Me.txtNavDBPath.Name = "txtNavDBPath"
        Me.txtNavDBPath.Size = New System.Drawing.Size(215, 20)
        Me.txtNavDBPath.TabIndex = 0
        '
        'frmNavigationOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(455, 567)
        Me.Controls.Add(Me.gbNavigationInformation)
        Me.Controls.Add(Me.gbNavigationStopAttributes)
        Me.Controls.Add(Me.gbNavigationStart)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.gbNavigationStopDistance)
        Me.Controls.Add(Me.btnNavigate)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 30)
        Me.Name = "frmNavigationOptions"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "NHDPlusHR VAA Navigation Options"
        Me.gbNavigationStopDistance.ResumeLayout(False)
        Me.gbNavigationStopDistance.PerformLayout()
        Me.gbNavigationStart.ResumeLayout(False)
        Me.gbNavigationStart.PerformLayout()
        Me.gbNavigationStopAttributes.ResumeLayout(False)
        Me.gbNavigationStopAttributes.PerformLayout()
        Me.gbNavigationInformation.ResumeLayout(False)
        Me.gbNavigationInformation.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbNavigationStopDistance As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents gbNavigationStart As System.Windows.Forms.GroupBox
    Friend WithEvents radWholeReach As System.Windows.Forms.RadioButton
    Friend WithEvents tbStartMeasure As System.Windows.Forms.TextBox
    Friend WithEvents radStartMeasure As System.Windows.Forms.RadioButton
    Friend WithEvents gbNavigationStopAttributes As System.Windows.Forms.GroupBox
    Friend WithEvents tbAttrValue As System.Windows.Forms.TextBox
    Friend WithEvents cbOperator As System.Windows.Forms.ComboBox
    Friend WithEvents cbAttrName As System.Windows.Forms.ComboBox
    Friend WithEvents lblAttrComparisonValue As System.Windows.Forms.Label
    Friend WithEvents lblAttrOperator As System.Windows.Forms.Label
    Friend WithEvents lblAttrName As System.Windows.Forms.Label
    'Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents lblStopAttributesMessage As System.Windows.Forms.Label
    Friend WithEvents gbNavigationInformation As System.Windows.Forms.GroupBox
    Friend WithEvents btnNavDBPathBrowse As System.Windows.Forms.Button
    Friend WithEvents lblNavDBPath As System.Windows.Forms.Label
    Friend WithEvents txtNavDBPath As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
#End Region 
End Class