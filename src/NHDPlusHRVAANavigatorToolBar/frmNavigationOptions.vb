Option Strict Off
Option Explicit On

Friend Class frmNavigationOptions
    Inherits System.Windows.Forms.Form

    Private gboolCancel As Boolean
    Private gboolNavigatePressed As Boolean

    Private Sub btnOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnNavigate.Click
        Dim strMessage As String = ""
        Dim dblDistance As Double
        Dim dblStartMeasure As Double
        Dim strAttrName As String
        Dim strAttrCompare As String
        Dim strAttrValue As String

        'VALIDATE EVERYTHING FROM NAVIGATION OPTIONS FORM
        If Me.txtNavDBPath.Text.Trim = "" Or Me.txtNavDBPath.Text.Trim.Contains(" ") Or Not System.IO.Directory.Exists(Me.txtNavDBPath.Text.Trim) Then
            If strMessage = "" Then
                strMessage = "Navigator Database Path must not be empty, not contain spaces, and be an existing local folder."
            Else
                strMessage = strMessage + vbCrLf + "Navigator Database Path must contain a value, not contain spaces, and be an existing local folder."
            End If
        End If

        dblDistance = Val(Me.txtMaxDistance.Text)

        If dblDistance < 0 Then
            If strMessage = "" Then
                strMessage = "Stop Distance must be >= 0."
            Else
                strMessage = strMessage + vbCrLf + "Stop Distance must be >= 0."
            End If
        End If
        If dblStartMeasure < 0 Or dblStartMeasure > 100 Then
            If strMessage = "" Then
                strMessage = "Start Measure must be between 0 and 100, inclusive."
            Else
                strMessage = strMessage + vbCrLf + "Start Measure must be between 0 and 100, inclusive."
            End If
        End If

        strAttrName = Me.cbAttrName.Text.Trim
        strAttrCompare = Me.cbOperator.Text.Trim
        strAttrValue = Me.tbAttrValue.Text.Trim

        If Not strAttrName = "" Or
           Not strAttrCompare = "" Or
           Not strAttrValue = "" Then
            If strAttrName = "" Or strAttrValue = "" Or strAttrCompare = "" Then
                If strMessage = "" Then
                    strMessage = "If an attribute name, operator, or comparison value is selected, ALL three pieces must be provided."
                Else
                    strMessage = strMessage + vbCrLf + "If an attribute name, operator, or comparison value is selected, ALL three pieces must be provided."
                End If
            End If
        End If

        If strAttrName > "" Then
            If Not Me.cbAttrName.Items.Contains(strAttrName) Then
                If strMessage = "" Then
                    strMessage = "Attribute Name must be a PlusFlowlineVAA field contained in the drop-down list on the Navigation Options form."
                Else
                    strMessage = strMessage + vbCrLf + "Attribute Name must be a PlusFlowlineVAA field contained in the drop-down list on the Navigation Options form."
                End If
            End If
        End If

        If strAttrCompare > "" Then
            If Not Me.cbOperator.Items.Contains(strAttrCompare) Then
                If strMessage = "" Then
                    strMessage = "Operator be a valid comparison operator contained in the drop-down list on the Navigation Options form."
                Else
                    strMessage = strMessage + vbCrLf + "Operator be a valid comparison operator contained in the drop-down list on the Navigation Options form."
                End If
            End If
        End If

        If strAttrValue > "" Then
            If Not IsNumeric(strAttrValue) Then
                If strMessage = "" Then
                    strMessage = "Attribute Comparison Value must be numeric."
                Else
                    strMessage = strMessage + vbCrLf + "Attribute Comparison Value must be numeric."
                End If
            End If
        End If

        If gnumFcodeValue = 56600 And (gstrNavigationType = "UPTRIB" Or gstrNavigationType = "DNDIV") Then
            'Coastline start - Do not allow uptrib navigations
            If strMessage = "" Then
                strMessage = "UP TRIB and DN Div navigations are not supported from Coastlines."
            Else
                strMessage = strMessage + vbCrLf + "UP TRIB and Dn Div navigations are not supported from Coastlines."
            End If
        End If

        If gnumFcodeValue = 56600 Then
            'Coastline start
            If dblDistance > 0 Or strAttrName > "" Then
                'boolInvalidStop = True
                If strMessage = "" Then
                    strMessage = "Coastline navigations with stop conditions or filters are not supported."
                Else
                    strMessage = strMessage + vbCrLf + "Coastline navigations with stop conditions or filters are not supported."
                End If
            End If
        End If


        If strMessage <> "" Then
            MsgBox(strMessage)
        Else
            Me.Hide()
        End If
        strMessage = Nothing

    End Sub

    Private Sub txtMaxDistance_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtMaxDistance.Validating
        Dim strMessage As String = ""

        If Not IsNumeric(txtMaxDistance.Text) Or Val(txtMaxDistance.Text) < 0 Then
            strMessage = "Navigation stop distance must be a numeric value greater than or equal to 0"
        End If
        If strMessage <> "" Then
            MsgBox(strMessage)
        End If
        strMessage = Nothing
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DoCancel = True
        Me.Hide()
    End Sub

    Public Property DoCancel() As Boolean
        Get
            Return gboolCancel
        End Get
        Set(ByVal Value As Boolean)
            gboolCancel = Value
        End Set
    End Property

    Public Property NavigatePressed() As Boolean
        Get
            Return gboolNavigatePressed
        End Get
        Set(ByVal Value As Boolean)
            gboolNavigatePressed = Value
        End Set
    End Property

    Private Sub radWholeReach_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radWholeReach.CheckedChanged
        If radWholeReach.Checked Then
            tbStartMeasure.Enabled = False
        End If
    End Sub

    Private Sub radStartMeasure_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radStartMeasure.CheckedChanged
        If radStartMeasure.Checked Then
            tbStartMeasure.Enabled = True
        End If
    End Sub

    Private Sub tbStartMeasure_Validating(ByVal eventsender As Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles tbStartMeasure.Validating

        Dim strMessage As String = ""
        ' If the value is a number larger than 10, keep the focus.

        If Not IsNumeric(tbStartMeasure.Text) Or Val(tbStartMeasure.Text) < 0 Or Val(tbStartMeasure.Text) > 100 Then

            strMessage = "Start Measure must be a numeric value greater than or equal to 0 and less than or equal to 100."
        End If
        If strMessage <> "" Then
            MsgBox(strMessage)
        End If
        strMessage = Nothing
    End Sub

    Private Sub btnNavDBPathBrowse_Click(sender As Object, e As System.EventArgs) Handles btnNavDBPathBrowse.Click
        FolderBrowserDialog1.Description = "Select the Navigator Database Path"
        FolderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer
        FolderBrowserDialog1.ShowNewFolderButton = False
        If FolderBrowserDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtNavDBPath.Text = FolderBrowserDialog1.SelectedPath
            If (InStr(txtNavDBPath.Text, " ") > 0) Then
                MsgBox("The Navigator Database Path contains at least one space. Please rename it so there are no spaces.")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub txtNavDBPath_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtNavDBPath.Validating

        Dim strErrMsg As String = ""
        'Dim intRet As Integer
        If txtNavDBPath.Text.Trim.Length > 0 And My.Computer.FileSystem.DirectoryExists(txtNavDBPath.Text) = False Then
            strErrMsg = "The Navigator Database Path (" & txtNavDBPath.Text & ") does not exist."
        End If
        If (InStr(txtNavDBPath.Text, " ") > 0) Then
            If strErrMsg = "" Then
                strErrMsg = "The Navigator Database Path contains at least one space. Please rename it so there are no spaces."
            Else
                strErrMsg = strErrMsg & vbCrLf & "The Navigator Database Path contains at least one space. Please rename it so there are no spaces."
            End If
            If Len(Trim(txtNavDBPath.Text)) = 0 Then
                strErrMsg = "Navigator Database Path must not be blank."
            End If
            If strErrMsg > "" Then
                MsgBox(strErrMsg)
            End If
        End If

        If strErrMsg <> "" Then
            MsgBox(strErrMsg)
        End If

        strErrMsg = Nothing
    End Sub

    Private Sub cbAttrName_Leave(sender As System.Object, e As System.EventArgs) Handles cbAttrName.Leave
        Dim strMessage As String = ""
        If cbAttrName.Text <> "" Then
            If Not cbAttrName.Items.Contains(cbAttrName.Text.ToUpper) Then
                strMessage = "Attribute Name must be a field contained in the drop-down list."
            End If

        End If

        If strMessage <> "" Then
            MsgBox(strMessage)
        End If
        strMessage = Nothing

    End Sub

    Private Sub cbOperator_Leave(sender As System.Object, e As System.EventArgs) Handles cbOperator.Leave
        Dim strMessage As String = ""
        If cbOperator.Items.Count > 0 Then
            If cbAttrName.Items.Contains(cbAttrName.Text.ToUpper) Then
                If Not cbOperator.Items.Contains(cbOperator.Text) Then
                    strMessage = "Operator must be a field contained in the drop-down list."
                End If
            End If
        End If

        If strMessage <> "" Then
            MsgBox(strMessage)
        End If

        strMessage = Nothing
    End Sub

End Class