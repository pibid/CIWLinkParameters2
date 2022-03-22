
Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Math
Imports System.Windows.Forms

Public Class Form_LinkParameters

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim _invApp As Inventor.Application = Marshal.GetActiveObject("Inventor.Application")

        ' Dim partFileDoc As PartDocument = _invApp.ActiveDocument
        Dim partFileDoc As Document = _invApp.ActiveDocument

        Dim newPartFileParams As UserParameters = partFileDoc.ComponentDefinition.Parameters.UserParameters

        Dim paramComment As String = Nothing
        Dim paramName As String = Nothing
        Dim paramExpression As String = Nothing


        For Each param In newPartFileParams

            paramComment = param.Comment
            paramName = param.Name
            paramExpression = param.Expression

            If paramName = "NOZZLE_ELEVATION" Or paramName = "MW_ELEVATION" Or paramName = "SG_ELEVATION" Then
                TextBox2.Enabled = True
                TextBox3.Text = paramExpression
                TextBox3.Enabled = True
            ElseIf paramName = "NOZZLE_ORIENTATION" Or paramName = "MW_ORIENTATION" Or paramName = "SG_ORIENTATION" Then
                TextBox2.Enabled = True
                TextBox4.Text = paramExpression
                TextBox4.Enabled = True
            ElseIf paramName = "NOZZLE_SIZE" Or paramName = "MW_SIZE" Or paramName = "SG_SIZE" Then
                TextBox2.Enabled = True
                TextBox5.Text = paramExpression
                TextBox5.Enabled = True
            ElseIf paramName = "NOZZLE_OFFSET" Or paramName = "MW_OFFSET" Or paramName = "SG_OFFSET" Then
                TextBox2.Enabled = True
                TextBox6.Text = paramExpression
                TextBox6.Enabled = True
            End If

        Next



        Dim oAsset As AssetLibrary
        oAsset = _invApp.AssetLibraries.Item("CIW_Materials")

        Dim oAssetMaterial As MaterialAsset
        'oAssetMaterial = oAsset.MaterialAssets.Item("304L SS")


        Dim materialList As New ArrayList


        For Each oAssetMaterial In oAsset.MaterialAssets
            materialList.Add(oAssetMaterial.DisplayName)
        Next

        materialList.Sort()
        ComboBox1.Items.AddRange(materialList.ToArray())


        ToolStripStatusLabel1.Text = "Ready"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim specfileString As String = Nothing

        'Open file dialog
        Using obj As New OpenFileDialog
            obj.Filter = "Autodesk Inventor Parts|*.ipt"
            obj.CheckFileExists = False
            obj.CheckPathExists = False
            obj.InitialDirectory = "C:\Work\Designs\Projects"
            'obj.FileName = "SPECIFICATION"
            obj.Title = "Open spec file"

            If obj.ShowDialog = Windows.Forms.DialogResult.OK Then

                specfileString = obj.FileName

                If specfileString.Contains("SPECIFICATION") Then
                    'MsgBox("Looks like a spec file!: " & specfileString)
                Else
                    MsgBox("File name does not include ""SPECIFICATION""." & vbCrLf & "Please verify this is a spec file.")
                End If


                TextBox1.Text = specfileString

                Button2.Enabled = True
                'Button3.Enabled = True

            End If
        End Using

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If String.IsNullOrEmpty(TextBox2.Text) And TextBox2.Enabled = True Then
            MsgBox("Please enter a nozzle tag.")
            Exit Sub
        ElseIf String.IsNullOrEmpty(ComboBox1.Text) Then
            MsgBox("Please select a Material.")
            Exit Sub
        End If

        Dim SW As New Stopwatch
        SW.Start()

        Dim _invApp As Inventor.Application = Marshal.GetActiveObject("Inventor.Application")

        Dim partFileDoc As PartDocument = _invApp.ActiveDocument
        Dim newPartFileParams As UserParameters = partFileDoc.ComponentDefinition.Parameters.UserParameters

        Dim SpecFileDoc As PartDocument = Nothing
        Dim specFileObj As New System.IO.FileInfo(TextBox1.Text)
        Dim specFileDirectory As String = specFileObj.Directory.FullName

        Dim partCompDef As PartComponentDefinition
        Dim specFilePath As String = TextBox1.Text
        SpecFileDoc = _invApp.Documents.Open(specFilePath, False)
        partCompDef = SpecFileDoc.ComponentDefinition
        Dim specFileUserParams As UserParameters = SpecFileDoc.ComponentDefinition.Parameters.UserParameters

        Dim oParamsToLink As ObjectCollection
        oParamsToLink = _invApp.TransientObjects.CreateObjectCollection

        Dim nozzleTag As String = TextBox2.Text
        Dim nozzleTagElevation As String = nozzleTag & "_ELEVATION"
        Dim nozzleTagOrientation As String = nozzleTag & "_ORIENTATION"
        Dim nozzleTagSize As String = nozzleTag & "_SIZE"
        Dim nozzleTagOffset As String = nozzleTag & "_OFFSET"

        Dim nozzleTagElev_Exist As Boolean = False
        Dim nozzleTagOrient_Exist As Boolean = False
        Dim nozzleTagSize_Exist As Boolean = False
        Dim nozzleTagOffset_Exist As Boolean = False
        Dim nozzleIsManway As Boolean = False
        Dim partIsNozzle As Boolean = False

        Dim param As Parameter
        Dim specParam As Parameter
        Dim newParam As Parameter
        Dim paramComment As String = Nothing
        Dim paramName As String = Nothing
        Dim paramExpression As String = Nothing
        Dim specParamName As String = Nothing

        Dim oDerivedParamTable As DerivedParameterTable


        For Each param In newPartFileParams
            paramName = param.Name

            If paramName = "MW_ELEVATION" Then
                nozzleIsManway = True
            ElseIf paramName = "NOZZLE_ELEVATION" Then
                partIsNozzle = True
            End If
        Next


        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("C:\TestFolder\test.txt")
            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.SetDelimiters(",")
            Dim currentRow As String()
            While Not MyReader.EndOfData
                Try
                    currentRow = MyReader.ReadFields()
                    Dim currentField As String
                    For Each currentField In currentRow
                        MsgBox(currentField)
                    Next
                Catch ex As Microsoft.VisualBasic.
                    FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
            End While
        End Using





        For Each param In newPartFileParams

            paramName = param.Name
            paramExpression = param.Expression

            ToolStripStatusLabel1.Text = "Evaluating parameter to link from spec file: " & paramName
            'Debug.Print("paramName: " & paramName)

            For Each specParam In specFileUserParams

                specParamName = specParam.Name

                If specParamName = paramName Then
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(paramName))
                    Catch ex As Exception
                        MsgBox("Unable to link parameter to part file: " & paramName)
                    End Try
                ElseIf specParamName = nozzleTagElevation Then
                    nozzleTagElev_Exist = True
                ElseIf specParamName = nozzleTagOrientation Then
                    nozzleTagOrient_Exist = True
                ElseIf specParamName = nozzleTagSize Then
                    nozzleTagSize_Exist = True
                ElseIf specParamName = nozzleTagOffset Then
                    nozzleTagOffset_Exist = True
                End If
            Next


            If paramName = "NOZZLE_ELEVATION" Or paramName = "MW_ELEVATION" Or paramName = "SG_ELEVATION" Then
                ToolStripStatusLabel1.Text = "Evaluating [NozzleTag]_ELEVATION to add to spec file"
                If nozzleTagElev_Exist = True Then
                    ToolStripStatusLabel1.Text = "Parameter exists in spec file, linking to part file | nozzleTagElevation: " & nozzleTagElevation
                    Debug.Print("Parameter exists in spec file, linking to part file | nozzleTagElevation: " & nozzleTagElevation)
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(nozzleTagElevation))
                    Catch ex As Exception
                        MsgBox("Unable to link [NozzleTag]_ELEVATION parameter to part file")
                    End Try
                Else
                    ToolStripStatusLabel1.Text = "Parameter does not exist in spec file, adding parameter | nozzleTagElevation: " & nozzleTagElevation
                    Debug.Print("Parameter does not exist in spec file, adding parameter | nozzleTagElevation: " & nozzleTagElevation)
                    Try
                        newParam = specFileUserParams.AddByExpression(nozzleTagElevation, TextBox3.Text, UnitsTypeEnum.kInchLengthUnits)
                        newParam.ExposedAsProperty = True
                        nozzleTagElev_Exist = True
                    Catch ex As Exception
                        'nozzleTagElev_Exist = False
                        MsgBox("Unable to add [NozzleTag]_ELEVATION parameter to spec file")
                    End Try

                    ToolStripStatusLabel1.Text = "Linking parameter to part file | nozzleTagElevation: " & nozzleTagElevation
                    Debug.Print("Linking parameter to part file | nozzleTagElevation: " & nozzleTagElevation)
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(nozzleTagElevation))
                    Catch ex As Exception
                        MsgBox("Unable to link [NozzleTag]_ELEVATION parameter to part file")
                    End Try
                End If

            ElseIf paramName = "NOZZLE_ORIENTATION" Or paramName = "MW_ORIENTATION" Or paramName = "SG_ORIENTATION" Then
                ToolStripStatusLabel1.Text = "Evaluating [NozzleTag]_ORIENTATION to add to spec file"
                If nozzleTagOrient_Exist = True Then
                    ToolStripStatusLabel1.Text = "Parameter exists in spec file, linking to part file | nozzleTagOrientation: " & nozzleTagOrientation
                    Debug.Print("Parameter exists in spec file, linking to part file | nozzleTagOrientation: " & nozzleTagOrientation)
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(nozzleTagOrientation))
                    Catch ex As Exception
                        MsgBox("Unable to link [NozzleTag]_ORIENTATION parameter to part file")
                    End Try
                Else
                    ToolStripStatusLabel1.Text = "Parameter does not exist in spec file, adding parameter | nozzleTagOrientation: " & nozzleTagOrientation
                    Debug.Print("Parameter does not exist in spec file, adding parameter | nozzleTagOrientation: " & nozzleTagOrientation)
                    Try
                        newParam = specFileUserParams.AddByExpression(nozzleTagOrientation, TextBox4.Text, UnitsTypeEnum.kDegreeAngleUnits)
                        newParam.ExposedAsProperty = True
                        nozzleTagOrient_Exist = True
                    Catch ex As Exception
                        'nozzleTagOrient_Exist = False
                        MsgBox("Unable to add [NozzleTag]_ORIENTATION parameter to spec file")
                    End Try

                    ToolStripStatusLabel1.Text = "Linking parameter to part file | nozzleTagOrientation: " & nozzleTagOrientation
                    Debug.Print("Linking parameter to part file | nozzleTagOrientation: " & nozzleTagOrientation)
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(nozzleTagOrientation))
                    Catch ex As Exception
                        MsgBox("Unable to link [NozzleTag]_ORIENTATION parameter to part file")
                    End Try
                End If

            ElseIf paramName = "NOZZLE_SIZE" Or paramName = "SG_SIZE" Then
                ToolStripStatusLabel1.Text = "Evaluating [NozzleTag]_SIZE to add to spec file"
                If nozzleTagSize_Exist = True And nozzleIsManway = False Then
                    ToolStripStatusLabel1.Text = "Parameter exists in spec file, linking to part file | nozzleTagSize: " & nozzleTagSize
                    Debug.Print("Parameter exists in spec file, linking to part file | nozzleTagSize: " & nozzleTagSize)
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(nozzleTagSize))
                    Catch ex As Exception
                        MsgBox("Unable to link [NozzleTag]_SIZE parameter to part file")
                    End Try
                ElseIf nozzleTagSize_Exist = False And nozzleIsManway = False Then
                    ToolStripStatusLabel1.Text = "Parameter does not exist in spec file, adding parameter | nozzleTagSize: " & nozzleTagSize
                    Debug.Print("Parameter does not exist in spec file, adding parameter | nozzleTagSize: " & nozzleTagSize)
                    Try
                        newParam = specFileUserParams.AddByExpression(nozzleTagSize, TextBox5.Text, UnitsTypeEnum.kInchLengthUnits)
                        newParam.ExposedAsProperty = True
                        nozzleTagSize_Exist = True
                    Catch ex As Exception
                        'nozzleTagSize_Exist = False
                        MsgBox("Unable to add [nozzle Tag]_SIZE parameter to spec file")
                    End Try

                    ToolStripStatusLabel1.Text = "Linking parameter to part file | nozzleTagSize: " & nozzleTagSize
                    Debug.Print("Linking parameter to part file | nozzleTagSize: " & nozzleTagSize)
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(nozzleTagSize))
                    Catch ex As Exception
                        MsgBox("Unable to link [NozzleTag]_SIZE parameter to part file")
                    End Try
                End If

            ElseIf paramName = "NOZZLE_OFFSET" Or paramName = "MW_OFFSET" Or paramName = "SG_OFFSET" Then
                ToolStripStatusLabel1.Text = "Evaluating [NozzleTag]_OFFSET to add to spec file"
                If nozzleTagOffset_Exist = True Then
                    ToolStripStatusLabel1.Text = "Parameter exists in spec file, linking to part file | nozzleTagOffset: " & nozzleTagOffset
                    Debug.Print("Parameter exists in spec file, linking to part file | nozzleTagOffset: " & nozzleTagOffset)
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(nozzleTagOffset))
                    Catch ex As Exception
                        MsgBox("Unable to link [NozzleTag]_OFFSET parameter to part file")
                    End Try
                Else
                    ToolStripStatusLabel1.Text = "Parameter does not exist in spec file, adding parameter | nozzleTagOffset: " & nozzleTagOffset
                    Debug.Print("Parameter does not exist in spec file, adding parameter | nozzleTagOffset: " & nozzleTagOffset)
                    Try
                        newParam = specFileUserParams.AddByExpression(nozzleTagOffset, TextBox6.Text, UnitsTypeEnum.kInchLengthUnits)
                        newParam.ExposedAsProperty = True
                        nozzleTagOffset_Exist = True
                    Catch ex As Exception
                        'nozzleTagOffset_Exist = false
                        MsgBox("Unable to add [nozzle Tag]_OFFSET parameter to spec file")
                    End Try

                    ToolStripStatusLabel1.Text = "Linking parameter to part file | nozzleTagOffset: " & nozzleTagOffset
                    Debug.Print("Linking parameter to part file | nozzleTagOffset: " & nozzleTagOffset)
                    Try
                        oParamsToLink.Add(partCompDef.Parameters.Item(nozzleTagOffset))
                    Catch ex As Exception
                        MsgBox("Unable to link [NozzleTag]_OFFSET parameter to part file")
                    End Try
                End If
            End If


        Next



        ToolStripStatusLabel1.Text = "Linking parameters from spec file..."

        Try
            oDerivedParamTable = partFileDoc.ComponentDefinition.Parameters.DerivedParameterTables.Add2(specFilePath, oParamsToLink)
        Catch ex As Exception
            MsgBox("Unable to add at least one linked parameter")
        End Try



        ToolStripStatusLabel1.Text = "Writing part file parameter expressions..."

        For Each param In newPartFileParams

            paramComment = param.Comment
            paramName = param.Name

            If paramName = "NOZZLE_ELEVATION" Or paramName = "MW_ELEVATION" Or paramName = "SG_ELEVATION" Then
                Try
                    newPartFileParams.Item(paramName).Expression = nozzleTagElevation
                Catch ex As Exception
                    MsgBox("Unable to write [NozzleTag]_ELEVATION expression to part file parameter.")
                End Try

            ElseIf paramName = "NOZZLE_ORIENTATION" Or paramName = "MW_ORIENTATION" Or paramName = "SG_ORIENTATION" Then
                Try
                    newPartFileParams.Item(paramName).Expression = nozzleTagOrientation
                Catch ex As Exception
                    MsgBox("Unable to write [NozzleTag]_ORIENTATION expression to part file parameter.")
                End Try

            ElseIf paramName = "NOZZLE_SIZE" Or paramName = "SG_SIZE" Then

                If nozzleIsManway = False Then
                    Try
                        newPartFileParams.Item(paramName).Expression = nozzleTagSize
                    Catch ex As Exception
                        MsgBox("Unable to write [NozzleTag]_SIZE expression to part file parameter.")
                    End Try
                End If

            ElseIf paramName = "NOZZLE_OFFSET" Or paramName = "MW_OFFSET" Or paramName = "SG_OFFSET" Then
                    Try
                        newPartFileParams.Item(paramName).Expression = nozzleTagOffset
                    Catch ex As Exception
                        MsgBox("Unable to write [NozzleTag]_OFFSET expression to part file parameter.")
                    End Try

                ElseIf paramComment.Contains("_1") Then
                    Try
                    newPartFileParams.Item(paramName).Expression = paramName & "_1"
                Catch ex As Exception
                    MsgBox("Unable to write expression to part file parameter: " & paramName)
                End Try

            End If

        Next

        Dim oAsset As AssetLibrary
        oAsset = _invApp.AssetLibraries.Item("CIW_Materials")

        Dim mAsset As MaterialAsset = oAsset.MaterialAssets.Item(ComboBox1.SelectedItem)
        partFileDoc.ActiveMaterial = mAsset


        SW.Stop()
        Label6.Text = (SW.ElapsedMilliseconds) / 1000 & " sec"
        Label6.Visible = True


        Dim partfileprefix As String = TextBox7.Text
        Dim filenamestring As String = ".ipt"
        Dim partFileString As String

        If partIsNozzle = True Then
            filenamestring = " NOZZLE " & nozzleTag & " .ipt"
        End If

        partFileString = partfileprefix & filenamestring

        'Save file dialog
        ToolStripStatusLabel1.Text = "Saving part file"

        Using obj As New SaveFileDialog
            obj.Filter = "Autodesk Inventor Parts|*.ipt"
            obj.CheckFileExists = False
            obj.CheckPathExists = False
            obj.InitialDirectory = specFileDirectory
            obj.FileName = partFileString
            obj.Title = "Save part file"

            Do
                If obj.ShowDialog = Windows.Forms.DialogResult.OK Then
                    partfileString = obj.FileName

                    'Save part file
                    _invApp.SilentOperation = True
                    partFileDoc.SaveAs(partfileString, False)
                    _invApp.SilentOperation = False

                    Exit Do
                Else
                    Dim result As DialogResult = MsgBox("File not saved, continue?" & vbCrLf & "Click No to go back and save the file, otherwise save will be bypassed", vbYesNo,)

                    If result = DialogResult.Yes Then
                        ToolStripStatusLabel1.Text = "Ready"
                        Exit Sub
                    End If

                End If
            Loop

        End Using



        SpecFileDoc.Close()

        ToolStripStatusLabel1.Text = "Complete."


        Button2.Enabled = False

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub


End Class

