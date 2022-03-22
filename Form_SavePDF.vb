Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Math
Imports System.Windows.Forms



Public Class Form_SavePDF


    Private Sub Form_SavePDF_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim _invApp As Inventor.Application = Marshal.GetActiveObject("Inventor.Application")

        Dim oDocument As Document
        oDocument = _invApp.ActiveDocument

        Dim oDocumentDWGSavePath As String
        Dim oDocumentDirectory As String
        Dim oDocumentFileName As String
        Dim oDocumentRevLevel As String = Nothing

        If CheckBox1.Checked = False Then
            If NumericUpDown1.Value < 0 Then
                oDocumentRevLevel = ""
            Else
                oDocumentRevLevel = " REV-" & CStr(NumericUpDown1.Value)
            End If
        ElseIf CheckBox1.Checked = True Then
            oDocumentRevLevel = " REV-" & CStr(ComboBox1.SelectedItem)
        End If


        oDocumentDWGSavePath = oDocument.FullDocumentName
        oDocumentDirectory = System.IO.Path.GetDirectoryName(oDocumentDWGSavePath)

        oDocumentFileName = System.IO.Path.GetFileNameWithoutExtension(oDocumentDWGSavePath)
        oDocumentFileName = oDocumentDirectory & "\" & oDocumentFileName & oDocumentRevLevel & ".pdf"

        If String.IsNullOrEmpty(oDocumentDWGSavePath) Then
            MsgBox("Please save drawing file.")
            Exit Sub
        Else
            TextBox2.Text = oDocumentDWGSavePath
            TextBox3.Text = oDocumentFileName
            Button4.Enabled = True
            'Debug.Print("oDocumentSavePath: " & oDocumentDWGSavePath)
        End If

        ToolStripStatusLabel1.Text = "Ready"

    End Sub



    Private Sub SavePDF4()

        Dim _invApp As Inventor.Application = Marshal.GetActiveObject("Inventor.Application")

        ' Get the PDF translator Add-In.
        Dim PDFAddIn As TranslatorAddIn
        PDFAddIn = _invApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")

        'Set a reference to the active document (the document to be published).
        Dim oDocument As Document
        oDocument = _invApp.ActiveDocument

        Dim oContext As TranslationContext
        oContext = _invApp.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As NameValueMap
        oOptions = _invApp.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As DataMedium
        oDataMedium = _invApp.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        ToolStripStatusLabel1.Text = "Setting save options..."
        If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
            ' Options for drawings... '
            'oOptions.Value("All_Color_AS_Black") = 0
            oOptions.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets

            'oOptions.Value("Remove_Line_Weights") = 0
            'oOptions.Value("Vector_Resolution") = 400
            'oOptions.Value("Custom_Begin_Sheet") = 2
            'oOptions.Value("Custom_End_Sheet") = 4
        End If

        'Set the destination file name
        ToolStripStatusLabel1.Text = "Setting the file name..."
        oDataMedium.FileName = TextBox3.Text

        'Publish document.
        ToolStripStatusLabel1.Text = "Saving PDF..."
        Try
            PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
            ToolStripStatusLabel1.Text = "Save complete."

            If CheckBox2.Checked = True Then
                ToolStripStatusLabel1.Text = "Opening PDF file."
                Process.Start(TextBox3.Text)
            End If

            ToolStripStatusLabel1.Text = "Complete."
        Catch ex As Exception
            ToolStripStatusLabel1.Text = "Unable to save PDF file."
            MsgBox("Unable to save PDF file.")
        End Try



    End Sub


    Private Sub UpdatePDFfilepath()

        If String.IsNullOrEmpty(TextBox2.Text) Then
            Exit Sub
        End If

        Dim oDocumentDWGSavePath As String
        Dim oDocumentDirectory As String
        Dim oDocumentFileName As String
        Dim oDocumentRevLevel As String = Nothing
        Dim oDocumentIssued As String = Nothing


        If CheckBox1.Checked = False And CheckBox3.Checked = False Then
            If NumericUpDown1.Value < 0 Then
                oDocumentRevLevel = ""
            Else
                oDocumentRevLevel = " REV-" & CStr(NumericUpDown1.Value)
            End If
        ElseIf CheckBox1.Checked = True And CheckBox3.Checked = False Then
            oDocumentRevLevel = " REV-" & ComboBox1.SelectedItem
        ElseIf CheckBox3.Checked = True Then
            oDocumentRevLevel = ""
        End If


        If CheckBox4.Checked = True Then
            oDocumentIssued = "ISSUED\"
        ElseIf CheckBox4.Checked = False Then
            oDocumentIssued = ""
        End If


        oDocumentDWGSavePath = TextBox2.Text
        oDocumentDirectory = System.IO.Path.GetDirectoryName(TextBox2.Text)

        oDocumentFileName = System.IO.Path.GetFileNameWithoutExtension(oDocumentDWGSavePath)
        oDocumentFileName = oDocumentDirectory & "\" & oDocumentIssued & oDocumentFileName & oDocumentRevLevel & ".pdf"

        TextBox3.Text = oDocumentFileName


    End Sub






    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        SavePDF4()

    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged

        UpdatePDFfilepath()

    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub


    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click

        If CheckBox1.Checked = True Then
            NumericUpDown1.Enabled = False
            ComboBox1.Enabled = True
        Else
            NumericUpDown1.Enabled = True
            ComboBox1.Enabled = False
        End If

        UpdatePDFfilepath()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        UpdatePDFfilepath()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            NumericUpDown1.Enabled = False
            CheckBox1.Enabled = False
            ComboBox1.Enabled = False
        Else
            NumericUpDown1.Enabled = True
            CheckBox1.Enabled = True
        End If

        UpdatePDFfilepath()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim filestring As String = Nothing
        Dim oDocumentDWGSavePath As String
        Dim oDocumentDirectory As String
        Dim oDocumentFileName As String

        oDocumentDWGSavePath = TextBox2.Text
        oDocumentDirectory = System.IO.Path.GetDirectoryName(oDocumentDWGSavePath)

        oDocumentFileName = System.IO.Path.GetFileNameWithoutExtension(TextBox3.Text)
        oDocumentFileName = oDocumentFileName & ".pdf"

        'Save file dialog
        Using obj As New OpenFileDialog
            obj.Filter = "DWG Files|*.dwg"
            obj.CheckFileExists = False
            obj.CheckPathExists = False
            obj.InitialDirectory = oDocumentDirectory
            obj.FileName = oDocumentFileName
            obj.Title = "Choose file save location"

            If obj.ShowDialog = Windows.Forms.DialogResult.OK Then

                filestring = obj.FileName
                TextBox3.Text = filestring

            End If

        End Using



    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged



        Dim oDocumentDWGSavePath As String
        Dim oDocumentDirectory As String
        Dim oDocumentDirectoryIssued As String


        oDocumentDWGSavePath = TextBox2.Text
        oDocumentDirectory = System.IO.Path.GetDirectoryName(oDocumentDWGSavePath)
        oDocumentDirectoryIssued = oDocumentDirectory & "\ISSUED"

        'Debug.Print("oDocumentDirectory: " & oDocumentDirectory & " | oDocumentDirectoryIssued: " & oDocumentDirectoryIssued)


        If System.IO.Directory.Exists(oDocumentDirectoryIssued) Then
            UpdatePDFfilepath()
        Else
            If CheckBox4.Checked = True Then
                MsgBox("ISSUED folder doesn't exist")
                CheckBox4.Checked = False
            End If
        End If




    End Sub
End Class