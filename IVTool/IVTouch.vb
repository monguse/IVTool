Module IVTouch
    Private ivApp As Inventor.Application
    Private pid As Integer
    Private tbDoc As Inventor.DrawingDocument
    Private tb1 As Inventor.TitleBlockDefinition
    Private tb2 As Inventor.TitleBlockDefinition
    'Private ivApp As Inventor.ApprenticeServerComponent
    Private Const BaseVaultPath = "C:\VaultWorkspace"
    Private Const BaseDWFSavePath = "C:\Projects\dwf"

    Public Sub LoadIV()
        Dim ivProcesses() As Process

        ivApp = CreateObject("Inventor.Application")
        ivApp.Visible = False
        ivApp.SilentOperation = True
        ivApp.ScreenUpdating = False
        ivProcesses = Process.GetProcessesByName("Inventor")
        pid = ivProcesses(ivProcesses.GetLowerBound(0)).Id
    End Sub

    Public Sub CloseIV()
        Process.GetProcessById(pid).Kill()
    End Sub

    Public Sub ProcessAssembly()
        Dim drawingPath, filePath As String
        Dim refDocs As Inventor.DocumentsEnumerator
        Dim refDrawings As New Collection
        Dim ivDoc As Inventor.Document
        Dim count As Integer

        filePath = frm_IVTool.tb_DocNumber.Text

        If Not My.Computer.FileSystem.FileExists(filePath) Then
            filePath = RecursiveFind(BaseVaultPath, frm_IVTool.tb_DocNumber.Text + ".iam")
            If filePath = "" Then
                filePath = RecursiveFind(BaseVaultPath, frm_IVTool.tb_DocNumber.Text + ".ipt")
            End If
        End If

        If My.Computer.FileSystem.FileExists(filePath) Then
            frm_IVTool.Cout("Document found!" & vbNewLine)
            frm_IVTool.Cout("+ " & My.Computer.FileSystem.GetName(filePath) & vbNewLine)
            frm_IVTool.Cout("Getting children..." & vbNewLine)
        Else
            frm_IVTool.Cout("Document " & frm_IVTool.tb_DocNumber.Text & " not found" & vbNewLine)
        End If

        ivDoc = ivApp.Documents.Open(filePath)
        refDocs = ivDoc.AllReferencedDocuments

        For Each refDoc As Inventor.Document In refDocs
            drawingPath = FindDrawing(refDoc.FullFileName)
            If drawingPath <> "" Then
                refDrawings.Add(drawingPath)
            End If
        Next
        drawingPath = FindDrawing(filePath)
        If drawingPath <> "" Then
            refDrawings.Add(drawingPath)
        End If

        ivDoc.Close()
        frm_IVTool.Cout("Found " & refDocs.Count & " child documents" & vbNewLine)
        frm_IVTool.Cout("Found " & refDrawings.Count & " drawing documents" & vbNewLine)
        frm_IVTool.Cout("Processing drawings..." + vbNewLine)

        tbDoc = ivApp.Documents.Open("C:\Projects\tbsource.idw")
        For Each tbdef As Inventor.TitleBlockDefinition In tbDoc.TitleBlockDefinitions
            Debug.Print(tbdef.Name)
        Next
        tb1 = tbDoc.TitleBlockDefinitions.Item("SL_Tblock_Main_Stampable")
        tb2 = tbDoc.TitleBlockDefinitions.Item("SL_Tblock_2_Stampable")
        count = 1
        For Each dwgPath In refDrawings
            ProcessDrawing(dwgPath, count, refDrawings.Count)
            count += 1
        Next
        count -= 1
        tbDoc.Close()
        frm_IVTool.Cout("Converted " & count & " drawings successfully" & vbNewLine)
        frm_IVTool.Cout("Failed to convert " & refDrawings.Count - count & " drawings")

    End Sub

    Private Sub ProcessDrawing(ByVal docPath As String, ByRef count As Integer, ByVal total As Integer)
        Dim oDoc As Inventor.Document

        Try
            frm_IVTool.Cout("+[" & count & " of " & total & "] " & My.Computer.FileSystem.GetName(docPath) & "...")
            oDoc = ivApp.Documents.Open(docPath)
            If frm_IVTool.cb_ReplaceTB.CheckState Then
                ReplaceTitleblock(oDoc)
            End If
            PublishDWF(oDoc)
            oDoc.Close()
            frm_IVTool.Cout("OK!" & vbNewLine)
        Catch ex As Exception
            Debug.Print(ex.Message)
            frm_IVTool.Cout("FAILED" & vbNewLine)
            count -= 1
        End Try

    End Sub

    Private Function FindDrawing(docPath As String) As String
        Dim baseDocName As String

        baseDocName = My.Computer.FileSystem.GetName(docPath)
        baseDocName = Left(baseDocName, Len(baseDocName) - 4)

        Return RecursiveFind(BaseVaultPath, baseDocName + ".idw")

    End Function

    Private Sub ReplaceTitleblock(oDoc As Inventor.DrawingDocument)
        Dim tba, tbb As Inventor.TitleBlockDefinition

        tba = tb1.CopyTo(oDoc)
        tbb = tb2.CopyTo(oDoc)

        For Each oSheet As Inventor.Sheet In oDoc.Sheets
            If Not oSheet.TitleBlock Is Nothing Then
                oSheet.TitleBlock.Delete()
            End If
            If oSheet.Name = "Sheet:1" Then
                oSheet.AddTitleBlock(tba)
            Else
                oSheet.AddTitleBlock(tbb)
            End If
        Next
    End Sub

    Private Function RecursiveFind(ByVal folderPath As String, ByVal fileName As String) As String
        Dim filePath As String

        filePath = folderPath + "\" + fileName
        RecursiveFind = ""

        If My.Computer.FileSystem.FileExists(filePath) Then
            RecursiveFind = filePath
        Else
            For Each foundDir As String In My.Computer.FileSystem.GetDirectories(folderPath)
                RecursiveFind = RecursiveFind + RecursiveFind(foundDir, fileName)
            Next
        End If
    End Function

    Private Sub PublishDWF(ByRef oDoc As Inventor.Document)
        ' Get the DWF translator Add-In.
        'Dim DWFAddIn As Inventor.TranslatorAddIn
        Dim DWFAddIn As Object
        DWFAddIn = DirectCast(ivApp.ApplicationAddIns.ItemById("{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}"), Inventor.TranslatorAddIn)

        'Set a reference to the active document (the document to be published).
        'Dim odocument As Document
        'Set odocument = ThisApplication.ActiveDocument

        'Copy model iproperties to drawing iproperties
        Dim propSetDU As Inventor.PropertySet
        propSetDU = oDoc.PropertySets("Inventor User Defined Properties")

        Dim propSetDS As Inventor.PropertySet
        propSetDS = oDoc.PropertySets("Inventor Summary Information")

        Dim propSetMU As Inventor.PropertySet
        propSetMU = oDoc.Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.PropertySets.Item("Inventor User Defined Properties")

        Dim propSetMS As Inventor.PropertySet
        propSetMS = oDoc.Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.PropertySets.Item("Inventor Summary Information")

        If Exists(propSetMU, "SL12_Process Code") Then ' eg: "01"
            If Exists(propSetDU, "SL12_Process Code") Then
                propSetDU.Item("SL12_Process Code").Value = propSetMU.Item("SL12_Process Code").Value
            Else
                Call propSetDU.Add(propSetMU.Item("SL12_Process Code").Value, "SL12_Process Code")
            End If
        End If

        If Exists(propSetMU, "SL02_Product_Code") Then ' eg: "D-10-SLICKLINE UNITS"
            If Exists(propSetDU, "SL02_Product_Code") Then
                propSetDU.Item("SL02_Product_Code").Value = propSetMU.Item("SL02_Product_Code").Value
            Else
                Call propSetDU.Add(propSetMU.Item("SL02_Product_Code").Value, "SL02_Product_Code")
            End If
        End If

        If Exists(propSetMU, "SL13_Product Category") Then ' eg: "D13"
            If Exists(propSetDU, "SL13_Product Category") Then
                propSetDU.Item("SL13_Product Category").Value = propSetMU.Item("SL13_Product Category").Value
            Else
                Call propSetDU.Add(propSetMU.Item("SL13_Product Category").Value, "SL13_Product Category")
            End If
        End If

        If Exists(propSetMU, "SL04_Drawing_Document_Type") Then ' eg: "Principal Assembly"
            If Exists(propSetDU, "SL04_Drawing_Document_Type") Then
                propSetDU.Item("SL04_Drawing_Document_Type").Value = propSetMU.Item("SL04_Drawing_Document_Type").Value
            Else
                Call propSetDU.Add(propSetMU.Item("SL04_Drawing_Document_Type").Value, "SL04_Drawing_Document_Type")
            End If
        End If

        If Exists(propSetMS, "Title") Then ' eg: "SPOOLER"
            If Exists(propSetDS, "Title") Then
                propSetDS.Item("Title").Value = propSetMS.Item("Title").Value
            Else
                Call propSetDS.Add(propSetMS.Item("Title").Value, "Title")
            End If
        End If

        Dim oContext As Inventor.TranslationContext
        oContext = ivApp.TransientObjects.CreateTranslationContext
        oContext.Type = Inventor.IOMechanismEnum.kFileBrowseIOMechanism

        ' Create a NameValueMap object
        Dim oOptions As Inventor.NameValueMap
        oOptions = ivApp.TransientObjects.CreateNameValueMap

        ' Create a DataMedium object
        Dim oDataMedium As Inventor.DataMedium
        oDataMedium = ivApp.TransientObjects.CreateDataMedium

        ' Check whether the translator has 'SaveCopyAs' options
        If DWFAddIn.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then


            oOptions.Value("Launch_Viewer") = 0

            ' Other options...
            oOptions.Value("Publish_All_Component_Props") = 0
            oOptions.Value("Publish_All_Physical_Props") = 0
            'oOptions.Value("Password") = 0

            If TypeOf oDoc Is Inventor.DrawingDocument Then

                ' Drawing options
                oOptions.Value("Publish_Mode") = Inventor.DWFPublishModeEnum.kCustomDWFPublish
                oOptions.Value("Publish_All_Sheets") = 1
            End If

        End If

        'Set the destination file name
        oDataMedium.FileName = frm_IVTool.tb_FolderPath.Text + _
                                "\" + _
                                My.Computer.FileSystem.GetName(oDoc.FullFileName) + _
                                ".dwf"

        'Publish document.
        Call DWFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)


    End Sub

    Private Function Exists(ByVal oCol As Inventor.PropertySet, ByVal vKey As String) As Boolean
        Dim o As Object
        Try
            o = oCol.Item(vKey)
            Exists = True
        Catch ex As Exception
            Exists = ExistsNonObject(oCol, vKey)
        End Try
    End Function

    Private Function ExistsNonObject(ByVal oCol As Inventor.PropertySet, ByVal vKey As String) As Boolean
        Dim v As Object
        Try
            v = oCol.Item(vKey)
            ExistsNonObject = True
        Catch ex As Exception
            ExistsNonObject = False
        End Try
    End Function

End Module
