Public Class clsinstallOld

    Public Const FormType = "InstallationAllocation"
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objRs As SAPbobsCOM.Recordset
    Dim i As Integer
    Dim contractid As Integer
    Dim strSQL As String

    Public Sub LoadScreen()
        Try
            objForm = objAddOn.objUIXml.LoadScreenXML("InstallationAllocation.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, FormType)
            objMatrix = objForm.Items.Item("4").Specific
            objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            Refresh(objForm.UniqueID)
            '  objMatrix.AddRow(1)
            objMatrix = objForm.Items.Item("5").Specific
        Catch Ex As Exception
            MsgBox(Ex.ToString)
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        If pVal.BeforeAction Then
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "101" Then
                        ServiceAdd()
                    ElseIf pVal.ItemUID = "7" Then
                        AddName(FormUID)
                    ElseIf pVal.ItemUID = "4" And pVal.ColUID = "0" Then
                        ShowName(FormUID)
                    End If

                    If pVal.ItemUID = "3" Then
                        objMatrix.Clear()
                        Refresh(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.ItemUID = "5" And pVal.ColUID = "2" Then
                        choose(FormUID, pVal)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "5" And pVal.ColUID = "2" Then

                        AddRow(FormUID)
                    End If
            End Select
        End If
    End Sub
    Private Sub AddRow(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("5").Specific
        If objMatrix.RowCount = 0 Then
            objMatrix.AddRow()
        ElseIf objMatrix.RowCount > 1 Then
            If objMatrix.Columns.Item("2").Cells.Item(objMatrix.RowCount).Specific.string <> String.Empty Then
                objForm.DataSources.DBDataSources.Item("@MIPLCNTC").Clear()
                objMatrix.AddRow()
            End If
        Else
            objForm.DataSources.DBDataSources.Item("@MIPLCNTC").Clear()
            objMatrix.AddRow()
        End If
    End Sub
    Private Sub choose(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim objDataTable As SAPbouiCOM.DataTable
        objCFLEvent = pval
        objDataTable = objCFLEvent.SelectedObjects
        ' Dim i As String
        Try
            If objDataTable Is Nothing Then
            Else

                objMatrix = objForm.Items.Item("5").Specific
                objMatrix.GetLineData(pval.Row)
                objForm.DataSources.DBDataSources.Item("@MIPLCNTC").SetValue("U_techname", 0, objDataTable.GetValue("firstName", 0))
                objForm.DataSources.DBDataSources.Item("@MIPLCNTC").SetValue("U_techcode", 0, objDataTable.GetValue("empID", 0))
                objMatrix.SetLineData(pval.Row)


            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub Refresh(ByVal FormUID As String)
        Dim strSQL As String
        Dim intLoop As Integer = 0
        Dim objRecordSet As SAPbobsCOM.Recordset
        Dim odatatabe As SAPbouiCOM.DataTable
        Dim ocons As SAPbouiCOM.Conditions
        Dim ocon As SAPbouiCOM.Condition
        Dim ocfl As SAPbouiCOM.ChooseFromList
        Dim i As Integer
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objAddOn.objApplication.SetStatusBarMessage("Loading Please Wait!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        objMatrix = objForm.Items.Item("4").Specific
        objMatrix.Clear()
        strSQL = "select distinct T2.contractid,T2.cstmrcode,T2.cstmrname from octr T2 left outer  join  oscl T1 on T1.contractid=t2.contractid where t1.contractid is null AND T2.EndDate >=GETDATE ()"
        objRecordSet = Nothing
        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRecordSet.DoQuery(strSQL)
        While Not objRecordSet.EoF
            objMatrix.AddRow()
            intLoop += 1
            objMatrix.Columns.Item("1").Cells.Item(intLoop).Specific.String = objRecordSet.Fields.Item("cstmrcode").Value
            objMatrix.Columns.Item("2").Cells.Item(intLoop).Specific.String = objRecordSet.Fields.Item("cstmrname").Value
            objMatrix.Columns.Item("3").Cells.Item(intLoop).Specific.String = objRecordSet.Fields.Item("contractid").Value
            objRecordSet.MoveNext()

        End While
        objRs = Nothing
        objAddOn.objApplication.StatusBar.SetText("Loading Completed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        'ChooseFrom List coding
        ocfl = objForm.ChooseFromLists.Item("CFL_4")
        If ocfl.GetConditions.Count = 0 Then

            Try
                odatatabe = objForm.DataSources.DataTables.Item("HEM6")
            Catch ex As Exception
                odatatabe = objForm.DataSources.DataTables.Add("HEM6")
            End Try
            odatatabe.ExecuteQuery("Select empId from HEM6 where RoleId = -2")
            ocons = ocfl.GetConditions

            For i = 0 To odatatabe.Rows.Count - 1
                ocon = ocons.Add()
                ocon.Alias = "empId"
                ocon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                ocon.CondVal = CInt(odatatabe.GetValue(0, i))
                If i < odatatabe.Rows.Count - 1 Then ocon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            Next
            ocfl.SetConditions(ocons)
        End If
    End Sub
    Private Sub ServiceAdd()
        Dim strSQL As String
        Dim strSQL1 As String
        Dim strSQL2 As String
        Dim strSQL3 As String
        Dim objRecordSet As SAPbobsCOM.Recordset
        Dim objService As SAPbobsCOM.ServiceCalls
        objService = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        Try
            objMatrix = objForm.Items.Item("4").Specific
            strSQL = "Select InternalSN,ItemCode from CTR1 Where ContractID= " & objMatrix.Columns.Item("3").Cells.Item(1).Specific.String & ""
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL)
            If Not objRecordSet.EoF Then
                objService.InternalSerialNum = objRecordSet.Fields.Item("InternalSN").Value
                objService.ItemCode = objRecordSet.Fields.Item("ItemCode").Value
                objService.CustomerName = objMatrix.Columns.Item("2").Cells.Item(1).Specific.string
                objService.ContractID = objMatrix.Columns.Item("3").Cells.Item(1).Specific.string
                objService.Subject = "Installation"
                strSQL1 = " select originID from OSCO Where Name='Installation'"
                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL1)
                objService.Origin = objRecordSet.Fields.Item("originID").Value
                strSQL2 = " select prblmTypID  from OSCP where Name='Installation'"
                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL2)
                objService.ProblemType = objRecordSet.Fields.Item("prblmTypID").Value
                strSQL3 = " select callTypeID  from OSCT Where Name='Installation Call'"
                objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecordSet.DoQuery(strSQL3)
                objService.CallType = objRecordSet.Fields.Item("callTypeID").Value
                If (objService.Add()) Then
                    MsgBox(objAddOn.objCompany.GetLastErrorDescription)
                End If
                MsgBox("Added")
            End If

        Catch ex As Exception
            MsgBox(objAddOn.objCompany.GetLastErrorDescription)
        End Try

        objRecordSet = Nothing
    End Sub
    Private Sub ShowName(ByVal FormUID As String)
        Dim contractno As Integer
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("4").Specific
        For i = 1 To objMatrix.RowCount()
            If (objMatrix.IsRowSelected(i)) Then
                contractno = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                Exit For
            End If
        Next
        TechExist(FormUID, CInt(contractno))
    End Sub
    Private Sub AddName(ByVal FormUID As String)
        Dim contractno As Integer
        Dim customername As String = ""
        Dim techname As String
        Dim techcode As Integer
        Dim i As Integer
        Dim intLoop As Integer
        Dim code As Integer
        Dim objRS As SAPbobsCOM.Recordset
        objMatrix = objForm.Items.Item("4").Specific

        For i = 1 To objMatrix.RowCount()
            If (objMatrix.IsRowSelected(i)) Then
                customername = objMatrix.Columns.Item("2").Cells.Item(i).Specific.string
                contractno = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                Exit For
            End If
        Next

        objMatrix = objForm.Items.Item("5").Specific
        strSQL = "delete [@MIPLCNTC] where U_cntrid=" & contractno
        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRS.DoQuery(strSQL)
        objRS = Nothing
        For intLoop = 1 To objMatrix.RowCount
            code = CInt(objAddOn.objGenFunc.GetCode("[@MIPLCNTC]"))
            techname = objMatrix.Columns.Item("2").Cells.Item(intLoop).Specific.string
            If Not techname = String.Empty Then
                techcode = CInt(objMatrix.Columns.Item("1").Cells.Item(intLoop).Specific.string)
                strSQL = "Insert into [@MIPLCNTC] (code,Name,U_custname,U_cntrid,U_techcode,U_techname) values (" & code & " ," & code & ",'" & customername & "' ," & contractno & "," & techcode & ",'" & techname & "')"
                objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS.DoQuery(strSQL)
                objRS = Nothing
            End If
        Next


    End Sub
    Private Function TechExist(ByVal FormUID As String, ByVal ContractID As Integer) As Boolean
        strSQL = "Select * from [@MIPLCNTC] where U_cntrid=" & ContractID
        objRs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objRs.DoQuery(strSQL)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("5").Specific
        objMatrix.Clear()
        If Not objRs.EoF Then
            While Not objRs.EoF
                AddRow(FormUID)
                objMatrix.Columns.Item("2").Cells.Item(objMatrix.RowCount).Specific.string = objRs.Fields.Item("U_techname").Value
                objRs.MoveNext()
            End While
        Else
            AddRow(FormUID)
            Return (False)
        End If
        AddRow(FormUID)

        Return True
    End Function

End Class
