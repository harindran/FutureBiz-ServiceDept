Imports ServiceDept.Servicesms
Public Class clsInstall
    Public Const FormType = "InstallationAllocation"
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim objRs As SAPbobsCOM.Recordset
    Dim i As Integer
    Dim contractid As Integer
    Dim strSQL As String
    Dim flag As Boolean = False
    Dim previousRow As Integer = 0
    Dim currentRow As Integer
    Dim custAsc As Boolean = False
    Dim contractAsc As Boolean = False
    Dim invoiceAsc As Boolean = False
    Dim sortColumn As Integer = 0
    Dim ChangedRows(200) As Integer
    Dim callid As String
    Dim RowIndex As Integer = 0
    Dim strsql6 As String
    Dim objrs3 As SAPbobsCOM.Recordset
    Dim intloop As Integer

    Public Sub LoadScreen()
        Try
            objForm = objAddOn.objUIXml.LoadScreenXML("InstallationAllocation.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, FormType)
            objMatrix = objForm.Items.Item("4").Specific
            objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            Refresh(objForm.UniqueID)
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        Catch Ex As Exception
            MsgBox(Ex.ToString)
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        If pVal.BeforeAction Then

        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        ServiceAdd(FormUID)
                    ElseIf pVal.ItemUID = "7" Then
                        AddName(FormUID)
                    ElseIf pVal.ItemUID = "4" And pVal.ColUID = "0" Then
                        currentRow = pVal.Row
                        If previousRow <> 0 And previousRow <> currentRow And flag Then
                            If objAddOn.objApplication.MessageBox("Assign Technicians?", , "Yes", "No") = 1 Then
                                AddName(FormUID)
                            End If
                        End If
                        previousRow = currentRow
                        flag = False
                        ShowName(FormUID)
                    ElseIf pVal.ItemUID = "4" And (pVal.ColUID = "2" Or pVal.ColUID = "3" Or pVal.ColUID = "4") And pVal.Row = 0 Then
                        sortColumn = pVal.ColUID
                        Refresh(FormUID)
                    ElseIf pVal.ItemUID = "3" Then
                        objMatrix.Clear()
                        Refresh(FormUID)
                    ElseIf pVal.ItemUID = "4" And pVal.ColUID = "6" And pVal.Row <> 0 Then
                        objMatrix = objForm.Items.Item("4").Specific
                        If objMatrix.Columns.Item("6").Cells.Item(pVal.Row).Specific.checked = True Then
                            AddRowIndex(pVal.Row)
                        Else
                            RemoveRowIndex(pVal.Row)
                        End If
                        'ElseIf pVal.ItemUID = "4" And pVal.ColUID = "6" And pVal.Row = 0 Then
                        '    objMatrix = objForm.Items.Item("4").Specific
                        '    SelectAll(FormUID)

                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.ItemUID = "5" And pVal.ColUID = "2" Then
                        choose(FormUID, pVal)
                    End If


                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                    If pVal.InnerEvent = False Then
                        ResizeMatrix(FormUID)
                    End If


            End Select
        End If
    End Sub
    Private Sub SelectAll(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("4").Specific
        If objMatrix.RowCount > 0 Then
            If objMatrix.Columns.Item("6").Cells.Item(1).Specific.checked Then
                For intloop = 1 To objMatrix.RowCount
                    objForm.DataSources.DataTables.Item("T1").SetValue(6, intloop - 1, "N")
                Next
                ChangedRows.SetValue(0, 0)
                RowIndex = 0
            Else
                For intloop = 1 To objMatrix.RowCount
                    objForm.DataSources.DataTables.Item("T1").SetValue(6, intloop - 1, "Y")
                    AddRowIndex(intloop)
                    If intloop = 199 Then
                        objAddOn.objApplication.MessageBox("Only 200 calls can be selected")
                        Exit For
                    End If
                Next
            End If
            objMatrix.LoadFromDataSource()
        End If
    End Sub

    Private Sub AddRowIndex(ByVal Row As Integer)
        Try
            If RowIndex <= 199 Then
                ChangedRows.SetValue(Row, RowIndex)
                RowIndex = RowIndex + 1
            Else
                objAddOn.objApplication.MessageBox("Only 200 calls can be selected")
            End If
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(CStr(RowIndex) + ex.ToString)
        End Try
    End Sub
    Private Sub RemoveRowIndex(ByVal Row As Integer)
        For i = 0 To RowIndex
            If ChangedRows(i) = Row Then
                Exit For
            End If
        Next
        For i = i To RowIndex
            ChangedRows(i) = ChangedRows(i + 1)
        Next
        RowIndex = RowIndex - 1
    End Sub
    Private Sub ResizeMatrix(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try
            objForm.Freeze(True)
            objForm.Items.Item("5").Top = objForm.Items.Item("4").Top + objForm.Items.Item("4").Height + 10
            objForm.Items.Item("5").Height = objForm.ClientHeight - (objForm.Items.Item("4").Top + objForm.Items.Item("4").Height + 50)
            objForm.Items.Item("5").Width = 300
        Catch ex As Exception
        Finally
            objForm.Freeze(False)
        End Try
    End Sub

    Private Sub AddRow(ByVal FormUID As String)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("5").Specific
        If objMatrix.RowCount = 0 Then
            objMatrix.AddRow()
            objMatrix.Columns.Item("0").Cells.Item(objMatrix.RowCount).Specific.string = objMatrix.RowCount
        ElseIf objMatrix.RowCount > 1 Then
            If objMatrix.Columns.Item("2").Cells.Item(objMatrix.RowCount).Specific.string <> String.Empty Then
                objForm.DataSources.DBDataSources.Item("@MIPLCNTC").Clear()
                objMatrix.AddRow()
                objMatrix.Columns.Item("0").Cells.Item(objMatrix.RowCount).Specific.string = objMatrix.RowCount
            End If
        Else
            objForm.DataSources.DBDataSources.Item("@MIPLCNTC").Clear()
            objMatrix.AddRow()
            objMatrix.Columns.Item("0").Cells.Item(objMatrix.RowCount).Specific.string = objMatrix.RowCount
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
                '  objForm.DataSources.DBDataSources.Item("@MIPLCNTC").SetValue("U_sno", 0, pval.Row)
                objForm.DataSources.DBDataSources.Item("@MIPLCNTC").SetValue("U_techname", 0, objDataTable.GetValue("lastName", 0) + "," + objDataTable.GetValue("firstName", 0))
                objForm.DataSources.DBDataSources.Item("@MIPLCNTC").SetValue("U_techcode", 0, objDataTable.GetValue("empID", 0))
                objMatrix.SetLineData(pval.Row)
                AddRow(FormUID)
                flag = True
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub Refresh(ByVal FormUID As String)
        Dim strSQL As String = ""
        Dim intLoop As Integer = 0
        Dim objRecordSet As SAPbobsCOM.Recordset
        Dim odatatabe As SAPbouiCOM.DataTable
        Dim objDataTable As SAPbouiCOM.DataTable
        Dim ocons As SAPbouiCOM.Conditions
        Dim ocon As SAPbouiCOM.Condition
        Dim ocfl As SAPbouiCOM.ChooseFromList
        Dim strROWID As String = ""
        Dim i As Integer
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("4").Specific
        objMatrix.Clear()

        strSQL = "T0.cstmrcode,T0.cstmrname,T0.contractid,convert(varchar,T0.startdate,103) as startdate,COUNT (T1.itemcode) as Qty,'N' from OCTR T0 join CTR1 T1 on T1.ContractID =T0.ContractID  " & _
        " left outer join OSCL T2 on T2.contractID =T0.ContractID join OINS T3 on T3.insid=T1.insid where(T0.EndDate >= GETDATE() And T2.contractID Is null)  and T3.U_instal='N'" & _
        " group by T0.contractid,T0.cstmrname,T0.startdate,T0.cstmrcode"

        If (sortColumn = 0 Or sortColumn = 3) And contractAsc = False Then
            strROWID = "select convert(int,ROW_NUMBER() OVER (ORDER BY T0.contractid ASC)) AS ROWID,"
            strSQL += " Order by T0.Contractid asc"
            If sortColumn = 3 Then
                contractAsc = True
            End If
        ElseIf (sortColumn = 0 Or sortColumn = 3) And contractAsc = True Then
            strROWID = "select convert(int,ROW_NUMBER() OVER (ORDER BY T0.contractid DESC)) AS ROWID,"
            strSQL += " Order by T0.Contractid desc"
            contractAsc = False
        ElseIf sortColumn = 2 And custAsc = False Then
            strROWID = "select convert(int,ROW_NUMBER() OVER (ORDER BY T0.cstmrname ASC)) AS ROWID,"
            strSQL += " Order by T0.cstmrname asc"
            custAsc = True
        ElseIf sortColumn = 2 And custAsc = True Then
            strROWID = "select convert(int,ROW_NUMBER() OVER (ORDER BY T0.cstmrname DESC)) AS ROWID,"
            strSQL += " Order by T0.cstmrname desc"
            custAsc = False
        ElseIf sortColumn = 4 And invoiceAsc = False Then
            strROWID = "select convert(int,ROW_NUMBER() OVER (ORDER BY T0.startdate ASC)) AS ROWID,"
            strSQL += " Order by T0.startdate asc"
            invoiceAsc = True
        ElseIf sortColumn = 4 And invoiceAsc = True Then
            strROWID = "select convert(int,ROW_NUMBER() OVER (ORDER BY T0.startdate DESC)) AS ROWID,"
            strSQL += " Order by T0.startdate desc"
            invoiceAsc = False
        End If

        strSQL = strROWID + strSQL
        objMatrix.FlushToDataSource()
        objDataTable = objForm.DataSources.DataTables.Item("T1")
        objDataTable.ExecuteQuery(strSQL)
        objMatrix.LoadFromDataSource()
        RowIndex = 0

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
        objMatrix = objForm.Items.Item("5").Specific
        objMatrix.Clear()

    End Sub
    Private Sub ServiceAdd(ByVal FormUID As String)
        Dim strSQL As String
        Dim strSQL1 As String
        Dim strSQL2 As String
        Dim strSQL3 As String
        Dim strsql4, strsql5 As String
        Dim Origin As Integer
        Dim ProblemType As Integer
        Dim CallType As Integer
        Dim objrs1, objrs2 As SAPbobsCOM.Recordset
        Dim objRecordSet As SAPbobsCOM.Recordset
        Dim objService As SAPbobsCOM.ServiceCalls
        Dim objCEC As SAPbobsCOM.CustomerEquipmentCards
        Dim Technician As String = ""
        Dim oneFlag As Boolean
        Dim intloop As Integer
        Dim cardnum As Integer
        Dim retCode As Long
        callid = ""
        objService = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try
            strSQL1 = " select originID from OSCO Where Name='Billing'"
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL1)
            Origin = objRecordSet.Fields.Item("originID").Value
            objRecordSet = Nothing

            strSQL2 = " select prblmTypID  from OSCP where Name='Installation'"
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL2)
            ProblemType = objRecordSet.Fields.Item("prblmTypID").Value
            objRecordSet = Nothing

            strSQL3 = " select callTypeID  from OSCT Where Name='Installation'"
            objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRecordSet.DoQuery(strSQL3)
            CallType = objRecordSet.Fields.Item("callTypeID").Value
            objRecordSet = Nothing
            objMatrix = objForm.Items.Item("4").Specific
            oneFlag = False
            '----------
            If RowIndex > 0 Then
                RowIndex = RowIndex - 1
            End If
            '----------
            ' For intloop = 1 To objMatrix.RowCount
            For i = 0 To RowIndex
                intloop = ChangedRows(i)
                If intloop <> 0 Then
                    If objMatrix.Columns.Item("6").Cells.Item(intloop).Specific.checked = True Then
                        oneFlag = True
                        objRs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strSQL = "select U_techcode from [@MIPLCNTC] where U_cntrid=" & objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string
                        objRs.DoQuery(strSQL)
                        If objRs.RecordCount = 0 Then
                            objAddOn.objApplication.MessageBox("Technician has not been assigned for Contract No:" & objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string)
                        End If
                        If objRs.RecordCount = 1 Then
                            objService.TechnicianCode = CInt(objRs.Fields.Item("U_techcode").Value)
                        End If
                        objRs = Nothing
                        strSQL = "Select Insid,InternalSN,ItemCode from CTR1 Where ContractID= " & objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.String & ""
                        objRecordSet = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRecordSet.DoQuery(strSQL)
                        If Not objRecordSet.EoF Then
                            While Not objRecordSet.EoF
                                objService.InternalSerialNum = objRecordSet.Fields.Item("InternalSN").Value
                                objService.ItemCode = objRecordSet.Fields.Item("ItemCode").Value
                                objService.CustomerCode = objMatrix.Columns.Item("1").Cells.Item(intloop).Specific.string
                                objService.CustomerName = objMatrix.Columns.Item("2").Cells.Item(intloop).Specific.string
                                objService.ContractID = objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string
                                objService.Subject = "Installation"
                                objService.Origin = Origin
                                objService.ProblemType = ProblemType
                                objService.CallType = CallType

                                retCode = objService.Add()
                                If retCode Then
                                    objAddOn.objApplication.MessageBox(retCode + "-" + objAddOn.objCompany.GetLastErrorDescription)
                                End If
                                objAddOn.objApplication.StatusBar.SetText("Call created", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                strsql4 = " SELECT callID,contractID,custmrName   FROM OSCL where callID =(SELECT mAX(callID) FROM OSCL WHERE subject ='Installation')"
                                objrs1 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objrs1.DoQuery(strsql4)
                                If objrs1.RecordCount > 0 Then
                                    If callid = "" Then
                                        callid = objrs1.Fields.Item("callID").Value.ToString()
                                    Else
                                        callid = callid + "," + objrs1.Fields.Item("callID").Value.ToString()
                                    End If
                                End If
                                ' update UDF U_instal with YES
                                objCEC = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCustomerEquipmentCards)
                                cardnum = CInt(objRecordSet.Fields.Item("Insid").Value)
                                If objCEC.GetByKey(cardnum) Then
                                    objCEC.UserFields.Fields.Item("U_INSTAL").Value = "Y"
                                    retCode = objCEC.Update
                                    If retCode Then
                                        objAddOn.objApplication.MessageBox(CStr(retCode) + "-" + objAddOn.objCompany.GetLastErrorDescription)
                                    End If

                                End If
                                objRecordSet.MoveNext()
                            End While
                            'MsgBox(callid)
                        End If
                    End If
                End If
                strsql4 = " SELECT callID,contractID,custmrName   FROM OSCL where callID =(SELECT mAX(callID) FROM OSCL WHERE subject ='Installation')"
                objrs1 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs1.DoQuery(strsql4)
                If objrs1.RecordCount > 0 Then
                    '       MsgBox("Call No: '" & objrs1.Fields.Item("callID").Value & "'  For Cust: '" & objrs1.Fields.Item("custmrName").Value & "'  Contact No:")
                    'Dim msg As String = "Call No:('" & callid & "')For Cust: '" & objrs1.Fields.Item("custmrName").Value & "'  Contact No:"
                    Dim msg As String = "Service Call Created For You.Call No:('" & callid & "').For Cust: '" & objrs1.Fields.Item("custmrName").Value & "'"
                    'MsgBox(msg)
                    strsql5 = "select B.mobile mobile from OHEM B join [@MIPLCNTC] A "
                    strsql5 += vbCrLf + " on A.U_techcode=B.empid where A.U_cntrid='" & objrs1.Fields.Item("contractID").Value & "'"
                    objrs2 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objrs2.DoQuery(strsql5)
                    While Not objrs2.EoF
                        Dim mobileno As String = objrs2.Fields.Item("mobile").Value
                        Dim message As Boolean = False
                        If mobileno.Length <> 10 Then
                            objAddOn.objApplication.SetStatusBarMessage("Invalid MobileNo", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        ElseIf Val(mobileno) = 0 Then
                            objAddOn.objApplication.SetStatusBarMessage("Invalid MobileNo", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        Else
                            message = True
                            objAddOn.objApplication.SetStatusBarMessage("Message Send Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        End If
                        '  MsgBox(objrs2.Fields.Item("mobile").Value)
                        Dim objsms As New Servicesms.Service
                        strsql6 = "select U_Username,U_password,U_senderid,U_Priority from [@MIPLLOGIN] "
                        objrs3 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objrs3.DoQuery(strsql6)
                        If objrs3.RecordCount > 0 Then
                            Dim username As String = objrs3.Fields.Item("U_Username").Value
                            Dim password As String = objrs3.Fields.Item("U_password").Value
                            Dim senderid As String = objrs3.Fields.Item("U_senderid").Value
                            Dim priority As Integer
                            If objrs3.Fields.Item("U_Priority").Value = "P" Then
                                priority = 1
                            Else
                                priority = 0
                            End If
                            If message = True Then
                                'MsgBox(username)
                                'MsgBox(password)
                                'MsgBox(mobileno)
                                'MsgBox(msg)
                                objsms.SendTextSMS(username, password, mobileno, msg, "SMSCntry")
                                Dim str As String = objAddOn.apicall("http://smsalertbox.com/api/sms.php?uid=" & username & "&pin=" & password & "&sender=" & senderid & "&route=" & priority & "mobile=" & mobileno & "&message=" & msg & "")

                            End If
                        End If
                        'If objsms.SendTextSMS("maharaja11", "raja123", objrs2.Fields.Item("mobile").Value, msg, "5678") Then
                        ' End If
                        objrs2.MoveNext()
                    End While
                End If
            Next
        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(objAddOn.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short)
        End Try
        If oneFlag = False Then
            objAddOn.objApplication.MessageBox("No contract selected for installation")
            Exit Sub
        End If
        Refresh(FormUID)

    End Sub
    Private Function getContractNo(ByVal FormUID As String) As Integer
        Dim contractno As Integer = 0
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("4").Specific
        For i = 1 To objMatrix.RowCount()
            If (objMatrix.IsRowSelected(i)) Then
                contractno = objMatrix.Columns.Item("3").Cells.Item(i).Specific.string
                Exit For
            End If
        Next
        Return contractno
    End Function
    Private Sub ShowName(ByVal FormUID As String)
        Dim contractno As Integer
        contractno = getContractNo(FormUID)
        If contractno <> 0 Then
            TechExist(FormUID, CInt(contractno))
        End If
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
                strSQL = "Insert into [@MIPLCNTC] (code,Name,U_sno,U_custname,U_cntrid,U_techcode,U_techname) values (" & code & " ," & code & "," & intLoop & ",'" & customername & "' ," & contractno & "," & techcode & ",'" & techname & "')"
                objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRS.DoQuery(strSQL)
                objRS = Nothing
            End If
        Next
        flag = False
        previousRow = currentRow

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
                objMatrix.GetLineData(objMatrix.RowCount)
                objForm.DataSources.DBDataSources.Item("@MIPLCNTC").SetValue("U_techname", 0, objRs.Fields.Item("U_techname").Value)
                objMatrix.SetLineData(objMatrix.RowCount)
                objRs.MoveNext()
            End While
            flag = False
        Else
            AddRow(FormUID)
            Return (False)
        End If
        AddRow(FormUID)
        Return True
    End Function

End Class
