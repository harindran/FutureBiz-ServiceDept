Public Class clscloseinstalll
    Public Const FormType = "CloseInstall"
    Dim objForm As SAPbouiCOM.Form
    Dim objMatrix As SAPbouiCOM.Matrix
    Public objRS As SAPbobsCOM.Recordset
    Dim objCombo As SAPbouiCOM.ComboBox
    Dim strSQL As String
    Dim intLoop As Integer
    Dim selectedRow As Integer
    Dim contractNos As String = ""
    Dim startIndex As Integer = 0
    Dim length As Integer = 0
    Dim custAsc As Boolean = False
    Dim contractAsc As Boolean = False
    Dim callAsc As Boolean = False
    Dim sortColumn As Integer = 0
    Dim ChangedRows(200) As Integer
    Dim RowIndex As Integer = 0
    Dim i As Integer

    Public Sub LoadScreen()
        Try
            objForm = objAddOn.objUIXml.LoadScreenXML("CloseInstall.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded, FormType)
            objMatrix = objForm.Items.Item("3").Specific
            objMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            strSQL = "select statusID,Name from OSCS"
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(strSQL)
            While Not objRS.EoF
                objMatrix.Columns.Item("7").ValidValues.Add(objRS.Fields.Item("statusID").Value, objRS.Fields.Item("Name").Value)
                objRS.MoveNext()
            End While
            Servicecalls(objForm.UniqueID)

            strSQL = "SELECT CallTypeID,Name from OSCT"
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(strSQL)
            While Not objRS.EoF
                objMatrix.Columns.Item("CallType").ValidValues.Add(objRS.Fields.Item("CallTypeID").Value, objRS.Fields.Item("Name").Value)
                objRS.MoveNext()
            End While

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        Catch Ex As Exception
            MsgBox(Ex.ToString)
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        If pVal.BeforeAction Then
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        Try
                            If contractNos <> "" Then
                                contractNos = contractNos.Remove(contractNos.LastIndexOf(", "), 2)
                                If objAddOn.objApplication.MessageBox("ContractID(s) : " & contractNos, , "Yes", "No") <> 1 Then
                                    BubbleEvent = False
                                End If
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.ItemUID = "3" And pVal.ColUID = "6" Then
                        setTechnicians(objForm.UniqueID, pVal.Row, "Select U_techcode  ,U_techname from [@MIPLCNTC] where U_cntrid=" & CInt(objMatrix.Columns.Item("2").Cells.Item(pVal.Row).Specific.string), "@MIPLCNTC")
                    End If
            End Select
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CLICK

                    If pVal.ItemUID = "1" And objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        serviceclose(FormUID)
                    ElseIf pVal.ItemUID = "13" Then
                        Servicecalls(FormUID)
                    ElseIf pVal.ItemUID = "3" And pVal.ColUID = "9" And pVal.Row <> 0 Then
                        changeStatus(FormUID, pVal.Row)
                    ElseIf pVal.ItemUID = "3" And pVal.ColUID = "9" And pVal.Row = 0 Then
                        SelectAll(FormUID)
                    ElseIf pVal.ItemUID = "3" And pVal.ColUID = "0" Then
                        selectedRow = pVal.Row
                    ElseIf pVal.ItemUID = "14" Then
                        NewActivity(FormUID)
                    ElseIf pVal.ItemUID = "3" And (pVal.ColUID = "1" Or pVal.ColUID = "2" Or pVal.ColUID = "3") And pVal.Row = 0 Then
                        sortColumn = pVal.ColUID
                        Servicecalls(FormUID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    If pVal.ItemUID = "7" Or pVal.ItemUID = "9" Or pVal.ItemUID = "5" Or (pVal.ItemUID = "3" And pVal.ColUID = "6") Then
                        choose2(FormUID, pVal)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                    If pVal.ItemUID = "9" Then
                        BPContract(FormUID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    If pVal.ItemUID = "3" And pVal.ColUID = "7" Then ' status
                        AddRowIndex(pVal.Row)
                        objMatrix = objForm.Items.Item("3").Specific
                        objForm.DataSources.DataTables.Item("T2").SetValue(9, pVal.Row - 1, objMatrix.Columns.Item("7").Cells.Item(pVal.Row).Specific.selected.value)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    If pVal.ItemUID = "3" And pVal.ColUID = "8" Then ' remarks
                        objMatrix = objForm.Items.Item("3").Specific
                        objForm.DataSources.DataTables.Item("T2").SetValue(10, pVal.Row - 1, objMatrix.Columns.Item("8").Cells.Item(pVal.Row).Specific.string)
                    End If
            End Select

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

    Private Sub BPContract(ByVal FormUID As String)
        Dim objDataTable As SAPbouiCOM.DataTable
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim i As Integer
        objForm = objAddOn.objApplication.Forms.Item(FormUID)

        oCFL = objForm.ChooseFromLists.Item("CFL_2")
        oCFL.SetConditions(Nothing)
        If oCFL.GetConditions.Count = 0 Then
            strSQL = "select contractid from oscl where OSCL.status <> '-1'AND contractID <>0  and problemTyp=1 "
            If objForm.Items.Item("5").Specific.string <> String.Empty Then
                strSQL += "  and custmrName='" & objForm.Items.Item("5").Specific.string & "'"
            End If
            Try
                objDataTable = objForm.DataSources.DataTables.Item("OCTR")
            Catch ex As Exception
                objDataTable = objForm.DataSources.DataTables.Add("OCTR")
            End Try
            objDataTable.ExecuteQuery(strSQL)
            oCons = oCFL.GetConditions

            For i = 0 To objDataTable.Rows.Count - 1
                oCon = oCons.Add()
                oCon.Alias = "contractID"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = CInt(objDataTable.GetValue(0, i))
                If i < objDataTable.Rows.Count - 1 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            Next
            oCFL.SetConditions(oCons)
        End If

    End Sub

    Private Sub NewActivity(ByVal FormUID As String)
        Dim ActivityForm As SAPbouiCOM.Form
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("3").Specific
        If selectedRow > 0 Then
            objAddOn.objApplication.Menus.Item("2563").Activate()
            ActivityForm = objAddOn.objApplication.Forms.ActiveForm
            If ActivityForm.TypeEx = "651" Then
                ActivityForm.Items.Item("67").Specific.select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                ActivityForm.Items.Item("9").Specific.string = objMatrix.Columns.Item("C1").Cells.Item(selectedRow).Specific.string
                ActivityForm.Items.Item("SID").Specific.string = objMatrix.Columns.Item("3").Cells.Item(selectedRow).Specific.string
                objAddOn.objApplication.SendKeys("{TAB}")
            End If
        End If
    End Sub

    Private Sub SelectAll(ByVal FormUID As String)

        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("3").Specific
        If objMatrix.RowCount > 0 Then
            If objMatrix.Columns.Item("9").Cells.Item(1).Specific.checked Then
                For intLoop = 1 To objMatrix.RowCount
                    Try
                        'objMatrix.Columns.Item("9").Cells.Item(intLoop).Specific.checked = False
                        objForm.DataSources.DataTables.Item("T2").SetValue(11, intLoop - 1, "N")
                        strSQL = "Select status,descrption from OSCL where callid=" & CInt(objMatrix.Columns.Item("3").Cells.Item(intLoop).Specific.string)
                        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRS.DoQuery(strSQL)
                        objForm.DataSources.DataTables.Item("T2").SetValue(9, intLoop - 1, CStr(objRS.Fields.Item("status").Value))
                        objForm.DataSources.DataTables.Item("T2").SetValue(10, intLoop - 1, objRS.Fields.Item("descrption").Value)
                        objRS = Nothing
                        If contractNos.Contains(CStr(objMatrix.Columns.Item("2").Cells.Item(intLoop).Specific.string)) Then
                            length = Len(CStr(objMatrix.Columns.Item("2").Cells.Item(intLoop).Specific.string)) + 2
                            startIndex = contractNos.IndexOf(CStr(objMatrix.Columns.Item("2").Cells.Item(intLoop).Specific.string))
                            contractNos = contractNos.Remove(startIndex, length)
                            ' remove row index from changedRows
                            ChangedRows.SetValue(0, 0)
                            RowIndex = 0
                        End If
                    Catch ex As Exception
                    End Try
                Next
            Else
                For intLoop = 1 To objMatrix.RowCount
                    ' objMatrix.Columns.Item("9").Cells.Item(intLoop).Specific.checked = True
                    objForm.DataSources.DataTables.Item("T2").SetValue(11, intLoop - 1, "Y")
                    objForm.DataSources.DataTables.Item("T2").SetValue(9, intLoop - 1, "-1")
                    objForm.DataSources.DataTables.Item("T2").SetValue(10, intLoop - 1, "Installed")
                    If contractNos.Contains(CStr(objMatrix.Columns.Item("2").Cells.Item(intLoop).Specific.string)) Then
                    Else
                        contractNos = contractNos + CStr(objMatrix.Columns.Item("2").Cells.Item(intLoop).Specific.string) + ", "
                    End If
                    'Add row index in ChangedRows
                    If intLoop = 199 Then
                        objAddOn.objApplication.MessageBox("Only 200 calls can be selected")
                        Exit For
                    End If
                    AddRowIndex(intLoop)
                Next

            End If
            objMatrix.LoadFromDataSource()

        End If
    End Sub

    Private Sub changeStatus(ByVal FormUID As String, ByVal Row As Integer)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("3").Specific
        If objMatrix.Columns.Item("9").Cells.Item(Row).Specific.checked Then
            objForm.DataSources.DataTables.Item("T2").SetValue(11, Row - 1, "Y")
            objForm.DataSources.DataTables.Item("T2").SetValue(9, Row - 1, "-1")
            objForm.DataSources.DataTables.Item("T2").SetValue(10, Row - 1, "Installed")
            objForm.Freeze(True)
            objMatrix.LoadFromDataSource()
            objForm.Freeze(False)
            If contractNos.Contains(CStr(objMatrix.Columns.Item("2").Cells.Item(Row).Specific.string)) Then
            Else
                contractNos = contractNos + CStr(objMatrix.Columns.Item("2").Cells.Item(Row).Specific.string) + ", "
            End If
            'Add rowindex to changedrows
            AddRowIndex(Row)
        Else
            strSQL = "Select status,descrption from OSCL where callid=" & CInt(objMatrix.Columns.Item("3").Cells.Item(Row).Specific.string)
            objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery(strSQL)
            objForm.DataSources.DataTables.Item("T2").SetValue(11, Row - 1, "N")
            objForm.DataSources.DataTables.Item("T2").SetValue(9, Row - 1, CStr(objRS.Fields.Item("status").Value))
            objForm.DataSources.DataTables.Item("T2").SetValue(10, Row - 1, objRS.Fields.Item("descrption").Value)
            objForm.Freeze(True)
            objMatrix.LoadFromDataSource()
            objForm.Freeze(False)
            objRS = Nothing
            Try
                If contractNos.Contains(CStr(objMatrix.Columns.Item("2").Cells.Item(Row).Specific.string)) Then
                    length = Len(CStr(objMatrix.Columns.Item("2").Cells.Item(Row).Specific.string)) + 2
                    startIndex = contractNos.IndexOf(CStr(objMatrix.Columns.Item("2").Cells.Item(Row).Specific.string))
                    contractNos = contractNos.Remove(startIndex, length)
                End If
                'remove from changedRows
                RemoveRowIndex(Row)
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub choose2(ByVal FormUID As String, ByVal pval As SAPbouiCOM.ItemEvent)
        Dim objCFLEvent As SAPbouiCOM.ChooseFromListEvent
        Dim objDataTable As SAPbouiCOM.DataTable
        Dim oDatatable As SAPbouiCOM.DataTable
        objCFLEvent = pval
        objDataTable = objCFLEvent.SelectedObjects
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Try
            If objDataTable Is Nothing Then
            Else

                Select Case objCFLEvent.ChooseFromListUID
                    Case "CFL_3"
                        objForm.DataSources.UserDataSources.Item("U13").ValueEx = objDataTable.GetValue("internalSN", 0)
                    Case "CFL_2"
                        objForm.DataSources.UserDataSources.Item("U14").ValueEx = objDataTable.GetValue("ContractID", 0)
                    Case "CFL_4"
                        objForm.DataSources.UserDataSources.Item("U15").ValueEx = objDataTable.GetValue("CardName", 0)
                    Case "CFL_5"
                        objMatrix = objForm.Items.Item("3").Specific
                        'objMatrix.GetLineData(pval.Row)
                        oDatatable = objForm.DataSources.DataTables.Item("T2")
                        oDatatable.SetValue(7, pval.Row - 1, objDataTable.GetValue("empID", 0))
                        oDatatable.SetValue(8, pval.Row - 1, objDataTable.GetValue("lastName", 0) + "," + objDataTable.GetValue("firstName", 0))
                        objMatrix.LoadFromDataSource()
                End Select
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub setTechnicians(ByVal FormUID As String, ByVal Row As Integer, ByVal Qry As String, ByVal Tablename As String)
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oDataTable As SAPbouiCOM.DataTable
        Dim i As Integer
        oCFL = objForm.ChooseFromLists.Item("CFL_5")

        If oCFL.GetConditions.Count > 0 Then
            oCFL.SetConditions(Nothing)
        End If

        Try
            oDataTable = objForm.DataSources.DataTables.Item(Tablename)
        Catch ex As Exception
            oDataTable = objForm.DataSources.DataTables.Add(Tablename)
        End Try
        oDataTable.ExecuteQuery(Qry)
        If oDataTable.Rows.Count > 0 Then
            Try
                oDataTable = objForm.DataSources.DataTables.Item("HEM6")
            Catch ex As Exception
                oDataTable = objForm.DataSources.DataTables.Add("HEM6")
            End Try
            Qry = "select empid from HEM6 where roleid=-2"
            oDataTable.ExecuteQuery(Qry)
        End If

        oCons = oCFL.GetConditions
        For i = 0 To oDataTable.Rows.Count - 1
            oCon = oCons.Add()
            oCon.Alias = "empId"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = CInt(oDataTable.GetValue(0, i))
            If i < oDataTable.Rows.Count - 1 Then oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
        Next
        oCFL.SetConditions(oCons)
        'End If
    End Sub

    Private Sub Servicecalls(ByVal FormUID As String)

        Dim intLoop As Integer = 0
        Dim status As Integer
        Dim customername As String
        Dim contractID As Integer = 0
        Dim objDataTable As SAPbouiCOM.DataTable
        Dim strRowID As String = ""
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        Dim SerialNo As String = ""
        objMatrix = objForm.Items.Item("3").Specific
        objMatrix.FlushToDataSource()
        customername = objForm.Items.Item("5").Specific.string
        SerialNo = objForm.Items.Item("7").Specific.string
        If objForm.Items.Item("9").Specific.string <> String.Empty Then
            contractID = CInt(objForm.Items.Item("9").Specific.string)
        End If
        objMatrix.Clear()
        strSQl = ""
        'strSQl = "select custmrName ,customer ,contractID ,callID,itemName ,internalSN,OSCL.status,technician, OHEM.lastName + char(44) + OHEM.firstName as TechName From OSCL left outer join OHEM on OSCL.technician = OHEM.empID  where OSCL.status <> '-1'AND contractID <>0  and OSCL.problemTyp=1"
        strSQL = "customer,custmrName  ,contractID ,callID,itemName ,internalSN,technician, OHEM.lastName + char(44) + OHEM.firstName as TechName,OSCL.status,descrption,'N',OSCL.createDate,OSCL.closeDate  From OSCL" &
                " left outer join OHEM on OSCL.technician = OHEM.empID  " &
                " where OSCL.status <> '-1'AND contractID <>0  and problemTyp=(SELECT prblmTypID  FROM OSCP WHERE Name='Installation')"
        If contractID = 0 And customername <> "" Then
            strSQL += " and custmrName='" & customername & "' "
        ElseIf customername = "" And contractID <> 0 Then
            strSQL += " AND contractID=" & contractID
        ElseIf customername <> "" And contractID <> 0 Then
            strSQL += "AND contractID=" & contractID & " AND custmrName='" & customername & "'"
        ElseIf SerialNo <> "" Then
            strSQL += " and internalSN='" & SerialNo & "'"
        End If
        If (sortColumn = 0 Or sortColumn = 2) And contractAsc = False Then
            strRowID = "select convert(int,ROW_NUMBER() over (order by contractID asc)) as Rowid,"
            strSQL += " Order by Contractid asc"
            contractAsc = True
        ElseIf (sortColumn = 0 Or sortColumn = 2) And contractAsc = True Then
            strRowID = "select convert(int,ROW_NUMBER() over (order by contractID desc)) as Rowid,"
            strSQL += " Order by Contractid desc"
            contractAsc = False
        ElseIf sortColumn = 1 And custAsc = False Then
            strRowID = "select convert(int,ROW_NUMBER() over (order by custmrName asc)) as Rowid,"
            strSQL += " Order by custmrName asc"
            custAsc = True
        ElseIf sortColumn = 1 And custAsc = True Then
            strRowID = "select convert(int,ROW_NUMBER() over (order by custmrName desc)) as Rowid,"
            strSQL += " Order by custmrName desc"
            custAsc = False
        ElseIf sortColumn = 3 And callAsc = False Then
            strRowID = "select convert(int,ROW_NUMBER() over (order by callID asc)) as Rowid,"
            strSQL += " Order by callID asc"
            callAsc = True
        ElseIf sortColumn = 3 And callAsc = True Then
            strRowID = "select convert(int,ROW_NUMBER() over (order by callID desc)) as Rowid,"
            strSQL += " Order by callID desc"
            callAsc = False
        End If
        strSQL = strRowID + strSQL
        objDataTable = objForm.DataSources.DataTables.Item("T2")
        objDataTable.ExecuteQuery(strSQl)
        objForm.Freeze(True)
        objMatrix.LoadFromDataSource()
        objForm.Freeze(False)
        RowIndex = 0
        strSQL = ""

    End Sub

    Private Sub serviceclose(ByVal FormUID As String)
        Dim objservice As SAPbobsCOM.ServiceCalls
        Dim intloop As Integer
        Dim i As Integer
        objservice = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objMatrix = objForm.Items.Item("3").Specific
        ' For intloop = 1 To objMatrix.RowCount
        If RowIndex > 0 Then
            RowIndex = RowIndex - 1
        End If
        For i = 0 To RowIndex
            intloop = ChangedRows(i)
            If intloop <> 0 Then
                If objservice.GetByKey(objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string) Then
                    Try
                        objservice.Status = objMatrix.Columns.Item("7").Cells.Item(intloop).Specific.selected.value
                        If objservice.Status = -1 Then
                            objservice.Resolution = objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string
                        Else
                            objservice.Description = objMatrix.Columns.Item("8").Cells.Item(intloop).Specific.string 'remarks
                        End If
                    Catch Ex As Exception
                    End Try

                    If objMatrix.GetCellSpecific("CallType", intloop).Selected.Value.ToString.Trim <> "" Then _
                       objservice.CallType = objMatrix.GetCellSpecific("CallType", intloop).Selected.Value

                    If objMatrix.GetCellSpecific("CallType", intloop).Selected.Value.ToString.Trim <> "" Then _
                        objservice.Subject = objMatrix.GetCellSpecific("CallType", intloop).Selected.Description

                    If objMatrix.GetCellSpecific("CallType", intloop).Selected.Value.ToString.Trim <> "" Then
                        Dim strsql = "SELECT prblmTypID  FROM OSCP WHERE Name='" & objMatrix.GetCellSpecific("CallType", intloop).Selected.Description & "'"
                        objRS = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objRS.DoQuery(strsql)
                        If objRS.RecordCount <> 0 Then
                            Dim ProblemTypeDesc = objRS.Fields.Item("prblmTypID").Value
                            objservice.ProblemType = ProblemTypeDesc
                        End If
                    End If

                    If objMatrix.Columns.Item("6C").Cells.Item(intloop).Specific.string <> String.Empty Then
                        If CInt(objMatrix.Columns.Item("6C").Cells.Item(intloop).Specific.string) <> 0 Then
                            objservice.TechnicianCode = objMatrix.Columns.Item("6C").Cells.Item(intloop).Specific.string
                        End If
                        If (objservice.Update()) Then
                            objAddOn.objApplication.MessageBox(objAddOn.objCompany.GetLastErrorDescription)
                        Else
                            objAddOn.objApplication.StatusBar.SetText("Service call updated!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    Else
                        objAddOn.objApplication.MessageBox("Technician has not been assigned to " & objMatrix.Columns.Item("3").Cells.Item(intloop).Specific.string)
                    End If
                End If
            End If
        Next

        Servicecalls(FormUID)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
