Public Class clsServicecall
    Public Const formtype = "ServiceCall"
    Dim objform As SAPbouiCOM.Form
    Dim strsql, strsql1 As String
    Dim objrs, objrs1 As SAPbobsCOM.Recordset
    Public technician As String
    Public username As String
    Public password As String
    Public senderid As String
    Public priority As Integer
    Dim retCode As Long
    Dim objcombo As SAPbouiCOM.ComboBox
    Dim objcombo1 As SAPbouiCOM.ComboBox
    Dim objitem As SAPbouiCOM.Item
    Dim objstat As SAPbouiCOM.StaticText

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.BeforeAction Then
                objform = objAddOn.objApplication.Forms.Item(FormUID)
            Else
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = "SMSADDR" Then
                            Dim objcombo As SAPbouiCOM.ComboBox
                            Dim str As String
                            Dim ObjNewForm As SAPbouiCOM.Form
                            ObjNewForm = objAddOn.objApplication.Forms.GetForm("60110", 1)
                            objcombo = ObjNewForm.Items.Item("SMSADDR").Specific
                            str = objcombo.Selected.Description
                            strsql = "select top 1 address,street,block,zipcode,city,country from crd1 where  address='" & objcombo.Selected.Description & "'"
                            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objrs.DoQuery(strsql)
                            If objrs.RecordCount > 0 Then
                                ObjNewForm = objAddOn.objApplication.Forms.GetForm(-(objform.Type), 1)
                                ObjNewForm.Items.Item("U_SMSDETAILS").Specific.string = objrs.Fields.Item("address").Value + "," + objrs.Fields.Item("block").Value + "," + objrs.Fields.Item("street").Value + "," + objrs.Fields.Item("city").Value + "," + objrs.Fields.Item("country").Value
                                objform.ActiveItem = 29
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        objform = objAddOn.objApplication.Forms.Item(FormUID)
                        createitems()
                    Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                        If pVal.ItemUID = "SMSADDR" Then
                            If objform.Items.Item("14").Specific.string <> "" Then
                                objform = objAddOn.objApplication.Forms.Item(FormUID)
                                objcombo = objform.Items.Item("SMSADDR").Specific
                                If objcombo.ValidValues.Count > 0 Then
                                    For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1
                                        objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                End If
                                strsql = "select  convert(int,row_Number() Over (order by cardcode)) as 'row', Address from CRD1  where CardCode='" & objform.Items.Item("14").Specific.string & "' and AdresType='S'"
                                objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                objrs.DoQuery(strsql)
                                While Not objrs.EoF
                                    For i As Integer = 1 To objrs.RecordCount
                                        If i <= objrs.RecordCount Then
                                            objcombo.ValidValues.Add(i, objrs.Fields.Item("Address").Value)
                                            objrs.MoveNext()
                                        End If
                                    Next
                                End While
                            End If
                        End If
                End Select
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


    Public Sub createitems()
        'objform = objAddOn.objApplication.Forms.Item(formuid)
        Dim OusrDS As SAPbouiCOM.UserDataSource
        OusrDS = objform.DataSources.UserDataSources.Add("addre", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        Dim objCmb As SAPbouiCOM.ComboBox
        objitem = objform.Items.Add("SMSLABEL", SAPbouiCOM.BoFormItemTypes.it_STATIC)
        objitem.Left = objform.Items.Item("106").Left
        objitem.Width = objform.Items.Item("106").Width
        objitem.Top = objform.Items.Item("106").Top + 15
        objitem.Height = objform.Items.Item("106").Height
        objstat = objform.Items.Item("SMSLABEL").Specific
        objstat.Caption = "SMS Address"

        objitem = objform.Items.Add("SMSADDR", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        objitem.Left = objform.Items.Item("107").Left
        objitem.Width = objform.Items.Item("107").Width
        objitem.Top = objform.Items.Item("107").Top + 15
        objitem.Height = objform.Items.Item("107").Height
        objCmb = objform.Items.Item("SMSADDR").Specific
        objform.Items.Item("SMSADDR").DisplayDesc = True
        objCmb.DataBind.SetBound(True, "OSCL", "U_Address")
        objCmb = objitem.Specific
    End Sub


    'Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
    '    If BusinessObjectInfo.BeforeAction Then
    '    Else
    '        If BusinessObjectInfo.ActionSuccess Then
    '            Dim objbutton As SAPbouiCOM.Button = objform.Items.Item("9").Specific
    '            MsgBox(objform.Items.Item("12").Specific.string)
    '            MsgBox(objform.Items.Item("14").Specific.string)
    '            MsgBox(objform.Items.Item("93").Specific.string)
    '        End If
    '    End If
    'End Sub
    Public Sub msgtotechnician()
        strsql = "select U_SMSSEND,(select lastname+', '+firstname from OHEM where empID=technician) "
        strsql += vbCrLf + " technician  from OSCL where callID ='" & objform.Items.Item("12").Specific.string & "'"
        objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objrs.DoQuery(strsql)
        If objrs.RecordCount = 0 Then
            Message()
        Else
            If objrs.Fields.Item("U_SMSSEND").Value <> "Y" Then
                Message()
            Else
                'MsgBox(technician)
                'MsgBox(objform.Items.Item("93").Specific.string)
                If technician = objform.Items.Item("93").Specific.string Then
                Else
                    Message()
                End If
            End If
        End If
    End Sub

    Public Sub Message()
        objform = objAddOn.objApplication.Forms.GetForm("60110", 1)
        techtocust()
        strsql = "select top 1 U_smstocust1,U_smstocust2,U_smstocust3,U_smstotech from [@miplsmscontrol]"
        objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objrs.DoQuery(strsql)
        If objrs.RecordCount > 0 Then
            If objrs.Fields.Item("U_smstotech").Value <> "Y" Then Exit Sub
        End If
        If objform.Items.Item("93").Specific.string <> "" Then
            Dim objService As SAPbobsCOM.ServiceCalls
            objService = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
            objService.GetByKey(objform.Items.Item("12").Specific.string)
            objService.UserFields.Fields.Item("U_SMSSEND").Value = "Y"
            retCode = objService.Update()
            If retCode Then
                objAddOn.objApplication.MessageBox(CStr(retCode) + "-" + objAddOn.objCompany.GetLastErrorDescription)
            End If
            getusernamepassword()
            If username = "" Then Exit Sub
            If password = "" Then Exit Sub
            If senderid = "" Then Exit Sub
            'MsgBox(message)
            strsql = "select mobile from ohem where lastname + ', ' + firstName='" & objform.Items.Item("93").Specific.string & "'"
            objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objrs.DoQuery(strsql)
            If objrs.RecordCount > 0 Then
                Dim mobileno As String = objrs.Fields.Item("mobile").Value
                If mobileno.Length <> 10 Then Exit Sub
                If Val(mobileno) = 0 Then Exit Sub
                Dim message As String = ""
                'message = "Service Call Created For You.Call No: " + objform.Items.Item("12").Specific.string + ". For Cust: " + objform.Items.Item("79").Specific.string
                message = "Call No: " + objform.Items.Item("12").Specific.string + ".For Cust: " + objform.Items.Item("79").Specific.string
                If objform.Items.Item("107").Specific.string = "" Then
                Else
                    message = " " & message & ".Contact No: " & objform.Items.Item("107").Specific.string & ""
                End If
                message = " " & message & ".Sr No:" & objform.Items.Item("85").Specific.string & ""
                objcombo1 = objform.Items.Item("44").Specific
                'objcombo1 objcombo1 = objform.Items.Item("44").Specific
                'MsgBox(objcombo1.Selected.Description)
                Try
                    'MsgBox(objcombo1.Selected.Value)
                    If (objcombo1.Selected Is Nothing) Then
                    Else
                        message = " " & message & ".Call Type:" & objcombo1.Selected.Description & ""
                    End If
                Catch ex As Exception
                    'MsgBox(ex.ToString)
                End Try

                message = " " & message & ".Subject:" & objform.Items.Item("6").Specific.string & ""

                objcombo = objform.Items.Item("SMSADDR").Specific
                ' MsgBox(objcombo.Selected.Description)
                Try

                    If objcombo.Selected Is Nothing Then
                        message = " " & message & ""
                    Else
                        strsql = "select Address,street,Block,city,ZipCode from crd1 where adrestype='S' and cardcode='" & objform.Items.Item("14").Specific.string & "' and address='" & objcombo.Selected.Description & "'"
                        objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objrs.DoQuery(strsql)

                        If objrs.RecordCount > 0 Then
                            Dim Address As String = ""
                            If objrs.Fields.Item("Address").Value <> "" Then Address = objrs.Fields.Item("Address").Value
                            If objrs.Fields.Item("Block").Value <> "" Then Address = "" & Address & "," & objrs.Fields.Item("Block").Value & ""
                            If objrs.Fields.Item("street").Value <> "" Then Address = "" & Address & "," & objrs.Fields.Item("street").Value
                            If objrs.Fields.Item("city").Value <> "" Then Address = "" & Address & "," & objrs.Fields.Item("city").Value & ""
                            If objrs.Fields.Item("ZipCode").Value <> "" Then Address = "" & Address & "," & objrs.Fields.Item("ZipCode").Value & ""
                            If Address.Length > 0 Then
                                message = " " & message & ", Customer Address:" & Address & ""
                            End If
                        End If
                    End If
                Catch ex As Exception
                    'MsgBox(ex.ToString)
                End Try
                Dim objsms As New Servicesms.Service
                objsms.SendTextSMS(username, password, mobileno, message, senderid)
                Dim str As String = objAddOn.apicall("http://smsalertbox.com/api/sms.php?uid=" & username & "&pin=" & password & "&sender=" & senderid & "&route=" & priority & "&mobile=" & mobileno & "&message=" & message & "")
                objAddOn.objApplication.SetStatusBarMessage("Message Send Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                'MsgBox(message.Length)
                'MsgBox(username)
                'MsgBox(password)
                'MsgBox(mobileno)
                'MsgBox(message)
            End If
        End If
    End Sub
    Private Sub techtocust()
        strsql = "select top 1 U_smstocust1,U_smstocust2,U_smstocust3,U_smstotech from [@miplsmscontrol]"
        objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objrs.DoQuery(strsql)
        If objrs.RecordCount > 0 Then
            If objrs.Fields.Item("U_smstocust2").Value <> "Y" Then Exit Sub
        End If
        If objform.Items.Item("93").Specific.string = "" Then Exit Sub
        getusernamepassword()
        If username = "" Then Exit Sub
        If password = "" Then Exit Sub
        If senderid = "" Then Exit Sub
        If objform.Items.Item("107").Specific.string = "" Then Exit Sub
        Dim mobileno As String = objform.Items.Item("107").Specific.string
        If mobileno.Length <> 10 Then Exit Sub
        If Val(mobileno) = 0 Then Exit Sub
        Dim techtocusmessage As String
        techtocusmessage = "Call No: " & objform.Items.Item("12").Specific.string & "Engineer Name:" & objform.Items.Item("93").Specific.string & ""
        'MsgBox(username)
        'MsgBox(password)
        'MsgBox(techtocusmessage)
        'MsgBox(mobileno)
        Dim objsms As New Servicesms.Service
        objsms.SendTextSMS(username, password, mobileno, techtocusmessage, senderid)
        Dim str As String = objAddOn.apicall("http://smsalertbox.com/api/sms.php?uid=" & username & "&pin=" & password & "&sender=" & senderid & "&route=" & priority & "&mobile=" & mobileno & "&message=" & techtocusmessage & "")
        objAddOn.objApplication.SetStatusBarMessage("Message Send Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub

    Public Sub gettechnicianname()
        technician = ""
        strsql = "select (select lastname+', '+firstname from OHEM where empID=technician) "
        strsql += vbCrLf + " technician  from OSCL where callID ='" & objform.Items.Item("12").Specific.string & "'"
        objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objrs.DoQuery(strsql)
        If objrs.RecordCount > 0 Then
            technician = objrs.Fields.Item("technician").Value
        End If
    End Sub

    Public Sub getusernamepassword()
        strsql1 = "select U_Username,U_password,U_SENDERID,U_Priority from [@MIPLLOGIN] "
        objrs1 = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objrs1.DoQuery(strsql1)
        If objrs1.RecordCount > 0 Then
            username = objrs1.Fields.Item("U_Username").Value
            password = objrs1.Fields.Item("U_password").Value
            senderid = objrs1.Fields.Item("U_SENDERID").Value
            If objrs1.Fields.Item("U_Priority").Value = "P" Then
                priority = 1
            Else
                priority = 0
            End If
        End If
    End Sub
    Public Sub msgtocustomer()
        'strsql = "select MAX(callid) as callid from oscl"
        'objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'objrs.DoQuery(strsql)
        'If objrs.RecordCount > 0 Then
        '    If objform.Items.Item("12").Specific.string = objrs.Fields.Item("callid").Value Then
        '    Else
        '        Exit Sub
        '    End If
        'End If
        strsql = "select isnull(U_SMSCUS,0) as sms from oscl where callid='" & objform.Items.Item("12").Specific.string & "'"
        objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objrs.DoQuery(strsql)
        'MsgBox(objform.Items.Item("12").Specific.string)
        If objrs.RecordCount > 0 Then
            If objrs.Fields.Item("sms").Value < 1 Then
            Else
                Exit Sub
            End If
        End If
        getusernamepassword()
        If username = "" Then Exit Sub
        If password = "" Then Exit Sub
        If senderid = "" Then Exit Sub
        Dim mobileno As String = objform.Items.Item("107").Specific.string
        If mobileno.Length <> 10 Then Exit Sub
        If Val(mobileno) = 0 Then Exit Sub
        Dim customermessage As String = "Thanks For Calling Mukesh Infoserve.Your Complaint No is:" & objform.Items.Item("12").Specific.string & ""
        Dim objsms As New Servicesms.Service
        'MsgBox(username)
        'MsgBox(password)
        'MsgBox(mobileno)
        'MsgBox(customermessage)
        objsms.SendTextSMS(username, password, mobileno, customermessage, senderid)
        Dim str As String = objAddOn.apicall("http://smsalertbox.com/api/sms.php?uid=" & username & "&pin=" & password & "&sender=" & senderid & "&route=" & priority & "&mobile=" & mobileno & "&message=" & customermessage & "")
        objAddOn.objApplication.SetStatusBarMessage("Message Send Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
        servicecallupdate(objform.Items.Item("12").Specific.string, 1)
    End Sub

    Private Sub servicecallupdate(ByVal callno As String, ByVal udfvalue As String)
        Dim objService As SAPbobsCOM.ServiceCalls
        objService = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        objService.GetByKey(callno)
        objService.UserFields.Fields.Item("U_SMSCUS").Value = udfvalue
        retCode = objService.Update()
        If retCode Then
            objAddOn.objApplication.MessageBox(CStr(retCode) + "-" + objAddOn.objCompany.GetLastErrorDescription)
        End If
    End Sub
    Public Sub finalmsgtocustomer()
        getusernamepassword()
        If username = "" Then Exit Sub
        If password = "" Then Exit Sub
        If senderid = "" Then Exit Sub
        Dim mobileno As String = objform.Items.Item("107").Specific.string
        If mobileno.Length <> 10 Then Exit Sub
        If Val(mobileno) = 0 Then Exit Sub
        objcombo = objform.Items.Item("38").Specific
        If objcombo.Selected.Description <> "Closed" Then Exit Sub
        Dim finalmsg As String = "Your Call No:" & objform.Items.Item("12").Specific.string & " has been Successfully Closed.Thanks for Your Support."
        Dim objsms As New Servicesms.Service
        'MsgBox(username)
        'MsgBox(password)
        'MsgBox(mobileno)
        'MsgBox(finalmsg)
        objsms.SendTextSMS(username, password, mobileno, finalmsg, senderid)
        Dim str As String = objAddOn.apicall("http://smsalertbox.com/api/sms.php?uid=" & username & "&pin=" & password & "&sender=" & senderid & "&route=" & priority & "&mobile=" & mobileno & "&message=" & finalmsg & "")
        objAddOn.objApplication.SetStatusBarMessage("Message Send Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub
End Class


'For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1
'    If objcombo.ValidValues.Count > 1 Then
'        '  objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
'    End If
'Next