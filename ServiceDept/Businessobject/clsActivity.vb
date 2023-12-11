Public Class clsActivity
    Dim objForm As SAPbouiCOM.Form
    Dim ServiceID As Integer
    Dim ActivityCode As Integer
    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        If pVal.BeforeAction Then
        Else
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    createEdit(FormUID)
            End Select
        End If

    End Sub
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        If BusinessObjectInfo.BeforeAction Then
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    Try
                        ServiceID = CInt(objForm.DataSources.DBDataSources.Item("OCLG").GetValue("U_servid", 0))
                    Catch ex As Exception
                        ServiceID = 0
                    End Try
            End Select
        Else
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.ActionSuccess Then
                        ActivityCode = getActivityID(BusinessObjectInfo)
                        If ActivityCode <> 0 Then
                            AddActivity(ActivityCode)
                        Else
                            objAddOn.objApplication.MessageBox("No Activity")
                        End If
                    End If
            End Select
        End If
    End Sub
    Private Sub createEdit(ByVal FormUID As String)
        Dim objItem As SAPbouiCOM.Item
        Dim objEdit As SAPbouiCOM.EditText
        objForm = objAddOn.objApplication.Forms.Item(FormUID)
        objItem = objForm.Items.Add("SID", SAPbouiCOM.BoFormItemTypes.it_EDIT)
        With objItem
            .Left = 100
            .Top = 200
            .Width = 80
            .Height = 15
            .Visible = False
        End With
        objEdit = objItem.Specific
        objEdit.DataBind.SetBound(True, "OCLG", "U_servid")

    End Sub
    Private Sub AddActivity(ByVal ActivityID As Integer)
        Dim objService As SAPbobsCOM.ServiceCalls
        Dim objActivity As SAPbobsCOM.Contacts
        objActivity = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oContacts)
        objService = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oServiceCalls)
        ' objAddOn.objApplication.SetStatusBarMessage(ServiceID)
        If objService.GetByKey(ServiceID) Then
            If objService.Activities.ActivityCode = 0 Then
                objService.Activities.Add()
                objService.Activities.SetCurrentLine(0)
            Else
                objService.Activities.Add()
            End If
            objService.Activities.ActivityCode = ActivityID
            If objService.Update Then
                objAddOn.objApplication.MessageBox(objAddOn.objCompany.GetLastErrorDescription & CStr(ActivityID))
            End If
        End If

    End Sub
    Private Function getActivityID(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo) As Integer
        Dim xmlString As String = BusinessObjectInfo.ObjectKey
        Dim sr As New System.IO.StringReader(xmlString)
        Dim doc As New Xml.XmlDocument
        doc.Load(sr)
        'or just in this case doc.LoadXML(xmlString)
        Dim reader As New Xml.XmlNodeReader(doc)
        While reader.Read()
            Select Case reader.NodeType
                Case Xml.XmlNodeType.Element
                    If reader.Name = "ContactCode" Then
                        Return reader.ReadElementString
                    End If
            End Select
        End While
        Return 0
    End Function
End Class
