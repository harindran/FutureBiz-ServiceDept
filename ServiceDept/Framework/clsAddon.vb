Imports System
Imports System.Collections.Generic
'using System.Linq
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Net
Imports System.IO
Imports SAPbouiCOM.Framework
Public Class clsAddOn
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Dim oProgBarx As SAPbouiCOM.ProgressBar
    Public objGenFunc As Mukesh.SBOLib.GeneralFunctions
    Public objUIXml As Mukesh.SBOLib.UIXML
    Public ZB_row As Integer = 0

    Public objinstall As clsInstall
    Public objCloseinstall As clscloseinstalll
    Public objservicecall As clsServicecall
    Public objActivity As clsActivity

    Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
    Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
    Dim ret As Long
    Dim str As String
    Dim objForm As SAPbouiCOM.Form
    Dim MenuCount As Integer = 0
    Dim oservice
    Dim oservice1
    Dim strsql As String
    Dim objrs As SAPbobsCOM.Recordset
    Public HWKEY() As String = New String() {"A0061802481", "L1653539483", "Q0319069806", "P1144285131", "X1211807750"}
    Private Sub CheckLicense()

    End Sub

    Function isValidLicense() As Boolean
        Try
            'If objApplication.Forms.ActiveForm.TypeCount > 0 Then
            '    For i As Integer = 0 To objApplication.Forms.ActiveForm.TypeCount - 1
            '        objApplication.Forms.ActiveForm.Close()
            '    Next
            'End If
            objApplication.Menus.Item("257").Activate()
            Dim CrrHWKEY As String = objApplication.Forms.ActiveForm.Items.Item("79").Specific.Value.ToString.Trim
            objApplication.Forms.ActiveForm.Close()
            For i As Integer = 0 To HWKEY.Length - 1
                If HWKEY(i).Trim = CrrHWKEY.Trim Then
                    Return True
                End If
            Next
            MsgBox("Add-on installation failed due to license mismatch", MsgBoxStyle.OkOnly, "License Management")
            Return False
        Catch ex As Exception
            objAddOn.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'MsgBox(ex.ToString)
        End Try
        Return True
    End Function

    Public Sub Intialize()
        Dim objSBOConnector As New Mukesh.SBOLib.SBOConnector
        objApplication = objSBOConnector.GetApplication(System.Environment.GetCommandLineArgs.GetValue(1))
        objCompany = objSBOConnector.GetCompany(objApplication)
        Try
            createTables()
            createObjects()
            loadMenu()
        Catch ex As Exception
            objAddOn.objApplication.MessageBox(ex.ToString)
            End
        End Try
        objApplication.SetStatusBarMessage("Service Add-On connected  successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, False)
    End Sub

    Public Sub Intialize(ByVal args() As String)
        Try
            Dim oapplication As Application
            If (args.Length < 1) Then oapplication = New Application Else oapplication = New Application(args(0))
            objapplication = Application.SBO_Application
            If isValidLicense() Then
                objApplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objCompany = Application.SBO_Application.Company.GetDICompany()
                Try
                    createObjects()
                    createTables()
                    'CreateUDOS()
                    loadMenu()
                Catch ex As Exception
                    objAddOn.objApplication.MessageBox(ex.ToString)
                    End
                End Try
                objApplication.StatusBar.SetText("Service Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                oapplication.Run()
            Else
                objApplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If

        Catch ex As Exception
            objAddOn.objApplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, True)
        End Try
    End Sub

    Private Sub UserTree()
    End Sub

    Private Sub CreateUDOS()
    End Sub

    Private Sub createObjects()
        'Library Object Initilisation
        objGenFunc = New Mukesh.SBOLib.GeneralFunctions(objCompany)
        objUIXml = New Mukesh.SBOLib.UIXML(objApplication)
        'Business Object Initialisation
        objinstall = New clsInstall
        objCloseinstall = New clscloseinstalll
        objservicecall = New clsServicecall
        objActivity = New clsActivity
    End Sub

    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try
            Select Case pVal.FormTypeEx

                Case clsInstall.FormType
                    objinstall.ItemEvent(FormUID, pVal, BubbleEvent)
                Case clscloseinstalll.FormType
                    objCloseinstall.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "651"
                    objActivity.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "60110"
                    objservicecall.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            MsgBox(ex.ToString)
            ' objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End Try
    End Sub

    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        Try
            If BusinessObjectInfo.FormTypeEx = "651" Then
                objActivity.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            ElseIf BusinessObjectInfo.FormTypeEx = "60110" And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = True And BusinessObjectInfo.ActionSuccess = False Then
                objservicecall.gettechnicianname()
            ElseIf BusinessObjectInfo.FormTypeEx = "60110" And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                strsql = "select top 1 U_smstocust1,U_smstocust2,U_smstocust3,U_smstotech from [@miplsmscontrol]"
                objrs = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then
                    If objrs.Fields.Item("U_smstocust1").Value = "Y" Then
                        objservicecall.msgtocustomer()
                    End If
                    If objrs.Fields.Item("U_smstocust3").Value = "Y" Then
                        objservicecall.finalmsgtocustomer()
                    End If
                    objservicecall.msgtotechnician()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Shared Function ErrorHandler(ByVal p_ex As Exception, ByVal objApplication As SAPbouiCOM.Application)
        Dim sMsg As String = Nothing
        If p_ex.Message = "Form - already exists [66000-11]" Then
            Return True
            Exit Function  'ignore error
        End If
        Return False
    End Function

    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        If pVal.BeforeAction Then
        Else

            Try
                Select Case pVal.MenuUID
                    Case clsInstall.FormType
                        objinstall.LoadScreen()
                    Case clscloseinstalll.FormType
                        objCloseinstall.LoadScreen()

                End Select

            Catch ex As Exception
                objAddOn.objApplication.StatusBar.SetText(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End Try
        End If
    End Sub

    Private Sub loadMenu()
        Dim objServiceMenu As SAPbouiCOM.MenuItem
        objServiceMenu = objApplication.Menus.Item("43520").SubMenus.Item("3584")
        If objServiceMenu.SubMenus.Exists("CloseInstall") Then Return
        MenuCount = objServiceMenu.SubMenus.Count
        CreateMenu("", MenuCount - 1, "Installation Allocation ", SAPbouiCOM.BoMenuType.mt_STRING, "InstallationAllocation", objServiceMenu)
        CreateMenu("", MenuCount, "Manage Installation Service Call", SAPbouiCOM.BoMenuType.mt_STRING, "CloseInstall", objServiceMenu)
    End Sub
   
    ' For Menu Creation
    Private Function CreateMenu(ByVal ImagePath As String, ByVal Position As Int32, ByVal DisplayName As String, ByVal MenuType As SAPbouiCOM.BoMenuType, ByVal UniqueID As String, ByVal ParentMenu As SAPbouiCOM.MenuItem) As SAPbouiCOM.MenuItem
        Try
            Dim oMenuPackage As SAPbouiCOM.MenuCreationParams
            oMenuPackage = objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            ' oMenuPackage.Image = ImagePath
            oMenuPackage.Position = Position
            oMenuPackage.Type = MenuType
            oMenuPackage.UniqueID = UniqueID
            oMenuPackage.String = DisplayName
            ParentMenu.SubMenus.AddEx(oMenuPackage)
        Catch ex As Exception
            objApplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
        End Try
        Return ParentMenu.SubMenus.Item(UniqueID)
    End Function

    Private Sub createTables()
        Dim objUDFEngine As New Mukesh.SBOLib.UDFEngine(objCompany)
        objAddOn.objApplication.SetStatusBarMessage("Creating Tables Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, False)
    

        objUDFEngine.CreateTable("MIPLCNTC", "Contract-Technicians", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.AddAlphaField("@MIPLCNTC", "custname", "Customer Name", 100)
        objUDFEngine.AddNumericField("@MIPLCNTC", "cntrid", "ContractID", 10)
        objUDFEngine.AddNumericField("@MIPLCNTC", "sno", "SerialNo", 10)
        objUDFEngine.AddNumericField("@MIPLCNTC", "techcode", "technician code", 10)
        objUDFEngine.AddAlphaField("@MIPLCNTC", "techname", "Technician Name", 120)

        objUDFEngine.AddNumericField("OCLG", "servid", "ServiceID", 10)

        objUDFEngine.CreateTable("MIPLLOGIN", "Login Information", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.AddAlphaField("@MIPLLOGIN", "USERNAME", "UserName", 100)
        objUDFEngine.AddAlphaField("@MIPLLOGIN", "PASSWORD", "Password", 100)
        objUDFEngine.AddAlphaField("@MIPLLOGIN", "SENDERID", "SENDERID", 100)
        objUDFEngine.addField("@MIPLLOGIN", "PRIORITY", "PRIORITY", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "N,P", "Normal,Priority", "N")

        ' objUDFEngine.AddAlphaField("OINS", "INSTAL", "INSTALL", 2)

        objUDFEngine.addField("OINS", "INSTAL", "INSTALL", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "N")

        objUDFEngine.AddAlphaField("OSCL", "SMSSEND", "SMS SEND", 2)
        objUDFEngine.AddAlphaField("OSCL", "SMSCUS", "SMS Customer", 2)
        objUDFEngine.AddAlphaField("OSCL", "SMSDETAILS", "SMSDetails", 250)
        objUDFEngine.AddAlphaField("OSCL", "Address", "SMSAddress", 250)

        objUDFEngine.CreateTable("MIPLSMSCONTROL", "SMS Control", SAPbobsCOM.BoUTBTableType.bott_NoObject)
        objUDFEngine.addField("MIPLSMSCONTROL", "SMSTOCUST1", "Call Created(Customer)", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.addField("MIPLSMSCONTROL", "SMSTOTech", "Tech Assigned(Tech)", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.addField("MIPLSMSCONTROL", "SMSTOCUST2", "Tech Assigned(Customer)", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")
        objUDFEngine.addField("MIPLSMSCONTROL", "SMSTOCUST3", "Call Closed(Customer)", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, "Y,N", "Yes,No", "Y")

        '*************************************************************************************************************************


    End Sub





    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        If eventInfo.BeforeAction Then
        End If
    End Sub

    Private Sub createUDO(ByVal tblname As String, ByVal udoname As String, ByVal childTable() As String, ByVal type As SAPbobsCOM.BoUDOObjType, Optional ByVal DfltForm As Boolean = False, Optional ByVal FindForm As Boolean = False)
        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim objform As SAPbouiCOM.Form
        Dim i As Integer
        'Dim c_Yes As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tYES
        Dim lRetCode As Long
        oUserObjectMD = objAddOn.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        If Not oUserObjectMD.GetByKey(udoname) Then
            oUserObjectMD.Code = udoname
            oUserObjectMD.Name = udoname
            oUserObjectMD.ObjectType = type
            oUserObjectMD.TableName = tblname
            oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            If DfltForm = True Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES

                oUserObjectMD.FormColumns.FormColumnAlias = "Code"
                oUserObjectMD.FormColumns.FormColumnDescription = "Code"
                oUserObjectMD.FormColumns.Add()
                oUserObjectMD.FormColumns.FormColumnAlias = "Name"
                oUserObjectMD.FormColumns.FormColumnDescription = "Name"
                oUserObjectMD.FormColumns.Add()
            End If
            If FindForm = True Then
                If type = SAPbobsCOM.BoUDOObjType.boud_MasterData Then
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                        Case "MIPLAGMAS"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLDM"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_ItemCode"
                            oUserObjectMD.FindColumns.ColumnDescription = "ItemCode"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_DesignCode"
                            oUserObjectMD.FindColumns.ColumnDescription = "DesignCode"
                            oUserObjectMD.FindColumns.Add()

                        Case "MIPLSC"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_WhsCode"
                            oUserObjectMD.FindColumns.ColumnDescription = "WarehouseCode"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_WhsName"
                            oUserObjectMD.FindColumns.ColumnDescription = "WarehouseName"
                            oUserObjectMD.FindColumns.Add()


                        Case "MIPLPM"
                            oUserObjectMD.FindColumns.ColumnAlias = "Code"
                            oUserObjectMD.FindColumns.ColumnDescription = "Code"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "Name"
                            oUserObjectMD.FindColumns.ColumnDescription = "Name"
                            oUserObjectMD.FindColumns.Add()

                        Case "MIPLGCPAUTH"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_UCode"
                            oUserObjectMD.FindColumns.ColumnDescription = "UCode"
                            oUserObjectMD.FindColumns.Add()


                    End Select
                Else
                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                    Select Case udoname
                        Case "MIPLCC"
                            oUserObjectMD.FindColumns.ColumnAlias = "DocNum"
                            oUserObjectMD.FindColumns.ColumnDescription = "DocNum"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_CardName"
                            oUserObjectMD.FindColumns.ColumnDescription = "CardName"
                            oUserObjectMD.FindColumns.Add()

                        Case "MIPLLRP"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_LRPNumber"
                            oUserObjectMD.FindColumns.ColumnDescription = "LRPNumber"
                            oUserObjectMD.FindColumns.Add()
                            oUserObjectMD.FindColumns.ColumnAlias = "U_CardCode"
                            oUserObjectMD.FindColumns.ColumnDescription = "CardCode"
                            oUserObjectMD.FindColumns.Add()

                        Case "MIPLGCPMAT"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_CardNo"
                            oUserObjectMD.FindColumns.ColumnDescription = "CardNo"
                            oUserObjectMD.FindColumns.Add()

                        Case "MIPLGPLAC"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_CardNo"
                            oUserObjectMD.FindColumns.ColumnDescription = "CardNo"
                            oUserObjectMD.FindColumns.Add()
                        Case "MIPLLRC"
                            oUserObjectMD.FindColumns.ColumnAlias = "U_CardNo"
                            oUserObjectMD.FindColumns.ColumnDescription = "Gcpcardnumber"
                            oUserObjectMD.FindColumns.Add()
                    End Select
                End If
            End If
            If childTable.Length > 0 Then
                For i = 0 To childTable.Length - 2
                    If Trim(childTable(i)) <> "" Then
                        oUserObjectMD.ChildTables.TableName = childTable(i)
                        oUserObjectMD.ChildTables.Add()
                    End If
                Next
            End If
            lRetCode = oUserObjectMD.Add()
            If lRetCode <> 0 Then
                MsgBox("error" + CStr(lRetCode))
                MsgBox(objAddOn.objCompany.GetLastErrorDescription)
            Else

            End If
            If DfltForm = True Then
                creationPackage = objAddOn.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
                ' Need to set the parameter with the object unique ID
                creationPackage.ObjectType = "1"
                creationPackage.UniqueID = udoname
                creationPackage.FormType = udoname
                creationPackage.BorderStyle = SAPbouiCOM.BoFormTypes.ft_Fixed
                objform = objAddOn.objApplication.Forms.AddEx(creationPackage)
            End If
        End If

    End Sub

    Private Sub objApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        If EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Or EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
            Try
                ' objUIXml.LoadMenuXML("RemoveMenu.xml", Mukesh.SBOLib.UIXML.enuResourceType.Embeded)
                If objCompany.Connected Then objCompany.Disconnect()
                objCompany = Nothing
                objApplication = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCompany)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objApplication)
                GC.Collect()
            Catch ex As Exception
            End Try
            End
        End If
    End Sub


    Private Sub applyFilter()
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        oFilters = New SAPbouiCOM.EventFilters
        'Item Master Data 
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)




        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)


    End Sub


    Public Function apicall(ByVal url As String) As String
        Dim httpreq As HttpWebRequest = DirectCast(WebRequest.Create(url), HttpWebRequest)
        Try
            Dim httpres As HttpWebResponse = DirectCast(httpreq.GetResponse(), HttpWebResponse)
            Dim sr As New StreamReader(httpres.GetResponseStream())
            Dim results As String = sr.ReadToEnd()

            sr.Close()
            Return results
        Catch
            Return "0"
        End Try
    End Function
End Class


