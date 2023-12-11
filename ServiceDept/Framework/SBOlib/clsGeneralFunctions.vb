Namespace Mukesh.SBOLib
    Public Class GeneralFunctions
        Private objCompany As SAPbobsCOM.Company
        Private strThousSep As String = ","
        Private strDecSep As String = "."
        Private intQtyDec As Integer = 3

        Public Sub New(ByVal Company As SAPbobsCOM.Company)
            Dim objRS As SAPbobsCOM.Recordset
            objCompany = Company

            objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery("SELECT * FROM OADM")
            If Not objRS.EoF Then
                strThousSep = objRS.Fields.Item("ThousSep").Value
                strDecSep = objRS.Fields.Item("DecSep").Value
                intQtyDec = objRS.Fields.Item("QtyDec").Value
            End If
        End Sub

        Public Function GetDateTimeValue(ByVal SBODaMIPLAGNTMASring As String) As DateTime
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            objBridge.Format_StringToDate("")
            Return objBridge.Format_StringToDate(SBODaMIPLAGNTMASring).Fields.Item(0).Value
        End Function
        Public Function GetSBODateString(ByVal DateVal As DateTime) As String

            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value
        End Function

        Public Function GetSBODaMIPLAGNTMASring(ByVal DateVal As DateTime) As String
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Return objBridge.Format_DateToString(DateVal).Fields.Item(0).Value
        End Function
        Public Function GetQtyValue(ByVal QtyString As String) As Double
            Dim dblValue As Double
            QtyString = QtyString.Replace(strThousSep, "")
            QtyString = QtyString.Replace(strDecSep, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
            dblValue = Convert.ToDouble(QtyString)
            Return dblValue
        End Function

        Public Function GetQtyString(ByVal QtyVal As Double) As String
            GetQtyString = QtyVal.ToString()
            GetQtyString.Replace(",", strDecSep)
        End Function

        Public Function GetCode(ByVal sTableName As String) As String
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRS.DoQuery("SELECT Top 1 Code FROM " & sTableName + " ORDER BY Convert(INT,Code) DESC")
            If Not objRS.EoF Then
                Return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString()) + 1
            Else
                GetCode = "1"
            End If
        End Function

        Public Function GetDocNum(ByVal sUDOName As String) As String
            Dim StrSQL As String
            Dim objRS As SAPbobsCOM.Recordset
            objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
            objRS.DoQuery(StrSQL)
            objRS.MoveFirst()
            If Not objRS.EoF Then
                Return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString())
            Else
                GetDocNum = "1"
            End If
        End Function
        'Public Function GetDocNum_Mbook(ByVal sUDOName As String) As String
        '    objRS = objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '    StrSQL = "select Autokey from onnm where objectcode='" & sUDOName & "'"
        '    objRS.DoQuery(StrSQL)
        '    objRS.MoveFirst()
        '    objAddOn.objApplication.MessageBox(objRS.RecordCount)
        '    If objRS.RecordCount > 0 Then
        '        Return objRS.Fields.Item(0).Value.ToString
        '    Else
        '        Return "1"
        '    End If
        'End Function
    End Class
End Namespace