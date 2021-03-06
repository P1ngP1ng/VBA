VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BBGCOMAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TAG: BBCOM WRAPPER
Option Explicit
'
' private constants for BCOM API
Private Const CONST_SERVICE_TYPE As String = "//blp/refdata"
Private Const CONST_REQUEST_TYPE_REFERENCE As String = "ReferenceDataRequest"
'
' private data structures
Private bInputSecurityArray() As Variant
Private bInputFieldArray() As Variant
Private bOutputArray() As Variant
'
' BCOM objects
Private bSession As blpapicomLib2.Session
Private bService As blpapicomLib2.Service
Private bRequest As blpapicomLib2.REQUEST
Private bSecurityArray As blpapicomLib2.Element
Private bFieldArray As blpapicomLib2.Element
Private bEvent As blpapicomLib2.Event
Private bIterator As blpapicomLib2.MessageIterator
Private bIteratorData As blpapicomLib2.Message
Private bSecurities As blpapicomLib2.Element
Private bSecurity As blpapicomLib2.Element
Private bSecurityName As blpapicomLib2.Element
Private bSecurityField As blpapicomLib2.Element
Private bFieldValue As blpapicomLib2.Element
Private bSequenceNumber As blpapicomLib2.Element
Private bFields As blpapicomLib2.Element
Private bField As blpapicomLib2.Element
'
Public Function getData(securities() As Variant, fields() As Variant) As Variant()
    '
    ' this is the public interface function
    bInputSecurityArray = securities
    bInputFieldArray = fields
    redimOutputArray
    '
    openSession
    sendRequest
    catchServerEvent
    releaseObjects
    getData = bOutputArray
End Function
'
Private Function redimOutputArray()
    '
    Dim nSecurities As Long, nFields As Long
    nSecurities = UBound(bInputSecurityArray)
    nFields = UBound(bInputFieldArray)
    ReDim bOutputArray(0 To nSecurities, 0 To nFields)
End Function
'
Private Function openSession()
    '
    Set bSession = New blpapicomLib2.Session
    bSession.Start
    bSession.OpenService CONST_SERVICE_TYPE
    Set bService = bSession.GetService(CONST_SERVICE_TYPE)
End Function
'
Private Function sendRequest()
    '
    Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_REFERENCE)
    Set bSecurityArray = bRequest.GetElement("securities")
    Set bFieldArray = bRequest.GetElement("fields")
    appendRequestItems
    bSession.sendRequest bRequest
End Function
'
Private Function appendRequestItems()
    '
    Dim nSecurities As Long: nSecurities = UBound(bInputSecurityArray)
    Dim nFields As Long: nFields = UBound(bInputFieldArray)
    Dim i As Long
    Dim nItems As Integer: nItems = getMax(nSecurities, nFields)
    For i = 0 To nItems
        If (i <= nSecurities) Then bSecurityArray.AppendValue CStr(bInputSecurityArray(i))
        If (i <= nFields) Then bFieldArray.AppendValue CStr(bInputFieldArray(i))
    Next i
End Function
'
Private Function getMax(a As Long, b As Long) As Long
    '
    getMax = a
    If (b > a) Then getMax = b
End Function
'
Private Function catchServerEvent()
    '
    Dim bExit As Boolean
    Do While (bExit = False)
        Set bEvent = bSession.NextEvent
        If (bEvent.EventType = PARTIAL_RESPONSE Or bEvent.EventType = RESPONSE) Then
            getServerData_reference
            If (bEvent.EventType = RESPONSE) Then bExit = True
        End If
    Loop
End Function
'
Private Function getServerData_reference()
    '
    Set bIterator = bEvent.CreateMessageIterator
    '
    Do While (bIterator.Next)
        '
        Set bIteratorData = bIterator.Message
        Set bSecurities = bIteratorData.GetElement("securityData")
        Dim offsetNumber As Long, nFields As Long, i As Long, j As Long
        Dim nSecurities As Long: nSecurities = bSecurities.Count
        '
        For i = 0 To (nSecurities - 1)
            '
            Set bSecurity = bSecurities.GetValue(i)
            Set bSecurityName = bSecurity.GetElement("security")
            Set bSecurityField = bSecurity.GetElement("fieldData")
            Set bSequenceNumber = bSecurity.GetElement("sequenceNumber")
            offsetNumber = CInt(bSequenceNumber.Value)
            '
            For j = 0 To UBound(bInputFieldArray)
                If (bSecurityField.HasElement(bInputFieldArray(j))) Then
                    Set bFieldValue = bSecurityField.GetElement(bInputFieldArray(j))
                    bOutputArray(offsetNumber, j) = bFieldValue.Value
                End If
            Next j
        Next i
    Loop
End Function
'
Private Function releaseObjects()
    '
    Set bFieldValue = Nothing
    Set bSequenceNumber = Nothing
    Set bSecurityField = Nothing
    Set bSecurityName = Nothing
    Set bSecurity = Nothing
    Set bSecurities = Nothing
    Set bIteratorData = Nothing
    Set bIterator = Nothing
    Set bEvent = Nothing
    Set bFieldArray = Nothing
    Set bSecurityArray = Nothing
    Set bRequest = Nothing
    Set bService = Nothing
    bSession.Stop
    Set bSession = Nothing
End Function
'
''TAG: MAIN PROGRAM
'Sub tester()
'    '
'    Dim b As New BBCOM_wrapper
'    Dim securities(0 To 1) As Variant: securities(0) = "EUR Curncy": securities(1) = "THB Curncy"
'    Dim fields(0 To 2) As Variant: fields(0) = "SECURITY_NAME": fields(1) = "PX_BID": fields(2) = "PX_ASK"
'    Dim data As Variant: data = b.getData(securities, fields)
'    Set b = Nothing
'End Sub

