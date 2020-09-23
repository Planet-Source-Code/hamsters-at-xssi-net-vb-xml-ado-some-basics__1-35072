VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' The DOMDocument object which holds the XML Document
' will be redimmed later, when exact class will be known
Dim xmlDoc As Object

' Constant which holds the generic object name for the DOMDocument
Private Const CONST_DOMDOCUMENT As String = "DOMDocument"

' Constants used to test-invoke the different MS XML parser version
Private Const CONST_INVOKE_DOMDOCUMENT_30 As String = "Msxml2.DOMDocument.3.0"  ' this can be invoked by MSXML30
Private Const CONST_INVOKE_DOMDOCUMENT_26 As String = "Msxml2.DOMDocument.2.6"  ' this can be invoked by MSXML30 & MSXML26
Private Const CONST_INVOKE_DOMDOCUMENT_20 As String = "Msxml.DOMDocument"       ' this can be invoked by MSXML30 & MSXML26 & MSXML20

' Module level variable which holds the .ParserVersion property
Dim mpavCurrentParserVersion As ParserVersionConstant


Public Enum ParserVersionConstant
    ' 0 = return value only, do not assign!
    pavV20 = 1
    pavV26 = 2
    pavV30 = 3
    pavAUTO = 99
End Enum

' Enumeration of possible node relationships the user may select
' with the .InsertNode method
Public Enum NodeRelationshipConstant
    norSIBLING = 0
    norchild = 1
End Enum

' Enumeration of possible XML document information the user may select
' with the .XMLDocumentInfo property
Public Enum XMLDocumentInfoConstant
    xdiVERSION = 0
    xdiENCODING = 1
    xdiSTANDALONE = 2
End Enum

' Enumeration of possible Constants used
' with the .OpenXML method
Private Enum OpenXMLConstant
    oxFILE = 0
    oxSTRING = 1
End Enum



Public Property Get XML() As String
    XML = xmlDoc.XML
End Property



Public Property Get XMLDocumentInfo(xdiDesiredInfo As XMLDocumentInfoConstant) As Variant
    On Error GoTo ErrorHandler
    
    Dim strDesiredInfo As String
    If xdiDesiredInfo = xdiVERSION Then
        strDesiredInfo = "version"
    ElseIf xdiDesiredInfo = xdiENCODING Then
        strDesiredInfo = "encoding"
    ElseIf xdiDesiredInfo = xdiSTANDALONE Then
        strDesiredInfo = "standalone"
    End If
    
    Dim xmlNode As Object
    Set xmlNode = xmlDoc.childNodes(0).Attributes.getNamedItem(strDesiredInfo)
    XMLDocumentInfo = xmlNode.Text
    Set xmlNode = Nothing
    
    Exit Property
ErrorHandler:
    XMLDocumentInfo = Null
    Set xmlNode = Nothing
End Property

Public Property Get XMLParserVersion() As ParserVersionConstant
    XMLParserVersion = mpavCurrentParserVersion
End Property



Public Function ReadNodeXML(ByVal strQueryString As String) As String
    On Error GoTo ErrorHandler
    
    Dim xmlNode As Object
    Set xmlNode = xmlDoc.documentElement.selectSingleNode(strQueryString)
  
    ReadNodeXML = xmlNode.XML
    
    Set xmlNode = Nothing
    Exit Function
ErrorHandler:
    ReadNodeXML = ""
    Set xmlNode = Nothing
   
End Function

Public Function Initialize(pavDesiredVersion As ParserVersionConstant) As ParserVersionConstant
    Select Case pavDesiredVersion ' Which parser does the user want ot invoke?
        Case pavAUTO: ' Autodetect parser, use latest version
            mpavCurrentParserVersion = AutoDetectParser()
            If mpavCurrentParserVersion = 0 Then ' Oops, no parser installed
                Initialize = 0
            Else
                 Initialize = mpavCurrentParserVersion
            End If
        Case Else: ' User selects parser himself. Better check, if really available
            ' Test invoke
            If InvokeParser(pavDesiredVersion) = True Then
                mpavCurrentParserVersion = pavDesiredVersion
                Initialize = mpavCurrentParserVersion
            Else
                mpavCurrentParserVersion = 0
                Initialize = 0
            End If
    End Select
End Function
Private Function OpenXML(ByVal oxFileOrString As OpenXMLConstant, ByVal strSource As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
    ' Create the DOMDocument
    Set xmlDoc = InitializeObject(CONST_DOMDOCUMENT)

    xmlDoc.resolveExternals = False
    xmlDoc.validateOnParse = True
    xmlDoc.async = False
    
    Dim bolLoadSuccess As Boolean
    Select Case oxFileOrString
        Case oxFILE:
            bolLoadSuccess = xmlDoc.Load(strSource) ' .Load = Load from File
        Case oxSTRING:
            bolLoadSuccess = xmlDoc.loadXML(strSource) ' .LoadXML = Load from String
    End Select
        
    Select Case bolLoadSuccess ' Loading successful?
        Case True:
            OpenXML = True
        Case False:
            On Error Resume Next
            Dim xmlParseError As Object
            Set xmlParseError = xmlDoc.parseError
            
            ' Can we apply the workaround? Is that the problem?
            If mpavCurrentParserVersion = pavV20 And oxFileOrString = oxSTRING And xmlParseError.errorCode = -1072896657 Then ' Guess so
                
                Dim strFile As String
                Dim intFileNumber As Integer

                strFile = App.Path & "\" & "tempgoxml.xml"
                intFileNumber = FreeFile()

                Open strFile For Output As intFileNumber
                Print #intFileNumber, strSource
                Close intFileNumber
                
                ' Try again, from file this time
                OpenXML = OpenFromFile(strFile, bolVerboseError)
            
                Kill (strFile)
            Else
                If bolVerboseError = True Then Call DisplayParseError
                OpenXML = False
                Set xmlDoc = Nothing
            End If
            
            Set xmlParseError = Nothing
        End Select
End Function

Public Function OpenFromFile(ByVal strFileName As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
    OpenFromFile = OpenXML(oxFILE, strFileName, bolVerboseError)
End Function


Public Function OpenFromString(ByVal strSource As String, Optional ByVal bolVerboseError As Boolean = False) As Boolean
    OpenFromString = OpenXML(oxSTRING, strSource, bolVerboseError)
End Function
Public Function NodeCount(ByVal strQueryString As String) As Long
    On Error GoTo ErrorHandler

    Dim xmlNodeList As Object
    Set xmlNodeList = xmlDoc.documentElement.selectNodes(strQueryString)
    NodeCount = xmlNodeList.length
    Set xmlNodeList = Nothing

    Exit Function
ErrorHandler:
    NodeCount = -1
    Set xmlNodeList = Nothing
End Function
Public Function ReadNode(ByVal strQueryString As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim xmlNode As Object
    Set xmlNode = xmlDoc.documentElement.selectSingleNode(strQueryString)
  
    ReadNode = xmlNode.Text
    
    Set xmlNode = Nothing
    Exit Function
ErrorHandler:
    ReadNode = Null
    Set xmlNode = Nothing
   
End Function
Public Function ReadAttribute(ByVal strQueryString As String, ByVal strAttributeName As String) As Variant
    On Error GoTo ErrorHandler
    
    Dim xmlAttributeNode As Object
    Set xmlAttributeNode = xmlDoc.documentElement.selectSingleNode(strQueryString)

    ReadAttribute = xmlAttributeNode.getAttribute(strAttributeName)

    Set xmlAttributeNode = Nothing
    Exit Function
ErrorHandler:
    ReadAttribute = Null
    Set xmlAttributeNode = Nothing
End Function
'.------------------------------------------------------------------------------
Public Function DeleteNode(ByVal strQueryString As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlNode As Object
    Set xmlNode = xmlDoc.documentElement.selectSingleNode(strQueryString)
    
    xmlNode.parentNode.removeChild xmlNode ' This one was tricky ;-)
    
    Set xmlNode = Nothing
    DeleteNode = True
    Exit Function
ErrorHandler:
    DeleteNode = False
    Set xmlNode = Nothing
End Function
Public Function WriteNode(ByVal strQueryString As String, ByVal vntNewNodeValue As Variant) As Boolean
    On Error GoTo ErrorHandler
   
    Dim xmlNode As Object
    Set xmlNode = xmlDoc.documentElement.selectSingleNode(strQueryString)
    
    xmlNode.Text = vntNewNodeValue
    
    Set xmlNode = Nothing
    WriteNode = True
    Exit Function
ErrorHandler:
    WriteNode = False
    Set xmlNode = Nothing
End Function
Public Function WriteAttribute(ByVal strQueryString As String, ByVal strAttributeName As String, ByVal vntNewAttributeValue As Variant) As Boolean
    On Error GoTo ErrorHandler

    Dim xmlAttributeNode As Object
    Set xmlAttributeNode = xmlDoc.documentElement.selectSingleNode(strQueryString)
    
    xmlAttributeNode.setAttribute strAttributeName, vntNewAttributeValue
    
    Set xmlAttributeNode = Nothing
    WriteAttribute = True
    Exit Function
ErrorHandler:
    WriteAttribute = False
    Set xmlAttributeNode = Nothing
End Function
Public Function WriteCDataSection(ByVal strQueryString As String, ByVal vntNewCDataValue As Variant) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlNode As Object
    Dim xmlCDATASection As Object
    Set xmlNode = xmlDoc.documentElement.selectSingleNode(strQueryString)

    Select Case xmlNode.childNodes.length
        Case 0: ' No Problem, go ahead
            Set xmlCDATASection = xmlDoc.createCDATASection(vntNewCDataValue)
            xmlNode.appendChild xmlCDATASection
        Case 1: ' UHOH, better check, if there is already a CData section in form of a CHILD Node present
            Select Case xmlNode.childNodes(0).nodeTypeString
                Case "cdatasection": ' Yes, CDATA section already exists, change Text-Property only!
                    xmlNode.childNodes.Item(0).Text = vntNewCDataValue
                Case Else: ' Probably a Text section, delete Text and replace with CData section
                    xmlNode.Text = ""
                    Set xmlCDATASection = xmlDoc.createCDATASection(vntNewCDataValue)
                    xmlNode.appendChild xmlCDATASection
            End Select
        Case Else: ' Raise Error, if more than 1 child node
            GoTo ErrorHandler
    End Select

    Set xmlCDATASection = Nothing
    Set xmlNode = Nothing
    WriteCDataSection = True
    Exit Function
ErrorHandler:
    WriteCDataSection = False
    Set xmlCDATASection = Nothing
    Set xmlNode = Nothing
End Function
Public Function InsertNode(ByVal strQueryString As String, ByVal strNewNodeName As String, Optional ByVal vntNewNodeValue As Variant, Optional ByVal strNewAttributeName As String, Optional ByVal vntNewAttributeValue As Variant, Optional ByVal norRelationship As NodeRelationshipConstant = norchild) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlNode As Object
    Dim xmlChildNode As Object
    Dim xmlAttributeNode As Object
    
    Select Case norRelationship
        Case norchild:
            Set xmlNode = xmlDoc.documentElement.selectSingleNode(strQueryString)
        Case norSIBLING:
            Set xmlNode = xmlDoc.documentElement.selectSingleNode(strQueryString).parentNode
    End Select
   
    Set xmlChildNode = xmlDoc.createElement(strNewNodeName)
            
    If Not IsMissing(vntNewNodeValue) Then
        xmlChildNode.Text = vntNewNodeValue
    End If

    Set xmlAttributeNode = xmlNode.appendChild(xmlChildNode)
            
    If Not IsMissing(strNewAttributeName) And Not IsMissing(vntNewAttributeValue) Then
        xmlAttributeNode.setAttribute strNewAttributeName, vntNewAttributeValue
    End If
            
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set xmlNode = Nothing
    InsertNode = True
    Exit Function
ErrorHandler:
    InsertNode = False
    Set xmlAttributeNode = Nothing
    Set xmlChildNode = Nothing
    Set xmlNode = Nothing
End Function
Public Function DeleteAttribute(ByVal strQueryString As String, ByVal strAttributeName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim xmlAttributeNode As Object
    Set xmlAttributeNode = xmlDoc.documentElement.selectSingleNode(strQueryString).Attributes.removeNamedItem(strAttributeName)
    
    Set xmlAttributeNode = Nothing
    DeleteAttribute = True
    Exit Function
ErrorHandler:
    DeleteAttribute = False
    Set xmlAttributeNode = Nothing
End Function
Public Function Reparse() As Boolean
    On Error GoTo ErrorHandler
    
    ' Create the DOMDocument
    Dim xmlDocReparse As Object
    Set xmlDocReparse = InitializeObject(CONST_DOMDOCUMENT)
    
    If xmlDocReparse.loadXML(xmlDoc.XML) = True Then
        Reparse = True
    Else
        Reparse = False
    End If
    
    Set xmlDocReparse = Nothing
    
    Exit Function
ErrorHandler:
    Reparse = False
    Set xmlDocReparse = Nothing
End Function
Public Function Save(ByVal strFileName As String) As Boolean
    On Error GoTo ErrorHandler
    
    xmlDoc.Save (strFileName)
    Save = True
    
    Exit Function
ErrorHandler:
    Save = False
End Function


Private Function InitializeObject(strObjectName As String) As Object
    On Error GoTo ErrorHandler
    Dim objDummy As Object
    Dim strClassName As String

    Select Case mpavCurrentParserVersion
        Case pavV20:
             strClassName = "Msxml"
        Case pavV26, pavV30:
            strClassName = "Msxml2"
        Case Else: ' currently not allowed = Error
            GoTo ErrorHandler
    End Select
    
    Dim strObjectToCreate As String
    strObjectToCreate = strClassName & "." & strObjectName
    
    Set objDummy = CreateObject(strObjectToCreate)
    
    Set InitializeObject = objDummy

    Set objDummy = Nothing
    Exit Function
ErrorHandler:
    If Err.Number = 429 Then
        Debug.Print "Can't CreateObject(" & strObjectToCreate & ")"
    End If
   
    Set InitializeObject = Nothing
    Set objDummy = Nothing
End Function
Private Function InvokeParser(pavDesiredVersion As ParserVersionConstant) As Boolean
    Dim objDummy As Object
    Dim strInvokation As String
    
    Select Case pavDesiredVersion
        Case pavV30:
            strInvokation = CONST_INVOKE_DOMDOCUMENT_30
        Case pavV26:
            strInvokation = CONST_INVOKE_DOMDOCUMENT_26
        Case pavV20:
            strInvokation = CONST_INVOKE_DOMDOCUMENT_20
    End Select
    
    On Error Resume Next
    Set objDummy = CreateObject(strInvokation)
    If Err.Number = 0 Then
        InvokeParser = True
    Else
        InvokeParser = False
    End If
    Set objDummy = Nothing
End Function
Private Function AutoDetectParser() As ParserVersionConstant
    ' Check latest versions first
    ' check for v30
    If InvokeParser(pavV30) = True Then
        AutoDetectParser = pavV30
        Exit Function
    End If
    
    ' check for v26
    If InvokeParser(pavV26) = True Then
        AutoDetectParser = pavV26
        Exit Function
    End If
       
    ' check for v20
    If InvokeParser(pavV20) = True Then
        AutoDetectParser = pavV20
        Exit Function
    End If
    
    ' No XML Parser detected
    AutoDetectParser = 0
End Function

Private Sub DisplayParseError()
    Dim xmlParseError As Object
    Set xmlParseError = xmlDoc.parseError
    
    MsgBox "Error loading XML" & vbCrLf & vbCrLf & _
        xmlParseError.errorCode & vbCrLf & _
        xmlParseError.reason & "Line " & xmlParseError.Line & _
        ", Position " & xmlParseError.linepos & vbCrLf & vbCrLf & _
        xmlParseError.srcText & vbCrLf _
        , vbExclamation, "XML Parse Error"
    
    Set xmlParseError = Nothing
End Sub

Private Sub Class_Terminate()
    Set xmlDoc = Nothing
End Sub













