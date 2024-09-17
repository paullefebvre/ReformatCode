Dim RC_Version As String = "0.18" 'Type rc_version to see this at runtime
Dim RC_DebugLevel As Integer' = 2 'Uncomment this if you want to debug preferences before RC_DebugLevel is set (see LoadPreferences)

'Set up script preferences
Sub LoadPreferences()
  'You can alter the defaults of this script by changing the value of the last parameter of ReadPreference below
  'However its best to do this in your project as an updated script will reset your changes
  
  'Set the default storage location for constants, we read this in first so they can be moved
  'We always read this in from App.ConstantStorageLocation (you can change the default below)
  ReadPreference_ConstStorageLocation("ConstantStorageLocation", RC_ConstStorageLocation, "App.")
  
  'Set the prefix for all constant names
  ReadPreference("Prefix", RC_Prefix, "")
  
  'Set the constant defaults - you can changes these in the IDE if you want, see docs
  ReadPreference("DebugLevel", RC_DebugLevel, 0)
  ReadPreference("ForceCase", RC_ForceCase, "") 'uppercase lowercase <empty>
  ReadPreference("DimVar", RC_DimVar, "") 'dim var <empty>
  ReadPreference("PadParInside", RC_PadParInside, False)
  ReadPreference("PadParOutside", RC_PadParOutside, False)
  ReadPreference("PadEmptyParInside", RC_PadEmptyParInside, False)
  ReadPreference("RemoveEmptyPar", RC_RemoveEmptyPar, False)
  ReadPreference("PadComma", RC_PadComma, True)
  ReadPreference("PadOperators", RC_PadOperators, True)
  ReadPreference("PadIIf", RC_PadIIf, False)
  ReadPreference("PadLineContinuation", RC_PadLineContinuation, True)
  ReadPreference("Comment", RC_Comment, "") '// ' rem
  ReadPreference("PadCommentBefore", RC_PadCommentBefore, "") 'true false <empty>
  ReadPreference("PadCommentAfter", RC_PadCommentAfter, "") 'true false <empty>
  ReadPreference("MessageComment", RC_MessageComment, "'") '// ' rem
  ReadPreference("MessageParMismatched", RC_MessageParMismatched, "MISMATCHED PARENTHESES")
  ReadPreference("MessageParOpening", RC_MessageParOpening, "MISSING OPENING PARENTHESIS")
  ReadPreference("MessageParClosing", RC_MessageParClosing, "MISSING CLOSING PARENTHESIS")
  ReadPreference("MessageQuoteMismatched", RC_MessageQuoteMismatched, "MISMATCHED QUOTES")
  ReadPreference("MessageInCodeBlock", RC_MessageInCodeBlock, " IN CODE BLOCK")
  
  'Set the macro storage location
  ReadPreference("MacroStorageLocation", RC_MacroStorageLocation, "Macro")
End Sub

'Words to replace
Dim RC_ReplaceList() As String = Array ("!", "Not", "endif", "End If", "endselect", "End Select", "endsub", "End Sub", "endfunction", "End Function", "endtry", "End Try")

'Keywords to CaPiTaLiSe
Dim RC_Words() As String = Array ( _
"AddHandler", "AddressOf", "Aggregates", "Alias", "And", "Array", "As", "Assigns", "Attributes", "Auto", _
"Boolean", "Break", "ByRef", "Byte", "ByVal", _
"Call", "Case", "Catch", "CGFloat", "CFStringRef", "CType", "Class", "Color", "Const", "Continue", "CString", "Currency", _
"Declare", "Delegate", "Dim", "Do", "Double", "DownTo", _
"Each", "Else", "ElseIf", "End", "Enum", "Event", "Exception", "Exit", "Extends", _
"False", "Finally", "For", "Function", _
"GetTypeInfo", "Global", "Goto", _
"Handles", "If", "Implements", "In", "Inherits", "Int8", "Int16", "Int32", "Int64", "Integer", "Interface", "Is", "IsA", _
"Lib", "Loop", _
"Me", "Mod", "Module", _
"Namespace", "New", "Next", "Nil", "Not", _
"Object", "Of", "Optional", "Or", "OSType", _
"Private", "Property", "Protected", "Ptr", "Public", _
"Raise", "RaiseEvent", "Redim", "Rem", "RemoveHandler", "Return", "Return", _
"Select", "Selector", "Self", "Shared", "Short", "Single", "Soft", "Static", "Step", "String", "Structure", "Sub", "Super", _
"Text", "Then", "To", "True", "Try", _
"UInt8", "UInt16", "UInt32", "UInt64", "UInteger", "Until", "Using", _
"Var", "Variant", _
"WeakAddressOf", "Wend", "While", "With", "WString", _
"Xor")

Dim RC_Compiler_Constants() As String = Array ("#If", "#Else", "#ElseIf", "#EndIf", "Target32Bit", "Target64Bit", "TargetARM", "TargetBigEndian", "TargetCocoa", "TargetLinux", "TargetLittleEndian", "TargetMacOS", "TargetWin32", "TargetWindows", "TargetX86", "TargetXojoCloud", "TargetConsole", "TargetDesktop", "TargetiOS", "TargetWeb", "CurrentMethodName", "DebugBuild", "TargetRemoteDebugger", "TargetXojoCloud", "XojoVersion", "XojoVersionString")

Dim RC_Pragma_Directives() As String = Array ("#Pragma", "BackgroundTasks", "BoundsChecking", "BreakOnExceptions", "DisableBackgroundTasks", "DisableBoundsChecking", "Error", "NilObjectChecking", "StackOverflowChecking", "Unused", "Warning", "X86CallingConvention")

Dim RC_Xojo_Core() As String = Array ("Xojo", "Core", "Date", "DateInterval", "Dictionary", "DictionaryEntry", "Iterable", "Iterator", "Locale", "MemoryBlock", "MutableMemoryBlock", "Point", "Rect", "Size", "StackFrame", "TextEncoding", "Timer", "TimeZone", "WeakRef", "Exceptions", "BadDataException", "ErrorException", "InvalidArgumentException", "IteratorException", "LogicException", "UnsupportedOperationException")

'TODO - reengineer this?
'Dim RC_Xojo_Core_Date() As String = Array("FormatStyles", "Day", "DayOfWeek", "DayOfYear", "Hour", "Minute", "Month", "Nanosecond", "Second", "SecondsFrom1970", "TimeZone", "Year", "ToText", "FromText", "Now", "Add", "Subtract")
'Dim RC_Xojo_Core_DateInterval() As String = Array("Days", "Hours", "Minutes", "Months", "NanoSeconds", "Seconds", "Years")
'Dim RC_Xojo_Core_Dictionary() As String = Array("Iterable", "CompareKeys", "Count", "Clone", "GetIterator", "HasKey", "Lookup", "RemoveAll", "Remove", "Value")
'Dim RC_Xojo_Core_DictionaryEntry() As String = Array("Key", "Value")
'Dim RC_Xojo_Core_Iterable() As String = Array("GetIterator")
'Dim RC_Xojo_Core_Iterator() As String = Array("MoveNext", "Value")
'Dim RC_Xojo_Core_Locale() As String = Array("Identifier", "CurrencySymbol", "GroupingSeparator", "DecimalSeparator", "Current", "Raw")
'Dim RC_Xojo_Core_MemoryBlock() As String = Array("LittleEndian", "Size", "Data", "Clone", "IndexOf", "Left", "Mid", "Right", "BooleanValue", "CurrencyValue", "CStringValue", "DoubleValue", "Int8Value", "Int16Value", "Int32Value", "Int64Value", "PtrValue", "SingleValue", "UInt8Value", "UInt16Value", "UInt32Value", "UInt64Value", "WStringValue")
'Dim RC_Xojo_Core_MutableMemoryBlock() As String = Array("LittleEndian", "Size", "Data", "Append", "Insert", "Left", "Mid", "Remove", "Right", "BooleanValue", "CurrencyValue", "CStringValue", "DoubleValue", "Int8Value", "Int16Value", "Int32Value", "Int64Value", "PtrValue", "SingleValue", "UInt8Value", "UInt16Value", "UInt32Value", "UInt64Value", "WStringValue")
'Dim RC_Xojo_Core_Point() As String = Array("X", "Y", "DistanceTo", "Translate", "Zero", "Add", "Compare", "Negate", "Subtract")
'Dim RC_Xojo_Core_Rect() As String = Array("Bottom", "Center", "Height", "Left", "Origin", "Right", "Size", "Top", "Width", "Contains", "Inset", "Intersection", "Offset", "Union", "Compare")
'Dim RC_Xojo_Core_Size() As String = Array("Height", "Width", "Zero", "Compare")

'Dim RC_Xojo_Crypto() As String = Array("HashAlgorithms", "BERDecodePrivateKey", "BERDecodePublicKey", "DEREncodePrivateKey", "DEREncodePublicKey", "GenerateRandomBytes", "Hash", "HMAC", "MD5", "PBKDF2", "RSADecrypt", "RSAEncrypt", "RSAGenerateKeyPair", "RSASign", "RSAVerifyKey", "RSAVerifySignature", "SHA1", "SHA256", "SHA512", "CryptoException", "MD5", "SHA1", "SHA256", "SHA512")
'Dim RC_Xojo_Data() As String = Array("GenerateJSON", "ParseJSON", "InvalidJSONException")
'I stopped here for the time being until I figure out what is the best way to progress.

Dim RC_WindowsDeclareList() As String = Array ( _
"W_ACCESS_MASK", "UInt32", _
"W_APIENTRY", "APIENTRY*", _
"W_ATOM", "UInt16", _
"W_BOOL", "Int32", _
"W_BOOLEAN", "UInt8", _
"W_BYTE", "Byte", _
"W_CALLBACK", "CALLBACK*", _
"W_CCHAR", "Int8", _
"W_CHAR", "Int8", _
"W_COLORREF", "UInt32/Color*", _
"W_CONST", "Const", _
"W_DOUBLE", "Double", _
"W_DWORD", "UInt32", _
"W_DWORDLONG", "UInt64", _
"W_DWORD_PTR", "UInteger", _
"W_DWORD32", "UInt32", _
"W_DWORD64", "UInt64", _
"W_EXECUTION_STATE", "UInt32", _
"W_FARPROC", "Integer", _
"W_FILEOP_FLAGS", "UInt16", _
"W_FLOAT", "Single", _
"W_HACCEL", "Integer", _
"W_HALF_PTR", "Int16/Int32*", _
"W_HANDLE", "Integer", _
"W_HBITMAP", "Integer", _
"W_HBRUSH", "Integer", _
"W_HCOLORSPACE", "Integer", _
"W_HCONV", "Integer", _
"W_HCONVLIST", "Integer", _
"W_HCURSOR", "Integer", _
"W_HDC", "Integer", _
"W_HDDEDATA", "Integer", _
"W_HDESK", "Integer", _
"W_HDROP", "Integer", _
"W_HDWP", "Integer", _
"W_HENHMETAFILE", "Integer", _
"W_HFILE", "Int32", _
"W_HFONT", "Integer", _
"W_HGDIOBJ", "Integer", _
"W_HGLOBAL", "Integer", _
"W_HHOOK", "Integer", _
"W_HICON", "Integer", _
"W_HINSTANCE", "Integer", _
"W_HKEY", "Integer", _
"W_HKL", "Integer", _
"W_HLOCAL", "Integer", _
"W_HMENU", "Integer", _
"W_HMETAFILE", "Integer", _
"W_HMODULE", "Integer", _
"W_HMONITOR", "Integer*", _
"W_HPALETTE", "Integer", _
"W_HPEN", "Integer", _
"W_HRESULT", "Int32", _
"W_HRGN", "Integer", _
"W_HRSRC", "Integer", _
"W_HSZ", "Integer", _
"W_HWINSTA", "Integer", _
"W_HWND", "Integer", _
"W_INT", "Int32", _
"W_INT_PTR", "Integer", _
"W_INT8", "Int8", _
"W_INT16", "Int16", _
"W_INT32", "Int32", _
"W_INT64", "Int64", _
"W_LANGID", "UInt16", _
"W_LCID", "UInt32", _
"W_LCTYPE", "UInt32", _
"W_LGRPID", "UInt32", _
"W_LONG", "Int32", _
"W_LONGLONG", "Int64", _
"W_LONG_PTR", "Integer", _
"W_LONG32", "Int32", _
"W_LONG64", "Int64", _
"W_LPARAM", "Integer", _
"W_LPBOOL", "Int32/Ptr/Integer*", _
"W_LPBYTE", "Byte/Ptr/Integer", _
"W_LPCOLORREF", "UInt32/Color*", _
"W_LPCSTR", "CString", _
"W_LPCTSTR", "WString", _
"W_LPCVOID", "UInt32/See PVOID*", _
"W_LPCWSTR", "WString", _
"W_LPDWORD", "UInt32", _
"W_LPHANDLE", "Integer", _
"W_LPINT", "Int32", _
"W_LPLONG", "Int32", _
"W_LPSTR", "CString", _
"W_LPTSTR", "WString/CString/Ptr*", _
"W_LPVOID", "See PVOID*", _
"W_LPWORD", "UInt16", _
"W_LPWSTR", "WString", _
"W_LRESULT", "Integer", _
"W_PBOOL", "Int32", _
"W_PBOOLEAN", "UInt8", _
"W_PBYTE", "Byte", _
"W_PCHAR", "CString", _
"W_PCSTR", "CString", _
"W_PCTSTR", "WString/CString*", _
"W_PCWSTR", "WString", _
"W_PDWORD", "UInt32", _
"W_PDWORDLONG", "UInt64", _
"W_PDWORD_PTR", "UInteger", _
"W_PDWORD32", "UInt32", _
"W_PDWORD64", "UInt64", _
"W_PFLOAT", "Single", _
"W_PHALF_PTR", "Int16/Int32*", _
"W_PHANDLE", "Integer", _
"W_PHKEY", "Integer", _
"W_PINT", "Int32", _
"W_PINT_PTR", "Integer", _
"W_PINT8", "Int8", _
"W_PINT16", "Int16", _
"W_PINT32", "Int32", _
"W_PINT64", "Int64", _
"W_PLCID", "UInt32", _
"W_PLONG", "Int32", _
"W_PLONGLONG", "Int64", _
"W_PLONG_PTR", "Integer", _
"W_PLONG32", "Int32", _
"W_PLONG64", "Int64", _
"W_POINTER_32", "Integer", _
"W_POINTER_64", "Integer", _
"W_POINTER_SIGNED", "Integer", _
"W_POINTER_UNSIGNED", "UInteger", _
"W_PSHORT", "Int16", _
"W_PSIZE_T", "UInteger", _
"W_PSSIZE_T", "Integer", _
"W_PSTR", "CString", _
"W_PTBYTE", "UInt16/Int8*", _
"W_PTCHAR", "UInt16/UInt8*", _
"W_PTSTR", "WString/CString", _
"W_PUCHAR", "UInt8", _
"W_PUHALF_PTR", "UInt16/UInt32*", _
"W_PUINT", "UInt32", _
"W_PUINT_PTR", "UInteger", _
"W_PUINT8", "UInt8", _
"W_PUINT16", "UInt16", _
"W_PUINT32", "UInt32", _
"W_PUINT64", "UInt64", _
"W_PULONG", "UInt32", _
"W_PULONGLONG", "UInt64", _
"W_PULONG_PTR", "UInteger", _
"W_PULONG32", "UInt32", _
"W_PULONG64", "UInt64", _
"W_PUSHORT", "UInt16", _
"W_PVOID", "Ptr*", _
"W_PWCHAR", "WString", _
"W_PWORD", "UInt16", _
"W_PWSTR", "WString/Ptr*", _
"W_QWORD", "UInt64", _
"W_REGSAM", "UInt32", _
"W_SC_HANDLE", "Integer", _
"W_SC_LOCK", "See PVOID", _
"W_SERVICE_STATUS_HANDLE", "Integer", _
"W_SHORT", "Int16", _
"W_SIZE_T", "UInteger", _
"W_SSIZE_T", "Integer", _
"W_TBYTE", "UInt16/Int8*", _
"W_TCHAR", "UInt16/UInt8*", _
"W_UCHAR", "UInt8", _
"W_UHALF_PTR", "UInt16/UInt32*", _
"W_UINT", "UInt32", _
"W_UINT_PTR", "UInteger", _
"W_UINT8", "UInt8", _
"W_UINT16", "UInt16", _
"W_UINT32", "UInt32", _
"W_UINT64", "UInt64", _
"W_ULONG", "UInt32", _
"W_ULONGLONG", "UInt64", _
"W_ULONG_PTR", "UInteger", _
"W_ULONG32", "UInt32", _
"W_ULONG64", "UInt64", _
"W_UNICODE_STRING", "Structure*", _
"W_USHORT", "UInt16", _
"W_USN", "Int64", _
"W_VARTYPE*", "UInt16", _
"W_VOID", "Ptr*", _
"W_WCHAR", "UInt16", _
"W_WINAPI", "WINAPI*", _
"W_WNDPROC", "Integer/Ptr*", _
"W_WORD", "UInt16", _
"W_WPARAM", "UInteger" _
)

Dim IDEVersion As String = ConstantValue("Xojo.IDEVersion")
Dim APIVersion As Integer = If(IDEVersion = "", 1, 2)

Debug("IDEVersion=>" + IDEVersion + "<", 2)
Debug("APIVersion=" + Str(APIVersion), 2)

Dim RC_Prefix As String

Dim RC_ConstStorageLocation As String

Dim RC_ForceCase As String
Dim RC_DimVar As String
Dim RC_PadParInside As Boolean
Dim RC_PadParOutside As Boolean
Dim RC_PadEmptyParInside As Boolean
Dim RC_RemoveEmptyPar As Boolean
Dim RC_PadComma As Boolean
Dim RC_PadOperators As Boolean
Dim RC_PadIIf As Boolean
Dim RC_PadLineContinuation As Boolean
Dim RC_Comment As String
Dim RC_PadCommentBefore As String 'We use a string here instead of a boolean because this is a tristate variable
Dim RC_PadCommentAfter As String 'We use a string here instead of a boolean because this is a tristate variable
Dim RC_MessageComment As String
Dim RC_MessageParMismatched As String
Dim RC_MessageParOpening As String
Dim RC_MessageParClosing As String
Dim RC_MessageQuoteMismatched As String
Dim RC_MessageInCodeBlock As String

Dim RC_MacroStorageLocation As String

Sub InitWords()
  RC_Words.Concat(RC_Compiler_Constants)
  RC_Words.Concat(RC_Pragma_Directives)
  RC_Words.Concat(RC_Xojo_Core)
  
  'Read in and process additional ReplaceWords, we'll add these to RC_ReplaceList
  Dim ReplaceWords As String
  ReadPreference("ReplaceWords", ReplaceWords, "")
  
  Debug("ReplaceWords=>" + ReplaceWords + "<", 2)
  Debug("RC_ReplaceList=>" + str(RC_ReplaceList.UBound) + "<", 2)
  
  If ReplaceWords <> "" Then
    Dim c As Integer
    c = ReplaceWords.CountFields(",")
    If c > 1 And c Mod 2 = 0 Then
      'Make sure that our comma separated list is in pairs
      Dim s() As String
      s = Split(ReplaceWords, ",")
      RC_ReplaceList.Concat(s)
    End If
  End If
  
  Debug("RC_ReplaceList=>" + str(RC_ReplaceList.UBound) + "<", 2)
  DebugArray(RC_ReplaceList, 2)
  
  'Merge in the list of windows delcares
  RC_ReplaceList.Concat(RC_WindowsDeclareList)
End Sub

'Heavy lifting is done here, get your safety gear on!
Sub CleanBlock()
  
  LoadPreferences
  
  InitWords 
  
  'Build error strings so we don't need to rebuild them when we use them
  Dim buildMessage As String = ""
  buildMessage = buildMessage + If(RC_PadCommentBefore = "true" Or RC_PadCommentBefore = "yes" Or RC_PadCommentBefore = "1" Or RC_MessageComment = "rem", " ", "")
  buildMessage = buildMessage + RC_MessageComment
  buildMessage = buildMessage + If(RC_PadCommentAfter = "true" Or RC_PadCommentAfter = "yes" Or RC_PadCommentAfter = "1" Or RC_MessageComment = "rem", " ", "")
  
  Dim expandedMessageParMismatched As String = If(RC_MessageParMismatched <> "", buildMessage + RC_MessageParMismatched, "")
  Dim expandedMessageParOpening As String = If(RC_MessageParOpening <> "", buildMessage + RC_MessageParOpening, "")
  Dim expandedRC_MessageParClosing As String = If(RC_MessageParClosing <> "", buildMessage + RC_MessageParClosing, "")
  Dim expandedMessageQuoteMismatched As String = If(RC_MessageQuoteMismatched <> "", buildMessage + RC_MessageQuoteMismatched, "")
  Dim expandedMessageInCodeBlock As String = If(RC_MessageInCodeBlock <> "", RC_MessageInCodeBlock, "")
  
  Debug("expandedMessageParMismatched=|" + expandedMessageParMismatched + "|", 2)
  Debug("expandedMessageParOpening=|" + expandedMessageParOpening + "|", 2)
  Debug("expandedRC_MessageParClosing=|" + expandedRC_MessageParClosing + "|", 2)
  Debug("expandedMessageQuoteMismatched=|" + expandedMessageQuoteMismatched + "|", 2)
  Debug("expandedMessageInCodeBlock=|" + expandedMessageInCodeBlock + "|", 2)
  
  Debug("Starting CleanBlock v" + RC_Version, 1)
  Debug("Lines detected=" + str(LineCount), 2)
  
  Dim currentLineNumber As Integer
  
  For currentLineNumber = 0 To LineCount - 1
    Debug("LINE=>" + Line(currentLineNumber) + "<", 2)
  Next
  
  Dim inContinuation As Boolean = False
  Dim wasInContinuation As Boolean = False
  'Parenthesis count across lines in the whole block
  Dim parCountBlock As Integer = 0
  Dim parMismatchBlock As Boolean = False
  
  Dim isInMSDNCodeBlock As Boolean = False
  
  Dim MSDN_Return As String = ""
  Dim MSDN_Call As String = ""
  
  If LineCount > 0 Then
    If Line(LineCount - 1) = ");" Then
      Debug("#~#~# INSIDE MSDN #~#~#", 2)
      isInMSDNCodeBlock = True
    End If
  End If
  
  For currentLineNumber = 0 To LineCount - 1
    Debug("Line=|" + Line(currentLineNumber) + "|", 1)
    
    Dim s As String = ""
    
    Dim skipToken As Boolean = False
    
    Dim allowNextSpace As Boolean = True
    
    StartTokenizer(Line(currentLineNumber))
    
    Dim TokenTextHistory() As String
    Dim TokenTypeHistory() As Integer
    
    'Keep track of and build up a tokenChain so we can easily detect macros
    Dim tokenChain As String = ""
    Dim tokenChainEnded As Boolean = False
    Dim tokenChainClear As Boolean = False
    
    'Keep track of the last found unicode literal so we can find it in Line(currentLineNumber) and check if it has quote around it
    Dim lastUnicodeLiteral As Integer = 0
    
    'Keep track of the number of parenthesis so we can check for a mismatch
    Dim parCount As Integer = 0
    'If we find a ) before a ( then set this to true
    Dim parMismatch As Boolean = False
    
    Dim MSDN_Type As String = ""
    Dim MSDN_Var As String = ""
    
    While NextToken = True
      
      Dim TokenTypeCurrent As Integer = TokenType
      Dim TokenTextCurrent As String = TokenText
      
      'Keep a track of the tokens so we can look back along the path at any time
      TokenTextHistory.Append(TokenTextCurrent)
      TokenTypeHistory.Append(TokenTypeCurrent)
      
      'Use some local variables so we don't need to jump around so much
      Dim TokenTextBackOne As String = GetTokenText(TokenTextHistory, 1)
      Dim TokenTextBackTwo As String = GetTokenText(TokenTextHistory, 2)
      Dim TokenTextBackThree As String = GetTokenText(TokenTextHistory, 3)
      
      Dim TokenTypeBackOne As Integer = GetTokenType(TokenTypeHistory, 1)
      Dim TokenTypeBackTwo As Integer = GetTokenType(TokenTypeHistory, 2)
      Dim TokenTypeBackThree As Integer = GetTokenType(TokenTypeHistory, 3)
      
      Debug("TT=>" + TokenTextCurrent + "<=" + Str(TokenTypeCurrent) + " s=>" + s + "< pTT=" + Str(TokenTypeBackOne) + " ppTT=" + Str(TokenTypeBackTwo) + " pppTT=" + Str(TokenTypeBackThree) + " aNS=" + If(allowNextSpace, "True", "False"), 1)
      
      'Main select
      Select Case TokenTypeCurrent
      Case 38 '&
        'Change & to + for users of other languages
        s = s + If(RC_PadOperators, " +", "+")
        allowNextSpace = RC_PadOperators
        skipToken = True
        
      Case 40 '(
        If isInMSDNCodeBlock And MSDN_Return <> "" And MSDN_Call <> "" And MSDN_Type = "" And MSDN_Var = "" Then
          'We're at the end of the first line so we can build the declare
          s = s + RC_Words_Formatted("Declare") + " " + If(MSDN_Return = "void", RC_Words_Formatted("Sub"), RC_Words_Formatted("Function")) + " " + MSDN_Call + " " + RC_Words_Formatted("Lib") + " ""REPLACE_ME.dll"" " + RC_Words_Formatted("Alias") + " """ + MSDN_Call + """ ( _"
          parCount = parCount + 1
          inContinuation = True
          allowNextSpace = False
          skipToken = True
          
        Else
          parCount = parCount + 1
          If RC_PadParOutside Then
            If TokenTypeBackOne = 45 Then '-
              'allow us to put the pad before the - so we end up with = -( not =- (
              If Mid(s, Len(s) - 1, 1) <> " " Then
                'check that we haven't already got a space before the - if not then put one in
                s = Left(s, Len(s) - 1) + " " + Right(s, 1)
              End If
              
            Else
              If TokenTypeBackOne = TOKEN_TK_IF And TokenTypeBackTwo > 0 Then
                'we are processing the ( of an IIf
                If RC_PadIIf Then
                  'allow space removal between inline if and the following (
                  s = s + " "
                End If
                
              Else
                s = s + " "
              End If
            End If
            
          Else
            If (TokenTypeBackOne = TOKEN_ASSIGN Or TokenTypeBackOne = 61) And RC_PadOperators Then '=
              '= ( if we are not padding ( but we are padding operators we pad = to keep things looking nice
              s = s + " "
              
            ElseIf TokenTypeBackOne = TOKEN_STRING Then
              'put a space between " and ( in a declare
              s = s + " "
              
            ElseIf TokenTypeBackOne = TOKEN_TK_STEP Then
              'put a space between step and (
              s = s + " "
              
            ElseIf TokenTypeBackOne = 45 And RC_PadOperators Then '-
              Select Case TokenTypeBackTwo
              Case 42, 43, 45, 47, 60, 61, 62, 92, 94, TOKEN_ASSIGN, TOKEN_GE_RELOP, TOKEN_LE_RELOP, TOKEN_NE_RELOP, TOKEN_TK_IF
                '  *   +   -   /   <   =   >   \   ^   =             >=              <=              <>              IF
                'put no space between - ( in ) >= -(
                'nor between - and ( in if -(
              Else
                'put a space between - and ( if we are RC_PadOperators and not RC_PadParOutside but not if we follow an = so we can do = -(
                s = s + " "
              End Select
              
            ElseIf TokenTypeBackOne = TOKEN_IDENTIFIER Then
              'don't put a space between a( if not RC_PadParOutside
              
            ElseIf TokenTypeBackOne = TOKEN_TK_CTYPE Then
              'don't put a space between ctype and ( if not RC_PadParOutside
              
            ElseIf TokenTypeBackOne = TOKEN_TK_IF And TokenTypeBackTwo > 0 Then
              'allow space removal between inline if and the following (
              If RC_PadIIf Then
                s = s + " "
              End If
              
            ElseIf allowNextSpace = True Then
              s = s + " "
            End If
          End If
          allowNextSpace = RC_PadParInside
        End If
        
        
      Case 41 ')
        parCount = parCount - 1
        If parCount < 0 Then parMismatch = True
        If parCountBlock < 0 Then parMismatchBlock = True
        If TokenTypeBackOne = 40 Then
          '()
          If RC_RemoveEmptyPar Then
            'strip out the ( as this is a ) and we want to remove empty ()
            s = Left(s, Len(s) - 1)
            If Right(s, 1) = " " Then
              'remove the trailing space if there is one, we could check RC_PadParOutside for this but we do it this way just to be sure
              s = Left(s, Len(s) - 1)
            End If
            
            skipToken = True
            
          Else
            If RC_PadEmptyParInside Then
              s = s + " "
            End If
            allowNextSpace = RC_PadParOutside
          End If
          
        Else
          If RC_PadParInside Then 
            s = s + " "
          End If
          allowNextSpace = RC_PadParOutside
        End If
        
      Case 44 ',
        If isInMSDNCodeBlock Then
          'Replace the , with a _ at the end of the msdn code block line
          s = s + ", _"
          inContinuation = True
          allowNextSpace = False
          skipToken = True
        Else
          allowNextSpace = RC_PadComma
        End If
        
      Case 46 '.
        allowNextSpace = False
        
      Case 42, 43, 45, 47, 60, 61, 62, 92, 94, TOKEN_ASSIGN, TOKEN_GE_RELOP, TOKEN_LE_RELOP, TOKEN_NE_RELOP
        '  *   +   -   /   <   =   >   \   ^   =             >=              <=              <>
        If ((TokenTypeCurrent = 43 And TokenTypeBackOne = 43) Or (TokenTypeCurrent = 45 And TokenTypeBackOne = 45)) And TokenTypeBackTwo = TOKEN_IDENTIFIER Then
          'handle a++ a--
          Debug("Handle a++ tokenChain=" + tokenChain, 2)
          Dim pad As String = If(RC_PadOperators, " ", "")
          Dim start As String = Left(s, Len(s) - (Len(pad) + 1) - Len(tokenChain))
          s = start + tokenChain + pad + "=" + pad + tokenChain + pad + TokenTextCurrent + pad + "1"
          skipToken = True
          
        ElseIf (TokenTypeCurrent = TOKEN_ASSIGN Or TokenTypeCurrent = 61) And (TokenTypeBackOne = 42 Or TokenTypeBackOne = 43 Or TokenTypeBackOne = 45 Or TokenTypeBackOne = 47) Then
          'handle a*=1 a+=1 a-=1 a/=1
          Debug("Handle a+=1 tokenChain=" + tokenChain, 2)
          Dim pad As String = If(RC_PadOperators, " ", "")
          Dim start As String = Left(s, Len(s) - (Len(pad) + 1) - Len(tokenChain))
          s = start + tokenChain + pad + "=" + pad + tokenChain + pad + TokenTextBackOne
          skipToken = True
          
        ElseIf TokenTypeCurrent = 61 And (TokenTypeBackOne = TOKEN_TK_NOT Or TokenTypeBackOne = 33) Then '!
          'convert Not = to <>
          s = Left(s, Len(s) - 3)
          If Right(s, 1) = " " Then s = Left(s, Len(s) - 1)
          If RC_PadOperators Or (RC_PadParOutside And TokenTypeBackTwo = 41) Then
            'put a space between ) and <>
            s = s + " "
          End If
          
          s = s + "<>"
          allowNextSpace = RC_PadOperators
          skipToken = True
          
        Else
          If RC_PadOperators Then
            If TokenTypeBackOne = 40 Then '(
              'Make sure we can do (-1
              If RC_PadParInside Then
                s = s + " "
              End If
              allowNextSpace = False
            Else
              s = s + " "
              allowNextSpace = True
            End If
            
          Else
            If TokenTypeBackOne = 41 And RC_PadParOutside Then
              ')
              s = s + " "
            End If
            
            If TokenTypeBackOne = 40 Then '(
              'Not RC_PadOperators and ( followed by select list above, i.e. if (-
              If RC_PadParInside Then
                s = s + " "
              End If
            ElseIf TokenTypeBackOne = TOKEN_TK_IF Then
              'Always put a space between if - even if RC_PadOperators is false
              s = s + " "
            ElseIf TokenTypeBackOne = TOKEN_TK_RETURN Then
              'Allow Return -1 and Return -a even if RC_PadOperators is false
              s = s + " "
            End If
            
            allowNextSpace = False
          End If
        End If
        
      Case 58 ':
        s = s + If(RC_PadOperators, " :", ":")
        allowNextSpace = RC_PadOperators
        skipToken = True
        
      Case 59 ';
        If isInMSDNCodeBlock And TokenTypeBackOne = 41 Then
          'Remove the ; from the end of the msdn code block and add the return
          If MSDN_Return <> "void" Then
            s = s + " " + RC_Words_Formatted("As") + " "
            'Inject the token so it can be handled by the conversion routine later on
            TokenTextCurrent = "w_" + MSDN_Return
            skipToken = False
            
          Else
            skipToken = True
          End If
          
        Else
          skipToken = False
        End If
        
      Case TOKEN_TK_AS '262
        If APIVersion = 1 And TokenTextHistory(0).Lowercase = "var" And (RC_DimVar = "Dim" Or RC_DimVar = "") Then
          'We have to do this here because API1 doesn't have a token for Var so when we find an As we check if there's a Var at the start
          s = RC_Words_Formatted("Dim") + Right(s, len(s) - 4)
        End If
      
        'Always put a space before an As no matter if it follows a ) like in a Declare
        s = s + " "
        allowNextSpace = True
        
      Case TOKEN_TK_DIM '265
        If APIVersion = 2 And RC_DimVar = "Var" Then
          'APIVersion checked here to stop us going to Var in pre API2 editions
          s = RC_Words_Formatted("Var")
          allowNextSpace = True
          skipToken = True
        End If
		
      Case TOKEN_TK_THEN '267 API2 - 266 API1
        'This is a special case for THEN as it seems to magically appear when hitting SHIFT+ENTER on an IF
        If s = "" Then s = s + " "
        s = s + " " + RC_Words_Formatted("Then")
        allowNextSpace = True
        skipToken = True
		
      Case 266 'TOKEN_TK_VAR '266 'All commented values after this are -1 when running on API1 not too confusing
        'We have to put this out of order and use a hard coded value of 266 because
        'a) TOKEN_TK_VAR doesn't exist in API1 if we checked against that in pre 2019r2 versions the script would fail
        'b) When 266 is actually hit above in API1 we don't really care about this one here because we find Var using TOKEN_TK_AS in API1
        'c) Xojo decided to inject 266, I assume to keep things neat at their end but it doesn't make my life easy, thanks!
        If APIVersion = 2 Then
          If RC_DimVar = "Dim" Then
            s = RC_Words_Formatted("Dim")
            allowNextSpace = True
            skipToken = True
          End If
        End If
        		
      Case TOKEN_TK_MOD, TOKEN_TK_AND, TOKEN_TK_OR, TOKEN_TK_XOR, TOKEN_TK_TO '264 277 278 345 283
        'Always put a space before these as they will end up merging if we don't
        s = s + " "
        allowNextSpace = True
        
      Case TOKEN_TK_STEP '298
        'Always put a space before a Step, it looks funny if we don't
        s = s + " "
        allowNextSpace = True
        
      Case TOKEN_TK_ISA '299
        'Always put a space before an IsA no matter if it follows a ) with RC_PadPasOutside false
        s = s + " "
        allowNextSpace = True
        
      Case TOKEN_IDENTIFIER '363
        'Part of the conversion of #define ABC 0x0001 To Const ABC = &h0001
        If TokenTypeBackOne = 35 And TokenTextBackOne = "#define" Then
          s = Left(s, Len(s) - 7) + RC_Words_Formatted("Const") + " " + TokenText + If(RC_PadOperators, " =", "=")
          allowNextSpace = RC_PadOperators
          skipToken = True
          
        ElseIf Left(TokenTextCurrent, 1) = "x" And TokenTypeBackOne = TOKEN_NUMBER And TokenTextBackOne = "0" Then
          'Convert hex formats from 0x0001 to &h0001
          If Right(TokenTextCurrent, 1) = "L" Then
            'Remove the L from the end of the hex value denoting a LONG const, this doesn't matter in Xojo
            TokenTextCurrent = Left(TokenTextCurrent, Len(TokenTextCurrent) - 1)
          End If
          s = Left(s, Len(s) - 1) + "&h" + Right(TokenTextCurrent, Len(TokenTextCurrent) - 1)
          skipToken = True
          
        ElseIf TokenTypeBackOne = 45 Then '-
          If RC_PadOperators Then
            Select Case TokenTypeBackTwo
            Case 44 ',
              'Drop the space after the - if we have , -a
            Case TOKEN_TK_IF '258
              If TokenTypeBackOne = 45 Then '-
                'if -a
                'Drop the space
              Else
                s = s + " "
              End If
              
            Case 42, 43, 45, 47, 60, 61, 62, 92, 94, TOKEN_ASSIGN, TOKEN_GE_RELOP, TOKEN_LE_RELOP, TOKEN_NE_RELOP, TOKEN_TK_NOT
              '  *   +   -   /   <   =   >   \   ^    =            >=              <=              <>              NOT
              'Drop the space if we follow the above so we can do a = a - -a
              
            Case TOKEN_TK_RETURN '291
              'Drop the space after a - that is following a Return so we can
              'Return -a
              
            Case Else
              If allowNextSpace Then
                s = s + " "
              End If
            End Select
          Else
            If allowNextSpace Then
              s = s + " "
            End If
          End If
          
        Else
          If isInMSDNCodeBlock = True Then
            'Tokens found inside an MSDN block
            If MSDN_Return = "" Then
              MSDN_Return = TokenTextCurrent
              skipToken = True
              
            ElseIf MSDN_Call = "" Then
              MSDN_Call = TokenTextCurrent
              skipToken = True
              
            ElseIf MSDN_Type = "" Then
              MSDN_Type = TokenTextCurrent
              Debug("MSDN_Type=>" + MSDN_Type, 2)
              skipToken = True
              
            ElseIf MSDN_Var = "" Then
              MSDN_Var = TokenTextCurrent
              Debug("MSDN_Var=>" + MSDN_Type, 2)
              s = s + MSDN_Var + " " + RC_Words_Formatted("As") + " "
              'Inject the token so it can be handled by the conversion routine later on
              TokenTextCurrent = "w_" + MSDN_Type
              inContinuation = True
              skipToken = False
            End If
            
          Else
            If TokenTextCurrent = "rc_version" Then
              s = s + "'Reformat Code Script Version v" + RC_Version
              skipToken = True
            Else
              'All generic tokens pass through here
              Debug("Inside generic token fallback ********", 2)
              If allowNextSpace Then
                s = s + " "
              End If
              allowNextSpace = True  
            End If
            
          End If
          
        End If
        
      Case TOKEN_STRING '365
        If allowNextSpace Then
          s = s + " "
        End If
        
        If Left(TokenTextCurrent, 2) = "&u" Then
          'Special case for &u because its considered a string!
          Dim found As Integer = InStr(lastUnicodeLiteral, Line(currentLineNumber), TokenTextCurrent)
          If found > 0 Then
            'We found an unicode literal...
            lastUnicodeLiteral = found + Len(TokenTextCurrent)
            If Mid(Line(currentLineNumber), found - 1, 1) = """" Then
              '...that had a quote before it
              s = s + """" + TokenTextCurrent + """"
            Else
              '...that had no quote before it
              s = s + TokenTextCurrent
            End If
          End If
          
        Else
          s = s + """" + ReplaceAll(TokenTextCurrent, """", """""") + """"
        End If
        skipToken = True
        
      Case TOKEN_UNMATCHEDQUOTES '374
        'Handle Unmatched Quotes
        
        Dim tmp As String = Line(currentLineNumber)
        If RC_MessageQuoteMismatched <> "" Then
          If tmp.instr(expandedMessageQuoteMismatched) = 0 Then
            'Add an error message onto the end of the line so we notice the problem if there isn't one there already
            tmp = tmp + expandedMessageQuoteMismatched
          End If
          Line(currentLineNumber) = tmp
        End If
        
        'Do nothing with the line if quotes are mismatched as we could end up borking things
        'Break out of the current line and move onto the next, this line makes me feel dirty
        Continue For
        
      Case TOKEN_LINE_CONTINUATION '376
        inContinuation = True
        s = s + If(RC_PadLineContinuation, " _", "_")
        skipToken = True
        
      Case Else
        Debug("Inside last Else ********", 2)
        If allowNextSpace Then
          If TokenTypeBackOne = 45 And TokenTypeBackTwo = TOKEN_TK_STEP Then '-
            'Leave out the space between - and 1 after a Step
            
          ElseIf TokenTypeBackTwo = TOKEN_ASSIGN And TokenTypeBackOne = 45 Then '-
            'Leave out the space between - and 1 if after a = so we can do this a = -1
            
          ElseIf TokenTypeBackOne = 45 And (TokenTypeCurrent = TOKEN_NUMBER Or TokenTypeCurrent = TOKEN_REALNUMBER) Then '-
            '-1 or -1.0
            If TokenTypeBackTwo = 41 Then ')
              'change ) -1 to ) - 1 or ) -1.0 to ) - 1.0
              s = s + " "
              
            ElseIf TokenTypeBackTwo = TOKEN_IDENTIFIER Then
              'change a -1 to a - 1 or a -1.0 to a - 1.0
              s = s + " "
              
            Else
              Debug("Inside last negative numeric fallback ********", 2)
              If RC_PadOperators Then
                Select Case TokenTypeBackTwo
                Case TOKEN_TK_IF, 44 ',
                  If TokenTypeBackOne = 45 Then '-
                    'if -1 or , -1
                    'Drop the space
                  Else
                    s = s + " "
                  End If
                  
                Case 42, 43, 45, 47, 60, 61, 62, 92, 94, TOKEN_ASSIGN, TOKEN_GE_RELOP, TOKEN_LE_RELOP, TOKEN_NE_RELOP, TOKEN_TK_NOT
                  '  *   +   -   /   <   =   >   \   ^    =            >=              <=              <>              NOT
                  'Drop the space if we follow the above so we can do a = 1 - -1
                  
                Case TOKEN_TK_RETURN '290
                  'Drop the space after a - that is following a Return so we can
                  'Return -1
                  
                Case Else
                  s = s + " "
                End Select
                
              End If
              
            End If
          Else
            Debug("Inside last generic fallback ########", 2)
            If allowNextSpace Then
              s = s + " "
            End If
            
          End If
        End If
        allowNextSpace = True
      End Select
      
      Debug("skipToken=" + If(skipToken, "True", "False"), 2)
      
      If Not skipToken Then
        
        Debug("inside Not SkipToken s=>" + s + "<", 2)
        
        Dim found As Integer = RC_Words.IndexOf(TokenTextCurrent)
        
        If found > -1 Then
          Debug("Found >" + TokenTextCurrent + "< inside RC_Words", 2)
          TokenTextCurrent = RC_Words(found) 
          s = s + RC_Words_Formatted(TokenTextCurrent)
        Else
          'If there's something we dont know (like a variable) then we just insert it without modification
          Debug("Found FAILED >" + TokenTextCurrent + "< inside RC_Words", 2)
          
          Dim cont As Boolean = False
          Dim pos As Integer
          
          'Implement RC_ReplaceList
          'Find the first occurance of a word in the list and use the next word in the list as a replacement
          Do
            pos = RC_ReplaceList.IndexOf(TokenTextCurrent, pos)
            If pos = -1 Then
              Debug("Found FAILED >" + TokenTextCurrent + "< inside RC_ReplaceList", 2)
              s = s + TokenTextCurrent
              cont = True
            Else
              Debug("Found >" + TokenTextCurrent + "< inside RC_ReplaceList", 2)
              If pos Mod 2 = 0 Then
                'Make sure the word is in the find slot (first of the pair)
                Select Case RC_ForceCase
                Case "Uppercase"
                  s = s + RC_ReplaceList(pos + 1).Uppercase
                Case "Lowercase"
                  s = s + RC_ReplaceList(pos + 1).Lowercase
                Case Else
                  s = s + RC_ReplaceList(pos + 1)
                End Select
                cont = True
              Else
                'If we found a word in the replace slot then skip it
                pos = pos + 1
              End If
            End If
          Loop Until cont = True
          
        End If
        
      Else
        skipToken = False
      End If
      
      'Keep track of the current tokenChain
      If TokenTypeCurrent = TOKEN_IDENTIFIER Or TokenTypeCurrent = 46 Or TokenTypeCurrent = TOKEN_TK_ME Or TokenTypeCurrent = TOKEN_TK_SELF Then '362 or . or Me or Self
        If tokenChainClear Then
          tokenChain = ""
          tokenChainClear = False
        End If
        tokenChain = tokenChain + TokenTextCurrent
        
      Else
        tokenChainEnded = True
      End If
      
      If Left(tokenChain, Len(RC_MacroStorageLocation)) = RC_MacroStorageLocation Then
        'Check if our tokenChain builder has RC_MacroStorageLocation at the start i.e. macro.test.test2
        'If it does then we're in a macro
        Debug("RENDER MACRO=|" + tokenChain + "| s=|" + s + "|", 2)
        If RenderMacro(tokenChain, s) Then
          'Render the macro and reset tokenChain as we've just replaced it
          tokenChain = ""
        End If
      End If
      
      If tokenChainEnded Then
        Debug("tokenChainEnded=|" + tokenChain + "|", 2)
        'tokenChain = ""
        tokenChainEnded = False
        tokenChainClear = True
      End If
      
    Wend
    
    If isInMSDNCodeBlock And s = "" Then
      'Handle empty lines in msdn code blocks, usually because the call has no parameters
      s = s + "_"
      inContinuation = True
    End If
    
    If isInMSDNCodeBlock And MSDN_Type <> "" And Right(s, 2) <> " _" Then
      'Add the _ on the last line before the ) of the msdn code block
      s = s + " _"
      inContinuation = True
    End If
    
    'Remove any error messages that exist at the end of the line
    Dim rl As String = RemainingLine()
    rl = rl.RemoveString(expandedMessageParMismatched)
    rl = rl.RemoveString(expandedMessageParOpening)
    rl = rl.RemoveString(expandedRC_MessageParClosing)
    rl = rl.RemoveString(expandedMessageQuoteMismatched)
    rl = rl.RemoveString(expandedMessageInCodeBlock)
    
    'Check for out of order )( or missing ()
    If ((parCount < 0 And RC_MessageParOpening <> "") Or (parCount > 0 And RC_MessageParClosing <> "") Or (parMismatch And RC_MessageParMismatched <> "" And parCount = 0)) And Not inContinuation And Not wasInContinuation Then
      Debug("MISMATCH", 1)
      Debug("parCount=" + str(parCount), 2)
      Debug("parMismatch=" + If(parMismatch, "True", "False"), 2)
      Debug("expandedMessageParOpening=|" + expandedMessageParOpening + "| expandedRC_MessageParClosing=|" + expandedRC_MessageParClosing + "| expandedMessageParMismatched=|" + expandedMessageParMismatched + "|", 2)
      
      If parCount > 0 Then
        s = s + expandedRC_MessageParClosing
      ElseIf parCount < 0 Then
        s = s + expandedMessageParOpening
      ElseIf (parCount = 0 And parMismatch) Then
        s = s + expandedMessageParMismatched
      End If
      
    End If
    
    'Add any left over characters to the line, like after a comment
    RenderRemainingLine(rl, s)
    
    Debug("***BEFORE inContinuation=" + If(inContinuation, "True", "False") + " wasInContinuation=" + If(wasInContinuation, "True", "False"), 1)
    
    If inContinuation Or wasInContinuation Then
      parCountBlock = parCountBlock + parCount
      Debug("parCountBlock=" + str(parCountBlock), 2)
      Debug("parMismatchBlock=" + str(parCountBlock), 2)
      
      If Not inContinuation And wasInContinuation Then
        'Just left a continuation so we can finish up
        If parCountBlock > 0 Then
          s = s + expandedRC_MessageParClosing + RC_MessageInCodeBlock
        ElseIf parCountBlock < 0 Then
          s = s + expandedMessageParOpening + RC_MessageInCodeBlock
        ElseIf (parCountBlock = 0 And parMismatchBlock) Then
          s = s + expandedMessageParMismatched + RC_MessageInCodeBlock
        End If
        
      End If
      
    End If
    
    'strip off the leading space, doing it this way keeps the code cleaner as we dont have to put it all in a huge if
    If Left(s, 1) = " " Then
      s = Right(s, Len(s) - 1)
    End If
    
    Debug("RENDERING LINE=|" + s + "|", 1)
    
    'Put the line that we've built into the IDE
    If s <> "" Then
      'We have to put this in an If as there is as a silent exception raised if we don't that we can't recover from <feedback://showreport?report_id=55859>
      Line(currentLineNumber) = s
    End If
    
    If inContinuation Then
      wasInContinuation = True 'remember that the last line was a continuation
      inContinuation = False 'reset this for the next loop around
    Else
      wasInContinuation = False 'if the current line didnt have a continuation then we remember that for the next line
    End If
    
    Debug("***AFTER inContinuation=" + If(inContinuation, "True", "False") + " wasInContinuation=" + If(wasInContinuation, "True", "False"), 1)
    
  Next
End Sub

'DebugLog handler
Sub Debug(message As String, level As Integer)
  'Note: DebugLog does not work in 2019r1 & 2019r1.1 due to #55388 - I put this comment here so I would stop forgetting this
  If RC_DebugLevel >= level Then DebugLog(message)
End Sub

Sub ReadPreference_ConstStorageLocation(name As String, ByRef value As String, defaultLocaiton As String)
  'Read in the setting and remove a . if its not got one
  value = ConstantValue(defaultLocaiton + name)
  'value = ConstantValue(defaultLocaiton + RC_Prefix + name)
  If Right(value, 1) = "." Then
    value = Left(value, Len(value) - 1)
  End If
  
  'If the setting doesnt exist above then it will be set to App.
  If value = "" Then
    value = "App"
  End If
  
  'If the setting is Me. then it will need to be blanked to read from the correct area
  If value = "Me" Then
    'value = "" + RC_Prefix
    value = ""
  Else
    'value = value + "." + RC_Prefix
    value = value + "."
  End If
End Sub

'ReadPreference for Boolean
Sub ReadPreference(name As String, ByRef value As Boolean, Optional defaultValue As Boolean = False)
  Select Case ConstantValue(RC_ConstStorageLocation + RC_Prefix + name)
  Case "1", "Yes", "True"
    Debug("ReadPreference " + RC_ConstStorageLocation + RC_Prefix + name + "=True", 1)
    value = True
    
  Case "0", "No", "False"
    Debug("ReadPreference " + RC_ConstStorageLocation + RC_Prefix + name + "=False", 1)
    value = False
    
  Case Else
    Debug("ReadPreference " + RC_ConstStorageLocation + RC_Prefix + name + "=" + If(defaultValue, "True", "False") + " (NOT FOUND USED DEFAULT)", 1)
    value = defaultValue
    
  End Select
End Sub

'ReadPreference for Integer
Sub ReadPreference(name As String, ByRef value As Integer, Optional default As Integer = 0)
  If ConstantValue(RC_ConstStorageLocation + RC_Prefix + name) <> "" Then
    value = Val(ConstantValue(RC_ConstStorageLocation + RC_Prefix + name))
    Debug("ReadPreference " + RC_ConstStorageLocation + RC_Prefix + name + "=" + Str(value), 1)
  Else
    value = default
    Debug("ReadPreference " + RC_ConstStorageLocation + RC_Prefix + name + "=" + Str(value) + " (NOT FOUND USED DEFAULT)", 1)
  End If
End Sub

'ReadPreference for String
Sub ReadPreference(name As String, ByRef value As String, Optional default As String = "")
  If ConstantValue(RC_ConstStorageLocation + RC_Prefix + name) <> "" Then
    If ConstantValue(RC_ConstStorageLocation + RC_Prefix + name) = "_" Then
      value = ""
      Debug("ReadPreference " + RC_ConstStorageLocation + RC_Prefix + name + "=_ (EMPTY_STRING)", 1)
    Else
      value = ConstantValue(RC_ConstStorageLocation + RC_Prefix + name)
      Debug("ReadPreference " + RC_ConstStorageLocation + RC_Prefix + name + "=" + value, 1)
    End If
    
  Else
    value = default
    Debug("ReadPreference " + RC_ConstStorageLocation + RC_Prefix + name + "=" + value + " (NOT FOUND USED DEFAULT)=>" + value + "<", 1)
  End If
End Sub

'Allow us to look back along the token type path
Function GetTokenType(ByRef TokenTypeHistory() As Integer, reversePos As Integer) As Integer
  If TokenTypeHistory.UBound > (reversePos - 1) Then
    Return TokenTypeHistory(TokenTypeHistory.UBound - reversePos)
  Else
    Return 0
  End If
End Function

'Allow us to look back along the token text path
Function GetTokenText(ByRef TokenTextHistory() As String, reversePos As Integer) As String
  If TokenTextHistory.UBound > (reversePos - 1) Then
    Return TokenTextHistory(TokenTextHistory.UBound - reversePos)
  Else
    Return ""
  End If
End Function

Sub RenderRemainingLine(rl As String, ByRef s As String)
  'If the line ends in a comment you will not get the token for a comment, but you will get the comment text in RemainingLine.
  If rl <> "" Then
    Debug("RemainingLine=|" + rl + "|", 1)
    
    Dim commentPos As Integer
    Dim comment As String
    Dim commentLen As Integer
    
    'Find the first comment in the RemainingLine
    Dim aPos As Integer = rl.InStr("'")
    Dim sPos As Integer = rl.InStr("//")
    Dim rPos As Integer = rl.InStr("REM ")
    
    Debug("aPos=" + Str(aPos), 2)
    Debug("sPos=" + Str(sPos), 2)
    Debug("rPos=" + Str(rPos), 2)
    
    If aPos > 0 And (aPos < sPos Or sPos = 0) And (aPos < rPos Or rPos = 0) Then
      commentPos = aPos + 1
      comment = If(RC_Comment = "", "'", RC_Comment)
      commentLen = 1
    ElseIf sPos > 0 And (sPos < aPos Or aPos = 0) And (sPos < rPos Or rPos = 0) Then
      commentPos = sPos + 2
      comment = If(RC_Comment = "", "//", RC_Comment)
      commentLen = 2
    ElseIf rPos > 0 And (rPos < aPos Or aPos = 0) And (rPos < sPos Or sPos = 0) Then
      commentPos = rPos + 4
      If RC_Comment = "" Then
        comment = RC_Words_Formatted("Rem")
      Else
        comment = RC_Comment
      End If
      commentLen = 4
    End If
    
    If comment = "rem" Then
      'Always surround Rem so we don't end up merging it with another word
      RC_PadCommentBefore = "true"
      RC_PadCommentAfter = "true"
    End If
    
    If commentPos > 0 Then
      'We found a comment
      Dim pre As String = Left(rl, commentPos - 1 - commentLen)
      Dim post As String = Right(rl, Len(rl) - commentPos + 1)
      
      Debug("pre=|" + pre + "|", 2)
      Debug("post=|" + post + "|", 2)
      
      Dim preSpace As String = ""
      Dim postSpace As String = ""
      
      'PRE
      If Right(pre, 1) = " " Then
        If RC_PadCommentBefore = "true" Or RC_PadCommentBefore = "yes" Or RC_PadCommentBefore = "1" Or RC_PadCommentBefore = "" Then
          'There's a space there already and we want it there
          preSpace = ""
        Else
          'Remove the space that is already there
          pre = Left(pre, Len(pre) - 1)
        End If
        
      Else
        If RC_PadCommentBefore = "true" Or RC_PadCommentBefore = "yes" Or RC_PadCommentBefore = "1" Then
          'There's no space there and we want one
          preSpace = " "
        End If
        
      End If
      
      'POST
      If Left(post, 1) = " " Then
        If RC_PadCommentAfter = "true" Or RC_PadCommentAfter = "yes" Or RC_PadCommentAfter = "1" Or RC_PadCommentAfter = "" Then
          'There's a space there already and we want it there
          postSpace = ""
        Else
          'Remove the space that is already there
          post = Right(post, Len(post) - 1)
        End If
        
      Else
        If RC_PadCommentAfter = "true" Or RC_PadCommentAfter = "yes" Or RC_PadCommentAfter = "1" Then
          'There's no space there and we want one
          postSpace = " "
        End If
      End If
      
      s = s + pre + preSpace + comment + postSpace + post
      
    Else
      s = s + rl
    End If
    
  End If
End Sub

'Render out the macro that is stored in the macroText ConstantValue
Function RenderMacro(macroText As String, ByRef s As String) As Boolean
  If ConstantValue(macroText) <> "" Then
    
    Dim macroContents As String = ConstantValue(macroText)
    
    If Left(macroContents, 1) = "'" Or Left(macroContents, 2) = "//" Then 'Left out REM so it doesn't clash with some words, add it here at your own risk
      'Allow a macro to have a comment on the first line that shows up in auto complete but isnt inserted into the code
      Dim newlinePos As Integer = InStr(macroContents, Chr(13))
      
      If newlinePos > 0 Then
        macroContents = Right(macroContents, Len(macroContents) - newlinePos)
      End If
      
    End If
    
    s = Left(s, Len(s) - Len(macroText)) + macroContents
    
    Return True
    
  End If
  
  Return False
  
End Function

'Merge two arrays of string into the first without creating a new array and optionally clean up the old array
Sub Concat(Extends ByRef toString() As String, ByRef fromString() As String, Optional clearFromArray As Boolean = True)
  'Store the size of the to array so we know where we need to start placing new data as the size is about to change
  Dim toSize As Integer = toString.Ubound + 1
  
  'Resize the to array to accommodate the new array
  Redim toString(toString.Ubound + fromString.Ubound + 1)
  
  'Overwrite the newly allocated space with the from array data
  For n As Integer = 0 To fromString.Ubound
    toString(toSize + n) = fromString(n)
  Next
  
  If clearFromArray Then
    'Clear the from array as we don't need it any more
    Redim fromString(-1)
  End If
End Sub

'Remove a string from inside another string, can remove first or all
Function RemoveString(Extends source As String, remove As String, removeAllMatches As Boolean = True) As String
  Dim c As Integer
  Dim s As String = source
  Do
    c = s.instr(remove)
    If c = 0 Then Return s 'no more hits
    s = left(s, c - 1) + right(source, len(s) - c - len(remove) + 1)
    If Not removeAllMatches Then Return s 'return after one hit otherwise we run until we get them all
  Loop Until False
End Function

Sub DebugArray(ByRef arr() As String, level As Integer)
  If RC_DebugLevel >= level Then
    Debug("DEBUGGING ARRAY -----START-----", 2)
    For n As Integer = 0 To arr.Ubound
      Debug(">" + arr(n) + "<", 2)
    Next
    Debug("DEBUGGING ARRAY ------END------", 2)
  End If
End Sub

Function RC_Words_Formatted(word As String) As String
  Select Case RC_ForceCase
  Case "Uppercase"
    Return RC_Words(RC_Words.IndexOf(word)).Uppercase
  Case "Lowercase"
    Return RC_Words(RC_Words.IndexOf(word)).Lowercase
  Case Else
    Return RC_Words(RC_Words.IndexOf(word))
  End Select
End Function
