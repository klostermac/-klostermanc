Option Explicit
Dim objRootDSE, strConfig, objConnection, objCommand, strQuery
Dim objRecordSet, objDC
Dim strDNSDomain, objShell, lngBiasKey, lngBias, k, arrstrDCs()
Dim strDN, dtmDate, objDate, lngDate, objList, strUser, strDomain
Dim strBase, strFilter, strAttributes, lngHigh, lngLow

'function to determine if account is disabled or not
'returns boolean value of true if the account is disabled, false if it is not
Function IsAccountDisabled( strDomain, strAccount )
   Dim objUser
   Set objUser = GetObject("WinNT://" & strDomain & "/" & strAccount & ",user")
   IsAccountDisabled = objUser.AccountDisabled
End Function


' Create dictionary to keep track of
' users and login times

Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare

' rmueller's script
' Obtain local Time Zone bias from machine registry.

Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
& "TimeZoneInformation\ActiveTimeBias")
If UCase(TypeName(lngBiasKey)) = "LONG" Then
lngBias = lngBiasKey
ElseIf UCase(TypeName(lngBiasKey)) = "VARIANT()" Then
lngBias = 0
For k = 0 To UBound(lngBiasKey)
   lngBias = lngBias + (lngBiasKey(k) * 256^k)
Next
End If

' Determine configuration context and DNS domain from RootDSE object.

Set objRootDSE = GetObject("LDAP://RootDSE")
strConfig = objRootDSE.Get("configurationNamingContext")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use ADO to search Active Directory for ObjectClass nTDSDSA.
' This will identify all Domain Controllers.

Set objCommand = CreateObject("ADODB.Command")
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
objCommand.ActiveConnection = objConnection
strBase = "<LDAP://" & strConfig & ">"
strFilter = "(objectClass=nTDSDSA)"
strAttributes = "AdsPath"
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
objCommand.CommandText = strQuery
objCommand.Properties("Page Size") = 100
objCommand.Properties("Timeout") = 60
objCommand.Properties("Cache Results") = False
Set objRecordSet = objCommand.Execute

' Enumerate parent objects of class nTDSDSA. Save Domain Controller
' AdsPaths in dynamic array arrstrDCs.

k = 0
Do Until objRecordSet.EOF
Set objDC = _
   GetObject(GetObject(objRecordSet.Fields("AdsPath")).Parent)
ReDim Preserve arrstrDCs(k)
arrstrDCs(k) = objDC.DNSHostName
k = k + 1
objRecordSet.MoveNext
Loop

' Retrieve lastLogon attribute for each user on each Domain Controller.

For k = 0 To Ubound(arrstrDCs)
strBase = "<LDAP://" & arrstrDCs(k) & "/OU=HP,OU=User Accounts,DC=crmprod,DC=w2k,DC=dtv,DC=cxo,DC=dec,DC=com>"
strFilter = "(&(objectCategory=person)(objectClass=user))"

'added sAMAccount name for readability in output - el

strAttributes = "distinguishedName,lastLogon,sAMAccountName"
strQuery = strBase & ";" & strFilter & ";" & strAttributes _
   & ";subtree"
objCommand.CommandText = strQuery
On Error Resume Next
Set objRecordSet = objCommand.Execute
If Err.Number <> 0 Then
   On Error GoTo 0
   Wscript.Echo "Domain Controller not available: " & arrstrDCs(k)
Else
   On Error GoTo 0
   Do Until objRecordSet.EOF
     strDN = objRecordSet.Fields("sAMAccountName")
     lngDate = objRecordSet.Fields("lastLogon")
   
     On Error Resume Next
     Set objDate = lngDate
     If Err.Number <> 0 Then
       On Error GoTo 0
       dtmDate = #1/1/1601#
     Else
       On Error GoTo 0
       'stored as a 64 bit integer that VBScript cannot handle, must
       'seperate it to work with it.
       lngHigh = objDate.HighPart
       lngLow = objDate.LowPart
       If lngLow < 0 Then
         lngHigh = lngHigh + 1
       End If
       If (lngHigh = 0) And (lngLow = 0 ) Then
         dtmDate = #1/1/1601#
       Else
         dtmDate = #1/1/1601# + (((lngHigh * (2 ^ 32)) _
           + lngLow)/600000000 - lngBias)/1440
       End If
     End If
     If objList.Exists(strDN) Then
       If dtmDate > objList(strDN) Then
         objList(strDN) = dtmDate
       End If
     Else
       objList.Add strDN, dtmDate
     End If
     objRecordSet.MoveNext
   Loop
End If
Next
Dim objNetwork,objExcel
dim intRow
intRow = 2
Set objNetwork = Wscript.CreateObject("Wscript.Network")
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

objExcel.Cells(1,1).value = "UserName"
objExcel.Cells(1,2).value = "Last Login"
objExcel.Cells(1,3).value = "Disabled"

' Output latest lastLogon date for each user.
Dim easydate, iDiff,strDays
strDays = InputBox("Enter number of days: ","Last Login", "60")
strDays = CInt(strDays)
For Each strUser In objList

'get rid of the timestamp, only need the date
easydate = objList(strUser)
easydate = FormatDateTime(easydate,2)

'calculate the difference between last login date
'and now.
iDiff = DateDiff("d",easydate,Now)

' check to see if account has been inactive for x number of days
If iDiff >= strDays Then

'output username and last login date
' change 1/1/1601 to Never because it looks nicer
If easydate = "1/1/1601" Then
 easydate = "Never"
End If 

'Else WScript.Echo strUser & " ; " & easydate
'End If
objExcel.Cells(intRow,1).value = strUser
objExcel.Cells(intRow,2).value = easydate

'populated disabled column
if IsAccountDisabled("CRMPROD",strUser) Then objExcel.Cells(intRow,3).value = "DISABLED"' End if
intRow = intRow + 1
End If

Next

'sort by login Date
Dim obrange, obrange2
Set obrange = objExcel.Range("A:K")
Set obrange2 = objExcel.Range("B2")
obrange.Sort obrange2,1,,,,,,1


' Clean up.
objConnection.Close
Set objRootDSE = Nothing
Set objConnection = Nothing
Set objCommand = Nothing
Set objRecordSet = Nothing
Set objDC = Nothing
Set objDate = Nothing
Set objList = Nothing
Set objShell = Nothing
Set obrange = Nothing
Set obrange2 = Nothing
Set objExcel = Nothing
WScript.Quit
