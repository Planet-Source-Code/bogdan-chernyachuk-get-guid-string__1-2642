<div align="center">

## Get GUID string


</div>

### Description

Generates Global Unique ID - when you want to create an identifier which will never repeat all over the word , you will find this usefull.
 
### More Info
 
Returns Global Unique ID in string format (the same as the format of ClassID written in registry).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bogdan Chernyachuk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bogdan-chernyachuk.md)
**Level**          |Unknown
**User Rating**    |4.7 (52 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bogdan-chernyachuk-get-guid-string__1-2642/archive/master.zip)

### API Declarations

```
Public Type GUID ' a structure for Global Uniq. ID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0, 7) As Byte
End Type
Declare Function CoCreateGuid Lib "ole32" (ByRef lpGUID As GUID) As Long
Declare Function StringFromGUID2 Lib "ole32" (ByRef lpGUID As GUID, ByVal lpStr As String, ByVal lSize As Long) As Long
```


### Source Code

```
Public Function GetNewGUIDStr() As String
Dim pGuid As GUID
Dim lResult As Long
Dim s As String
  'this is a buffer string to be passed in API function
  '100 chars will be enough
  s = String(100, " ")
  'creating new ID and obtaining result in pointer to GUID
  lResult = CoCreateGuid(pGuid)
  'converting GUID structure to string
  lResult = StringFromGUID2(pGuid, s, 100)
  'removing all trailing blanks
  s = Trim(s)
  'converting a sting from unicode
  GetNewGUIDStr = StrConv(s, vbFromUnicode)
End Function
```

