'JPG EXIF to XML
'Takes a hierarchal directory of geotagged photos and makes an XML file describing them

'Specifically to read geotags from JPGs and write to album.xml for MultiLevelPhotoMap PHP program, http://tympanus.net/codrops/2011/09/27/multi-level-photo-map/
'reads GPS tags of EXIF data using Windows Image Acquisition library (the .dll should be installed in Windows version Vista and above)
'Attribution: EXIF reading is based on a VB6 program by "dilettante": http://www.vbforums.com/showthread.php?752887-VB6-Extract-JPEG-EXIF-Data-with-WIA-2-0
'             Some of my code (the WIA objects and Select..Case structure) remain directly copied from their code, which was posted without license information, on a public forum.
'My code may be used freely under a Creative Commons Attribution Share Alike License  http://creativecommons.org/licenses/by-sa/4.0/

'XML structure that this code formats as:
'<album>
'	<name>Thailand 2011</name>
'	<description>Some description</description>
'	<places>
    '<place>
    '	<name>Bangkok</name>		
    '	<location>
    '		<lat>13.696693336737654</lat>
    '		<lng>100.57159423828125</lng>
    '	</location>
    '	<photos>
    '		<photo>
    '			<thumb>photos/Bangkok/thumbs/1.jpg</thumb>
    '			<source>photos/Bangkok/1.jpg</source>
    '			<description>Some description</description>
    '			<location>
    '				<lat>13.710035342476681</lat>
    '				<lng>100.52043914794922</lng>
    '			</location>
    '		</photo>
    '		<photo>
    '			...
    '		</photo>
    '	</photos>
    '</place>
'	</places>
'</album>



Option Explicit

'FileSystemObject
Dim FSO
Dim objFolder
Dim objSubFolder
Dim objFile
Dim strFolderPath
Dim strThumbsDir
Dim strExt
Dim strThumbPostfix
Dim operatingSystemFolderSlash
Dim strNotGeotagged
Dim fileNotGeotagged
Dim flagAnyFileNotGeotagged
Dim strFileXML
Dim fileXML
Dim strIndent

'properties for user to change
strFileXML = "album.xml"
strFolderPath = "photos"
strThumbsDir = "thumbs"
strExt = "jpg"
strThumbPostfix = "_s"
strNotGeotagged = "log_not_geotagged.txt" 'written only if there are files that the user must go back and add manual geotags later
flagAnyFileNotGeotagged = false
operatingSystemFolderSlash = "/"

'set up file objects
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = FSO.GetFolder(strFolderPath)
'backup existing "album.xml" file
If FSO.FileExists(strFileXML) Then FSO.CopyFile strFileXML, strFileXML & "." & Year(Now()) & Month(Now()) & Day(Now()) & Hour(Now()) & Minute(Now()) & Second(Now()) & ".bak"
Set fileNotGeotagged = FSO.OpenTextFile(strNotGeotagged, 2, True)
Set fileXML = FSO.OpenTextFile(strFileXML, 2, True)
strExt = LCase(strExt)

'WIA
Dim objWIAImageFile
Set objWIAImageFile = CreateObject("WIA.ImageFile")
Dim propEach

'GPS
Dim decLat, decLng
Dim foundNS, foundEW, foundLat, foundLng
Dim latFactor, lngFactor

        
strIndent = "" 'stores the current indentation "prefix", ie. "  " two spaces or "    " four spaces, etc
fileXML.WriteLine strIndent & "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
fileXML.WriteLine strIndent & "<album>"
strIndent = indent(strIndent)
fileXML.WriteLine strIndent & "<name></name>"
fileXML.WriteLine strIndent & "<description></description>"
fileXML.WriteLine strIndent & "<places>"
strIndent = indent(strIndent)

For Each objSubFolder in objFolder.SubFolders
  fileXML.WriteLine strIndent & "<place>"
  strIndent = indent(strIndent)
  fileXML.WriteLine strIndent & "<name>" & objSubFolder.Name & "</name>"
  fileXML.WriteLine strIndent & "<location>"
  strIndent = indent(strIndent)
  fileXML.WriteLine strIndent & "<lat></lat>"
  fileXML.WriteLine strIndent & "<lng></lng>"
  strIndent = outdent(strIndent)
  fileXML.WriteLine strIndent & "</location>"
  fileXML.WriteLine strIndent & "<photos>"
  strIndent = indent(strIndent)

  For Each objFile in objSubFolder.Files
    'MsgBox "Name: """ & objFile.Name & """, ParentFolder: """ & objFile.ParentFolder & """"
    'MsgBox "objFolder.Name: """ & objFolder.Name & """" & "objSubFolder.Name: """ & objSubFolder.Name & """" & "objFileName: """ & objFile.Name & """"
    
    If LCase(FSO.GetExtensionName(objFile.Name)) = strExt Then
      If Not FSO.FileExists(objFile.ParentFolder & "\" & objFile.Name) Then
        MsgBox "File doesn't exist! " & vblf & objFile.ParentFolder & "\" & objFile.Name
      Else
        objWIAImageFile.LoadFile objFile.ParentFolder & "\" & objFile.Name
        'On Error Resume Next
        If Err.Number <> 0 Then
          MsgBox "Error (" & CStr(Err.Number) & ") in " & Err.Source
          MsgBox Err.Description
          Err.Clear
          Wscript.quit
        End If
        
        foundNS = false
        foundEW = false
        foundLat = false
        foundLng = false
        latFactor = 1
        lngFactor = 1
        
        For Each propEach In objWIAImageFile.Properties
          Select Case propEach.Name
            Case "GpsLatitudeRef"
              foundNS = true
              If propEach.Value = "S" Then latFactor = -1
              'MsgBox "GpsLatitudeRef = " & propEach.Value & ", latFactor: " & latFactor
            Case "GpsLatitude"
              'msgbox "Found GpsLatitude"
              'strLat = DecodeVofRAngle(propEach.Value)
              foundLat = true
              decLat = DegMinSecToDecimalGPS(propEach.Value)
            Case "GpsLongitudeRef"
              foundEW = true
              If propEach.Value = "W" Then lngFactor = -1
              'MsgBox "GpsLongitudeRef = " & propEach.Value & ", longFactor: " & lngFactor
            Case "GpsLongitude"
              'msgbox "Found GpsLongitude"
              'strLong = DecodeVofRAngle(propEach.Value)
              foundLng = true
              decLng = DegMinSecToDecimalGPS(propEach.Value)
          End Select
          
          'If Err Then
          '  Err.Clear
          '  MsgBox propEach.Name & " *complex property* " & TypeName(propEach.Value) & ", WiaImagePropertyType = " & CStr(propEach.Type)
          'End If 
        Next

        If Not foundLng Then
          fileNotGeotagged.WriteLine objFile.Name
          flagAnyFileNotGeotagged = true
        ElseIf Not foundLat Then
          fileNotGeotagged.WriteLine objFile.Name
          flagAnyFileNotGeotagged = true
        ElseIf Not foundNS Then
          'MsgBox objFile.Name & " - Did not find N/S. Latitude imprecise!"
          fileNotGeotagged.WriteLine objFile.Name & " - Did not find N/S. Latitude imprecise!"
          flagAnyFileNotGeotagged = true
        ElseIf Not foundEW Then
          'MsgBox objFile.Name & " - Did not find E/W. Longitude imprecise!"
          fileNotGeotagged.WriteLine objFile.Name & " - Did not find N/S. Longitude imprecise!"
          flagAnyFileNotGeotagged = true
        Else
          decLat = decLat * latFactor
          decLng = decLng * lngFactor
          fileXML.WriteLine strIndent & "<photo>"
          strIndent = indent(strIndent)
          fileXML.WriteLine strIndent & "<thumb>" & objFolder.Name & operatingSystemFolderSlash & objSubFolder.Name & operatingSystemFolderSlash & strThumbsDir & operatingSystemFolderSlash & FSO.GetBaseName(objFile.Name) & strThumbPostfix & "." & strExt & "</thumb>"
          fileXML.WriteLine strIndent & "<source>" & objFolder.Name & operatingSystemFolderSlash & objSubFolder.Name & operatingSystemFolderSlash & FSO.GetBaseName(objFile.Name) & "." & strExt & "</source>"
          fileXML.WriteLine strIndent & "<description></description>"
          fileXML.WriteLine strIndent & "<location>"
          strIndent = indent(strIndent)
          fileXML.WriteLine strIndent & "<lat>" & decLat & "</lat>"
          fileXML.WriteLine strIndent & "<lng>" & decLng & "</lng>"
          strIndent = outdent(strIndent)
          fileXML.WriteLine strIndent & "</location>"
          strIndent = outdent(strIndent)
          fileXML.WriteLine strIndent & "</photo>"
        End If
        'MsgBox objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & strThumbPostfix & "." & strExt
      End If
    End If
  Next 'file
  
  strIndent = outdent(strIndent)
  fileXML.WriteLine strIndent & "</photos>"
  strIndent = outdent(strIndent)
  fileXML.WriteLine strIndent & "</place>"
Next 'folder

strIndent = outdent(strIndent)
fileXML.WriteLine strIndent & "</places>"
strIndent = outdent(strIndent)
fileXML.WriteLine strIndent & "</album>"

fileNotGeotagged.Close
If flagAnyFileNotGeotagged Then MsgBox "One or more photos were not geotagged. See log file: " & strNotGeotagged

fileXML.Close
MsgBox "Script done!" & vblf & strFileXML & " written."

'seems like cr/lf are all proper: msgbox "check cf/lf in N++"

Function indent(strIndent)
  indent = strIndent & "  "
End Function

Function outdent(strIndent)
  outdent = Left(strIndent, Len(strIndent) - 2)
End Function

Function DegMinSecToDecimalGPS(wiaVector)
  Dim i
  Dim num
  For i = 1 to wiaVector.Count
    'Set rational = wiaVector.Item(i)
    If i = 1 Then
      num = CStr(wiaVector.Item(i).Value)
    ElseIf i = 2 Then
      num = num + CStr(wiaVector.Item(i).Value) / 60
    ElseIf i = 3 Then
      num = num + CStr(wiaVector.Item(i).Value) / 3600
    End If
  Next
  DegMinSecToDecimalGPS = num
End Function
