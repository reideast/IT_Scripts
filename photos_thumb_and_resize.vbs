'Create Thumbnails by Andrew East
'Description: a script to traverse a photos directory and create thumbnails
'prereq: a folder in the same directory as this script, called strFolderName
'prereq: pictures have extension "strExt", usually .jpg
'will create a folder called strThumbsDir within each subfolder
'only traverses one level of subfolders! (ie. is not recursive, and this functionality is not needed for my web site projects. I could add this by making the function call recursive? if needed)
'will name thumbs "original file name" & strThumbPostfix & . & strExt, eg. picture.jpg -> thumbs\picture_s.jpg
'   use strThumbPostfix to set the suffix to "_s" or "_thumb" or whatever
'pictures are resized to at most maxWidth x maxHeight, and maintains aspect ratio
'this code may be used freely under a Creative Commons Attribution Share Alike License  http://creativecommons.org/licenses/by-sa/4.0/


Option Explicit

Const maxWidth = 200
Const maxHeight = 200
Const createWebSized = True
Const webSizeWidth = 1000
Const webSizeHeight = 1000

Dim Img 'As ImageFile
Dim IP 'As ImageProcess
Dim FSO
Dim objFolder
Dim objSubFolder
Dim objFile
Dim strFolderName
Dim strThumbsDir
Dim strExt
Dim strThumbPostfix

strFolderName = "photos"
strThumbsDir = "thumbs"
strExt = "jpg" 'script matches this extension and only this 'TODO: make into an array of extensions
strThumbPostfix = "_s"

Set FSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = FSO.GetFolder(strFolderName)

strExt = LCase(strExt)

'main folder pictures
thumbFilesInFolder(objFolder)

'all subfolder pictures
'TODO: make the function call recursive or iterative over subfolders, so this single-sublevel repetitive isn't necessary
For Each objSubFolder in objFolder.SubFolders
  thumbFilesInFolder(objSubFolder)
Next 'folders

msgbox "Thumbnail Photos script done."

Function thumbFilesInFolder(objFolder)
  'do not process folders name "thumbs"
  If objFolder.Name = strThumbsDir Then Exit Function 
  
  'create thumbs dir
  If Not FSO.FolderExists(objFolder.Path & "\" & strThumbsDir) Then
    FSO.CreateFolder(objFolder.Path & "\" & strThumbsDir)
  End If
  
  'iterate all files
  For Each objFile in objFolder.Files
    'MsgBox "DEBUG: Name: """ & objFile.Name & """, ParentFolder: """ & objFile.ParentFolder & """"

    'create thumbnail if file has proper extension
    If LCase(FSO.GetExtensionName(objFile.Name)) = strExt Then
      'create Windows Image Acquisition objects
      Set Img = CreateObject("WIA.ImageFile")
      Set IP = CreateObject("WIA.ImageProcess")
    
      Img.LoadFile objFile
      
      'set up and apply "Scale" filter
      IP.Filters.Add IP.FilterInfos("Scale").FilterID
      IP.Filters(1).Properties("MaximumWidth") = maxWidth
      IP.Filters(1).Properties("MaximumHeight") = maxHeight
      Set Img = IP.Apply(Img)
      
      'save modified image object as a file inside the thumbs directory
      'MsgBox "DEBUG: " & objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & strThumbPostfix & "." & strExt
      If FSO.FileExists(objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & strThumbPostfix & "." & strExt) Then
        FSO.DeleteFile(objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & strThumbPostfix & "." & strExt)
      End If
      Img.SaveFile objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & strThumbPostfix & "." & strExt
      
      'create web sized copy of image
      If createWebSized Then
        Set Img = CreateObject("WIA.ImageFile")
        Set IP = CreateObject("WIA.ImageProcess")
      
        Img.LoadFile objFile
        
        IP.Filters.Add IP.FilterInfos("Scale").FilterID
        IP.Filters(1).Properties("MaximumWidth") = webSizeWidth
        IP.Filters(1).Properties("MaximumHeight") = webSizeHeight
        Set Img = IP.Apply(Img)
        
        'MsgBox objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & strThumbPostfix & "." & strExt
        If FSO.FileExists(objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & "." & strExt) Then
          FSO.DeleteFile(objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & "." & strExt)
        End If
        Img.SaveFile objFile.ParentFolder & "\" & strThumbsDir & "\" & FSO.GetBaseName(objFile.Name) & "." & strExt
      End If

    End If  
  Next 'For Each objFile in objFolder.Files
End Function