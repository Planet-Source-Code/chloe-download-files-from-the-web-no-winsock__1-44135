VERSION 5.00
Begin VB.UserControl Downloader 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Downloader.ctx":0000
   ScaleHeight     =   2385
   ScaleWidth      =   3480
   ToolboxBitmap   =   "Downloader.ctx":030A
End
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------
' This User control can easily be added to any
' application to give it quick and easy download
' capability without the use of Winsock.  This
' user control can easily be made into a stand
' alone OCX or can be left the way it is to allow
' it to be embeded into any EXE to eliminate the
' need to distribute an OCX.
'------------------------------------------------------------
'------------------------------------------------------------
' If you copy this CONTROL file and its matching
' CTX file to C:\Program Files\Microsoft Visual
' Studio\VB98\Template\Userctls you can have it
' in your list when you add a user control to your
' projects for quick insertion later.
'------------------------------------------------------------


'------------------------------------------------------------
' Also note, this control can download MULTIPLE
' files at one time and the download progress and
' download complete events pass which file it is
' talking about.  You will have to code accordingly
' in your EXE to give any user feedback if you
' wish.
'------------------------------------------------------------

Option Explicit
'------------------------------------------------------------
' Events we will raise to give the developer feedback
'------------------------------------------------------------
Event DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
Event DownloadError(SaveFile As String)
Event DownloadComplete(MaxBytes As Long, SaveFile As String)
'------------------------------------------------------------
' When a download is done, the UserControl fires
' this event, so we use the PropertyName of the
' download [set to the filename to save to inside
' BeginDownload] to save the file then fire our
' DownloadComplete event to tell the developer.
'------------------------------------------------------------
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error Resume Next
    Dim f() As Byte, fn As Long
    If AsyncProp.BytesMax <> 0 Then
        fn = FreeFile
        f = AsyncProp.Value
        Open AsyncProp.PropertyName For Binary Access Write As #fn
        Put #fn, , f
        Close #fn
    Else
        RaiseEvent DownloadError(AsyncProp.PropertyName)
    End If
    RaiseEvent DownloadComplete(CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
End Sub
'------------------------------------------------------------
' This usercontrol event fires to give progress
' so we just use it to fire off our custom events
' back to the developer so he/she can use it for
' user feedback
'------------------------------------------------------------
Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    On Error Resume Next
    If AsyncProp.BytesMax <> 0 Then
        RaiseEvent DownloadProgress(CLng(AsyncProp.BytesRead), CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
    End If
End Sub
Private Sub UserControl_Initialize()
    SizeIt
End Sub
Private Sub UserControl_Resize()
    SizeIt
End Sub
'------------------------------------------------------------
' Start a download using the AsyncRead method of
' the UserControl.  Set the propertyname to the
' name of the file we are downloading so we can
' use it later and to keep track of files.
'------------------------------------------------------------
Public Sub BeginDownload(url As String, SaveFile As String)
    On Error GoTo ErrorBeginDownload
    UserControl.AsyncRead url, vbAsyncTypeByteArray, SaveFile, vbAsyncReadForceUpdate
    Exit Sub
ErrorBeginDownload:
    MsgBox Err & ":Error in call to BeginDownload()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
'------------------------------------------------------------
' Dont let the control be bigger than the icon
' on the control.  It is hidden at runtime so no
' need to allow sizing.
'------------------------------------------------------------
Public Sub SizeIt()
    On Error GoTo ErrorSizeIt
    With UserControl
        .Width = ScaleX(32, vbPixels, vbTwips)
        .Height = ScaleY(32, vbPixels, vbTwips)
    End With
    Exit Sub
ErrorSizeIt:
    MsgBox Err & ":Error in call to SizeIt()." _
    & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
