
'-----------------------------------------------------------'
'   iTunes COM Playlist VBScript     Ver 0.1                '
'           2015/05/25                      (@pcjpnet)      '
'-----------------------------------------------------------'

Option Explicit

' Playlist Name
Const PlayListName = "end"

' iTunes Handle
Dim iTunesApp
Set iTunesApp = WScript.CreateObject("iTunes.Application")

' Playlist Handle
Dim UserList
On Error Resume Next
Set UserList = iTunesApp.LibrarySource.Playlists.ItemByName(PlayListName)
On Error Goto 0

If UserList is Nothing Then
WScript.Echo("�I�����ꂽ�v���C���X�g�͑��݂��܂���B")
WScript.Quit(0)
End If

If iTunesApp.PlayerState = 1 Then
UserList.SongRepeat = 2		'RepeatMode = Repeat playlist(2)
UserList.PlayFirstTrack		'Play
Else
WScript.Echo("�Ȃ��Đ�����Ă��܂���B�v���C���X�g�̍Đ��͍s���܂���B")
WScript.Quit(0)
End If

WScript.Quit(0)
