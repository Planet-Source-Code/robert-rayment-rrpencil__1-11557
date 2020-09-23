Attribute VB_Name = "IView"
'IView.bas

'Using FREEWARE i_view32.exe by Irfan Skiljan
'See iview32 folder

Option Base 1
DefInt A-Q  'a-q integers
DefSng R-Z  'rst, uvw, xyz real

Public Sub SaveAsJPG()
Title$ = "Save JPG file"
Choice$ = "JPG files(*.jpg)|*.jpg"
InitDir$ = CurrentDirec$
SFile$ = ""
OpenSaveDialog Title$, Choice$, SaveFileSpec$, InitDir$, SFile$
If SaveFileSpec$ <> "" Then
   
   FixFileExtension SaveFileSpec$, "jpg"
   CurrentDirec$ = ExtractPath(SaveFileSpec$)
   
   tempsave$ = App.Path & "\~~temp.bmp"
   SavePicture Form1.picCanvas.Image, tempsave$
   H$ = Trim$(Str$(Form1.picCanvas.Height))
   wh$ = "600," + H$ + ")"
   'NB .Picture saves as original size, .Image as whole picture box
   aString$ = App.Path & "\i_view32.exe " & tempsave$
   bString$ = " /convert=" & SaveFileSpec$ & "/crop=(0,0," & wh$ ' & "/ killmesoftly"
   cString$ = aString$ & bString$
   'Critical note on spaces:
   'after .exe, before /convert
   res& = Shell(cString$, vbNormalFocus)
   Kill tempsave$
End If
End Sub

Public Sub PrintFromIView()
tempsave$ = App.Path & "\~~temp.bmp"
SavePicture Form1.picCanvas.Image, tempsave$
aString$ = App.Path & "\i_view32.exe " & tempsave$
bString$ = " /print"
cString$ = aString$ & bString$
'Critical note on spaces:
'Space after exe & before /print
res& = Shell(cString$, vbNormalFocus)
Kill tempsave$
End Sub
