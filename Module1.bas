Attribute VB_Name = "Module1"

Option Explicit

' ======= USER SETTINGS =======
Private Const audioFolder As String = "C:\Users\1300772\Downloads\narrations\"  ' <-- change to your folder (must end with \)
Private Const paddingSeconds As Double = 3#             ' extra seconds after audio before advancing
' =============================

' Main: insert audio per slide based on alphabetical order of MP3s in the folder
Public Sub InsertAudioByFolderOrder()
    Dim files() As String
    Dim fileCount As Long
    Dim sld As Slide
    Dim i As Long

    ' Validate folder path ending
    If Right$(audioFolder, 1) <> "\" Then
        MsgBox "audioFolder must end with a backslash (\). Current: " & audioFolder, vbExclamation
        Exit Sub
    End If

    ' Collect and sort MP3s alphabetically
    files = GetMp3FilesSorted(audioFolder, fileCount)
    If fileCount = 0 Then
        MsgBox "No MP3 files found in: " & audioFolder, vbExclamation
        Exit Sub
    End If

    ' Loop slides and insert audio by folder order
    i = 1
    For Each sld In ActivePresentation.Slides
        If i > fileCount Then Exit For
        Call InsertAudioAndTiming(sld, files(i))
        i = i + 1
    Next sld

    ' Ensure slideshow uses timings
    ActivePresentation.SlideShowSettings.AdvanceMode = ppSlideShowUseSlideTimings

    ActivePresentation.Save
    MsgBox "Done: inserted audio and set timings for " & (i - 1) & " slide(s).", vbInformation
End Sub

' Insert one audio file to a slide and set timing (audio length + paddingSeconds)
Private Sub InsertAudioAndTiming(ByVal sld As Slide, ByVal mp3Path As String)
    Dim audioShape As Shape
    Dim effect As effect
    Dim audioLenMs As Single
    Dim audioLenSec As Double

    ' Embed audio
    Set audioShape = sld.Shapes.AddMediaObject2(mp3Path, False, True, 10, 10)

    ' Autoplay + hide icon during show
    Set effect = sld.TimeLine.MainSequence.AddEffect( _
        audioShape, msoAnimEffectMediaPlay, , msoAnimTriggerWithPrevious)
    effect.MoveTo 1
    effect.EffectInformation.PlaySettings.HideWhileNotPlaying = True

    ' Timing = audio length + padding
    audioLenMs = audioShape.MediaFormat.Length    ' milliseconds
    audioLenSec = audioLenMs / 1000#

    With sld.SlideShowTransition
        .AdvanceOnTime = msoTrue
        .AdvanceTime = audioLenSec + paddingSeconds
    End With

    Debug.Print "Slide " & sld.SlideIndex & " ? " & Format$(audioLenSec, "0.00") & _
                "s + " & paddingSeconds & "s = " & Format$(audioLenSec + paddingSeconds, "0.00") & "s"
End Sub

' Return sorted list of *.mp3 files from a folder
Private Function GetMp3FilesSorted(ByVal folderPath As String, ByRef countOut As Long) As String()
    Dim f As String
    Dim files() As String
    Dim n As Long

    ' Gather files using Dir
    f = Dir$(folderPath & "*.mp3")
    Do While Len(f) > 0
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = folderPath & f
        f = Dir$
    Loop

    countOut = n
    If n > 1 Then Call SortStringsAscending(files)
    GetMp3FilesSorted = files
End Function


' Simple in-place alphabetical sort for string array
Private Sub SortStringsAscending(ByRef arr() As String)
    Dim i As Long, j As Long
    Dim tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            ' Use the real < operator, not &lt;
            If UCase$(arr(j)) < UCase$(arr(i)) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
End Sub

