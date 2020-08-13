Option Explicit

' replace base path below
Dim basePath : basePath = "C:\Users\Ruth\Documents\GitHub"

Dim sounds(3)
sounds(0) = basePath & "\Tordh-Clock\audio\tordh_icelandic.mp3"
sounds(1) = basePath & "\Tordh-Clock\audio\tordh_italian.mp3"
sounds(2) = basePath & "\Tordh-Clock\audio\tordh_korean.mp3"
sounds(3) = basePath & "\Tordh-Clock\audio\tordh_spanish.mp3"

Dim max, min
max = UBound(sounds) ' last position of sounds array
min = 0
Randomize

Dim rand : rand = Int((max-min+1)*Rnd+min)

Dim sound : sound = sounds(rand)

Dim o : Set o = CreateObject("wmplayer.ocx")
With o
    .url = sound
    .controls.play
    While .playstate <> 1
        wscript.sleep 100
    Wend
    .close
End With

Set o = Nothing