Attribute VB_Name = "Speech"
'-- The Enums:

Enum SpeechVoiceSpeakFlags
SVSFDefault = 0
SVSFlagsAsync = 1 '--speak asynchronously.
SVSFPurgeBeforeSpeak = 2 '--purge pending speech.
SVSFIsFilename = 4
SVSFIsXML = 8
SVSFIsNotXML = 16
SVSFPersistXML = 32
SVSFNLPSpeakPunc = 64 '--speak punctuation as word: "blah blah period"
'-- Masks
SVSFNLPMask = 64
SVSFVoiceMask = 127
SVSFUnusedFlags = -128
End Enum

'-- This flag may not be what you want but it's a sanmple.
'-- It will make TTS keep up with events rather than finish speaking it's
'-- buffer before moving on. For instance, if you set it to speak 500
'-- names and then tell it halfway through to speak 500 numbers instead,
'-- ASync will cause it to stop reading names and switch to numbers.
'-- Not-Async will cause it to finish with the names first.

Const SPEAK_FLAGS_1 = SVSFlagsAsync Or SVSFPurgeBeforeSpeak Or SVSFIsNotXML

Public SpkVoice As New SpVoice

'-- Public sub that can be called to speak provided string:

Public Sub SpeakString(sText As String)
'On Error Resume Next
If (Len(sText) = 0) Then Exit Sub
SpkVoice.Speak sText, SVSFlagsAsync
End Sub



