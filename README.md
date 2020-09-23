<div align="center">

## Sounds


</div>

### Description

Here is what I did to make my PC speaker beep

at the frequency and length of time I want,

using hardware direct control.

It works fine in Win95 and Win98. Not in WinNT.
 
### More Info
 
Read comments of Win95IO.dll from SoftCircuits


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jorge Loubet](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jorge-loubet.md)
**Level**          |Unknown
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jorge-loubet-sounds__1-4069/archive/master.zip)

### API Declarations

In code, as private declarations


### Source Code

```
'*****************************************************************
'  October 17 1999- By Jorge Loubet
'  jorgeloubet@yahoo.com
'  Durango, Dgo. Mexico.
'  Hola amigos !
'  Here is what I did to make my PC speaker beep
'  at the frequency and length of time I want,
'  using hardware direct control.
'  It works fine in Win95 and Win98. Not in WinNT.
'  (Revenge against beep() function in NT ? )
'  Just follow these steps:
'  1) Download the library WIN95IO.DLL from
'    http://www.softcircuits.com (Free software)
'  2) Copy this DLL to your System folder
'  3) Put a command buton on your form named cmdStartSound
'  4) Put a timer on your form and name it as TimerSound
'  5) Copy all of this code to your form
'  6) Run it !!!
'
'  Have a nice sound and make your own fiesta with tequila and señoritas...!
'  If you think this is good for you, let me know that, sending me
'  your comments to my e-mail.
'*****************************************************************
Option Explicit
Dim SoundEnd As Boolean
'If you wish, put this declarations on a module, deleting "Private"
'Write a byte to port:
Private Declare Sub vbOut Lib "WIN95IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
'Read a byte from port:
Private Declare Function vbInp Lib "WIN95IO.DLL" (ByVal nPort As Integer) As Integer
'These are standard freqs of music. You can set any freq.
Const C = 523    'Do in spanish
Const D = 587.33  'Re
Const E = 659.26  'Mi
Const F = 698.46  'Fa
Const G = 783.99  'Sol
Const A = 880    'La
Const B = 987.77  'Si
Private Sub cmdStartSound_Click()
  Dim i As Integer
  'This is all you have to do to simulate a phone ring sound.
  For i = 1 To 12
    Sounds C, 20  'Sounds 523 Hz in 20 miliseconds
    Sounds F, 20  'Sounds 698.46 Hz in 20 miliseconds
  Next i
  'Need to go up an octave? Just double the frequency or viceversa.
  ' example:
  'Sounds C * 2, 500  'An octave up
  'Sounds C / 2, 500  'An octave down
  'Yes, you can do a funny piano using your programming skills !
End Sub
Private Sub Sounds(Freq, Length)
Dim LoByte As Integer
Dim HiByte As Integer
Dim Clicks As Integer
Dim SpkrOn As Integer
Dim SpkrOff As Integer
'  "I didn't tested if this is exactly the frequency,
'  but it's ok to start here. I you wish more precision,
'  try with a piano or another reference to adjust the clicks.
'  For example, "A" has a frequency of 880 Hertz. If you have
'  a good ear, it may be adjusted very close by
'  changing the 1193280 number up or down.
'  Of course, you can use a frequency meter.
'  I didn't tested the frequency limits too. Test it by yourself."
'  Length precision is the same as the timer control precision.
'Ports 66, 67, and 97 control timer and speaker
'Divide clock frequency by sound frequency
'to get number of "clicks" clock must produce.
  Clicks = CInt(1193280 / Freq)
  LoByte = Clicks And &HFF
  HiByte = Clicks \ 256
'Tell timer that data is coming
  vbOut 67, 182
'Send count to timer
  vbOut 66, LoByte
  vbOut 66, HiByte
'Turn speaker on by setting bits 0 and 1 of PPI chip.
  SpkrOn = vbInp(97) Or &H3
  vbOut 97, SpkrOn  'My speaker is sounding !
'Leave speaker on (while timer runs)
  SoundEnd = False        'Do not finish yet
  TimerSound.Interval = Length  'Time to sound
  TimerSound.Enabled = True    'Begin to count time
  Do While Not SoundEnd
    'Let processor do other tasks
    DoEvents
  Loop
'Turn speaker off resetting bit 0 and 1.
  SpkrOff = vbInp(97) And &HFC
  vbOut 97, SpkrOff
End Sub
Private Sub TimerSound_Timer()
  'Time is over
  SoundEnd = True   'Finish sound now
  TimerSound.Enabled = False
End Sub
```

