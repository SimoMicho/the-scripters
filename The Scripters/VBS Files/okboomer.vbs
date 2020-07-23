On Error Resume Next

Set FSO=CreateObject("Scripting.FileSystemObject")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_CDROMDrive")


boomer = 0
For Each objItem in colItems
	boomer = 1
Next
If boomer = 1 then
    MsgBox("OK BOOMER")
Else
    a1 = MsgBox("Nope", vbOkCancel)
    If a1 = 1 then
        a2 = MsgBox("I said nope", vbOkCancel)
    End If
    If a2 = 1 then
        a3 = MsgBox("You need to stop", vbOkCancel)
    End If
    If a3 = 1 then
        a4 = MsgBox("There is nothing to see", vbOkCancel)
    End If
    If a4 = 1 then
        a5 = MsgBox("Nothing. Go away", vbOkCancel)
    End If
    If a5 = 1 then
        a6 = MsgBox("Still nothing", vbOkCancel)
    End If
    If a6 = 1 then
        a7 = MsgBox("You are losing your time", vbOkCancel)
    End If
    If a7 = 1 then
        a8 = MsgBox("Stop there is nothing", vbOkCancel)
    End If
    If a8 = 1 then
        a9 = MsgBox("Still the same", vbOkCancel)
    End If
    If a9 = 1 then
        b1 = MsgBox("Yeah i know it's disapointing", vbOkCancel)
    End If
    If b1 = 1 then
        b2 = MsgBox("Ok.", vbOkCancel)
    End If
    If b2 = 1 then
        b3 = MsgBox("You like loosing your time?", vbYesNo)
    End If
    If b3 = 7 then
        nope = MsgBox("Nice, cya!", vbOkCancel)
    End If
    If b3 = 6 then
        b4 = MsgBox("Ok, i'm going to help you with that", vbOkCancel)
    End If
    If b4 = 1 then
        b5 = MsgBox("Are you bored ?", vbYesNo)
    End If
    If b5 = 7 then
        b6 = MsgBox("Still not bored ?", vbYesNo)
    End If
    If b5 = 6 then
        nope = MsgBox("Nice, now go do something else!", vbOkCancel)
    End If
    If b6 = 7 then
        nope = MsgBox("Gotcha, now go find a job!", vbOkCancel)
    End If
    If b6 = 6 then
        b7 = MsgBox("Ok.", vbOkCancel)
    End If
    If b7 = 1 then
        b8 = MsgBox("I'm going to tell you something.", vbOkCancel)
    End If
    If b8 = 1 then
        b9 = MsgBox("There is still nothing.", vbOkCancel)
    End If
    If b9 = 1 then
        c1 = MsgBox("You are still hanging on ?", vbYesNo)
    End If
    If c1 = 6 then
        c2 = MsgBox("Ok.", vbOkCancel)
    End If
    If b9 = 7 then
        c1 = MsgBox("Oh, okay cya!", vbOkCancel)
    End If
    If c2 = 1 then
        c3 = MsgBox("I'm going to tell you something", vbOkCancel)
    End If
    If c3 = 1 then
        c4 = MsgBox("But not now", vbOkCancel)
    End If
    If c4 = 1 then
        c5 = MsgBox("Still not now", vbOkCancel)
    End If
    If c5 = 1 then
        c6 = MsgBox("You know what?", vbOkCancel)
    End If
    If c6 = 1 then
        c7 = MsgBox("Almost not now", vbOkCancel)
    End If
    If c7 = 1 then
        c8 = MsgBox("Time will come", vbOkCancel)
    End If
    If c8 = 1 then
        c9 = MsgBox("Patience is the greatest of virtues", vbOkCancel)
    End If
    If c9 = 1 then
        d1 = MsgBox("Be patient", vbOkCancel)
    End If
    If d1 = 1 then
        d2 = MsgBox("Keep it like that", vbOkCancel)
    End If
    If d2 = 1 then
        d3 = MsgBox("Ok.", vbOkCancel)
    End If
    If d3 = 1 then
        d4 = MsgBox("I'm going to tell you now.", vbOkCancel)
    End If
    If d4 = 1 then
        d5 = MsgBox("You are rly determinated", vbOkCancel)
    End If
    If d5 = 1 then
        d6 = MsgBox("Or totaly bored", vbOkCancel)
    End If
    If d6 = 1 then
        d7 = MsgBox("In both cases, i'm going to give you a present.", vbOkCancel)
    End If
    If d7 = 1 then
        Set objFile = FSO.CreateTextFile("./present",True)
        Randomize
        random = Int(Rnd)
        values = Array("15564", "78945")
        data = values(random)
        objFile.Write data & vbCrLf
        objFile.Close
        d8 = MsgBox("Here it is: 'present'", vbOkCancel)
    End If
    If d8 = 1 then
        d9 = MsgBox("You can send me the file created via discord and you will get a cookie.", vbOkCancel)
    End If
    If d9 = 1 then
        e1 = MsgBox("You want more ?", vbYesNo)
    End If
    If e1 = 7 then
        nope = MsgBox("Good Choice.", vbOkCancel)
    End If
    If e1 = 6 then
        e2 = MsgBox("Sure ?", vbYesNo)
    End If
    If e2 = 7 then
        nope = MsgBox("You almost did the wrong thing.", vbOkCancel)
    End If
    If e2 = 6 then
        e3 = MsgBox("Rly ?", vbYesNo)
    End If
    If e3 = 7 then
        nope = MsgBox("What a dangerous world, good thing you stopped here.", vbOkCancel)
    End If
    If e3 = 6 then
        e4 = MsgBox("Don't do it.", vbOkCancel)
    End If
    If e4 = 2 then
        nope = MsgBox("UFFF. That was close to the end.", vbOkCancel)
    End If
    If e4 = 1 then
        e5 = MsgBox("Plz stop it's bad !", vbOkCancel)
    End If
    If e5 = 2 then
        nope = MsgBox("One step closer and it was ended.", vbOkCancel)
    End If
    If e5 = 1 then
        e6 = MsgBox("Ok.", vbOkCancel)
    End If
    If e6 = 2 then
        Set objFile = FSO.CreateTextFile("./Playing with fire",True)
        Randomize
        random = Int(Rnd)
        values = Array("66618", "77742")
        data = values(random)
        objFile.Write data & vbCrLf
        objFile.Close
        nope = MsgBox("You like to play with fire, don't ya? you almost did the bad choice and it is rly bad to do so don't do it.", vbOk)
        nope = MsgBox("You got the achievement: 'Playing with fire', redeem your prize by sending the file created to me via discord.")
    End If
    If e6 = 1 then
        baddd = 0
        While baddd <= 10
            nope = MsgBox("Now you are trapped", vbOkCancel)
            if nope = 2 Then
                baddd = baddd + 1
            end if
        Wend
        nope = MsgBox("Just kidding, it's ended, bye! Made by Shyne#3230")
    End If
End If