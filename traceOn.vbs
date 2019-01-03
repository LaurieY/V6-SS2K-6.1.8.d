dim c
set c = CreateObject("ASCOM.utilities.Util")
msgBox("ASCOM Platform Version " & c.Platformversion)


dim u

'set u = CreateObject("ASCOM.utilities.Util")

'u.SerialTrace=True
dim t
set t = CreateObject("SS2K.Telescope")
't.SetTracing=1

