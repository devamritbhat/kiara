^w::
{
	Run "\\192.168.10.3\MwApp\Exe\Emr.exe"
	Sleep 1000
	Send "KJ"
	Send "{tab}"
	Send "4ST01"
	Send "{tab}"
	Send "4ST01{Enter}"
}


^1::
{
	JobCard()
	{

		JC := InputBox("Please Scan the Second Bar Code in your Job Card for Filing", "Filing", "W500 H100").Value
		If InStr(JC, "/")
		{
			JC := SubStr(JC, 6)
			return JC
		}
		Else If (JC = "")
		{
			Exit
		}
		JobCard()
	}

	WorkerCode()
	{
		Loop
		{
			WC := InputBox("Please Scan your Worker Bar Code", "Worker Bar Code", "W500 H100").Value
			If Not InStr(WC, "/") && WC != ""
			{
				return WC
			}
		}
	}

	AddNewEntry(JC)
	{
		Send "MFL"
		Send "{Tab}{F2}{Enter}{Tab}{Tab}{Enter}"
		Sleep 5000
		Loop 6
			{
				Send "{Tab}"
				Sleep 100
			}
		Send "^{n}{Tab}{Tab}{Tab}" JC "{Tab}"
		Sleep 500
	}

	SaveEntry()
	{
		Send "^{s}"
		Sleep 2000
		Send "^{x}"
		Sleep 2000
	}

	BagMovement(JC)
	{
		Click "170 -11"
		Sleep 500
		Send "{Down}{Down}{Right}{Down}{Enter}"
		Sleep 5000
		Send "{Tab}{Tab}"
		Sleep 1000
		AddNewEntry(JC)	
	}

	DailyTransaction(JC, WC)
	{
		Sleep 1000
		Click "170 -11"
		Sleep 500
		Send "{Down}{Down}{Right}{Enter}"
		Sleep 5000
		Send "{Tab}{Tab}"
		Sleep 1000
		AddNewEntry(JC)
		Send "^{Delete}"
		Send WC "!{l}"
		Sleep 500
		Send "!{l}"
		Send "^{Delete}"
		Sleep 500
		Send "ZSELF{Tab}"
		Send "{Tab}{Tab}{Tab}{Tab}"
		
	}

	ProcessEntry(JC, WC)
	{
		BagMovement(JC)
		Send "^{Delete}"
		Sleep 500
		Send "MPFIL"
		Sleep 1000
		Send "{Tab}{Tab}{Tab}"
		SaveEntry()

		DailyTransaction(JC, WC)
		SaveEntry()

		BagMovement(JC)
		Send "PSECD"
		Send "{Tab}{Tab}"
		SaveEntry()
	}

	
	JC := JobCard()
	WC := WorkerCode()
	ProcessEntry(JC, WC)
}

^2::
{
	JobCard()
	{

		JC := InputBox("Please Scan the Second Bar Code in your Job Card for Setting", "Setting", "W500 H100").Value
		If InStr(JC, "/")
		{
			JC := SubStr(JC, 6)
			return JC
		}
		Else If (JC = "")
		{
			Exit
		}
		JobCard()
	}

	WorkerCode()
	{
		Loop
		{
			WC := InputBox("Please Scan your Worker Bar Code", "Worker Bar Code", "W500 H100").Value
			If Not InStr(WC, "/") && WC != ""
			{
				return WC
			}
		}
	}

	AddNewEntry(JC)
	{
		Send "MST"
		Send "{Tab}{F2}{Enter}{Tab}{Tab}{Enter}"
		Sleep 5000
		Loop 6
		{
			Send "{Tab}"
			Sleep 100
		}
		Send "^{n}{Tab}{Tab}{Tab}" JC "{Tab}"
		Sleep 500
	}

	SaveEntry()
	{
		Send "^{s}"
		Sleep 1500
		Send "^{x}"
		Sleep 1500
	}

	BagMovement(JC)
	{
		Click "170 -11"
		Sleep 500
		Send "{Down}{Down}{Right}{Down}{Enter}"
		Sleep 5000
		Send "{Tab}{Tab}"
		Sleep 1000
		AddNewEntry(JC)	
	}

	DailyTransaction(JC, WC)
	{
		Click "170 -11"
		Sleep 500
		Send "{Down}{Down}{Right}{Enter}"
		Sleep 5000
		Send "{Tab}{Tab}"
		Sleep 1000
		AddNewEntry(JC)
		Send "^{Delete}"
		Send WC "!{l}"
		Sleep 500
		Send "!{l}"
		Send "^{Delete}"
		Sleep 500
		Send "ZSELF{Tab}"
		Send "{Tab}{Tab}{Tab}{Tab}"
		
	}

	ProcessEntry(JC, WC)
	{
		BagMovement(JC)
		Send "^{Delete}"
		Sleep 500
		Send "MPSET"
		Sleep 1000
		Send "{Tab}{Tab}{Tab}"
		SaveEntry()

		DailyTransaction(JC, WC)
		SaveEntry()

		BagMovement(JC)
		Send "PSECD"
		Send "{Tab}{Tab}"
		SaveEntry()
	}

	
	JC := JobCard()
	WC := WorkerCode()
	ProcessEntry(JC, WC)
}

^3::
{
	JobCard()
	{

		JC := InputBox("Please Scan the Second Bar Code in your Job Card for Polishing", "Polishing", "W500 H100").Value
		If InStr(JC, "/")
		{
			JC := SubStr(JC, 6)
			return JC
		}
		Else If (JC = "")
		{
			Exit
		}
		JobCard()
	}

	WorkerCode()
	{
		Loop
		{
			WC := InputBox("Please Scan your Worker Bar Code", "Worker Bar Code", "W500 H100").Value
			If Not InStr(WC, "/") && WC != ""
			{
				return WC
			}
		}
	}

	Weight()
	{
		Loop
		{
			WT := InputBox("Please Type the Weight", "Weight", "W500 H100").Value
			If Not InStr(WT, "/") && WT != ""
			{
				return WT
			}
		}
	}

	AddNewEntry(JC)
	{
		Send "SCD"
		Send "{Tab}{F2}{Enter}{Tab}{Tab}{Enter}"
		Sleep 5000
		Loop 6
		{
			Send "{Tab}"
			Sleep 100
		}
		Send "^{n}{Tab}{Tab}{Tab}" JC "{Tab}"
		Sleep 500
	}

	SaveEntry()
	{
		Send "^{s}"
		Sleep 1500
		Send "^{x}"
		Sleep 1500
	}

	DailyTransaction(JC, WC, WT)
	{
		Click "170 -11"
		Sleep 500
		Send "{Down}{Down}{Right}{Enter}"
		Sleep 5000
		Send "{Tab}{Tab}"
		Sleep 1000
		AddNewEntry(JC)
		Send "^{Delete}"
		Send WC "!{l}"
		Sleep 500
		Send WT
		Send "!{g}"
		Send "^{Delete}"
		Sleep 500
		Send "ZSELF{Tab}"
		Send "{Tab}{Tab}{Tab}{Tab}"
		
	}
	
	ProcessEntry(JC, WC, WT)
	{
		DailyTransaction(JC, WC, WT)
		SaveEntry()
	}

	JC := JobCard()
	WC := WorkerCode()
	WT := Weight()
	ProcessEntry(JC, WC, WT)
}

^0::
{
	ExitApp
}
