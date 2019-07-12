#SingleInstance
; <COMPILER: v1.0.47.6>
Gui, Add, CheckBox, x26 y10 w190 h20 vHCopy, Print Hardcopies on Default Printer?
Gui, Add, CheckBox, x26 y40 w130 h20 vRecursive, Include Sub Folders?
Gui, Add, CheckBox, x26 y70 w130 h20 vExcel, Include Excel Files?*
Gui, Add, Text, x36 y100 w160 h30 , *Must have PDF XChange as 
Gui, Add, Text, x36 y120 w160 h30 , default printer if saving to PDF
Gui, Add, Button, x126 y150 w80 h30 , OK
Gui, Add, Button, x16 y150 w80 h30 , Cancel
GuiControl,, Recursive, 1
SetTitleMatchMode 3

FileSelectFolder, folder,, 0, Select Folder Containing HCS+ Files
;folder := "C:\Program Files (x86)\HCS7 Demo\HCS\Examples"
if (errorlevel)
{
  exitapp
}
Gui, Show, x531 y410 h200 w230, HCS+ PDFer

Ctrl & Q::
BlockInput, Off
exitapp

ButtonCancel:
exitapp

ButtonOK:
Gui, Submit
BlockInput, On
loop, %folder%\*.*, 0, %Recursive%
{
  if (A_LoopFileExt = "xls" or A_LoopFileExt = "xlsx")
  {
    if (Excel)
    {
      PrintExcel(A_LoopFileLongPath)
    }
  }
  else if (A_LoopFileExt = "xhf")
  {
    PrintHCS(A_LoopFileLongPath, "Freeways")
  }
  else if (A_LoopFileExt = "xuf")
  {
    PrintHCS(A_LoopFileLongPath, "Freeways")
  }
  else if (A_LoopFileExt = "xhr")
  {
    PrintHCS(A_LoopFileLongPath, "Ramps")
  }
  else if (A_LoopFileExt = "xhw")
  {
    PrintHCS(A_LoopFileLongPath, "Weaving")
  }
  else if (A_LoopFileExt = "syn")
  {
    PrintSynchro10(A_LoopFileLongPath)
  }
  else if (A_LoopFileExt = "hcf")
  {
    PrintHCS(A_LoopFileLongPath, "Freeways")
  }
  else if (A_LoopFileExt = "hcr")
  {
    PrintHCS(A_LoopFileLongPath, "Ramps")
  }
  else if (A_LoopFileExt = "hcw")
  {
    PrintHCS(A_LoopFileLongPath, "Weaving")
  }
}

BlockInput, Off
sleep 3000
process, close, Freeways.exe
process, close, Ramps.exe
process, close, Weaving.exe
process, close, synchro7.exe
process, close, Excel.exe
exitapp

PrintHCS(file, type)
{
  global HCopy
  ifnotexist, %file%
  {

    return
  }
  if (!HCopy)
  {
    stringtrimright, pdfpath, file, 3
    pdfpath = %pdfpath%pdf
    ifexist, %pdfpath%
    {
      return
    }
  }
  run, %file%

  process, wait, %type%.exe, 10
  HCSPID = %ErrorLevel%
  if (HCSPID = 0)
  {
    return
  }
  WinWaitActive, ahk_pid %HCSPID%,,10
  if (!ErrorLevel)
  {
    ;PostMessage, 0x111, 61550
    Sleep 1000
  if WinExist("Demo Version Limitations")
  {
	WinActivate
	Click, 300, 320
  }
    ShowPrintDialog(%HCSPID%)
	WinWaitActive, Print
 if WinExist("Print"){
	WinActivate
	SendInput, !p
 }
    if (!HCopy)
    {
      ;WinActivate, Print
      ;ControlSend, FolderView, primo, Print
      ;Sleep 1000
      ;ControlClick, &Print, Print
      SavePDFXchange(pdfpath, "HCS", %HCSPID%)
	  ;SavePrimoPDF(pdfpath, "HCS")
    }
    ;else
    ;{
      ;ControlClick, Button13, Print
      ;sleep 2000
    ;}
	
	
	;*************
	
	;*************
	
	;msgbox, Closing Freeways.exe
	WinActivate, ahk_pid %ParentPID%
	SendInput !{f4}
    winclose, ahk_pid %HCSPID%,,5
  }
}

PrintSynchro10(file)
{
  global HCopy
  ifnotexist, %file%
  {

    return
  }
  if (!HCopy)
  {
    stringtrimright, pdfpath, file, 4
    pdfpath1 = %pdfpath%-Synchro.pdf
    pdfpath2 = %pdfpath%-HCM.pdf
    ifexist, %pdfpath1%
    {

      return
    }
    ifexist, %pdfpath2%
    {

      return
    }
  }
  run, %file%
  process, wait, synchro10.exe, 10
  SynchroPID = %ErrorLevel%
  if (SynchroPID = 0)
  {

    return
  }
  WinWaitActive, ahk_pid %SynchroPID%,,10
  if (!ErrorLevel)
  {
    if (!HCopy)
    {
      Synchro10Print(pdfpath1, SynchroPID)
    }
  }
}

Synchro10Print(pdfpath, ParentPID)
{
	;Send Alt
	Send, {Alt}
	;Arrow Right four times to Report tab
	Send, {Right 4}
	WinWaitActive, Create Report,,3
	;Arrow down to select Intersections
	Send, {Down}
	;Send {enter} to activate Intersections report
	Send, {Enter}
	;Open Print dialog
    Send, !p
	;WinWaitActive, Print,,3
	;Print to PDFXchange
	SavePDFXchange(pdfpath, "Synchro", ParentPID)
}

PrintExcel(file)
{
  global HCopy
  ifnotexist, %file%
  {
    return
  }
  if (!HCopy)
  {
    stringtrimright, pdfpath, file, 3
    pdfpath = %pdfpath%pdf
    ifexist, %pdfpath%
    {
      return
    }
  }
  run, %file%
  process, wait, Excel.exe, 10
  ExcelPID = %ErrorLevel%
  if (ExcelPID = 0)
  {
    return
  }
  WinWait, ahk_pid %ExcelPID%,,10
  if (!ErrorLevel)
  {
    WinActivate
    WinWaitActive
    ;ShowPrintDialog(%ExcelPID%)
	SendInput ^p
	;WinActivate, NetUIHWND1, Print
	;WinWaitActive
	ControlClick, NetUIHWND1, Print
	SendInput, {Enter}
    if (!HCopy)
    {
      ;SavePrimoPDF(pdfpath)
	  SavePDFXchange(pdfpath, "Excel", %ExcelPID%)
    }
    winclose, ahk_pid %ExcelPID%,,5
  }
}

ShowPrintDialog(ParentPID)
{
  WinActivate, ahk_pid %ParentPID%
  SendInput ^p
  
  ;winwait, Print,,5
  WinWaitActive, Print
  if (ErrorLevel)
  {
    ShowPrintDialog(%ParentPID%)
	;msgbox, Error: Loopback in ShowPrintDialog
  }  
}
	
SavePDFXchange(path, type, ParentPID){
	;msgbox, %path%
		WinWaitActive, Save As,,5
	 if WinExist("Save As"){
		WinActivate
		WinWaitActive, Save As,,5
		; If type is HCS ControlSetText if type is Synchro then don't
		if(%type% = "HCS"){
			ControlSetText, Edit1, %path%
		}
		;ControlClick, &Save, Save As
		Send, !s
		;Close parent process
		
	}
	;If trying to save in a restricted location, this warning appears to save in My Documents
	if WinExist("Save As"){
		WinWaitActive, Save As,,5
		;ControlClick, &Yes, Save As
		Send, !y
		Sleep 1000
		;Focus goes back to the Save As dialog, so Save again
		;ControlClick, &Save, Save As
		Send, !s
	}

	;If the file already exists, this dialog opens
	;PDF-XChange Standard ahk_class UIX:WindowNC ahk_exe pdfSaver.exe ahk_pid 11628 
	If WinExist("PDF-XChange Standard"){
		WinActivate
		WinWaitActive, PDF-XChange Standard,,3
		Send, {Enter}
		;Enter to default action, Overwrite
	}
	if(%type% = "Synchro"){
		; Close Synchro window
		winclose, ahk_pid %ParentPID%,,3
		; alternatively: Alt+F, Y9
	}
}