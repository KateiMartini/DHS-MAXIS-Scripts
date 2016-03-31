

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================


BeginDialog Pact_Dialog, 0, 0, 326, 245, "Pact Panel Closure"
  EditBox 65, 5, 80, 15, Case_number
  DropListBox 65, 65, 175, 15, "Select"+chr(9)+"Appl W/Drawn Per Client Req"+chr(9)+"Prog Closed Per Client Req"+chr(9)+"Refused/Failed Required Info", Snap_Pact_reason
  DropListBox 65, 90, 175, 20, "Select"+chr(9)+"Appl W/Drawn Per Client Req"+chr(9)+"Prog Closed Per Client Req"+chr(9)+"Refused/Failed Required Info", HC_Pact_reason
  DropListBox 65, 115, 175, 15, "Select"+chr(9)+"Appl W/Drawn Per Client Req"+chr(9)+"Prog Closed Per Client Req"+chr(9)+"Refused/Failed Required Info", Cash_Pact_reason
  DropListBox 65, 140, 175, 15, "Select"+chr(9)+"Appl W/Drawn Per Client Req"+chr(9)+"Prog Closed Per Client Req"+chr(9)+"Refused/Failed Required Info", GRH_Pact_reason
  ButtonGroup ButtonPressed
    OkButton 270, 90, 50, 15
    CancelButton 270, 110, 50, 15
  CheckBox 5, 60, 30, 15, "SNAP", FS_check_box
  CheckBox 5, 85, 50, 15, "Heatlh Care", HC_check_box
  CheckBox 5, 110, 35, 15, "Cash", Cash_check_box
  CheckBox 5, 135, 35, 15, "GRH", GRH_check_box
  Text 5, 5, 50, 10, "Case Number"
  Text 65, 30, 175, 15, "Select what action you would like below"
  CheckBox 5, 175, 175, 15, "Check box if Pact Panel used for returned mail.", Returned_mail
  CheckBox 5, 220, 150, 15, "Check box to Tilk to 10 day return for Snap", tikl_check_box
 
EndDialog


'The script-------------------------------------------
EMConnect ""
	
Call maxis_case_number_finder (case_number)




call check_for_maxis (True)

'look to see what programs are open
call navigate_to_Maxis_screen ("stat", "prog")

'read what programs are actv or pending
EmReadScreen cash_I_status, 4, 6, 74 
EmReadScreen cash_II_status, 4,7, 74
EmReadScreen GRH_status, 4, 9, 74
EmReadScreen FS_status, 4, 10, 74
EmReadScreen HC_status, 4, 12, 74




'Filling in the check boxs accourding Actv or Pending programs.
If FS_status = "ACTV" or FS_status = "PEND" then FS_check_box = checked
IF cash_I_status = "ACTV" or cash_I_status = "PEND" then Cash_check_box = checked 
If cash_II_status = "ACTV" or cash_II_status = "PEND" then Cash_check_box = checked 
If HC_status = "ACTV" or  HC_status = "PEND" then HC_check_box = checked
IF GRH_status = "ACTV"  or HC_status = "PEND" then GRH_check-box = checked

Do 
err_msg = ""
dialog Pact_Dialog
cancel_confirmation
If FS_check_box = checked and Snap_Pact_reason = "Select" then err_msg = err_msg & vbCr &" *You selected Food Support box, but didn't select a reason."
If FS_check_box = unchecked and Snap_Pact_reason <> "Select" then err_msg = err_msg & vbCr &" *You didn't selected a Food Support box, but did select a reason."

IF err_msg <> "" then msgbox "****NOTICE!!!****" & vbCr &  err_msg & vbCr & vbCr & "Resolve above to continue."
Loop until err_msg = ""

'this code for all programs listed and 



	
'account for people not selectiong prog, and selecting reasons.  




'navigate to panel
call navigate_to_Maxis_screen ("stat", "pact")

'looking to if pact panel is needed or if new panel needs to be created
EMReadScreen pact_look, 1, 2, 78
IF pact_look = "1" Then
	pf9
ElseIf pact_look = "0" then 
	EMWritescreen "nn", 20, 79
	transmit
End If 

're-establishing variables to code into MAXIS for application withdraw 
If Cash_Pact_reason = "Appl W/Drawn Per Client Req" then cash_Pact_reason_code = "Withdraw"
If Snap_Pact_reason = "Appl W/Drawn Per Client Req" then snap_Pact_reason_code = "Withdraw"
If HC_Pact_reason = "Appl W/Drawn Per Client Req" then hc_Pact_reason_code = "Withdraw"
If GRH_Pact_reason = "Appl W/Drawn Per Client Req" then GRH_Pact_reason_code = "Withdraw"


'According to new policy pending applications need to denied off PND2.
If Cash_check_box = checked AND cash_Pact_reason_code = "Withdraw" and (cash_I_status = "PEND" or cash_II_status = "PEND")then Msgbox "Pending Cash application CAN'T be denied off PACT, due to new policy change.  Go to PND2 to deny application"
If FS_check_box = checked AND snap_Pact_reason_code = "Withdraw" and FS_status = "PEND" then Msgbox "Pending Snap application CAN'T be denied off PACT, due to new policy change.   Go to Pnd2 to deny"
'*****add GRH and HC


'*****notes needed 
If Cash_Pact_reason = "Prog Closed Per Client Req" then Cash_Pact_reason_code = "2"
If Snap_Pact_reason = "Prog Closed Per Client Req" then Snap_Pact_reason_code = "2"
If HC_Pact_reason = "Prog Closed Per Client Req" then HC_Pact_reason_code = "2"
If GRH_Pact_reason ="Prog Closed Per Client Req" then GRH_Pact_reason_code = "2"



If Cash_Pact_reason = "Refused/Failed Required Info" then Cash_Pact_reason_code = "3"
If Snap_Pact_reason = "Refused/Failed Required Info" then Snap_Pact_reason_code = "3"
If HC_Pact_reason = "Refused/Failed Required Info" then HC_Pact_reason_code = "3"

If GRH_Pact_reason ="Refused/Failed Required Info" then GRH_Pact_reason_code = "3"





'****program pact reason code
'updating pact panel if program actv and adding pact_reason
If cash_Pact_reason_code <> "Withdraw" then 
	If cash_I_status = "ACTV" then EmwriteScreen Cash_Pact_reason, 6, 58

	
	If cash_II_status = "ACTV" then EmwriteScreen Cash_Pact_reason, 8, 58
End if 

If GRH_status = "ACTV" then EmwriteScreen GRH_Pact_reason, 10, 58


If FS_status = "ACTV" then 
	If FS_Pact_reason_code = "3" then FS_pact_reason_code = "4"

	EmwriteScreen FS_Pact_reason, 12, 58

End If



'extra steps are needed to access HC opitions. 
If HC_status = "ACTV" then 

	EMWritescreen "x", 16, 25

	transmit

	EMWritescreen Pact_reason, 9, 41

	
	PF3
End If




'close panel
transmit

'sending case through back ground.
pf3




 transmit

'ph3 to send case through background. 
pf3
pf3
pf3


'If FS_status = "ACTV" or "PEND" then FS_check_box = checked
'IF cash_I_status = "ACTV" or "PEND" then Cash_check_box + checked 
'If cash_II_status,= "ACTV" or  
 'GRH_status, 4, 9, 74
'FS_status, 4, 10, 74
 'HC_status, 4, 12, 74




