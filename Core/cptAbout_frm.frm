VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} cptAbout_frm 
   Caption         =   "The ClearPlan Toolbar"
   ClientHeight    =   4410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   OleObjectBlob   =   "cptAbout_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "cptAbout_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<cpt_version>v1.9.6</cpt_version>
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub lblScoreBoard_Click()
  Dim strScoreboard As String

  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b1" EWR > MSY '3/22/19 = 1
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b2" MSY > EWR '3/24/19 = 2
  strScoreboard = "2019-03-22 EWR > MSY" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b1" & vbCrLf
  strScoreboard = strScoreboard & "2019-03-24 MSY > EWR" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b2" & vbCrLf
  
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b3" 'EWR > SAN '10/25/19 = 3
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b4" 'SAN > EWR '10/27/19 = 4
  strScoreboard = strScoreboard & "2019-10-25 EWR > SAN" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b3" & vbCrLf
  strScoreboard = strScoreboard & "2019-10-27 SAN > EWR" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b4" & vbCrLf
  
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b5" 'EWR > NAS '2/17/20 = 5
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b6" 'NAS > EWR '2/20/20 = 6
  strScoreboard = strScoreboard & "2020-02-17 EWR > NAS" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b5" & vbCrLf
  strScoreboard = strScoreboard & "2020-02-20 NAS > EWR" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b6" & vbCrLf
  
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b7" 'EWR > SAV '6/3/22 = 7
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b8" 'EWR > SAV '6/5/22 = 8
  strScoreboard = strScoreboard & "2022-06-03 EWR > SAV" & vbTab & "WIN " & vbTab & "t0" & vbTab & "b7" & vbCrLf
  strScoreboard = strScoreboard & "2022-06-05 SAV > EWR" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b8" & vbCrLf
  
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b9" 'EWR > DFW '5/16/25
  'myAbout_frm.lblScoreBoard.Caption = "t0" & vbtab & "b10" 'DFW > EWR '5/18/25
  strScoreboard = strScoreboard & "2025-05-16 EWR > DFW" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b9" & vbCrLf
  strScoreboard = strScoreboard & "2025-05-18 DFR > EWR" & vbTab & "WIN" & vbTab & "t0" & vbTab & "b10" & vbCrLf
  MsgBox strScoreboard, vbOKOnly, "#Winning"

End Sub

Private Sub lblURL_Click()

  If cptErrorTrapping Then On Error GoTo err_here Else On Error GoTo 0

  If cptInternetIsConnected Then Application.FollowHyperlink ("http://" & Me.lblURL.Caption)

exit_here:
  On Error Resume Next

  Exit Sub
err_here:
  Call cptHandleErr("cptAbout_frm", "lblURL", Err, Erl)
  Resume exit_here
End Sub
