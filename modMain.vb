'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Guanzon Software Engineering Group
' Guanzon Group of Companies
' Perez Blvd., Dagupan City
'
'    Quick Match Result
'
' Copyright 2020 and Beyond
' All Rights Reserved
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  Jheff [ 01/07/2020 03:53 pm ]
'       start creasting this object
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit On

Imports ggcAppDriver
Imports ggcGOCAS
Imports rmjGOCAS
Imports Newtonsoft.Json.Linq

Module modMain
    Private Const pxeCSSNumber As String = "09158683181" 'default

    Private p_oApp As New GRider("LRTrackr")
    Private p_sCreatedx As String
    Private p_sTransNox As String
    Private p_sLastName As String
    Private p_sFrstName As String
    Private p_sMiddName As String
    Private p_sSuffixNm As String

    Sub Main()
        If Not p_oApp.LoadEnv() Then
            MsgBox("Unable to load configuration file!")
            Exit Sub
        End If

        If Not p_oApp.LogUser("M001111122") Then
            MsgBox("Unable to load configuration file!")
            Exit Sub
        End If


        Call loadCreditOnline()
        End
    End Sub

    Public Sub loadCreditOnline()
        Dim loQMResult As ggcGOCAS.QMResult
        Dim lsQuickMatch As String
        Dim lsSQL As String
        Dim loDta As DataTable
        Dim loJSONAppInfo As JObject
        Dim loJSONResInfo As JObject
        Dim loJSONSpoInfo As JObject
        Dim loJSONSpoResx As JObject
        Dim loJSONCoMaker As JObject

        Dim instance As GOCASCalculator
        Dim lnUnitTpye As String
        Dim lnDownPaym As Double
        Dim loDTBranch As DataTable
        Dim lsGOCASNOx As String
        Dim lsMessagex As String

        'mac 2021.05.20
        Dim lsCoMakerQM1 As String
        Dim lsCoMakerQM2 As String

        Dim lsCSSNumber As String
        lsCSSNumber = p_oApp.getConfiguration("CSSNmbr") 'SMS receiving mobile number of CSS Department
        If lsCSSNumber = "" Then lsCSSNumber = pxeCSSNumber 'assign the pre-defined number if configuration was empty

        lsSQL = "SELECT *" & _
                " FROM Credit_Online_Application" & _
                " WHERE sSourceCd = 'APP'" & _
                    " AND cTranStat IS NULL" & _
                " ORDER BY dTransact"

        Debug.Print(lsSQL)
        loDta = New DataTable

        loDta = ExecuteQuery(lsSQL, p_oApp.Connection)

        If loDta.Rows.Count = 0 Then Exit Sub

        For i = 0 To loDta.Rows.Count - 1
            p_oApp.BeginTransaction()
            Debug.Print(loDta(i)("sTransNox"))
            loJSONAppInfo = CType(JObject.Parse(loDta(i)("sDetlInfo"))("applicant_info"), JObject)
            loJSONResInfo = CType(JObject.Parse(loDta(i)("sDetlInfo"))("residence_info"), JObject)
            loJSONSpoInfo = CType(JObject.Parse(loDta(i)("sDetlInfo"))("spouse_info"), JObject)
            loJSONCoMaker = CType(JObject.Parse(loDta(i)("sDetlInfo"))("comaker_info"), JObject)

            loQMResult = New ggcGOCAS.QMResult

            With loQMResult
                .AppDriver = p_oApp
                .Branch = loDta(i)("sBranchCd")
                .ApplicationNo = loDta(i)("sTransNox")

                .InitTransaction()
                'Set the Applicant info
                .Applicant("sClientID") = ""
                If loJSONAppInfo.GetValue("sLastName") = "" Then GoTo voidTrans
                .Applicant("sLastName") = CStr(loJSONAppInfo.GetValue("sLastName")) & IIf(IFNull(CStr(loJSONAppInfo.GetValue("sSuffixNm"))) = "", "", " " & CStr(loJSONAppInfo.GetValue("sSuffixNm")))
                .Applicant("sFrstName") = CStr(loJSONAppInfo.GetValue("sFrstName"))
                .Applicant("sMiddName") = CStr(loJSONAppInfo.GetValue("sMiddName"))
                .Applicant("dBirthDte") = CStr(loJSONAppInfo.GetValue("dBirthDte"))
                .Applicant("sBirthPlc") = CStr(loJSONAppInfo.GetValue("sBirthPlc"))
                .Applicant("sTownIDxx") = CStr(JObject.Parse(loJSONResInfo.GetValue("present_address").ToString).GetValue("sTownIDxx"))

                p_sCreatedx = loDta(i)("sCreatedx")
                p_sTransNox = loDta(i)("sTransNox")
                p_sLastName = CStr(loJSONAppInfo.GetValue("sLastName"))
                p_sFrstName = CStr(loJSONAppInfo.GetValue("sFrstName"))
                p_sMiddName = CStr(loJSONAppInfo.GetValue("sMiddName"))
                p_sSuffixNm = CStr(loJSONAppInfo.GetValue("sSuffixNm"))

                'Set the spouse info
                If Not IsNothing(loJSONSpoInfo) Then
                    If IFNull(CStr(JObject.Parse(loJSONSpoInfo.GetValue("personal_info").ToString).GetValue("sLastName"))) <> "" Then
                        .Spouse("sClientID") = ""

                        .Spouse("sLastName") = CStr(JObject.Parse(loJSONSpoInfo.GetValue("personal_info").ToString).GetValue("sLastName"))
                        .Spouse("sFrstName") = CStr(JObject.Parse(loJSONSpoInfo.GetValue("personal_info").ToString).GetValue("sFrstName")) & IIf(IFNull(CStr(JObject.Parse(loJSONSpoInfo.GetValue("personal_info").ToString).GetValue("sSuffixNm"))) = "", "", " " & CStr(JObject.Parse(loJSONSpoInfo.GetValue("personal_info").ToString).GetValue("sSuffixNm")))
                        .Spouse("sMiddName") = CStr(JObject.Parse(loJSONSpoInfo.GetValue("personal_info").ToString).GetValue("sMiddName"))
                        .Spouse("dBirthDte") = CStr(JObject.Parse(loJSONSpoInfo.GetValue("personal_info").ToString).GetValue("dBirthDte"))
                        .Spouse("sBirthPlc") = CStr(JObject.Parse(loJSONSpoInfo.GetValue("personal_info").ToString).GetValue("sBirthPlc"))

                        loJSONSpoResx = CType(JObject.Parse(CType(JObject.Parse(loDta(i)("sDetlInfo"))("spouse_info"), JObject).ToString)("residence_info"), JObject)
                        .Spouse("sTownIDxx") = CStr(JObject.Parse(loJSONSpoResx.GetValue("present_address").ToString).GetValue("sTownIDxx"))
                    End If
                End If

                'mac 2021.02.04
                '   added comaker info on QM validation
                If Not IsNothing(loJSONCoMaker) Then
                    If IFNull(CStr(loJSONCoMaker.GetValue("sLastName"))) <> "" Then
                        .CoMaker("sClientID") = ""

                        .CoMaker("sLastName") = CStr(loJSONCoMaker.GetValue("sLastName"))
                        .CoMaker("sFrstName") = CStr(loJSONCoMaker.GetValue("sFrstName")) & IIf(IFNull(CStr(loJSONCoMaker.GetValue("sSuffixNm"))) = "", "", " " & CStr(loJSONCoMaker.GetValue("sSuffixNm")))
                        .CoMaker("sMiddName") = CStr(loJSONCoMaker.GetValue("sMiddName"))
                        .CoMaker("dBirthDte") = CStr(loJSONCoMaker.GetValue("dBirthDte"))
                        .CoMaker("sBirthPlc") = CStr(loJSONCoMaker.GetValue("sBirthPlc"))

                        loJSONSpoResx = CType(JObject.Parse(CType(JObject.Parse(loDta(i)("sDetlInfo"))("comaker_info"), JObject).ToString)("residence_info"), JObject)

                        If Not TypeName(loJSONSpoResx) = "Nothing" Then
                            If IFNull(CStr(JObject.Parse(loJSONSpoResx.GetValue("present_address").ToString).GetValue("sAddress1"))) <> "" Then
                                .CoMaker("sAddressx") = CStr(JObject.Parse(loJSONSpoResx.GetValue("present_address").ToString).GetValue("sAddress1"))
                            End If

                            If IFNull(CStr(JObject.Parse(loJSONSpoResx.GetValue("present_address").ToString).GetValue("sAddress2"))) <> "" Then
                                .CoMaker("sAddressx") = .CoMaker("sAddressx") & " " & CStr(JObject.Parse(loJSONSpoResx.GetValue("present_address").ToString).GetValue("sAddress2"))
                            End If

                            .CoMaker("sAddressx") = Trim(.CoMaker("sAddressx"))
                            .CoMaker("sBrgyIDxx") = CStr(JObject.Parse(loJSONSpoResx.GetValue("present_address").ToString).GetValue("sBrgyIDxx"))
                            .CoMaker("sTownIDxx") = CStr(JObject.Parse(loJSONSpoResx.GetValue("present_address").ToString).GetValue("sTownIDxx"))
                        End If
                    End If
                End If
                'end - mac 2021.02.04

                .Term("sModelIDx") = CStr(CType(JObject.Parse(loDta(i)("sDetlInfo")), JObject).GetValue("sModelIDx"))
                .Term("nDownPaym") = CStr(CType(JObject.Parse(loDta(i)("sDetlInfo")), JObject).GetValue("nDownPaym"))
                .Term("nAcctTerm") = CStr(CType(JObject.Parse(loDta(i)("sDetlInfo")), JObject).GetValue("nAcctTerm"))

                'Execute quickmatch here
                lsQuickMatch = .QuickMatch

                'mac 2021.05.20
                lsCoMakerQM1 = .QuickMatchResult("comaker1")
                lsCoMakerQM2 = .QuickMatchResult("comaker2")

                instance = New GOCASCalculator
                instance.setAppDriver = p_oApp
                instance.setJSON = IFNull(loDta(i)("sCatInfox"), loDta(i)("sDetlInfo"))

                lnDownPaym = getDownpayment(CStr(CType(JObject.Parse(loDta(i)("sDetlInfo")), JObject).GetValue("cUnitAppl")), _
                                    lnUnitTpye, _
                                    CStr(CType(JObject.Parse(loDta(i)("sDetlInfo")), JObject).GetValue("sModelIDx")), _
                                    instance.Compute(), _
                                    CStr(CType(JObject.Parse(loDta(i)("sDetlInfo")), JObject).GetValue("dAppliedx")))

                If lsQuickMatch <> "" Then
                    lsSQL = "SELECT * " & _
                            " FROM Branch_Mobile" & _
                            " WHERE sBranchCd = " & strParm(loDta(i)("sBranchCd"))

                    loDTBranch = New DataTable
                    loDTBranch = ExecuteQuery(lsSQL, p_oApp.Connection)

                    Select Case Trim(Left(lsQuickMatch, 2))
                        Case "DA", "BA"
                            Call getModel(CStr(CType(JObject.Parse(loDta(i)("sDetlInfo")), JObject).GetValue("sModelIDx")), True, True, "", lnUnitTpye)

                            lsSQL = "UPDATE Credit_Online_Application SET" & _
                                        "  sQMatchNo = " & strParm(lsQuickMatch) & _
                                        ", sCoMkrRs1 = " & strParm(lsCoMakerQM1) & _
                                        ", sCoMkrRs2 = " & strParm(lsCoMakerQM2) & _
                                        ", cTranStat = '3'" & _
                                        ", nDownPaym = 90" & _
                                        ", cWithCIxx = '1'" & _
                                        ", sCredInvx = " & strParm(getCreditInvestigator(.Applicant("sTownIDxx"), .Branch)) & _
                                        ", cEvaluatr = '1'" & _
                                        ", sVerified = 'M001180003'" & _
                                        ", dVerified = " & dateParm(p_oApp.getSysDate) & _
                                        ", sGOCASNox = " & strParm(createGOCAS(True, 100)) & _
                                    " WHERE sTransNox = " & strParm(loDta.Rows(i)("sTransnox"))
                            '", cEvaluatr = '0'"
                            For lnCtr As Integer = 0 To loDTBranch.Rows.Count - 1
                                'mac 2021-05-26
                                '   css requested to change the message format
                                lsMessagex = "GOCAS #: " & createGOCAS(True, 100) & vbCrLf & _
                                                "Application of Mr./Ms. " & loDta(i).Item("sClientNm") & " is on Process." & vbCrLf & _
                                                "Valid Until 60 days upon application." & vbCrLf & _
                                                "REF. #: " & loDta(i).Item("sTransNox") & vbCrLf & _
                                                "-GUANZON Group"
                                Call createReply(lsMessagex, loDTBranch(lnCtr)("sMobileNo"), loDta(i).Item("sTransNox"))
                                Call createReply(lsMessagex, lsCSSNumber, loDta(i).Item("sTransNox"))
                            Next
                        Case "SA", "SV", "PA"
                            lsSQL = "UPDATE Credit_Online_Application SET" & _
                                        "  sQMatchNo = " & strParm(lsQuickMatch) & _
                                        ", sCoMkrRs1 = " & strParm(lsCoMakerQM1) & _
                                        ", sCoMkrRs2 = " & strParm(lsCoMakerQM2) & _
                                        ", cEvaluatr = '1'" & _
                                        ", cTranStat = '0'" & _
                                    " WHERE sTransNox = " & strParm(loDta.Rows(i)("sTransnox"))
                            '", cEvaluatr = '0'"
                        Case "CI"
                            'mac 2021.05.20
                            lsSQL = "1"
                            Select Case Trim(Left(lsCoMakerQM1, 2))
                                Case "SA", "SV", "PA", "DA", "BA"
                                    'lsSQL = "0"
                                    lsSQL = "1"
                            End Select

                            If lsSQL = "1" Then
                                Select Case Trim(Left(lsCoMakerQM2, 2))
                                    Case "SA", "SV", "PA", "DA", "BA"
                                        'lsSQL = "0"
                                        lsSQL = "1"
                                End Select
                            End If
                            'end - mac 2021.05.20

                            lsSQL = "UPDATE Credit_Online_Application SET" & _
                                        "  sQMatchNo = " & strParm(lsQuickMatch) & _
                                        ", sCoMkrRs1 = " & strParm(lsCoMakerQM1) & _
                                        ", sCoMkrRs2 = " & strParm(lsCoMakerQM2) & _
                                        ", cTranStat = '0'" & _
                                        ", cWithCIxx = '1'" & _
                                        ", sCredInvx = " & strParm(getCreditInvestigator(.Applicant("sTownIDxx"), .Branch)) & _
                                        ", cEvaluatr = " & strParm(lsSQL) & _
                                    " WHERE sTransNox = " & strParm(loDta.Rows(i)("sTransnox"))
                        Case "AP"
                            'mac 2021.05.20
                            lsSQL = "1"
                            Select Case Trim(Left(lsCoMakerQM1, 2))
                                Case "SA", "SV", "PA", "DA", "BA"
                                    lsSQL = "0"
                            End Select

                            If lsSQL = "1" Then
                                Select Case Trim(Left(lsCoMakerQM2, 2))
                                    Case "SA", "SV", "PA", "DA", "BA"
                                        lsSQL = "0"
                                End Select
                            End If
                            'end - mac 2021.05.20

                            If lsSQL = "0" Then
                                lsSQL = "UPDATE Credit_Online_Application SET" & _
                                            "  sQMatchNo = " & strParm(lsQuickMatch) & _
                                            ", sCoMkrRs1 = " & strParm(lsCoMakerQM1) & _
                                            ", sCoMkrRs2 = " & strParm(lsCoMakerQM2) & _
                                            ", cEvaluatr = '1'" & _
                                            ", cTranStat = '0'" & _
                                        " WHERE sTransNox = " & strParm(loDta.Rows(i)("sTransnox"))

                                '", cEvaluatr = '0'"
                            Else
                                lsGOCASNOx = createGOCAS(False, 200)
                                lsSQL = "UPDATE Credit_Online_Application SET" & _
                                            "  sQMatchNo = " & strParm(lsQuickMatch) & _
                                            ", sCoMkrRs1 = " & strParm(lsCoMakerQM1) & _
                                            ", sCoMkrRs2 = " & strParm(lsCoMakerQM2) & _
                                            ", sGOCASNox = " & strParm(lsGOCASNOx) & _
                                            ", nDownPaym = 200" & _
                                            ", cWithCIxx = '0'" & _
                                            ", cEvaluatr = '0'" & _
                                            ", cTranStat = '1'" & _
                                            ", sVerified = 'M001180003'" & _
                                            ", dVerified = " & dateParm(p_oApp.getSysDate) & _
                                        " WHERE sTransNox = " & strParm(loDta.Rows(i)("sTransnox"))

                                For lnCtr As Integer = 0 To loDTBranch.Rows.Count - 1
                                    'mac 2021-05-26
                                    '   css requested to change the message format
                                    lsMessagex = "GOCAS #: " & lsGOCASNOx & vbCrLf & _
                                                    "Application of Mr./Ms. " & loDta(i).Item("sClientNm") & " is on Process." & vbCrLf & _
                                                    "Valid Until 60 days upon application." & vbCrLf & _
                                                    "REF. #: " & loDta(i).Item("sTransNox") & vbCrLf & _
                                                    "-GUANZON Group"
                                    Call createReply(lsMessagex, loDTBranch(lnCtr)("sMobileNo"), loDta(i).Item("sTransNox"))
                                    Call createReply(lsMessagex, lsCSSNumber, loDta(i).Item("sTransNox"))
                                Next
                            End If
                    End Select
                    GoTo saveTrans

                    'invalid entry
voidTrans:
                    lsSQL = "UPDATE Credit_Online_Application SET" & _
                                " cTranStat = '4'" & _
                            " WHERE sTransNox = " & strParm(loDta.Rows(i)("sTransnox"))


saveTrans:
                    If p_oApp.Execute(lsSQL, "Credit_Online_Application") <= 0 Then
                        p_oApp.RollBackTransaction()
                        Exit Sub
                    End If
                End If
            End With
            Call saveHistory()
            p_oApp.CommitTransaction()
        Next
    End Sub

    Function getDownpayment(ByVal fcLoanType As String, _
                            ByVal fcUnitType As String, _
                            ByVal fsModelIDx As String, _
                            ByVal fnCredtScr As Double, _
                            ByVal fdTransact As Date) As Long
        Dim lsSQL As String
        Dim loDT As DataTable

        lsSQL = "SELECT" & _
                    " IFNULL(b.nDownPaym, a.nDownPaym) nDownPaym" & _
                " FROM Credit_Score_By_Model a" & _
                    " LEFT JOIN Credit_Score_By_Model_History b" & _
                        " ON a.sCSBMIDxx = b.sCSBMIDxx" & _
                        " AND " & dateParm(fdTransact) & " BETWEEN b.dDateFrom AND b.dDateThru" & _
                " WHERE a.sModelIDx = " & strParm(fsModelIDx) & _
                    " AND a.cLoanType = " & strParm(fcLoanType) & _
                    " AND " & fnCredtScr & " BETWEEN a.nScoreFrm AND a.nScoreThr"

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count > 0 Then
            Return loDT(0)("nDownPaym")
        End If

        lsSQL = "SELECT" & _
                    " IFNULL(b.nDownPaym, a.nDownPaym) nDownPaym" & _
                " FROM Credit_Score_By_Type a" & _
                    " LEFT JOIN Credit_Score_By_Type_History b" & _
                        " ON a.sCSBTIDxx = b.sCSBTIDxx" & _
                        " AND " & dateParm(fdTransact) & " BETWEEN b.dDateFrom AND b.dDateThru" & _
                " WHERE a.cUnitType = " & strParm(fcUnitType) & _
                    " AND a.cLoanType = " & strParm(fcLoanType) & _
                    " AND " & fnCredtScr & " BETWEEN a.nScoreFrm AND a.nScoreThr"

        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count > 0 Then
            Return loDT(0)("nDownPaym")
        End If

        Return 0
    End Function

    Function getModel(ByVal sValue As String, ByVal bSearch As Boolean, ByVal bByCode As Boolean, ByRef sModelIDx As String, ByRef cUnitType As String) As String
        Dim lsCondition As String
        Dim lsProcName As String
        Dim lsSQL As String
        Dim loDataRow As DataRow

        lsProcName = "getModel"

        lsCondition = String.Empty

        If sValue <> String.Empty Then
            If bByCode = False Then
                If bSearch Then
                    lsCondition = "sModelNme LIKE " & strParm("%" & sValue & "%")
                Else
                    lsCondition = "sModelNme = " & strParm(sValue)
                End If
            Else
                lsCondition = "sModelIDx = " & strParm(sValue)
            End If
        ElseIf bSearch = False Then
            GoTo endWithClear
        End If

        lsSQL = AddCondition(getSQL_Model, lsCondition)
        Debug.Print(lsSQL)

        Dim loDT As DataTable
        loDT = New DataTable
        loDT = p_oApp.ExecuteQuery(lsSQL)

        If loDT.Rows.Count = 0 Then
            GoTo endWithClear
        ElseIf loDT.Rows.Count = 1 Then
            getModel = loDT(0)("sModelNme")
            sModelIDx = loDT(0)("sModelIDx")
        Else
            loDataRow = KwikSearch(p_oApp, _
                                lsSQL, _
                                "", _
                                "sModelIDx»sModelNme", _
                                "ID»Model", _
                                "", _
                                "sModelIDx»sModelNme", _
                                1)

            If Not IsNothing(loDataRow) Then
                getModel = loDataRow("sModelNme")
                sModelIDx = loDataRow("sModelIDx")
            Else : GoTo endWithClear
            End If
        End If
        loDT = Nothing

endProc:
        Exit Function
endWithClear:
        getModel = ""
        sModelIDx = ""
        GoTo endProc
errProc:
        MsgBox(Err.Description)
    End Function

    Private Function getSQL_Model() As String
        Return "SELECT" & _
                    "  sModelIDx" & _
                    ", sModelNme" & _
                    ", cMotorTyp" & _
                " FROM MC_Model" & _
                " WHERE cRecdStat = " & strParm(xeLogical.YES)
    End Function

    Private Function createGOCAS(ByVal fbIsCINeeded As Boolean, _
                         ByVal fnDownPaymnt As Long) As String
        Dim instance As GOCASCodeGen

        instance = New GOCASCodeGen

        With instance
            .UserID = p_sCreatedx 'created
            .TransactionNo = p_sTransNox 'table transaction number
            .LastName = p_sLastName
            .FirstName = p_sFrstName
            .MiddleName = p_sMiddName
            .SuffixName = p_sSuffixNm
            .IsCINeeded = fbIsCINeeded 'is CI needed
            .DownPayment = fnDownPaymnt 'approved downpayment
            .Encode() 'generate code
        End With

        Return instance.GOCASApprvl
    End Function

    Private Sub createReply(ByVal fsMessages As String, _
                              ByVal fsMobileNo As String, _
                              ByVal fsTransNox As String)
        Dim lsSQL As String

        lsSQL = "INSERT INTO HotLine_Outgoing SET" & _
                    "  sTransNox = " & strParm(GetNextCode("HotLine_Outgoing", "sTransNox", True, p_oApp.Connection, True, p_oApp.BranchCode)) & _
                    ", dTransact = " & dateParm(p_oApp.SysDate) & _
                    ", sDivision = " & strParm("MC") & _
                    ", sMobileNo = " & strParm(fsMobileNo) & _
                    ", sMessagex = " & strParm(fsMessages) & _
                    ", cSubscrbr = " & strParm(classifyMobileNo(fsMobileNo)) & _
                    ", dDueUntil = " & dateParm(DateAdd(DateInterval.Day, 10, p_oApp.SysDate)) & _
                    ", cSendStat = " & strParm("0") & _
                    ", nNoRetryx = " & strParm("0") & _
                    ", sUDHeader = " & strParm("") & _
                    ", sReferNox = " & strParm(fsTransNox) & _
                    ", sSourceCd = " & strParm("APP1") & _
                    ", cTranStat = " & strParm("0") & _
                    ", nPriority = 0" & _
                    ", sModified = " & strParm(p_oApp.UserID) & _
                    ", dModified = " & dateParm(p_oApp.SysDate)

        p_oApp.ExecuteActionQuery(lsSQL)
    End Sub

    Private Function classifyMobileNo(ByVal MobileNo As String) As Integer
        '0 = GLOBE
        '1 = SMART
        Select Case Left(MobileNo, 4)
            Case "0817", "0917", "0994", "0904", "0905", "0906", "0915", "0916", "0917", "0973"
                classifyMobileNo = 0
            Case "0925", "0926", "0927", "0935", "0978", "0979", "0936", "0996", "0997", "0999"
                classifyMobileNo = 0
            Case "0956", "0975", "0965", "0976", "0937", "0966", "0977", "0995", "0945", "0967"
                classifyMobileNo = 0
            Case Else
                classifyMobileNo = 1
        End Select
    End Function

    Private Function saveHistory() As Boolean
        Dim lsSQL As String
        Dim loDT As DataTable

        loDT = New DataTable
        loDT = ExecuteQuery("SELECT * FROM Credit_Online_Application_Verification_History" & _
                                " WHERE sTransNox = " & strParm(p_sTransNox), p_oApp.Connection)

        lsSQL = "INSERT INTO Credit_Online_Application_Verification_History SET" & _
                    "  sTransNox = " & strParm(p_sTransNox) & _
                    ", nEntryNox = " & CDbl(loDT.Rows.Count + 1) & _
                    ", sModified = " & strParm(p_oApp.UserID) & _
                    ", dModified = " & dateParm(p_oApp.SysDate)

        If p_oApp.Execute(lsSQL, "Credit_Online_Application_Verification_History", p_oApp.BranchCode) = 0 Then
            MsgBox("Unable to Save History Info!!!", vbCritical, "Warning")
            Return False
        End If

        Return True
    End Function

    Private Function getCreditInvestigator(ByVal lsValue As String, ByVal lsBranch As String) As String
        Dim lsSQL As String

        lsSQL = "SELECT" & _
                    "  a.sCredInvx" & _
                    ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) sFullName" & _
                " FROM Route_Area a" & _
                        " LEFT JOIN Route_Area_Town b ON a.sRouteIDx = b.sRouteIDx" & _
                    ", Employee_Master001 c" & _
                        " LEFT JOIN Client_Master d ON c.sEmployID = d.sClientID" & _
                " WHERE a.sCredInvx = c.sEmployID" & _
                    " AND a.cTranStat = '1'" & _
                    " AND c.cRecdStat = '1'" & _
                    " AND b.sTownIDxx = " & strParm(lsValue) & _
                    " AND a.sBranchCd = " & strParm(lsBranch) & _
                " GROUP BY a.sCredInvx" & _
                " ORDER BY c.dHiredxxx" & _
                " LIMIT 1"

        Dim loRS As DataTable = p_oApp.ExecuteQuery(lsSQL)

        If loRS.Rows.Count <> 0 Then Return loRS(0)("sCredInvx")

        lsSQL = "SELECT" & _
                    "  a.sCredInvx" & _
                    ", CONCAT(d.sLastName, ', ', d.sFrstName, ' ', d.sMiddName) sFullName" & _
                " FROM Route_Area a" & _
                        " LEFT JOIN Route_Area_Town b ON a.sRouteIDx = b.sRouteIDx" & _
                    ", Employee_Master001 c" & _
                        " LEFT JOIN Client_Master d ON c.sEmployID = d.sClientID" & _
                " WHERE a.sCredInvx = c.sEmployID" & _
                    " AND a.cTranStat = '1'" & _
                    " AND c.cRecdStat = '1'" & _
                    " AND b.sTownIDxx =  " & strParm(lsValue) & _
                " GROUP BY a.sCredInvx" & _
                " ORDER BY c.dHiredxxx" & _
                " LIMIT 1"

        loRS = p_oApp.ExecuteQuery(lsSQL)

        If loRS.Rows.Count = 0 Then
            Return ""
        Else
            Return loRS(0)("sCredInvx")
        End If
    End Function
End Module
