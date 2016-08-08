Sub main()
   Dim oC2, oC5, oC7, oC10, oC11, oC12, oCR, oC13 As Range
   Dim lLR2, lLR5, lLR10, llR11, lLR13, lI2, lI5, lI11, lI13, lStart As Long
   Dim sNumber, sSupplier, sSD, sTo, sGCM, sCC As String
   Dim oOL As New Outlook.Application
   Dim oMI As MailItem
   Dim sBody As String
   Application.ScreenUpdating = False
   With Sheet2
      Set oC2 = .Cells(1, 1)
      If .FilterMode Then
         .ShowAllData
      End If
      lLR2 = .Cells(.Rows.Count, 1).End(xlUp).Row
   End With
   With Sheet5
      Set oC5 = .Cells(1, 1)
      If .FilterMode Then
         .ShowAllData
      End If
      lLR5 = .Cells(.Rows.Count, 1).End(xlUp).Row
   End With
   With Sheet10
      .Cells.Delete
      .Cells.Clear
      Set oC10 = .Cells(1, 1)
   End With
   With Sheet7
      Set oC7 = .Cells(1, 1)
      If .FilterMode Then
         .ShowAllData
      End If
      oC7.CurrentRegion.AdvancedFilter _
         Action:=xlFilterCopy, CriteriaRange:=Sheet9.Range("A1:H2"), CopyToRange:=oC10, Unique:=xlYes
      .Cells.AutoFilter
   End With
   With oC10
      .CurrentRegion.Sort Key1:=.Offset(0, 6), Key2:=.Offset(0, 3), Header:=xlYes
   End With
   With Sheet11
      .Cells.Delete
      .Cells.Clear
      Set oC11 = .Cells(1, 1)
   End With
   With Sheet10
      lLR10 = .Cells(.Rows.Count, 1).End(xlUp).Row
   End With
   Range(oC10, oC10.Offset(lLR10 - 1, 0)).AdvancedFilter _
      Action:=xlFilterCopy, CopyToRange:=oC11, Unique:=xlYes
   oC11.CurrentRegion.Sort Key1:=oC11, Header:=xlYes
   With Sheet11
      llR11 = .Cells(.Rows.Count, 1).End(xlUp).Row
   End With
   With Sheet12
      Set oC12 = .Cells(2, 1)
      Set oCR = .Range("A1:G2")
   End With
   Set oOL = CreateObject("Outlook.Application")
   For lI11 = 1 To llR11 - 1
      With Sheet13
         .Cells.Delete
         .Cells.Clear
         Set oC13 = .Cells(1, 1)
      End With
      sNumber = oC11.Offset(lI11, 0).Value
      oC12.Value = sNumber
      oC10.CurrentRegion.AdvancedFilter _
         Action:=xlFilterCopy, CriteriaRange:=oCR, CopyToRange:=oC13, Unique:=xlYes
      With Sheet13
         lLR13 = .Cells(.Rows.Count, 1).End(xlUp).Row
      End With
      For lI2 = 1 To lLR2 - 1
         If oC2.Offset(lI2, 0).Value = sNumber Then
            sSD = oC2.Offset(lI2, 23).Value
            Exit For
         End If
      Next lI2
      For lI5 = 1 To lLR5 - 1
         sSupplier = ""
         sTo = ""
         sGCM = ""
         sCC = ""
         If oC5.Offset(lI5, 0).Value = oC12.Value Then
            sSupplier = oC5.Offset(lI5, 1).Value
            sTo = oC5.Offset(lI5, 4).Value
            sGCM = oC5.Offset(lI5, 3).Value
            sCC = oC5.Offset(lI5, 5).Value
            lStart = InStrRev(sGCM, "(") + 1
            sGCM = Mid(sGCM, lStart, Len(sGCM) - lStart)
            sGCM = sGCM & "@amazon.com"
            Exit For
         End If
      Next lI5
      Set oMI = oOL.CreateItem(olMailItem)
      With oMI
         Debug.Print .Session
         .To = sTo
         .CC = "conflictminerals@amazon.com;sleu@sourceintelligence.com;lab126@sourceintelligence.com;" & sCC
         .Subject = "Amazon Conflict Minerals Follow Up - " & sNumber & " - " & sSupplier & " - Quality Control & Due Diligence issues"
         sBody = "<body style='color:rgb(31,73,125);font:13px arial'>"
         
         sBody = sBody & "Dear " & sSupplier & " Team,<br /><br />"
         
         sBody = sBody & "Thank you for providing Lab126/Amazon.com, Inc. with your most recent Conflict Minerals Reporting Template (CMRT).<br /><br />"
         
         sBody = sBody & "Our team has reviewed your responses for quality control and due diligence.  Based on the latest CMRT you have submitted to Source Intelligence (dated "
         
         sBody = sBody & sSD & "), we have found the following issues with the CMRT information you have provided.<br /><br />"
         
         sBody = sBody & "1. Table below identifies the entries and the corresponding issues we have with the information provided. Please refer to the &quot;Issues Classification table&quot; to know more about them and actions required from your end."
         
         sBody = sBody & "<table border='1' cellspacing='0' style='width:70%;color:rgb(31,73,125);font:0.9em calibri'>"
         
         sBody = sBody & "<tr style='color:black'><th style='width:17%;background:rgb(214,220,228)'>Metal</th><th style='width:45%;background:rgb(214,220,228)'>Smelter Name<br />(As in CMRT)</th><th style='width:38%;background:rgb(214,220,228)'>Issue identified</th></tr>"
         
         For lI13 = 1 To lLR13 - 1
            sBody = sBody & "<tr><td style='text-align:center'>" & oC13.Offset(lI13, 2).Value & "</td><td>" & _
               oC13.Offset(lI13, 3).Value & "</td><td>" & oC13.Offset(lI13, 6).Value & "</td></tr>"
         
         Next lI13
         
         sBody = sBody & "</table><br />"
         
         sBody = sBody & "<span style='font:bold;text-decoration:underline;'>Issues Classification Table:</span>"
         
         sBody = sBody & "<table border='1' cellspacing='0' style='width:70%;color:rgb(31,73,125);font:0.85em arial'>"
         
         sBody = sBody & "<tr style='color:black;font:14px calibri'><th style='background:rgb(214,220,228)'>Issue</th><th colspan='3' style='background:rgb(214,220,228)'>Action Required</th></tr>"
         
         sBody = sBody & "<tr><td style='width:16%;font:bold'>Metal is not known to be processed at this smelter</td><td style='width:41%;'>Provide evidence that metal is processed at this smelter</td><td rowspan='3' style='width:3%;font:bold;text-align:center'>or</td><td rowspan='3' style='width:35%;'><br />If you find that the entry is an error, remove the reference and re-submit a new CMRT with the updated information"
         
         sBody = sBody & "<br /><br /><span style='font:bold 11px'>Note: Many a times supplier's CMRT is a consolidation of sub-tier sourcing information and these issues arise from wrong input on sub-supplier's CMRTs. We suggest that you review your sub-supplier CMRTs for these errors (and reach back to them if necessary) so that future occurrences can be prevented.</span><br /><br /></td></tr>"
         
         sBody = sBody & "<tr><td style='font:bold'>Entry not proven to be a smelter</td><td>Provide evidence that entry is a smelter</td></tr>"
         
         sBody = sBody & "<tr><td style='font:bold'>Group company</td><td>Remove reference to Group company and provide details of underlying smelters via an updated CMRT</td></tr>"
         
         sBody = sBody & "</table>"
         
         sBody = sBody & "<span style='font:italic bold 11px'>* Evidence should be in the form of an industry accepted certification<br />"
         
         sBody = sBody & "** CMRT - <a href='http://www.conflictfreesourcing.org/conflict-minerals-reporting-template/'>Conflict Minerals Reporting Template</a></span><br /><br />"
         
         sBody = sBody & "If you have any questions, or you are not the person within your organization responsible for conflict minerals reporting, please forward this email to the correct person and copy conflictminerals@amazon.com, sleu@sourceintelligence.com, lab126@sourceintelligence.com.<br /><br />"
         
         sBody = sBody & "<span style='font:bold'>Thank you,<br />Sam Leu<br />Program Manager, Team Lead | Source Intelligence</span><br /><br />"
         
         sBody = sBody & "The Amazon's Conflict Minerals policy can be found in our website as part of Supply Chain's Supplier Code of Conduct. <a href='http://www.amazon.com/gp/help/customer/display.html?nodeId=200885140'>http://www.amazon.com/gp/help/customer/display.html?nodeId=200885140</a><br /><br />"
         
         sBody = sBody & "<span style='font:italic bold'>&quot;Amazon is committed to avoiding the use of minerals that have fueled conflict in the Democratic Republic of the Congo or an adjoining country. We expect suppliers to support our effort to identify the origin of designated minerals used in our products.&quot;</span>"
         
         sBody = sBody & "</body>"
         .HTMLBody = sBody
         .Display
      End With
      Set oMI = Nothing
   Next lI11
   Set oOL = Nothing
   Application.ScreenUpdating = True
   'MsgBox ("Well done!")
End Sub
