# VBA-Material_Data_Pre_Check
a VBA script that searches for existing data regarding the material number of every open PR. 

![image](https://user-images.githubusercontent.com/50633734/150686872-cfe4e24a-8cde-41ce-acd9-0e7a9b2edc1b.png)


Sub Run()
'Private Sub Auto_Open()

'Private Sub Auto_open () makes the macro run as soon as the excel file is open


Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If


'set date variables
Today = Format(Sheets("Info").Range("a1"), "DD.MM.YYYY")
yesterday = Format(Sheets("Info").Range("B1"), "DD.MM.YYYY")

With SAPsession
        .findById("wnd[0]/tbar[0]/okcd").Text = "/n/BASF/Tbox_toolbox"
        .findById("wnd[0]").sendVKey 0
        
'Runs Variant

        .findById("wnd[0]/tbar[1]/btn[17]").press
        .findById("wnd[1]/usr/txtENAME-LOW").Text = ""
        .findById("wnd[1]/usr/txtENAME-LOW").SetFocus
        .findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
        .findById("wnd[1]").sendVKey 0
        .findById("wnd[1]/usr/txtV-LOW").Text = "Commentsmacro"
        .findById("wnd[1]/usr/txtV-LOW").SetFocus
        .findById("wnd[1]/usr/txtV-LOW").caretPosition = 13
        .findById("wnd[1]").sendVKey 0
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/ctxtS_FRGDT-HIGH").Text = Today
        .findById("wnd[0]/usr/ctxtS_FRGDT-LOW").Text = yesterday
        .findById("wnd[0]/tbar[1]/btn[8]").press
            
            
 'Copy to clipboard the table after being filtered
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectContextMenuItem "&PC"
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
        .findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
        .findById("wnd[1]/tbar[0]/btn[0]").press

    
        End With
        
    
    
'Paste information into Excel and fill in Overview
Sheets("Comments").Activate
    Sheets("Comments").Select
    Sheets("Comments").Cells.ClearContents
    
    Sheets("Comments").Select
    Sheets("Comments").Range("A1").Select
    
    Application.Wait (Now + TimeValue("0:00:05"))
    ActiveSheet.Paste
    
    FinalRow = Sheets("Comments").Cells(Rows.Count, 1).End(xlUp).Row

    Sheets("Comments").Range("A1:A" & FinalRow).Select
        
'Changes the format Text to Columns
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), Array(23, 1)), _
        TrailingMinusNumbers:=True
    Selection.Delete Shift:=xlToLeft
    
'Deletes uneccesary rows
    Sheets("Comments").Range("1:3").Delete
    Rows(2).EntireRow.Delete
    Rows(3).EntireRow.Delete
    

'Look for previous PO data
'ME2M
lastitem = Sheets("Comments").Cells(Rows.Count, 13).End(xlUp).Row


Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If

 
 lastitem = Sheets("Comments").Cells(Rows.Count, 13).End(xlUp).Row
   
 With SAPsession
 
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nME2M"
    .findById("wnd[0]").sendVKey 0
    
    
   For e = 2 To lastitem
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nME2M"
    .findById("wnd[0]").sendVKey 0
    
    Material = Sheets("Comments").Range("M" & e).Value

    .findById("wnd[0]/usr/ctxtEM_MATNR-LOW").Text = Material
    .findById("wnd[0]/usr/ctxtSELPA-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtEM_EKORG-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_EKGRP-LOW").Text = ""
    .findById("wnd[0]/usr/txtS_EAN11-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_BEDAT-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_LIFNR-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_RESWK-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtEM_EKORG-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_BSART-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = ""
    .findById("wnd[0]/tbar[1]/btn[8]").press
    If SAPsession.findById("wnd[0]/sbar").Text = "No existe ningún documento de compras adecuado" Then 'add error if no found in english
        Sheets("Comments").Range("R" & e).Value = ""
    Else
        
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "EBELN"
    
 'copy po number
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"


Application.Wait (Now + TimeValue("0:00:01"))

    Sheets("Comments").Activate
    Sheets("Comments").Range("R" & e).PasteSpecial
   
   .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "BEDAT"
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"

   
    Sheets("Comments").Range("s" & e).PasteSpecial
  
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "SUPERFIELD"
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"

   
    Sheets("Comments").Range("T" & e).PasteSpecial
  
  
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "NETPR"
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"

   
    Sheets("Comments").Range("U" & e).PasteSpecial
  
    End If
    
    Next e
    
End With

'///////////////////////////////////////////////////////////

'Look for existing OA

Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If
lastitem = Sheets("Comments").Cells(Rows.Count, 13).End(xlUp).Row
 With SAPsession
 


    For e = 2 To lastitem
    
    .findById("wnd[0]/tbar[0]/okcd").Text = "/nME3M"
    .findById("wnd[0]").sendVKey 0
    
    Material = Sheets("Comments").Range("M" & e).Value

    .findById("wnd[0]/usr/ctxtEM_MATNR-LOW").Text = Material
    .findById("wnd[0]/usr/ctxtEM_EKORG-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = ""
    
    .findById("wnd[0]/tbar[1]/btn[8]").press
    If SAPsession.findById("wnd[0]/sbar").Text = "No existe ningún documento de compras adecuado" Then
        Sheets("Comments").Range("V" & e).Value = ""
    Else
    
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 1, "EBELN"
    
    
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"

   Application.Wait (Now + TimeValue("0:00:05"))
    Sheets("Comments").Range("V" & e).PasteSpecial
   
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 1, "EBELP"
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"

   
    Sheets("Comments").Range("W" & e).PasteSpecial
    
    
'Check if there are more than 1 row
'To decide if you want to develop this chunk of code


    If SAPsession.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "2" Then
        
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 2, "EBELN"
    
    
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"

   Application.Wait (Now + TimeValue("0:00:01"))
    Sheets("Comments").Range("X" & e).PasteSpecial
    
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 2, "EBELP"
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
    .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"

   Application.Wait (Now + TimeValue("0:00:01"))
    Sheets("Comments").Range("Y" & e).PasteSpecial
    
        End If
    

        End If
        Next e
        
    End With

Call GetOldMro
Call SearchDataOldMRO
Call Potential_Vendors
Call AddComments

End Sub

Sub Potential_Vendors()

'If No PO history or OAs, then look for potential suppliers
Dim Matnum As Variant

Today = Format(Sheets("Info").Range("a1"), "DD.MM.YYYY")
yesterday = Format(Sheets("Info").Range("B1"), "DD.MM.YYYY")

Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If
lastitem = Sheets("Comments").Cells(Rows.Count, 13).End(xlUp).Row
 With SAPsession
 


    For e = 2 To lastitem
    
If Sheets("Comments").Range("R" & e).Value = "" And Sheets("Comments").Range("v" & e).Value = "" And Sheets("Comments").Range("Y" & e).Value = "" Then

    .findById("wnd[0]/tbar[0]/okcd").Text = "/nME2M"
    .findById("wnd[0]").sendVKey 0
    
    Material = Sheets("Comments").Range("M" & e).Value
    Matnum = Sheets("Comments").Range("D" & e).Value
    Plant = Sheets("Comments").Range("F" & e).Value
    
    Application.DisplayAlerts = False
    On Error Resume Next
    
    .findById("wnd[0]/usr/ctxtEM_MATNR-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtSELPA-LOW").Text = ""
    .findById("wnd[0]/usr/ctxtS_MATKL-LOW").Text = Matnum
    
    'format If statement because we (Spain) have one plant that starts with a "3"... LOL
    'check if other regions have plants/sites that start with a number
    If Sheets("Comments").Range("F" & e).Value = "30000000000" Then
        .findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = "3e10"
    Else
    .findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = Plant
    End If
    
    
    .findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").Text = Today
    
    .findById("wnd[0]/usr/ctxtS_BEDAT-LOW").Text = "01.01.2020" 'Maybe set to variable date (?),
    
    .findById("wnd[0]/usr/ctxtS_EKGRP-LOW").Text = ""
    .findById("wnd[0]/tbar[1]/btn[8]").press
    If SAPsession.findById("wnd[0]/sbar").Text = "No existe ningún documento de compras adecuado" Then GoTo Nextiteration
    
    
  'code to scroll
  
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").lastVisibleRow = 20
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 97
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 132
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 167
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 190
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 210
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 240
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 270
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 300
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 330
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 360
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 390
        On Error Resume Next
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").firstVisibleRow = 420
        
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "SUPERFIELD"
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"
        On Error Resume Next
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        
    
'do all the sorting and stuff to get the top 3 suppliers
Sheets("Data").Activate
    With Sheets("Data")
    
        Sheets("Data").Columns("A").ClearContents
        
        
        
        Sheets("Data").Columns("A:z").Select
        Selection.Delete Shift:=xlToLeft
        Sheets("Data").Columns("A:z").Select
        Selection.Delete Shift:=xlToLeft
        Sheets("Data").Range("A1").Select
        ActiveSheet.Paste
        Sheets("Data").Columns("A:A").Select
        Selection.Copy
        Sheets("Data").Range("F1").Select
        ActiveSheet.Paste
        Sheets("Data").Rows("1:1").Select
        Application.CutCopyMode = False
        Sheets("Data").Rows("1:1").Select
        Application.CutCopyMode = False
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets("Data").Columns("F:F").Select
        ActiveSheet.Range("$F$1:$F$50").RemoveDuplicates Columns:=1, Header:=xlNo
        Sheets("Data").Columns("F:F").EntireColumn.AutoFit
        Sheets("Data").Range("G2").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],RC[-1])"
        Sheets("Data").Range("G2").Select
        Selection.AutoFill Destination:=Range("G2:G32"), Type:=xlFillDefault
        Sheets("Data").Range("G2:G32").Select
        Sheets("Data").Range("H2").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=SUM(C[-1])"
        Sheets("Data").Range("H2").Select
        Selection.AutoFill Destination:=Range("H2:H32")
        Sheets("Data").Range("H2:H32").Select
        Sheets("Data").Range("I2").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[-2]/RC[-1]"
        Sheets("Data").Range("I2").Select
        Selection.AutoFill Destination:=Range("I2:I32")
        Sheets("Data").Range("I2:I32").Select
        Sheets("Data").Columns("I:I").Select
        Selection.Style = "Percent"
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        Selection.AutoFilter
        Range("F1:I37").Select
        Selection.AutoFilter
        Selection.AutoFilter
        ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "I2:I37"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Sheets("Data").Range("I1").Select
        ActiveCell.FormulaR1C1 = "p"
        Sheets("Data").Columns("F:I").Select
        Selection.AutoFilter
        Selection.AutoFilter
        ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
            "I1:I50"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Sheets("Data").Range("J2").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = _
            "=RC[-4]&"" "" &RC[-1]&"" ""&R[1]C[-4]&"" ""& R[1]C[-1]&"" ""&R[2]C[-4]&"" ""&R[2]C[-1]"
    
        Sheets("Data").Range("K7").Select
        ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""%"")"
        Sheets("Data").Range("K7").Select
        ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""0,0%"")"
        Sheets("Data").Range("L7").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=RC[-6]&RC[-1]"
        Sheets("Data").Columns("J:J").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Sheets("Data").Range("L7").Select
        Selection.Cut Destination:=Range("J2")
        Sheets("Data").Range("J2").Select
        ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""0,0%"")"
        Sheets("Data").Range("J2").Select
        Selection.AutoFill Destination:=Range("J2:J4"), Type:=xlFillDefault
        Sheets("Data").Range("J2:J4").Select
        Sheets("Data").Range("K2").Select
        ActiveCell.FormulaR1C1 = _
            "=RC[-5]&"" "" &RC[-1]&"" ""&R[1]C[-5]&"" ""& R[1]C[-1]&"" ""&R[2]C[-5]&"" ""&R[2]C[-1]"
        Sheets("Data").Range("L14").Select
        ActiveCell.FormulaR1C1 = ""
        Sheets("Data").Range("J17").Select
        
        
        Sheets("comments").Range("x" & e).Value = Sheets("Data").Range("k2").Value
    
    End With
    
    
    
End If



Nextiteration:
        Next e
        
    End With
   


End Sub

Sub GetOldMro()

'/////////////////////////////////



Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If
lastitem = Sheets("Comments").Cells(Rows.Count, 13).End(xlUp).Row


 With SAPsession
 
    For e = 2 To lastitem
    MRO = Sheets("Comments").Range("M" & e).Value
    
    
'If No PO history or OAs, then look for old MRO field
    
    
If Sheets("Comments").Range("R" & e).Value = "" And Sheets("Comments").Range("v" & e).Value = "" Then
            
            
            .findById("wnd[0]/tbar[0]/okcd").Text = "/nmm03"
            .findById("wnd[0]").sendVKey 0
                       
                
            .findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = MRO
            .findById("wnd[0]").sendVKey 0
            .findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = -1
            .findById("wnd[1]/tbar[0]/btn[6]").press
            
            'take the material and paste it in the excel
            Sheets("Comments").Range("Y" & e).Value = SAPsession.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLYRSPA01_DPTK:2001/txtMARA-BISMT").Text
            
            .findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLYRSPA01_DPTK:2001/txtMARA-BISMT").SetFocus
            .findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLYRSPA01_DPTK:2001/txtMARA-BISMT").caretPosition = 0

            .findById("wnd[0]/tbar[0]/btn[3]").press
                                   
    
    
    
    
End If



Nextiteration:
        Next e
        
    End With

End Sub

Sub SearchDataOldMRO()

Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If

 
lastitem = Sheets("Comments").Cells(Rows.Count, 13).End(xlUp).Row
   
With SAPsession
 

    
    
   For e = 2 To lastitem
       OldMRO = Sheets("Comments").Range("Y" & e).Value
       
       'if statement to check if the sub above found a OldMRo
       If Sheets("Comments").Range("Y" & e).Value > 1 Then
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nME2M"
        .findById("wnd[0]").sendVKey 0
        
        
    
        .findById("wnd[0]/usr/ctxtEM_MATNR-LOW").Text = OldMRO
        .findById("wnd[0]/usr/ctxtSELPA-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtEM_EKORG-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_EKGRP-LOW").Text = ""
        .findById("wnd[0]/usr/txtS_EAN11-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_BEDAT-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_LIFNR-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_RESWK-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtEM_EKORG-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtS_BSART-LOW").Text = ""
        .findById("wnd[0]/usr/ctxtEM_WERKS-LOW").Text = ""
        .findById("wnd[0]/tbar[1]/btn[8]").press
        
            If SAPsession.findById("wnd[0]/sbar").Text = "No existe ningún documento de compras adecuado" Then  'add english
                Sheets("Comments").Range("R" & e).Value = ""
            Else
            
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "EBELN"
        
      
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"
    
    
        Application.Wait (Now + TimeValue("0:00:01"))
        Sheets("Comments").Activate
        Sheets("Comments").Range("Z" & e).PasteSpecial
       
       .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "BEDAT"
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"
    
       
        Sheets("Comments").Range("AA" & e).PasteSpecial
      
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "SUPERFIELD"
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"
    
       
        Sheets("Comments").Range("AB" & e).PasteSpecial
      
      
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 0, "NETPR"
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
        .findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItemByPosition "0"
    
       
        Sheets("Comments").Range("AC" & e).PasteSpecial
      
            
        End If

        End If
      
        Next e
   

End With


End Sub



Sub AddComments()

'///////////////////////
'Adding Comments in Toolbox

Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If


 With SAPsession
 
 lastitem = Sheets("Comments").Cells(Rows.Count, 13).End(xlUp).Row
 For e = 2 To lastitem
 
    SolPe = Sheets("Comments").Range("k" & e).Value
    PreviousPO = Sheets("Comments").Range("R" & e).Value
    PP_Date = Sheets("Comments").Range("S" & e).Value
    PP_Supp = Sheets("Comments").Range("T" & e).Value
    PP_Price = Sheets("Comments").Range("U" & e).Value
    OA = Sheets("Comments").Range("V" & e).Value
    OA_Pos = Sheets("Comments").Range("W" & e).Value
    PotentialSupp = Sheets("Comments").Range("X" & e).Value
    OldMRO = Sheets("Comments").Range("Y" & e).Value
    OldMRO_PreviousPO = Sheets("Comments").Range("Z" & e).Value
    OldMRO_PP_Date = Sheets("Comments").Range("AA" & e).Value
    OldMRO_Supp = Sheets("Comments").Range("AB" & e).Value
    OldMRO_PP_Price = Sheets("Comments").Range("AC" & e).Value
 
 
    .findById("wnd[0]/tbar[0]/okcd").Text = "/n/BASF/Tbox_toolbox"
    .findById("wnd[0]").sendVKey 0
    
    .findById("wnd[0]/usr/ctxtS_BANFN-LOW").Text = SolPe
    
    '"Declick selection boxes
    .findById("wnd[0]/usr/chkP_PROA").Selected = -1
    .findById("wnd[0]/usr/chkP_MANU").Selected = -1
    .findById("wnd[0]/usr/chkP_PROCES").Selected = -1

'layout is Hard codded :/
    .findById("wnd[0]/usr/ctxtP_VARI").Text = "CommentsMac"
    

    .findById("wnd[0]/tbar[1]/btn[8]").press
    
    
        
'Add comment

        On Error Resume Next
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").currentCellColumn = "RTEXT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").doubleClickCurrentCell
        .findById("wnd[1]/usr/subSUB1:SAPLYMM1_UI_COMPONENTS:0100/cntlTEXTEDIT/shellcont/shell").Text = "CHECK: #" & vbCr & "Previous PO: " & PreviousPO & vbCr & "Date: " & PP_Date & vbCr & "Supp: " & PP_Supp & vbCr & "Price: " & PP_Price & vbCr & "OA: " & OA & vbCr & "Pos: " & OA_Pos & vbCr & "Old MRO: " & OldMRO & vbCr & "Old MRO Previous PO: " & OldMRO_PreviousPO & vbCr & "Old MRO Previous PO Date: " & OldMRO_PP_Date & vbCr & "Old MRO Previous PO Supplier: " & OldMRO_Supp & vbCr & "Old MRO Previous PO Price: " & OldMRO_PP_Price & vbCr & "Potential Supps: " & PotentialSupp & vbCr
        .findById("wnd[1]/usr/subSUB1:SAPLYMM1_UI_COMPONENTS:0100/cntlTEXTEDIT/shellcont/shell").setSelectionIndexes 4, 4
        .findById("wnd[1]/tbar[0]/btn[0]").press
    


        Next e
            
    End With
 
            

End Sub
Sub Shut_itDown()

    ThisWorkbook.Saved = True
    Application.Quit
End Sub


Sub Create_PO_Conditions()

'this part of the macro is not completed, still for development

Dim rng As Range
Sheets("Comments").Activate

'Variables
        FinalRow = Sheets("Comments").Cells(Rows.Count, 1).End(xlUp).Row

'Condition to change Column S to Date
    
        Columns("S:S").Select
        Selection.TextToColumns Destination:=Range("S1"), DataType:=xlFixedWidth, _
            FieldInfo:=Array(0, 4), TrailingMinusNumbers:=True
            
    
    
        Range("S:S").NumberFormat = "mm/DD/yyyy"


        
'Condition to change column U into numbers to only convert POs below 10K

        Columns("U:U").Select
        Selection.TextToColumns Destination:=Range("U1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
        Selection.NumberFormat = "0.00"
        
    
'The next conditions are crammed together
'Condition if data available in column R is on *CELL AD
'Condition if date from column S is less than 6 months *CELL AE
'Condition if value is less than 10K *CELL AF
'Condition to check if OA exist to not convert into PO *CELL AG

        Range("AD2").Select
        Application.CutCopyMode = False
        ActiveCell.FormulaR1C1 = "=IF(R[1]C[-12] >1, ""YES"",""NO"")"
        Range("AE2").Select
        ActiveCell.FormulaR1C1 = "=IF(TODAY()-RC[-12]>180,""NO"",""YES"")"
        Range("AF2").Select
        ActiveCell.FormulaR1C1 = "=IF(RC[-11]<10000,""YES"",""NO"")"
        Range("AG2").Select
        ActiveCell.FormulaR1C1 = "=IF(RC[-11]>1,""NO"",""YES"")"
        Range("AH2").Select
        ActiveCell.FormulaR1C1 = _
            "=IF(AND(RC[-4]=""YES"",RC[-3]=""YES"",RC[-2]=""YES"",RC[-1]=""YES""),""CONVERT2PO"",""DONT"")"
        Range("AH3").Select


'Condition to drag and drop condition columns to lastvalue
        Range("AD2").Select
        Selection.AutoFill Destination:=Range("AD2:AD" & FinalRow)
        Range("AE2").Select
        Selection.AutoFill Destination:=Range("AE2:AE" & FinalRow)
        Range("AF2").Select
        Selection.AutoFill Destination:=Range("AF2:AF" & FinalRow)
        Range("AG2").Select
        Selection.AutoFill Destination:=Range("AG2:AG" & FinalRow)
        Range("AH2").Select
        Selection.AutoFill Destination:=Range("AH2:AH" & FinalRow)
    
'Change Column AH to only values for condition

        For Each rng In Range("AH:AH")
            If rng.HasFormula Then
                rng.Formula = rng.Value
                
            End If
        Next rng
        


End Sub


Sub CreatePO()

'this part of the macro is not completed, still for development

'condition for PRs with multiple positions
Sheets("Comments").Activate

    Columns("T:T").Select
    Selection.Copy
    Columns("AI:AI").Select
    ActiveSheet.Paste
    Range("AI11").Select

    Columns("AI:AI").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("AI1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
        
        
        
'            Columns("U:U").Select
'   Selection.TextToColumns Destination:=Range("U1"), DataType:=xlDelimited, _
  '      TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
   '     Semicolon:=False, Comma:=False, Space:=True, Other:=False, OtherChar _
    '    :="|", FieldInfo:=Array(1, 2), TrailingMinusNumbers:=True

Set Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
If Not IsObject(session) Then
   Set SAPsession = Connection.Children(0)
End If


 With SAPsession
 
 lastitem = Sheets("Comments").Cells(Rows.Count, 13).End(xlUp).Row
 For e = 2 To lastitem


    PR = Sheets("Comments").Range("K" & e).Value
    PurchOrg = Sheets("Comments").Range("G" & e).Value
    Supplier = Sheets("Comments").Range("AI" & e).Value
    Price = Sheets("Comments").Range("U" & e).Value
    LastPO = Sheets("Comments").Range("R" & e).Value
    LastPODate = Sheets("Comments").Range("S" & e).Value
    
If Sheets("Comments").Range("AH" & e).Value = "CONVERT2PO" Then
.findById("wnd[0]/tbar[0]/okcd").Text = "/n/BASF/Tbox_toolbox"

.findById("wnd[0]").sendVKey 0

.findById("wnd[0]/usr/ctxtS_BANFN-LOW").Text = PR
.findById("wnd[0]/usr/ctxtP_VARI").Text = "/DIEGO"
.findById("wnd[0]/tbar[1]/btn[8]").press

'selecting all the columns in the toolbox:

        
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").setCurrentCell -1, ""
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "LVO_RELE"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ERDAT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ATTACH"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "MATKL"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "EKGRP"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "WERKS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BUKRS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BNFPO"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BSART"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ZISSUE"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BANFN"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "FLIEF"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "MATNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "QUALCODE"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "S2CACTION"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "MAKTX"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ANFNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "RTEXT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ZREASON"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ZTOT_PR_EUR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "RFQ_DEADLINE"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ZACTION"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "HEAD_TEXT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "LOEKZ"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BADAT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BSART_PO"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "STATU"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "FRGGR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "FRGKZ"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "FRGDT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "FRGST"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "AGING_DAYS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "EKNAM"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "TXZ01"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "LGORT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "MENGE"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "MEINS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "PREIS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "PEINH"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "WAERS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "LAND1"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BEDNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "WGBEZ60"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "LFDAT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "PSTYP"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "KNTTP"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "LIFNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "NAME"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "F_NAME"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "EKORG"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "KONNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "KTPNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ERFUE"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "INFNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "DISPO"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "EBELN"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "EBELP"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ANFPS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BEDAT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "MATKL_PO"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "EKGRP_PO"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "AFNAM"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "RLWRT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "LBATCH_DT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "LBATCH_TIME"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BTCH_ST"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ERROR_CNT"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BLOCKBY"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BLOCK"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "SRM_CONTRACT_ID"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "SRM_CONTRACT_ITM"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ERNAM"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "EBAKZ"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ESTKZ"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "FIXKZ"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "PLIFZ"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "E_CLASS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ZEQUNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "FORDN"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "FORDP"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "WEPOS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "WEUNB"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "REPOS"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "MFRPN"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BSMNG"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "GSFRG"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "BANPR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ZPROCESSID"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ZVSBED"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "IDNLF"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "RESPONSE"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "ZTOT_PR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectColumn "RFQ_LIFNR"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").selectedRows = "0"
        .findById("wnd[0]/usr/tabsTAB_PR/tabpTAB_PR_FC1/ssubTAB_PR_SCA:/BASF/TBOX_P2P_TOOLBOX:0501/cntlCUSTOM_501/shellcont/shell").pressToolbarButton "ME21N"
        .findById("wnd[1]/usr/ctxtEKKO-EKORG").Text = PurchOrg
        .findById("wnd[1]/usr/ctxtEKKO-EKORG").SetFocus
        .findById("wnd[1]/usr/ctxtEKKO-EKORG").caretPosition = 4
        .findById("wnd[1]/tbar[0]/btn[5]").press



.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").SetFocus
.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = Supplier
.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = Supplier
.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").caretPosition = 0







.findById("wnd[0]").sendVKey 0

'# to addddddd
' i think the problem with the supplier thingy is the ir, to check
'if IR existen then go to next
' There should be a way how to in case of error then go to else. Maybe add a way to identify. Or find the reason why
' is different

.findById("wnd[0]").sendVKey 0

.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,0]").Text = ""
.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,0]").Text = Price
.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,0]").SetFocus
.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,0]").caretPosition = 14
.findById("wnd[0]").sendVKey 0


'if error (tab not shown) open it


.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/cntlTEXT_TYPES_0100/shell").selectedNode = "F26"
.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/cntlTEXT_TYPES_0100/shell").topNode = "F24"

.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").Text = "Price taken from last PO " & LastPO & " from the " & LastPODate & " Made by Macro" & vbCr & "" & vbCr & "" & vbCr & "" & vbCr & ""
.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").Text = "Price taken from last PO " & LastPO & " from the " & LastPODate & " Made by Macro" & vbCr & "" & vbCr & "" & vbCr & "" & vbCr & ""
.findById("wnd[0]/tbar[0]/btn[11]").press


End If
        Next e
            
    End With

End Sub

