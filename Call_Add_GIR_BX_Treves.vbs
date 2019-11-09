

SrcFolder 		= "D:\EDI2\PostProcess\Add_GIR_BX_Treves\Input\"
DstFolder 		= "D:\EDI2\PostProcess\Add_GIR_BX_Treves\Output\"
ArcFolder	 	= "D:\EDI2\PostProcess\Add_GIR_BX_Treves\Input\Archive\"


Set mfso = CreateObject("Scripting.FileSystemObject")
	Set Folder = mfso.GetFolder(srcfolder)
	
	For Each File In Folder.Files
		Set ts = mfso.OpenTextFile(File,1)
    	strDocument = ts.ReadAll
    	ts.Close
		
		
		strDocument = Add_GIR_BX_Treves(strDocument)
		SeqFiles = SeqFiles + 1
		newfile = DstFolder & "Result_" & File.Name 
		Set writefl = mfso.CreateTextFile(newfile, true)
		writefl.Write strDocument
		writefl.Close()
		
	'	File.Copy ArcFolder, True
	'	File.Delete
	Next
	






Function Add_GIR_BX_Treves(SrcStr)
'=========================================================================================================================================
'------- Created By : Rudy Pringadi <hansip101@yahoo.com>
'------- Date 		: Feb 16, 2019
'=========================================================================================================================================
'------- Function reMap GIR segment
'=========================================================================================================================================


    subsep = ":"
    dump = ""
    GetSepDel SrcStr, sep, del
    write = True 
    If del <> "" Then
        rows = Split(SrcStr, del)
        For Each row In rows
            If row <> "" Then
            	changed = False
                cols = Split(row, sep)
                Select Case cols(0)         
              	Case "UNB"
              		subsep = Mid(cols(1), 5, 1)
                Case "UNH"
                    SE01 = 0
					CPS01 = -1        
					SE01x = 0
					BOXe = 0
					before = ""
                Case "CPS"
                	dump = Replace(dump, "{BOXe}", BOXe)
                	BOXe = 0
                	cols(3) = "1"
                	changed = True
                	LIN = False
                	ReDim mGIR(-1)
                Case "PAC"
              		cols(1) = "{BOXe}"
              		changed = True
                Case  "UNT"
                	dump = Replace(dump, "{BOXe}", BOXe)
                	SE01 = SE01 + 1
                	cols(1) = SE01 
                	changed = True
                Case "PCI"
                	row = ""
                	PCI01 = cols(1)
                Case "RFF"
                	dmps = Split(cols(1), subsep)
                	If dmps(0) = "AAT" And before = "PCI" Then
                		row = ""
                		RFF_ATT0102 = dmps(1)
                	ElseIf dmps(0) = "ON" Then
                		For x = 0 To UBound(GIR)
                			cls = Split(GIR(x), sep)
                			dcls = Split(cls(2), subsep)
                			ddcls = Split(dcls(0), ";")
                			dx = UBound(ddcls)
                			Select Case x
                			Case 0 
                				
                				ReDim mGIR(dx)
                				For xx = 0 To dx
                					mGIR(xx) = "PCI" & sep & 17 & del & "RFF" & sep & dcls(1) & subsep & ddcls(xx) & del
                					SE01 = SE01 + 2
                				Next
                			Case 1
                				For xx = 0  To dx
                					dddcls = Split(ddcls(xx), ",")
                					For xxx = 0 To UBound(dddcls)
                						mGIR(xx) = mGIR(xx) & "GIR" & sep & "3" & sep & dddcls(xxx) & subsep & dcls(1) & "{dcl-" & xxx & "}" & del
                						SE01 = SE01 + 1
                						BOXe = BOXe + 1
                					Next
                				Next
                			Case 2
                				For xx = 0  To dx
                					dddcls = Split(ddcls(xx), ",")
                					For xxx = 0 To UBound(dddcls)
                						mGIR(xx) = Replace(mGIR(xx), "{dcl-" & xxx & "}", sep & sep & sep & sep & dddcls(xxx) &  subsep & dcls(1))
                					Next
                				Next
                			End Select
                			
                		Next
                		dds = ""
                		For x = 0 To UBound(mGIR)
                			dds = dds & mGIR(x)
            			Next
            			dump = Replace(dump, "{GIR}", dds)
                	End If
                Case "LIN"
                	LIN = True
                Case "GIR"
                	If LIN Then
                		GIRX = GIRX + 1
                		ReDim Preserve GIR(GIRX)
                		GIR(GIRX) = row
                	Else
                		GIRX = -1
                		ReDim GIR(GIRX)
                		dump = dump & "{GIR}"
                	
                	End If
                	row = ""
                End Select
                If changed Then
                	row = cols(0)
                	For x = 1 To UBound(cols)
                		row = row & sep & cols(x)
                	Next
                End If
                If row <> "" Then
            	    SE01 = SE01 + 1
                    dump = dump & row & del
                
                End If
                before = cols(0)
            End If
        Next
      	SrcStr = dump
   	
    End If
    Add_GIR_BX_Treves = SrcStr
	
End Function


Function GetSepDel(strAll, sep, del) 'New function that can handle X12 EDIFACT
	mchar = ""
	sep = ""
	del = ""
	For x = 1 To Len(strAll)
		nych = Mid(strAll, x, 1)
		If (64 < Asc(nych) And Asc(nych) < 91) Then
			mchar = mchar & nych
		Else
			cols = Split(strAll, nych)
			If Right(mchar, 3) = "UNB" Then
				For yy = 2 To 16
					If Right(cols(yy), 3) = "UNH" Then
						del = Replace(cols(yy), "UNH", "")
						For yyy = yy To 20
							If Right(cols(yyy), 3) = "BGM" Then
								dd = Replace(cols(yyy), "BGM", "")
								xx = 1
								Do While Right(del , xx) = Right(dd, xx)
									mdel = Right(del, xx)
									xx = xx + 1
								Loop
								del = mdel
								Exit For
							End If
						Next
						sep = nych
						Exit For
					End If
				Next
				Exit For
			ElseIf Right(mchar, 3) = "ISA" Then
				For yy = 10 To 19
					If (Right(cols(yy), 2)) = "GS" Then
						del = Replace(cols(yy), "GS", "")
						xx = 1
						Do While Right(del , xx) = Right(strAll, xx)
							mdel = Right(del, xx)
							xx = xx + 1
						Loop
						del = mdel
						sep = nych
						Exit For
					End If
				Next
				Exit For
			End If
		End If
		If x > 50 Then
			Exit For
		End If
	Next
	
End Function

Function ReIndexEDIFACT_CPS(strval)
	'On Error Resume Next
	
    	brs = -1
    	GetSepDel strval, sep, del
    	
    	dump = ""
    	If del <> "" Then
    		pos = InStr(strval, "UNOA")
	    	If pos > 0 Then
	    		subsep = Mid(strval, pos + 4, 1)
	    	Else
	    		subsep = ":"
	    	End If
    		rows = Split(strval, del)
    		write = True
    		For Each row In rows
    			If row <> "" Then
    				col = Split(row, sep)
		    		Select Case col(0)
		    		Case "UNH"
		    			SE01 = 0
		    			QTYFirst = False
		    			RFFFirst = False
		    			Seq = 0
		    		Case "CPS"
		    			Seq = Seq + 1
		    			row = col(0)
		    			col(1) = Seq
		    			For x = 1 To UBound(col)
		    				row = row & sep & col(x)
		    			Next
		    		End Select
    				If row <> "" And write Then
    					SE01 = SE01 + 1
    					dump = dump & row & del
    				End If
    			End If
    		Next
    		strval = dump	
    	End If
    	
    	ReIndexEDIFACT_CPS = strval
    	
End Function
