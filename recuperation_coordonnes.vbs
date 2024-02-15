' F3_TEST REFERENCIEL
' 
' Fill depending entry field after changes to the selection field (1. field)
'
'
Option Explicit
On Error Goto 0

'------------------------------------------------------------------------------
' Consts
'------------------------------------------------------------------------------
Const adStateClosed     = 0 'Indicates that the object is closed.
Const adStateOpen       = 1 'Indicates that the object is open.
Const adStateConnecting = 2 'Indicates that the object is connecting.
Const adStateExecuting  = 4 'Indicates that the object is executing a command.
Const adStateFetching   = 8 'Indicates that the rows of the object are being retrieved.    

'------------------------------------------------------------------------------
' Definitions
'------------------------------------------------------------------------------
Dim tab, xls, cs, Conn
  tab= "RPG_Pilote_EDF"
  xls= "C:\Program Files (x86)\JuK\DpuScan\UDD\RPG_Pilote_EDF.xlsx"
  cs = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+xls+";Extended Properties='Excel 12.0 Xml;HDR=YES'"

Set Conn= nothing
 
'------------------------------------------------------------------------------
' EntryPoint
'------------------------------------------------------------------------------
 
  '%(D.SMARTGEC_PRENOM) = "prenom quelconque"
  ' essai de dalil  *******************/ 
  
  
Function ExtFunction (REF_BENEF)
Dim arRet(2) ' Array for the return values
Dim res


  res= GetRecord (REF_BENEF,cs,"","",tab)
  

	arRet(0)="[84]" 		 ' Special command 84 "Fill the UDD, start with next line"
	arRet(1)="||*||*||Prenom|"+res(2)+"||Identifian Dest |"+res(0)+" ||Email |"+res(1)+"||Code Corbeil | ||*||*||*||*||*||*||*||*"

'MsgBox arRet(1)
  ExtFunction=arRet	'Assign the array
  
  
  ' debug de dalil  *******************/ 
  'MsgBox "Le contenu de ExtFunction! 2 -> " + "||identifiant = "+res(0)+"||mail = "+res(1)+"||nom = "+res(2)+"|| 2eme id = "+res(3)+""
  ' essai de dalil  *******************/ 
  	
End Function

'------------------------------------------------------------------------------
' XLS-Query by DB functions
'------------------------------------------------------------------------------

Function GetRecord (IREF_BENEF,cs,user,pass,table)
dim sql, oRs
Dim i,res(20)

'MsgBox "IREF_BENEF = "+IREF_BENEF



On Error Resume Next

  For i=0 TO UBOUND(res)
    res(i)= ""
  Next

  Err.Clear
  If Conn Is nothing Then Set Conn   = CreateObject("ADODB.Connection")
  If (Conn.State And adStateOpen) <> adStateOpen Then
    Conn.Open cs, user, pass
  End If
  If Err.Number <> 0 Then
    MsgBox Err.Description
  Else
    'sql= "SELECT * FROM ["+table+"$] WHERE F3 ='" & IREF_BENEF & "'"
    sql= "SELECT * FROM ["+table+"$] WHERE NOM ='" & IREF_BENEF & "'"
	'MsgBox "la requette sql contien: "&sql
'MsgBox sql,,"SQL"
    Err.Clear
    set oRs= Conn.Execute(sql)
    If Err.Number <> 0 Then
      'MsgBox Err.Description,,"Error"
      MsgBox " on a une erreur ici de type : "&Err.Description
    Else
		If IsObject(oRs) Then
				'MsgBox " l'objet oRs : "
		  If oRs.Fields.Count > 1 Then
			  If 1=1 Then
				'MsgBox "email : "+ oRs(4).Value + "nom :" + oRs("2").Value
				'MsgBox oRs(4).Value 
				IF Not IsNull(oRs(0).Value) Then res(0)= CStr(oRs(0).Value) ' Identifiant											
				IF Not IsNull(oRs(4).Value) Then res(1)= CStr(oRs(4).Value) ' email
				IF Not IsNull(oRs(3).Value) Then res(2)= CStr(oRs(3).Value) ' nom		
				' IF Not IsNull(oRs(16).Value) Then res(3)= CStr(oRs(16).Value) ' 2eme Identifian		
			  Else
				Dim f
				For each f in oRs.Fields
				 MsgBox "Key=" & f.Name & " --- Value=" & f.Value
				Next
			  End If
		  Else
		  End If
		  
		  'MsgBox " l'oRs  n'est pas un objet: "
		  set oRs= Nothing
		End If
    End If
	
		  'MsgBox " on ferme la connexion: "
    Conn.Close
  End If

'  MsgBox Join(res,vbCrLf)
   GetRecord= res

End Function

'-------------------------------------------------------------------------------
' EntryFunction for ScriptDLL inserted by DpuScan at mer., 07.02.2024 13:24
'-------------------------------------------------------------------------------

Const UDD_F3_KEY       = 1
Const UDD_FIELD_ENTER  = 2
Const UDD_FIELD_LEAVE  = 3
Const UDD_DIALOG_OPEN  = 4
Const UDD_DIALOG_CLOSE = 5
Dim   gF3Event 'Global variable for use in EntryPoint

Function ExtFunctionWrapper(vF3Arg,vF3Event)
'$$ DPUSCAN_DECLARATION_BEGIN
'$$ MODE Process
'$$ VAR vF3Arg   = %(V.F3Arg)
'$$ VAR vF3Event = %(V.F3Event)
'$$ VAR ExtFunctionWrapper = %(V.F3Res), %(V.F3Arg)
'$$ DPUSCAN_DECLARATION_END
Dim vF3Res
Dim res


  ' debug de dalil  *******************/ 
 ' MsgBox "Le contenu de de la fonction ExtFunctionWrapper 1 "
  ' essai de dalil  *******************/ 
  
  
  
    'Convert to number and make it global
    gF3Event = CInt(vF3Event)

    Select Case gF3Event
      Case UDD_F3_KEY
          ' Insert your code here ...
      Case UDD_FIELD_ENTER
      Case UDD_FIELD_LEAVE
      Case UDD_DIALOG_OPEN
      Case UDD_DIALOG_CLOSE
    End Select

    res = ExtFunction(vF3Arg)

    If IsArray(res) Then
        If UBound(res)>0 Then
            vF3Res = "" 
            vF3Arg = res(0) & res(1)
        End If
    Else
        vF3Res = res
    End If

    ExtFunctionWrapper = Array( vF3Res, vF3Arg )
End Function
