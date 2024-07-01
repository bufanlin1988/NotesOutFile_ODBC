%REM
	Agent FileOutSave[附件转存上传数据]
	Created 2019-10-08 by ID Developer/CN/Schneider
	Description: Comments for Agent
%END REM
Option Public
Option Declare
UseLSX "*LSXODBC"
Sub Initialize      
	On Error GoTo MsgError
	MsgBox "Leap-附件转存上传数据代理开始。。。"	
	Dim odbcConn As New ODBCConnection
	Dim odbcQuery As New ODBCQuery
	Dim odbcRS As New ODBCResultSet	
	Dim strCommand As String
	If Not (odbcConn.ConnectTo ("test","note_adm","9N*o1T6eo")) Then
		MsgBox "OBDC连接错误，请联系统管理员！"
		Exit sub
	Else
		Set odbcQuery.Connection =odbcConn
	'连接正确
		
	Dim session As New NotesSession
	Dim doc As NotesDocument
	Dim docTemp As NotesDocument
	Dim view As NotesView
	Dim bolResult As Boolean 
	Dim FilePath As String
	Dim fileNames As Variant
	Dim CREATEDATE As String
	Dim objEmbed As NotesEmbeddedObject
	Dim db As NotesDatabase
	Dim dbSystem As NotesDatabase  
	Dim x As Integer 
	Set db = session.CurrentDatabase
		Set dbSystem=session.GetDatabase(Db.Server,"China\wf\wf000084.NSF")
	Dim DataCome As String
	DataCome=CStr("wf000084.nsf")
	'Set view = dbSystem.GetView("ErrorLeapFileALL_Show")  '索引需要转存附件的视图
	Set view = dbSystem.GetView("LeapFileALL")  '索引需要转存附件的视图
	
	Set doc = view.GetFirstDocument
	Dim SystemApp As String
	Dim SystemDc As String
     SystemDc="c:\"
     SystemApp="HR_E_CAR"	
		Call SetMainMd(SystemDc,SystemApp) 				'创建外侧主文件夹
	While Not doc Is Nothing
		
		Set docTemp=view.GetNextDocument(doc)
		
		If doc.HasItem("$FILE") Then
			fileNames=Evaluate("@AttachmentNames",doc)  '获取文档所有附件名称
			filePath=SystemDc+SystemApp+"\"	
			If(IsArray(fileNames)) Then
				For x = 0 To UBound(fileNames)
					filePath=SystemDc+SystemApp+"\"+doc.UniversalID+"_"+CStr(x)+"#"+fileNames(x)
					'赋值创建时间
					CREATEDATE=CStr(Format(doc.Created,"YYYY-MM-DD"))
					Set objEmbed = doc.GetAttachment(CStr(fileNames(x)))
					'bolResult=SetUploadSql("0","NOTE_ADM","DIRECT_PAYMENT_FILE",doc.UniversalID,IllegalCharacter(CStr(fileNames(x))),SystemDc+SystemApp,doc.UniversalID+"_"+IllegalCharacter(CStr(fileNames(x))),doc.ParentUNID(0),doc.REQUESTNO(0),doc.unid(0),DATACOME,odbcQuery,odbcRS,CREATEDATE)
					If Not objEmbed Is  Nothing Then 
						Call objEmbed.ExtractFile(filePath)'将文档中的附件导出 
						'执行数据库操作最后参数1为第一次导入
			
						bolResult=SetUploadSql("1","NOTE_ADM","HR_E_CAR_FILE",doc.UniversalID,IllegalCharacter(CStr(fileNames(x))),SystemDc+SystemApp,doc.UniversalID+"_"+CStr(x)+"#"+IllegalCharacter(CStr(fileNames(x))),DATACOME,odbcQuery,odbcRS,doc.PARENTUNID(0),doc.SELFUNID(0),doc.REQUESTNO(0),doc.UNID(0),doc.SAVEOPTIONS(0),CREATEDATE)
				
					Else
						Dim LogInfo As String 
						Dim STATUS As String  
						Dim EXPLAIN As String  
						STATUS="2" ' 状态1(普通数据),2 为 导入(附件数据).
						LogInfo= "当前附件导出失败,附件名:"+doc.UniversalID+"_"+CStr(x)+"#"+fileNames(x)+"附件："+fileNames(x) 
						EXPLAIN=doc.UniversalID+"_"+CStr(x)+"#"+fileNames(x)
						Call  SetErrorLog("NOTE_ADM","LOGINFO",doc.UniversalID,IllegalCharacter(CStr(LogInfo)),STATUS,DATACOME,odbcQuery,odbcRS,EXPLAIN)
						doc.IsErrorFileUpload="1"
						Call doc.save(True,False)	
					End If
				Next   
			End If	
		Else
			MsgBox "当前文档:"+doc.UniversalID+"没有附件。"

		End If
		
		
		doc.IsFileUpload="1"
		Call doc.save(True,False)	
		
Goto1:   
		Set doc = docTemp
	Wend    
		odbcRS.Close(DB_CLOSE)	
		odbcConn.Disconnect	
	MsgBox "Leap-附件转存上传数据代理完成。。。"	 
	End If  
	Exit Sub
MsgError:
	MsgBox "Error:"+Error$+" onLine:"+Cstr(Erl)
	GoTo Goto1 
End Sub
Public Function SetMainMd(filePath As String ,filefolder As String)
	Dim Wsh As Variant
	'MsgBox "创建主文件夹开始。。。"
	Set Wsh = CreateObject("WScript.Shell")
	Wsh.run  "C:\Windows\System32\cmd.exe /k md "+filePath+"\"+filefolder,True
	'MsgBox "创建主文件夹完成。。。"
	Exit Function
End Function
Function IllegalCharacter(strSource As String) As String
	On Error GoTo ERR_HANDLER	
	IllegalCharacter=Replace(strSource,"'","''")
	Exit Function
ERR_HANDLER:
	IllegalCharacter=strSource
End Function
Public Function SetMd(filePath As String ,Uid As String)
	Dim Wsh As Variant
	'MsgBox "创建文件夹开始。。。"
	Set Wsh = CreateObject("WScript.Shell")
	Wsh.run  "C:\Windows\System32\cmd.exe /k md "+filePath+"\"+Uid,True
	'MsgBox "创建文件夹完成。。。"
	Exit Function
End Function
%REM
	Function PrintInsureQuotationFull
	Description: Comments for Function
%END REM

Function SetUploadSql(GeCount As String ,OracleUser As String ,Table As String,UNID As String,FILENAME As String ,UPLOADDIRECTORY As String ,UPLOADFILE As String,DataCome As String,odbcQuery As ODBCQuery,odbcRS As ODBCResultSet,PARENTUNID As String ,SELFUNID As String ,REQUESTNO As String ,FILEUNID As String ,SAVEOPTIONS As String ,CREATEDDATE As String )As Boolean 
	On Error GoTo MsgError
	Dim strCommand As String
	odbcQuery.Sql =" SELECT * FROM "+OracleUser+"."+Table+" where UNID='"+UNID+"' and FILENAME='"+FILENAME+"' and UPLOADFILE='"+UPLOADFILE+"'"
	'MsgBox " SELECT * FROM "+OracleUser+"."+Table+" where UNID='"+UNID+"' and FILENAME='"+FILENAME+"' and UPLOADFILE='"+UPLOADFILE+"'"
	Set odbcRS.Query =odbcQuery	
	Call odbcRS.Execute
	
	If (odbcRS.IsResultSetAvailable) Then				
		'重复文档不做插入操作'
		strCommand="update "+OracleUser+"."""+Table+""" Set ""UPLOADDIRECTORY""='"+UPLOADDIRECTORY+"',"&_
		"UNID='"+UNID+"',"&_
		"FILENAME='"+FILENAME+"',"&_
		"PARENTUNID='"+PARENTUNID+"',"&_
		"SELFUNID='"+SELFUNID+"',"&_
		"REQUESTNO='"+REQUESTNO+"',"&_
		"FILEUNID='"+FILEUNID+"',"&_
		"SAVEOPTIONS='"+SAVEOPTIONS+"',"&_
		"UPLOADFILE='"+UPLOADFILE+"',"&_
		"CREATEDDATE='"+CREATEDDATE+"',"&_
		"UPLOADDATE=TO_CHAR(SYSDATE,'YYYYMMDD HH24:MI:SS') ,"&_
		"DataCome='"+DataCome+"'"&_
		" WHERE UNID='"+UNID+"' and FILENAME='"+FILENAME+"' and UPLOADFILE='"+UPLOADFILE+"'"
		'" WHERE PARENTUNID='"+PARENTUNID+"' and FILENAME='"+FILENAME+"'"
		'MsgBox strCommand
		odbcQuery.Sql=strCommand
		Set odbcRS.Query =odbcQuery	
		odbcRS.Execute	
		
		
		
	Else
		strCommand="insert into "+OracleUser+"."""+Table+""" VALUES"&_
		"('"+UNID+"','"+FILENAME+"','"+UPLOADDIRECTORY+"','"&_
		UPLOADFILE+"',TO_CHAR(SYSDATE,'YYYYMMDD HH24:MI:SS'),'"+PARENTUNID+"','"+SELFUNID+"','"+REQUESTNO+"','"+FILEUNID+"','"+SAVEOPTIONS+"','"+DataCome+"','"+CREATEDDATE+"')"
		'MsgBox strCommand
		odbcQuery.Sql=strCommand
		Set odbcRS.Query =odbcQuery	
		odbcRS.Execute	
		

	End If	

	SetUploadSql=True
	
	Exit Function
MsgError:
	SetUploadSql=False
	MsgBox "Error:"+Error$+" onLine:"+Cstr(Erl)+strCommand
End Function
Function SetErrorLog(OracleUser As String ,Table As String,UNID As String,LOGINFO As String ,STATUS As String ,DATACOME As String,odbcQuery As ODBCQuery,odbcRS As ODBCResultSet,EXPLAIN As String )
	On Error GoTo MsgError

	Dim strCommand As String

		odbcQuery.Sql =" SELECT * FROM "+OracleUser+"."+Table+" where LOGINFO='"+LOGINFO+"' and UNID='"+UNID+"' and DATACOME='"+DATACOME+"'"
		'MsgBox " SELECT * FROM "+OracleUser+"."+Table+" where LOGINFO='"+LOGINFO+"' and UNID='"+UNID+"' and DATACOME='"+DATACOME="'"
		Set odbcRS.Query =odbcQuery	
		Call odbcRS.Execute
		
		If (odbcRS.IsResultSetAvailable) Then
			
			'重复文档不做插入操作'
			strCommand="update "+OracleUser+"."""+Table+""" Set ""STATUS""='"+STATUS+"',"&_
			"EXPLAIN='"+EXPLAIN+"',"&_
			"UPLOADDATE=TO_CHAR(SYSDATE,'YYYYMMDD HH24:MI:SS') "&_
			" where LOGINFO='"+LOGINFO+"' and UNID='"+UNID+"' and DATACOME='"+DATACOME+"'"
			'MsgBox strCommand
			odbcQuery.Sql=strCommand
			Set odbcRS.Query =odbcQuery	
			odbcRS.Execute	
		Else
			strCommand="insert into "+OracleUser+"."""+Table+""" VALUES"&_
			"('"+UNID+"','"+LOGINFO+"','"+STATUS+"','"&_
			DATACOME+"',TO_CHAR(SYSDATE,'YYYYMMDD HH24:MI:SS'),'"+EXPLAIN+"')"
			'MsgBox strCommand
			odbcQuery.Sql=strCommand
			Set odbcRS.Query =odbcQuery	
			odbcRS.Execute	
			
		End If
	Exit Function
MsgError:
	MsgBox "Error:"+Error$+" onLine:"+Cstr(Erl)+strCommand
End Function
