<% class directory
	public AbsoluteBasePath
	public HttpBasePath
	public RelativePath

	
	Private Sub Class_Initialize()
		if cInteger(request("FILEMAN_AZ_ID"))>0 then
			Session("FILEMAN_AZ_ID") = cInteger(request("FILEMAN_AZ_ID"))
		end if
		
		AbsoluteBasePath = Application("IMAGE_PATH") & Session("FILEMAN_AZ_ID") & "\"
		HttpBasePath = Application("IMAGE_SERVER") & "/" & Session("FILEMAN_AZ_ID") 
	end sub	
	
	
	Public Property Get RelativeDirPath()
		RelativeDirPath = RelativePath
	end Property
	
	
	Public Property Let RelativeDirPath(value)
		value = replace(value, "/", "\")
		if instr(1, value, "\", vbTextCompare)=1 OR (not instr(1, value, "\", vbTextCompare)) then
			RelativePath = value
		else
			RelativePath = "\" & value
		end if
	end Property
	
	
	'restituisce la pagina corrente
	Public Property Get DIRPath()
		DIRPath = AbsoluteBasePath & "\" & RelativePath & "\"
		DIRPath = replace(DIRPath, "/", "\")
		DIRPath = IIF(Left(DIRPath, 2) = "\\", "\", "") & replace(DIRPath, "\\", "\")
	End Property
	
	
	Public Property Get URLPath(filename)
	   	URLPath = HttpBasePath & Replace(RelativePath,"\","/") & "/" & filename
	End Property
	
	
	'ritorna il percorso relativo a partire dalla directory di base (images, objects, testi)
	Public Property Get RelativeURL(filename)
		RelativeURL = RelativePath & IIF(filename<>"", "\", "") & filename
		RelativeURL = right(RelativeURL, 1 + len(RelativeURL) - instr(2, RelativeURL, "\", vbTextCompare))
		RelativeURL = replace(RelativeURL, "\", "/")
		if instrrev(RelativeURL, "/", vbTrue, vbTextCompare)=1 then
			'file nella directory principale: toglie la barra iniziale
			RelativeURL = right(RelativeUrl, len(RelativeURL)-1)
		end if
		relativeURL = replace(relativeURL, "//", "/")
	End Property
	
	
	'restituisce la pagina corrente
	Public Property Get FILEPath(filename)
	    	FILEPath = AbsoluteBasePath & RelativePath & "\" & filename
	End Property
	
	
	'ritorna la directory padre della directory richiesta
	Public Property Get ParentPath()
		if instrrev(RelativePath, "\", vbTrue, vbTextCompare)>1 then
			'livello interno
			ParentPath = left(RelativePath, (instrrev(RelativePath, "\", vbTrue, vbTextCompare)-1))
		else
			ParentPath = RelativePath
		end if
	end Property
	
	
	Public function FolderName(path)
		path = replace(path, "/", "\")
		path = replace(path, "\\", "\")
		FolderName = right(path, len(path) - (instrrev(path, "\", vbTrue, vbTextCompare)))
	end function
	
	
	Public Function CreateFolder(name)
	   Dim fso, f,path
	   path = DIRPath() & name
	   Set fso = CreateObject("Scripting.FileSystemObject")
	   Set f = fso.CreateFolder(path)
	   CreateFolder = f.Path
	End Function
	
	
	public function IsEmptyFolder(fldr)
		Dim fso, msg, path, f
		path = DIRPath() & fldr
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		If (fso.FolderExists(path)) Then
			Set f = fso.GetFolder(path)
			if f.Subfolders.Count > 0 OR f.Files.Count > 0 then
				IsEmptyFolder = false
			else
				IsEmptyFolder = true
			end if
		else
			IsEmptyFolder = true
		End If
	end function
	
	
	public function ExistFolder(fldr)
		Dim fso, msg,path
		path = DIRPath() & fldr
		Set fso = CreateObject("Scripting.FileSystemObject")
	   
		ExistFolder = fso.FolderExists(path)
	end function
	
	
	Public Function RemoveFolder(name)
	   Dim fso, path
	   path = DIRPath() & name
	   Set fso = CreateObject("Scripting.FileSystemObject")
	   
	   RemoveFolder = FolderRemove(fso, path, true)
	   
	   set fso = nothing
	End Function

end class

%>

