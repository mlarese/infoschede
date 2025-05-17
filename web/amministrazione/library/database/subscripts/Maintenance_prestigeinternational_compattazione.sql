 
 USE master 

 ALTER DATABASE Prestigeinternational SET Recovery Simple WITH NO_WAIT
 USE prestigeinternational
 
 DBCC SHRINKFILE (N'24investimenti_Log' , 1)
 
 USE [master] 
 ALTER DATABASE Prestigeinternational SET Recovery Full WITH NO_WAIT
 USE prestigeinternational