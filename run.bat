rmdir /S /Q tool\vba\Input_temp
rd /s /Q tool\vba\Input_temp

REM Find OS architecture bit
SET osBit=%PROCESSOR_ARCHITECTURE%
SET pdftotextProg="tool\xpdf\bin32\pdftotext.exe"
If %osBit% NEQ x86 (
 SET pdftotextProg="tool\xpdf\bin64\pdftotext.exe"
)


mkdir "tool\vba\Input_temp"


:run_xpdf
	REM 判斷目次表數量
	SET /A count=0     
	REM 執行 pdf to text
	for /R %%f in ("Input\*") do ( 
		%pdftotextProg% -cfg "tool\xpdf\xpdfrc" -enc Big5 "%%f" "tool\vba\Input_temp\%%~nf.txt"
		SET /A count+=1
	)

If %count% NEQ 16 (
	tool\ErrorPrompt.vbs "ERROR: There should be 16 pdf files!"
	goto exit
)

:runVba
	tool\vba\run.vbs

:exit
REM For prevent accidently unremoved dir
rem	rmdir /S /Q tool\vba\Input_temp
rem	rd /s /Q tool\vba\Input_temp   	

	