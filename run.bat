
rem For prevent accidently unremoved dir
rmdir /S /Q tool\vba\Input_temp

rem Find OS architecture bit
SET osBit=%PROCESSOR_ARCHITECTURE%
SET pdftotextProg="tool\xpdf\bin32\pdftotext.exe"
If %osBit% NEQ x86 (
 SET pdftotextProg="tool\xpdf\bin64\pdftotext.exe"
)


mkdir "tool\vba\Input_temp"


:run_xpdf
rem 判斷目次表數量
	SET /A count=0     
rem 對於單卷檔案，執行 pdfToText 
	for %%g in ("Input\*.pdf") do ( 
		%pdftotextProg% -cfg "tool\xpdf\xpdfrc" -enc Big5 "%%g" "tool\vba\Input_temp\%%~ng.txt"
		SET /A count+=1
		@echo on
	)
	
rem 對於多卷檔案，執行 pdfToText 
	SET /A count2=0     
	for /d %%k in ("Input\*") do ( 
		mkdir tool\vba\Input_temp\%%~nk
		for %%g in ("Input\%%~nk\*.pdf") do ( 
		%pdftotextProg% -cfg "tool\xpdf\xpdfrc" -enc Big5 "%%g" "tool\vba\Input_temp\%%~nk\%%~ng.txt"
		SET /A count2+=1
		@echo on
		)
		SET /A count+=1
	)

If %count% NEQ 16 (
	tool\ErrorPrompt.vbs "ERROR: There should be 16 pdf files!"
	goto exit
)

:runVba
	tool\vba\run.vbs

:exit
	rmdir /S /Q tool\vba\Input_temp

	