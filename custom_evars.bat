rem --------------------------------------------------
rem Below are the standard project initialisation calls, to add other projects
rem or override environment variables, add them to this file.
rem --------------------------------------------------


rem Call projects explicitly
rem --------------------------------------------------
if exist "%projects_dir%AvevaCatalogue\evarsAvevaCatalogue.bat" call "%projects_dir%AvevaCatalogue\evarsAvevaCatalogue.bat"
if exist "%projects_dir%AvevaPlantSample\evarsAvevaPlantSample.bat" call "%projects_dir%AvevaPlantSample\evarsAvevaPlantSample.bat"
if exist "%projects_dir%AvevaMarineSample\evarsAvevaMarineSample.bat" call "%projects_dir%AvevaMarineSample\evarsAvevaMarineSample.bat"
if exist "%projects_dir%mdu\evarsmdu.bat" call "%projects_dir%mdu\evarsmdu.bat" "%projects_dir%"
if exist "%projects_dir%psl\evarspsl.bat" call "%projects_dir%psl\evarspsl.bat" "%projects_dir%"
if exist "%projects_dir%lis\evarslis.bat" call "%projects_dir%lis\evarslis.bat" "%projects_dir%"
if exist "%projects_dir%cpl\evarscpl.bat" call "%projects_dir%cpl\evarscpl.bat" "%projects_dir%"
REM if exist "C:\Users\Public\Documents\AVEVA\Projects\RAB\evarsRAB.bat" call "C:\Users\Public\Documents\AVEVA\Projects\RAB\evarsRAB.bat" "%projects_dir%"
if exist "D:\AVEVA_DABACON_PROJECT\RAB\evarsRAB.bat" call "D:\AVEVA_DABACON_PROJECT\RAB\evarsRAB.bat" "%projects_dir%"

rem Additional user variables can be added below
rem --------------------------------------------------
rem Add any additional projects into the file projects.bat.
REM CALL "C:\Users\Public\Documents\AVEVA\Projects\RAB\evarsRAB.bat"


if exist "\\als.local\NOC\Data\Appli\DigitalAsset\EngData\E3D\E3D3.1\custom_evars.bat" call "\\als.local\NOC\Data\Appli\DigitalAsset\EngData\E3D\E3D3.1\custom_evars.bat"
