@echo off

set gitRepositoryUrl=https://example.com/dummy.git
set temporaryFolder=%TEMP%\%RANDOM%

git clone -q %gitRepositoryUrl% "%temporaryFolder%"

rem Don't add /s option because it will delete files even under the .git folder
del /f /q "%temporaryFolder%" > nul 2>&1

copy /b /y "%~n0.xlsm" "%temporaryFolder%\%~n0.xlsm"

cscript "%~n0.vbs" "%temporaryFolder%\%~n0.xlsm" "%temporaryFolder%\vba" //nologo

cd %temporaryFolder%
git add .
git commit -aq -m "By %USERNAME% at %TIME% on %DATE%"
git push -q

cd ..
rd /q /s %temporaryFolder% > nul 2>&1

start %gitRepositoryUrl%
