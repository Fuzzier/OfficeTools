@ECHO OFF

CALL :AddRoot
CALL :AddTrustedPublisher

PAUSE
EXIT /B

:AddRoot
CertUtil -verifystore Root "Fuzzier's Root Certificate Authority" 2>NUL 1>NUL
IF %ERRORLEVEL% NEQ 0 (
  CertUtil -user -addstore Root Root.cer
)
EXIT /B

:AddTrustedPublisher
CertUtil -verifystore TrustedPublisher "Fuzzier's Certificate for VBA" 2>NUL 1>NUL
IF %ERRORLEVEL% NEQ 0 (
  CertUtil -user -addstore TrustedPublisher VBA.cer NoRoot
)
EXIT /B
