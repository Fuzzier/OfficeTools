CHCP 65001
REM 1. 安装证书.
REM    进入Cert目录, 运行"install.cmd".
REM 2. 安装`Visio`.
REM    版本 >= 2016.
REM 3. 安装模具.
REM    将模具拷贝到"%USERPROFILE%\Documents\我的形状".
REM 4. 打开`Visio`.
REM    菜单中选择"更多形状", 选择"我的形状".
REM    勾选模具.

SET MyShapesPath=
CALL :GetVisioMyShapePath 16.0
CALL :CopyShapes

PAUSE
EXIT /B

:GetVisioMyShapePath
SET VER=%1
SET KEY=HKCU\SOFTWARE\Microsoft\Office\%VER%\Visio\Application
FOR /F "tokens=2,*" %%a IN ('REG QUERY %KEY% /v MyShapesPath 2^>NUL') DO (
  SET MyShapesPath=%%b
)
IF "%MyShapesPath%" EQU "" (
  SET MyShapesPath=%USERPROFILE%\Documents\我的形状
)
EXIT /B

:CopyShapes
IF DEFINED MyShapesPath (
  REM 创建目录.
  CMD /E:ON /C MKDIR "%MyShapesPath%"
  REM 拷贝模具.
  COPY /Y *.vssm "%MyShapesPath%"
  COPY /Y *.vssx "%MyShapesPath%"
)
EXIT /B