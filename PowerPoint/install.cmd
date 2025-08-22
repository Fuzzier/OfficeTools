CHCP 65001
REM 1. 安装证书.
REM    进入Cert目录, 运行"install.cmd".
REM 2. 安装`PowerPoint`.
REM    版本 >= 2016.
REM 3. 安装加载项.
REM    拷贝"自定义工具.ppam"到"%APPDATA%\Microsoft\Addins".
REM 4. 打开`PowerPoint`.
REM    菜单中选择"文件", 点击"选项".
REM    点击"加载项".
REM    在"管理"列表中选择"PowerPoint "加载项", 点击[转到...]按钮.
REM    在"加载项对话框"中, 点击[添加...]按钮.
REM    选择"自定义工具.ppam", 点击[打开]按钮.
REM    关闭"加载项对话框".

REM 创建目录.
CMD /E:ON /C MKDIR "%APPDATA%\Microsoft\Addins"
REM 拷贝加载项.
COPY /Y "自定义工具.ppam" "%APPDATA%\Microsoft\Addins"

CALL :RegAddin 16.0

PAUSE
EXIT /B

:RegAddin
SET VER=%1
SET KEY=HKCU\SOFTWARE\Microsoft\Office\%VER%\PowerPoint
REG QUERY %KEY% 2>&1 >NUL
IF %ERRORLEVEL% EQU 0 (
  REG ADD "%KEY%\Addins\自定义工具" /v Path /t REG_SZ /d "自定义工具.ppam" /f
  REG ADD "%KEY%\Addins\自定义工具" /v AutoLoad /t REG_DWORD /d 1 /f
)

EXIT /B
