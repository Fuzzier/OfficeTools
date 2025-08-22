CHCP 65001
REM 1. 安装证书.
REM    进入Cert目录, 运行`install.cmd`
REM 2. 安装`Word`.
REM    版本 >= 2016.
REM 3. 安装`MathType`.
REM    版本 >= 6.
REM 4. 安装加载项.
REM    拷贝`自定义工具.dotm`到`%APPDATA%\Microsoft\Word\STARTUP`.

REM 创建目录.
CMD /E:ON /C MKDIR "%APPDATA%\Microsoft\Word\STARTUP"
REM 拷贝加载项.
COPY /Y *.dotm "%APPDATA%\Microsoft\Word\STARTUP"

PAUSE
EXIT /B
