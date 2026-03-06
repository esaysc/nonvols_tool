@echo off
SETLOCAL EnableDelayedExpansion

:: ========== 配置项（无需手动改） ==========
set VENV_DIR=venv
:: ===========================

echo ==========================================
echo 项目环境初始化脚本
echo ==========================================
echo [1/5] 自动检测 Conda 安装路径...

:: 1. 自动查找 Conda 路径（精准解析）
set "CONDA_EXE_PATH="
set "CONDA_ROOT_PATH="

:: 第一步：优先找 conda.exe（避免 bat 文件）
for /f "tokens=*" %%i in ('where conda.exe 2^>nul') do (
    set "CONDA_EXE_PATH=%%i"
    set "CONDA_ROOT_PATH=!CONDA_EXE_PATH:\Scripts\conda.exe=!"
    goto FOUND_CONDA
)

:: 第二步：兜底找 conda.bat 并提取根路径
for /f "tokens=*" %%i in ('where conda.bat 2^>nul') do (
    set "CONDA_BAT_PATH=%%i"
    set "CONDA_ROOT_PATH=!CONDA_BAT_PATH:\condabin\conda.bat=!"
    goto FOUND_CONDA
)

:: 第三步：常见路径兜底
if exist "D:\ProgramData\miniconda3\Scripts\conda.exe" (
    set "CONDA_ROOT_PATH=D:\ProgramData\miniconda3"
    goto FOUND_CONDA
)
if exist "%USERPROFILE%\AppData\Local\miniconda3\Scripts\conda.exe" (
    set "CONDA_ROOT_PATH=%USERPROFILE%\AppData\Local\miniconda3"
    goto FOUND_CONDA
)

echo [ERROR] 错误: 未检测到 Conda 安装！
echo 请先安装 Miniconda/Anaconda，安装时勾选 "Add to PATH" 选项。
pause
exit /b 1

:FOUND_CONDA
:: 清理路径末尾反斜杠
if "!CONDA_ROOT_PATH:~-1!" == "\" set "CONDA_ROOT_PATH=!CONDA_ROOT_PATH:~0,-1!"
echo [OK] 自动找到 Conda 根路径: !CONDA_ROOT_PATH!

:: 2. 获取并解析 Conda 环境列表（纯批处理兼容逻辑）
echo.
echo [2/5] 检测到以下 Conda 环境：
echo ------------------------------------------
set "ENV_LIST_FILE=%temp%\conda_envs.txt"
:: 直接执行 conda env list，不添加重定向符号
"!CONDA_ROOT_PATH!\Scripts\conda.exe" env list > "%ENV_LIST_FILE%"

:: 解析环境列表（兼容带 * + 标记的格式）
set "ENV_COUNT=0"
set "ENV_LIST="
for /f "skip=3 tokens=1 delims= " %%i in (%ENV_LIST_FILE%) do (
    :: 跳过空行、注释、base、*、+ 等无效内容
    if not "%%i" == "" (
        if not "%%i" == "#" if not "%%i" == "base" if not "%%i" == "*" if not "%%i" == "+" (
            set /a ENV_COUNT+=1
            :: 用临时变量存储环境名，避免嵌套解析错误
            call set "ENV_NAME_%%ENV_COUNT%%=%%i"
            call echo %%ENV_COUNT%%. %%i
            set "ENV_LIST=!ENV_LIST! %%i"
        )
    )
)
del "%ENV_LIST_FILE%"

:: 无可用环境处理
if !ENV_COUNT! equ 0 (
    echo [ERROR] 未检测到除 base 外的 Conda 环境，请先创建环境：
    echo 示例: !CONDA_ROOT_PATH!\Scripts\conda.exe create -n py312 python=3.12
    pause
    exit /b 1
)

:: 3. 选择 Conda 环境（添加退出选项 + 输入验证）
echo ------------------------------------------
echo 0. 退出脚本
echo ------------------------------------------
set "SELECTED_ENV_NUM="
:SELECT_ENV
set /p "SELECTED_ENV_NUM=请选择要使用的 Conda 环境序号 [0-%ENV_COUNT%]（0=退出）: "
:: 验证输入是否为数字且在范围内
set "IS_VALID=0"
:: 支持 0 退出
if "!SELECTED_ENV_NUM!" == "0" (
    echo [INFO] 你选择退出脚本，程序结束。
    pause
    exit /b 0
)
:: 验证 1~ENV_COUNT 范围
for /l %%n in (1,1,!ENV_COUNT!) do (
    if "!SELECTED_ENV_NUM!" == "%%n" set "IS_VALID=1"
)
if !IS_VALID! equ 0 (
    echo [ERROR] 输入无效，请输入 0-%ENV_COUNT% 之间的数字！
    goto SELECT_ENV
)

:: 核心修复：正确获取选中的环境名（用 call 解析嵌套变量）
call set "SELECTED_ENV_NAME=%%ENV_NAME_!SELECTED_ENV_NUM!%%"
echo [OK] 你选择了 Conda 环境: !SELECTED_ENV_NAME!

:: 4. 激活 Conda 环境
echo.
echo [3/5] 激活 Conda 环境: !SELECTED_ENV_NAME!...
call "!CONDA_ROOT_PATH!\Scripts\activate.bat" !SELECTED_ENV_NAME!
if !ERRORLEVEL! equ 0 (
    echo [OK] Conda 环境激活成功。
) else (
    echo [ERROR] 错误: 无法激活 Conda 环境 !SELECTED_ENV_NAME!
    pause
    exit /b 1
)

:: 5. 创建/管理项目虚拟环境（支持跳过/重建）
echo.
echo [4/5] 处理项目虚拟环境...
if exist "%VENV_DIR%" (
    echo [INFO] 检测到虚拟环境 %VENV_DIR% 已存在！
    echo ------------------------------------------
    echo 1. 跳过（使用现有虚拟环境）
    echo 2. 删除并重建虚拟环境
    echo 3. 退出脚本
    echo ------------------------------------------
    set "VENV_OPTION="
    :VENV_CHOICE
    set /p "VENV_OPTION=请选择操作 [1-3]: "
    if "!VENV_OPTION!" == "1" (
        echo [INFO] 跳过虚拟环境创建，使用现有环境。
    ) else if "!VENV_OPTION!" == "2" (
        echo [INFO] 正在删除现有虚拟环境 %VENV_DIR%...
        rmdir /s /q "%VENV_DIR%" >nul 2>&1
        echo [INFO] 正在重新创建虚拟环境 %VENV_DIR%...
        python -m venv "%VENV_DIR%"
        if !ERRORLEVEL! equ 0 (
            echo [OK] 虚拟环境重建成功。
        ) else (
            echo [ERROR] 错误: 虚拟环境重建失败！
            pause
            exit /b 1
        )
    ) else if "!VENV_OPTION!" == "3" (
        echo [INFO] 你选择退出脚本，程序结束。
        pause
        exit /b 0
    ) else (
        echo [ERROR] 输入无效，请输入 1-3 之间的数字！
        goto VENV_CHOICE
    )
) else (
    echo [INFO] 正在创建项目虚拟环境 %VENV_DIR%...
    python -m venv "%VENV_DIR%"
    if !ERRORLEVEL! equ 0 (
        echo [OK] 虚拟环境创建成功。
    ) else (
        echo [ERROR] 错误: 虚拟环境创建失败！
        pause
        exit /b 1
    )
)

:: 6. 安装依赖
echo.
echo [5/5] 安装项目依赖...
call "%VENV_DIR%\Scripts\activate.bat"
python -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple >nul 2>&1
if exist "requirements.txt" (
    echo [INFO] 正在安装 requirements.txt 中的依赖...
    python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    if !ERRORLEVEL! equ 0 (
        echo [OK] 依赖安装完成！
    ) else (
        echo [ERROR] 依赖安装过程中出现错误，请检查 requirements.txt 是否正确。
    )
) else (
    echo [WARN] 警告: 未找到 requirements.txt，跳过依赖安装。
)

echo.
echo ==========================================
echo 环境初始化完成！
echo 激活项目虚拟环境命令: call %VENV_DIR%\Scripts\activate.bat
echo ==========================================
pause

ENDLOCAL