@echo off
:: Excel MCP Server - Windows Batch Script
:: Replaces Makefile functionality for Windows environments
:: Supports Python -m pytest as required

setlocal enabledelayedexpansion

:: Configuration
set PROJECT_NAME=Excel MCP Server
set PYTHON_CMD=python
set PYTEST_CMD=python -m pytest
set SRC_DIR=src
set TESTS_DIR=tests
set COVERAGE_DIR=htmlcov
set CLEAN_PATTERNS=*.pyc __pycache__ .pytest_cache %COVERAGE_DIR% .coverage coverage.xml

:: Colors for output (Windows ANSI support)
set "INFO=[36m"
set "SUCCESS=[32m"
set "WARNING=[33m"
set "ERROR=[31m"
set "RESET=[0m"

:: Initialize environment
call :init_environment

:: Main command handler
if "%1"=="" goto help
if "%1"=="test-unit" goto test_unit
if "%1"=="test-integration" goto test_integration
if "%1"=="test-performance" goto test_performance
if "%1"=="test-coverage" goto test_coverage
if "%1"=="test-all" goto test_all
if "%1"=="format" goto format_code
if "%1"=="lint" goto lint_code
if "%1"=="clean" goto clean
if "%1"=="help" goto help
if "%1"=="--help" goto help
if "%1"=="/?" goto help

:: Default to help if unknown command
echo %ERROR%Unknown command: %1%RESET%
echo.
goto help

:init_environment
    :: Set Python path
    if defined PYTHONPATH (
        set "PYTHONPATH=%PYTHONPATH%;%CD%\%SRC_DIR%"
    ) else (
        set "PYTHONPATH=%CD%\%SRC_DIR%"
    )

    :: Check if Python is available
    %PYTHON_CMD% --version >nul 2>&1
    if errorlevel 1 (
        echo %ERROR%Error: Python is not installed or not in PATH%RESET%
        exit /b 1
    )

    :: Check if we're in the right directory
    if not exist "%SRC_DIR%" (
        echo %ERROR%Error: Source directory '%SRC_DIR%' not found%RESET%
        echo %ERROR%Please run this script from the project root directory%RESET%
        exit /b 1
    )

    :: Check if tests directory exists
    if not exist "%TESTS_DIR%" (
        echo %ERROR%Error: Tests directory '%TESTS_DIR%' not found%RESET%
        exit /b 1
    )

    echo %INFO%Environment initialized successfully%RESET%
    goto :eof

:test_unit
    echo %INFO%Running unit tests...%RESET%
    echo.

    :: Run specific unit test modules
    %PYTEST_CMD% ^
        %TESTS_DIR%/test_api_excel_operations.py ^
        %TESTS_DIR%/test_core.py ^
        %TESTS_DIR%/test_server.py ^
        %TESTS_DIR%/test_utils.py ^
        %TESTS_DIR%/test_error_handler.py ^
        -v ^
        --tb=short ^
        --no-header

    if errorlevel 1 (
        echo %ERROR%Unit tests failed%RESET%
        exit /b 1
    ) else (
        echo %SUCCESS%Unit tests passed successfully%RESET%
    )
    goto :eof

:test_integration
    echo %INFO%Running integration tests...%RESET%
    echo.

    :: Run integration and end-to-end tests
    %PYTEST_CMD% ^
        %TESTS_DIR%/test_features.py ^
        %TESTS_DIR%/test_new_features.py ^
        %TESTS_DIR%/test_new_apis.py ^
        %TESTS_DIR%/test_excel_operations_extended.py ^
        -v ^
        --tb=short ^
        --no-header

    if errorlevel 1 (
        echo %ERROR%Integration tests failed%RESET%
        exit /b 1
    ) else (
        echo %SUCCESS%Integration tests passed successfully%RESET%
    )
    goto :eof

:test_performance
    echo %INFO%Running performance tests...%RESET%
    echo.

    :: Run performance-related tests
    %PYTEST_CMD% ^
        %TESTS_DIR%/ -k "performance or benchmark" ^
        -v ^
        --tb=short ^
        --no-header ^
        --durations=10

    if errorlevel 1 (
        echo %WARNING%No specific performance tests found, running basic timing...%RESET%
        echo.

        :: Run all tests with timing information
        %PYTEST_CMD% ^
            %TESTS_DIR%/test_api_excel_operations.py::TestExcelOperations::test_get_range_success_flow ^
            %TESTS_DIR%/test_core.py::TestExcelReader::test_read_workbook_success ^
            -v ^
            --tb=short ^
            --no-header ^
            --durations=0

        if errorlevel 1 (
            echo %ERROR%Performance tests failed%RESET%
            exit /b 1
        ) else (
            echo %SUCCESS%Performance tests completed%RESET%
        )
    ) else (
        echo %SUCCESS%Performance tests passed successfully%RESET%
    )
    goto :eof

:test_coverage
    echo %INFO%Running coverage tests...%RESET%
    echo.

    :: Check if coverage is installed
    %PYTHON_CMD% -c "import coverage" >nul 2>&1
    if errorlevel 1 (
        echo %WARNING%Coverage package not found. Installing...%RESET%
        %PYTHON_CMD% -m pip install coverage pytest-cov
        if errorlevel 1 (
            echo %ERROR%Failed to install coverage package%RESET%
            exit /b 1
        )
    )

    :: Run tests with coverage
    %PYTEST_CMD% ^
        %TESTS_DIR%/ ^
        --cov=%SRC_DIR% ^
        --cov-report=html:%COVERAGE_DIR% ^
        --cov-report=term-missing ^
        --cov-report=xml:coverage.xml ^
        -v ^
        --tb=short

    if errorlevel 1 (
        echo %ERROR%Coverage tests failed%RESET%
        exit /b 1
    ) else (
        echo %SUCCESS%Coverage tests completed successfully%RESET%
        echo %INFO%Coverage report generated in %COVERAGE_DIR%/ directory%RESET%

        :: Display coverage summary
        if exist "%COVERAGE_DIR%\index.html" (
            echo %INFO%Open %COVERAGE_DIR%\index.html in your browser to view detailed coverage%RESET%
        )
    )
    goto :eof

:test_all
    echo %INFO%Running all tests...%RESET%
    echo.

    :: First run unit tests
    call :test_unit
    if errorlevel 1 exit /b 1

    echo.
    echo %INFO%=====================================%RESET%

    :: Then integration tests
    call :test_integration
    if errorlevel 1 exit /b 1

    echo.
    echo %INFO%=====================================%RESET%

    :: Finally performance tests
    call :test_performance
    if errorlevel 1 exit /b 1

    echo.
    echo %INFO%=====================================%RESET%

    :: Generate coverage report
    call :test_coverage
    if errorlevel 1 exit /b 1

    echo.
    echo %SUCCESS%All tests completed successfully!%RESET%
    goto :eof

:format_code
    echo %INFO%Formatting code...%RESET%
    echo.

    :: Check if black is installed
    %PYTHON_CMD% -c "import black" >nul 2>&1
    if errorlevel 1 (
        echo %WARNING%Black formatter not found. Installing...%RESET%
        %PYTHON_CMD% -m pip install black
        if errorlevel 1 (
            echo %ERROR%Failed to install black formatter%RESET%
            exit /b 1
        )
    )

    :: Format Python files
    echo %INFO%Formatting Python files in src/ directory...%RESET%
    %PYTHON_CMD% -m black %SRC_DIR%/ --line-length=100

    echo %INFO%Formatting Python files in tests/ directory...%RESET%
    %PYTHON_CMD% -m black %TESTS_DIR%/ --line-length=100

    echo %INFO%Formatting Python files in scripts/ directory...%RESET%
    if exist "scripts\" (
        %PYTHON_CMD% -m black scripts/ --line-length=100
    )

    echo %SUCCESS%Code formatting completed%RESET%
    goto :eof

:lint_code
    echo %INFO%Running code quality checks...%RESET%
    echo.

    :: Check if pylint is installed
    %PYTHON_CMD% -c "import pylint" >nul 2>&1
    if errorlevel 1 (
        echo %WARNING%Pylint not found. Installing...%RESET%
        %PYTHON_CMD% -m pip install pylint
        if errorlevel 1 (
            echo %ERROR%Failed to install pylint%RESET%
            exit /b 1
        )
    )

    :: Lint source code
    echo %INFO%Linting source code...%RESET%
    %PYTHON_CMD% -m pylint %SRC_DIR%/ --disable=R0902,R0903,R0913,C0114,C0115,C0116

    if errorlevel 1 (
        echo %WARNING%Source code linting found issues%RESET%
    ) else (
        echo %SUCCESS%Source code linting passed%RESET%
    )

    echo.

    :: Lint test code with relaxed rules
    echo %INFO%Linting test code...%RESET%
    %PYTHON_CMD% -m pylint %TESTS_DIR%/ --disable=R0902,R0903,R0913,C0114,C0115,C0116,R0801

    if errorlevel 1 (
        echo %WARNING%Test code linting found issues%RESET%
    ) else (
        echo %SUCCESS%Test code linting passed%RESET%
    )

    :: Check for import issues
    echo.
    echo %INFO%Checking for import issues...%RESET%
    %PYTHON_CMD% -c "
import sys
sys.path.insert(0, 'src')
try:
    import server
    import api.excel_operations
    import core.excel_reader
    print('✅ All imports successful')
except ImportError as e:
    print(f'❌ Import error: {e}')
    sys.exit(1)
"

    if errorlevel 1 (
        echo %ERROR%Import checks failed%RESET%
        exit /b 1
    ) else (
        echo %SUCCESS%Import checks passed%RESET%
    )

    echo.
    echo %SUCCESS%Code quality checks completed%RESET%
    goto :eof

:clean
    echo %INFO%Cleaning test artifacts and temporary files...%RESET%
    echo.

    :: Clean Python cache files
    echo %INFO%Removing Python cache files...%RESET%
    for /d /r . %%d in (__pycache__) do (
        if exist "%%d" (
            echo "Removing %%d"
            rmdir /s /q "%%d" 2>nul
        )
    )

    :: Remove .pyc files
    del /s /q *.pyc 2>nul

    :: Remove pytest cache
    if exist ".pytest_cache" (
        echo %INFO%Removing pytest cache...%RESET%
        rmdir /s /q .pytest_cache
    )

    :: Remove coverage files
    if exist "%COVERAGE_DIR%" (
        echo %INFO%Removing coverage report directory...%RESET%
        rmdir /s /q %COVERAGE_DIR%
    )

    if exist ".coverage" (
        echo %INFO%Removing coverage data file...%RESET%
        del .coverage
    )

    if exist "coverage.xml" (
        echo %INFO%Removing coverage XML report...%RESET%
        del coverage.xml
    )

    :: Remove temporary test files
    echo %INFO%Removing temporary test files...%RESET%
    del /q test_*.xlsx 2>nul
    del /q test_*.xlsm 2>nul
    del /q test_*.csv 2>nul
    del /q test_*.json 2>nul

    :: Remove log files
    del /q *.log 2>nul

    :: Remove build artifacts
    if exist "build" (
        echo %INFO%Removing build directory...%RESET%
        rmdir /s /q build
    )

    if exist "dist" (
        echo %INFO%Removing dist directory...%RESET%
        rmdir /s /q dist
    )

    echo %SUCCESS%Cleanup completed%RESET%
    goto :eof

:help
    echo %INFO%%PROJECT_NAME% - Windows Batch Script%RESET%
    echo.
    echo Usage: test.bat [command]
    echo.
    echo Available commands:
    echo.
    echo   %SUCCESS%test-unit%RESET%         Run unit tests (API, Core, Server, Utils modules)
    echo   %SUCCESS%test-integration%RESET% Run integration and end-to-end tests
    echo   %SUCCESS%test-performance%RESET%  Run performance and benchmark tests
    echo   %SUCCESS%test-coverage%RESET%     Run tests with coverage reporting
    echo   %SUCCESS%test-all%RESET%          Run all test suites
    echo   %SUCCESS%format%RESET%            Format code using Black
    echo   %SUCCESS%lint%RESET%              Run code quality checks using Pylint
    echo   %SUCCESS%clean%RESET%             Clean test artifacts and temporary files
    echo   %SUCCESS%help%RESET%              Show this help message
    echo.
    echo Examples:
    echo.
    echo   test.bat test-unit          # Run only unit tests
    echo   test.bat test-coverage      # Run tests with coverage report
    echo   test.bat test-all           # Run all test suites
    echo   test.bat format             # Format all Python code
    echo   test.bat lint               # Check code quality
    echo   test.bat clean              # Clean up temporary files
    echo.
    echo Notes:
    echo   - Uses 'python -m pytest' instead of direct pytest command
    echo   - Automatically installs missing dependencies
    echo   - Supports Windows environment variables and paths
    echo   - Coverage report is generated in htmlcov/ directory
    echo   - PYTHONPATH is automatically set to include src/ directory
    echo.
    echo Environment:
    echo   - Python: %PYTHON_CMD%
    echo   - Tests directory: %TESTS_DIR%
    echo   - Source directory: %SRC_DIR%
    echo   - PYTHONPATH: %PYTHONPATH%
    echo.
    goto :eof

endlocal