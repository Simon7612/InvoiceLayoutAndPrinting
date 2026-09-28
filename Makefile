SHELL := cmd
APP=发票排版与打印
# 版本号唯一来源：pyproject.toml 的 version；Nuitka 接受最多 4 段数字，三位可直接使用
PYPROJECT_VERSION := $(word 3,$(subst ",,$(shell findstr /B /R ^version pyproject.toml)))
ifeq ($(PYPROJECT_VERSION),)
PYPROJECT_VERSION := 0.0.0
endif
APP_VERSION = $(PYPROJECT_VERSION)
ICON=icon.ico
OUTDIR=dist
ENTRY=main.py
VENV=.venv
PY=$(VENV)\\Scripts\\python.exe

.PHONY: venv install package run clean version

version:
	@echo $(APP_VERSION)

venv:
	if exist "$(PY)" (echo venv ready) else ( where py >nul 2>nul && (py -m venv "$(VENV)") || (python -m venv "$(VENV)") )

install: venv
	where uv >nul 2>nul && (uv sync) || ( "$(PY)" -m pip install -U pip && "$(PY)" -m pip install nuitka pyqt6 pypdf )

package: install
	powershell -NoProfile -ExecutionPolicy Bypass -Command "& \"$(PY)\" -m nuitka --onefile --mingw64 --windows-console-mode=disable --enable-plugin=pyqt6 --windows-icon-from-ico=\"$(ICON)\" --include-data-files=\"$(ICON)=$(ICON)\" --windows-company-name=\"Simon Chan\" --windows-product-name=\"InvoiceLayoutAndPrinting\" --windows-file-version=\"$(APP_VERSION)\" --windows-product-version=\"$(APP_VERSION)\" --output-filename=\"$(APP)\" --output-dir=\"$(OUTDIR)\" \"$(ENTRY)\""

run: install
	powershell -NoProfile -ExecutionPolicy Bypass -Command "& \"$(PY)\" \"$(ENTRY)\""

clean:
	if exist build rd /s /q build
	if exist "$(OUTDIR)" rd /s /q "$(OUTDIR)"
	if exist *.build del /q *.build
	if exist *.dist del /q *.dist