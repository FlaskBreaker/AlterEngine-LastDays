# Wine environment
WINE_ENV_PATH := $(HOME)/.wine-ae

# Libs
LIBS_PATH=../libs
LIB1=${LIBS_PATH}/1.exe
LIB2=${LIBS_PATH}/2.exe
LIB3=${LIBS_PATH}/Eclipse\ Library\ Installer\ x86.exe

DLL_TO_CHECK=${WINE_ENV_PATH}/drive_c/windows/system32/MSVBVM60.DLL

# Apps
APP_CLIENT=../cliente/Cliente.exe
APP_SERVER=../servidor/Servidor.exe
APP_SCRIPT_EDITOR=../servidor/Editor.exe

all: run

$(WINE_ENV_PATH):
	@echo "Creating a new Wine environment at $(WINE_ENV_PATH)..."
	WINEARCH=win32 WINEPREFIX=$(WINE_ENV_PATH) wineboot
	@echo "New Wine environment created at $(WINE_ENV_PATH)."

install-libs: $(WINE_ENV_PATH)
	@if [ ! -f "$(DLL_TO_CHECK)" ]; then \
		echo "Installing dependencies at $(WINE_ENV_PATH)..."; \
		WINEPREFIX=$(WINE_ENV_PATH) wine $(LIB1); \
		WINEPREFIX=$(WINE_ENV_PATH) wine $(LIB2); \
		WINEPREFIX=$(WINE_ENV_PATH) wine $(LIB3) || true; \
		echo "Dependencies successfully installed at $(WINE_ENV_PATH)."; \
	else \
		echo "Dependencies already installed at $(WINE_ENV_PATH). Skipping installation."; \
	fi

run: run-server

run-server: install-libs
	@echo "Runing Alter Engine server..."
	WINEPREFIX=$(WINE_ENV_PATH) wine $(APP_SERVER)

run-client: install-libs
	@echo "Runing Alter Engine client..."
	WINEPREFIX=$(WINE_ENV_PATH) wine $(APP_CLIENT)

run-editor: install-libs
	@echo "Runing Alter Engine script editor..."
	WINEPREFIX=$(WINE_ENV_PATH) wine $(APP_SCRIPT_EDITOR)

clean:
	@echo "Removing Wine environment at $(WINE_ENV_PATH)..."
	rm -rf $(WINE_ENV_PATH)
	@echo "Wine environment removed."
