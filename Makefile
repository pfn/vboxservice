VBOX_HOME	= C:\Program Files\Oracle\VirtualBox
VBOX_COM_DLL	= "$(VBOX_HOME)\VBoxC.dll"
CC		= csc /nologo
LDFLAGS		= /l:VirtualBox.dll

all: vboxservice.exe

vboxservice.exe: vboxservice.cs VirtualBox.dll
	@$(CC) $(LDFLAGS) vboxservice.cs

VirtualBox.dll: $(VBOX_COM_DLL)
	@tlbimp /nologo $(VBOX_COM_DLL)

clean:
	@del VirtualBox.dll vboxservice.exe *.InstallLog *.InstallState
