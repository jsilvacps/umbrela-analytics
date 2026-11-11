; ═══════════════════════════════════════════════════════════════════
;  Umbrela Analytics — Inno Setup Script
;  Gera: UmbrelaAnalytics_Setup_v7.0.exe
; ═══════════════════════════════════════════════════════════════════

#define AppName    "Umbrela Analytics"
#define AppVersion "7.0"
#define AppPublisher "DMS Empresarial"
#define AppURL     "https://dmsempresarial.com.br"
#define AppExeName "UmbrelaAnalytics.exe"
#define SrcDir     "C:\Temp\UmbrelaAnalyticsBuild\dist"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} v{#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=C:\Temp\UmbrelaAnalyticsBuild\setup
OutputBaseFilename=UmbrelaAnalytics_Setup_v{#AppVersion}
SetupIconFile=
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
WizardResizable=no
DisableProgramGroupPage=no
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} v{#AppVersion}
VersionInfoVersion=7.0.0.0
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Installer
VersionInfoProductName={#AppName}
VersionInfoProductVersion={#AppVersion}
ChangesAssociations=no
DisableWelcomePage=no
WizardImageFile=compiler:WizModernImage.bmp
WizardSmallImageFile=compiler:WizModernSmallImage.bmp

[Languages]
Name: "brazilianportuguese"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "Criar ícone na Área de Trabalho"; GroupDescription: "Atalhos:"; Flags: unchecked
Name: "quicklaunchicon"; Description: "Adicionar à barra de tarefas (fixar)"; GroupDescription: "Atalhos:"; Flags: unchecked

[Files]
Source: "{#SrcDir}\{#AppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\Desinstalar {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon
Name: "{userstartup}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Iniciar {#AppName} agora"; Flags: nowait postinstall skipifsilent

[Messages]
WelcomeLabel1=Bem-vindo ao instalador do {#AppName}
WelcomeLabel2=Este assistente irá instalar o {#AppName} v{#AppVersion} no seu computador.%n%nFerramenta de análise de produtividade Makrolock para geração de relatórios e classificação de atividades.%n%nFeche outras aplicações antes de continuar.
FinishedHeadingLabel=Instalação do {#AppName} concluída!
FinishedLabel=O {#AppName} foi instalado com sucesso.%n%nClique em Concluir para fechar o assistente.
