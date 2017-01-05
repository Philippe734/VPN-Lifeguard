
[CustomMessages]
eng.AppName=VPN Lifeguard
[Setup]
UsePreviousLanguage=no
AppName={cm:AppName} 1.4.12
VersionInfoVersion=1.4.12
AppVerName={cm:AppName}
DefaultDirName={pf}\{cm:AppName}
DefaultGroupName={cm:AppName}
; Définir le nom du programme d'installation
OutputBaseFilename=Setup_VPN_Lifeguard
; Définir le répertoire d'enregistrement du programme d'installaton compilé
OutputDir=D:\Documents and Settings\Philippe\Mes documents\Visual basic\Programmes VB\VPN Lifeguard\VPN Lifeguard 1.4.12\Setup_VPN_Lifeguard
InternalCompressLevel=max
VersionInfoCompany=http://sourceforge.net/projects/vpnlifeguard/
VersionInfoDescription=VPN Lifeguard
VersionInfoCopyright=(C)2010 philippe734 - GNU/GPL
Compression=lzma
VersionInfoProductName=VPN Lifeguard
LicenseFile=D:\Documents and Settings\Philippe\Mes documents\Visual basic\Programmes VB\VPN Lifeguard\VPN Lifeguard 1.4.12\License GPL.txt
RestartIfNeededByRun=true
PrivilegesRequired=admin
AllowNoIcons=true
[Files]
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\asycfilt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile noregerror regserver 32bit
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\olepro32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver noregerror 32bit
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\oleaut32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver noregerror 32bit
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver noregerror 32bit
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\VB6FR.DLL; DestDir: {sys}; Flags: sharedfile uninsneveruninstall regserver restartreplace noregerror 32bit
Source: VpnLifeguard.exe; DestDir: {app}; Flags: ignoreversion
Source: License GPL.txt; DestDir: {app}
Source: Lisez-moi.rtf; DestDir: {app}; Flags: ignoreversion
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\iphlpapi.dll; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\wshom.ocx; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\winhttp.dll; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\ws2_32.dll; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\mscomctl.ocx; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\advpack.dll; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: ..\..\..\..\..\..\..\DLL pour Setup VB\psapi.dll; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: English.lng; DestDir: {app}; Flags: ignoreversion
Source: French.lng; DestDir: {app}; Flags: ignoreversion
Source: LangSetting.ini; DestDir: {app}; Flags: ignoreversion
[Icons]
Name: {group}\VPN Lifeguard; Filename: {app}\VpnLifeguard.exe; WorkingDir: {app}; IconIndex: 0
Name: {userdesktop}\VPN Lifeguard; Filename: {app}\VpnLifeguard.exe; WorkingDir: {app}; IconIndex: 0
[Languages]
Name: eng; MessagesFile: compiler:Default.isl
[Code]
procedure creationconstante;
begin
	ExpandConstant('{cm:AppName}');
end;

// Variables Globales
//var
  //Edit: TEdit;
  //Memo: TMemo;




// Procédure de construction des pages personnelles
procedure CreateTheWizardPages;
// variables locales
var
  Page: TWizardPage;
  Lbl: TLabel;
begin


                //wpInfoAfter           //wpWelcome
  Page := CreateCustomPage(wpInfoAfter , 'Rights require', 'Run as administrator');


  //Edit := TEdit.Create(Page);
  //Edit.Top := ScaleY(8);
  //Edit.Width := Page.SurfaceWidth div 2 - ScaleX(8);
  //Edit.Text := 'Edit.Text';
  //Edit.Parent := Page.Surface;

  //Memo := TMemo.Create(Page);
  //Memo.Top := ScaleY(8);
  //Memo.Width := Page.SurfaceWidth;
  //Memo.Height := ScaleY(209);
  //Memo.Font.Size := 15;
  //Memo.Readonly := True;
  //Memo.ScrollBars := ssVertical;
  //Memo.Text := 'Attention' + #13 + 'You must run the program' + #13 + 'as administrator.' + #13 + 'If not, then show error' + #13 + 'on loading.';
  //Memo.Parent := Page.Surface;

  Lbl := TLabel.Create(Page);
  Lbl.Top := ScaleY(8);
  Lbl.Font.Size := 15;
  Lbl.Font.Name := 'Arial Black';
  Lbl.Caption := 'Attention' + #13 + 'You must run the program' + #13 + 'as administrator.' + #13 + 'If not, then show error' + #13 + 'on loading.';
  Lbl.Width := Page.SurfaceWidth;
  Lbl.AutoSize :=True;
  Lbl.Parent := Page.Surface;
end;



// Initialisation
procedure InitializeWizard();
begin
  CreateTheWizardPages;
end;
