; 德育分自动化系统 · Inno Setup 安装脚本（方案 3.3 节，A 路线）
; 编译前提：
;   1. 安装 Inno Setup 6（https://jrsoftware.org/isinfo.php 或 winget install JRSoftware.InnoSetup）
;   2. 先跑 python tools/make_payload.py 组装出 installer\payload\ 完整目录
; 编译命令（ISCC 在 PATH 里时）：
;   ISCC installer\setup.iss
; 产物：installer\Output\德育分自动化系统-安装程序-1.7.0.exe
;
; 用户数据保护：
;   · 名单模板.xlsx / 更新配置.json 用 onlyifdoesntexist，绝不覆盖已有文件
;   · 卸载时把 名单模板.xlsx、20??年*月\、更新配置.json 暂存到 %TEMP% 再放回，
;     卸载完成后用户数据仍在（Inno 默认会删它自己装过的 名单模板.xlsx，必须绕）

#define MyAppName "德育分自动化系统"
#define MyAppVersion "1.8.2"
#define MyAppExeName "德育分自动化系统.exe"
#define MyAppId "{{8E7C1A52-7F3B-4C21-9A55-1D2C7B0E4F11}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher=moral-edu-automation
DefaultDirName={localappdata}\Programs\{#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
SetupIconFile=payload\软件说明\doro.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
OutputDir=Output
OutputBaseFilename=德育分自动化系统-安装程序-{#MyAppVersion}
LanguageDetectionMethod=none

[Languages]
; 中文语言包（官方 istrans 的非官方翻译，Inno 自身不带），
; 已随仓库放在脚本同目录，缺失时删除本节即可退回英文向导
Name: "chinese"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式(&D)"; GroupDescription: "附加任务:"

[Files]
; 程序文件：升级语义（覆盖）
Source: "payload\德育分自动化系统.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "payload\update.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "payload\AI对接\*"; DestDir: "{app}\AI对接"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "payload\软件说明\*"; DestDir: "{app}\软件说明"; Flags: ignoreversion recursesubdirs createallsubdirs
; 用户数据：只在不存在时写入，绝不覆盖
Source: "payload\名单模板.xlsx"; DestDir: "{app}"; Flags: onlyifdoesntexist
Source: "payload\更新配置.json"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

; 这两个文件是 [Code] 在安装时生成的，Inno 的卸载清单里没有它们，需显式声明清理
[UninstallDelete]
Type: files; Name: "{app}\AI对接\MCP配置.json"
Type: files; Name: "{app}\安装信息.json"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即运行{#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
var
  UserDataStash: string;

procedure CurStepChanged(CurStep: TSetupStep);
var
  S: string;
begin
  if CurStep = ssPostInstall then
  begin
    // MCP 配置：绝对路径随安装目录变化，必须运行时生成
    S := '{"mcpServers":{"deyufen":{"command":"' + ExpandConstant('{app}') +
         '\AI对接\mcp_server.exe","args":[]}}}';
    SaveStringToFile(ExpandConstant('{app}\AI对接\MCP配置.json'), S, False);
    // 版本记录（更新器回滚时读它）
    S := '{"version":"' + '{#MyAppVersion}' + '","install_dir":"' +
         ExpandConstant('{app}') + '"}';
    SaveStringToFile(ExpandConstant('{app}\安装信息.json'), S, False);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  SR: TFindRec;
begin
  if CurUninstallStep = usUninstall then
  begin
    // 卸载删除之前，把用户数据搬到 %TEMP% 暂存
    UserDataStash := ExpandConstant('{tmp}') + '\deyufen-keep';
    CreateDir(UserDataStash);
    if FileExists(ExpandConstant('{app}\名单模板.xlsx')) then
      FileCopy(ExpandConstant('{app}\名单模板.xlsx'),
               UserDataStash + '\名单模板.xlsx', False);
    if FileExists(ExpandConstant('{app}\更新配置.json')) then
      FileCopy(ExpandConstant('{app}\更新配置.json'),
               UserDataStash + '\更新配置.json', False);
    if FindFirst(ExpandConstant('{app}\20*年*月'), SR) then
    begin
      repeat
        if SR.Attributes and FILE_ATTRIBUTE_DIRECTORY <> 0 then
          // 目录整体搬移：同盘 RenameFile 即可
          RenameFile(ExpandConstant('{app}') + '\' + SR.Name,
                     UserDataStash + '\' + SR.Name);
      until not FindNext(SR);
      FindClose(SR);
    end;
  end;
  if CurUninstallStep = usPostUninstall then
  begin
    // 卸载完成后放回原目录
    if DirExists(UserDataStash) then
    begin
      if FindFirst(UserDataStash + '\*.*', SR) then
      begin
        repeat
          RenameFile(UserDataStash + '\' + SR.Name,
                     ExpandConstant('{app}') + '\' + SR.Name);
        until not FindNext(SR);
        FindClose(SR);
      end;
      if FileExists(UserDataStash + '\名单模板.xlsx') then
        FileCopy(UserDataStash + '\名单模板.xlsx',
                 ExpandConstant('{app}\名单模板.xlsx'), False);
      if FileExists(UserDataStash + '\更新配置.json') then
        FileCopy(UserDataStash + '\更新配置.json',
                 ExpandConstant('{app}\更新配置.json'), False);
      DelTree(UserDataStash, True, True, False);
      MsgBox('卸载完成。名单模板与月度数据已保留在原目录。', mbInformation, MB_OK);
    end;
  end;
end;
