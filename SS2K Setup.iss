;
; Script generated by the ASCOM Driver Installer Script Generator 6.0.0.0
; Generated by Laurie on 07/07/2011 (UTC)
;
[Setup]
AppName=ASCOM SS2K Telescope Driver
AppVerName=ASCOM SS2K Telescope Driver 6.1.8d_Move
AppVersion=6.1.8d_Move
AppPublisher=Laurie <laurie@lyates.com>
AppPublisherURL=mailto:laurie@lyates.com
AppSupportURL=http://tech.groups.yahoo.com/group/ASCOM-Talk/
AppUpdatesURL=http://ascom-standards.org/
VersionInfoVersion=1.0.0
MinVersion=0,5.0.2195sp4
DefaultDirName="{cf}\ASCOM\Telescope"
DisableDirPage=yes
DisableProgramGroupPage=yes
OutputDir="."
OutputBaseFilename="SS2K Setup"
Compression=lzma
SolidCompression=yes
; Put there by Platform if Driver Installer Support selected
WizardImageFile="C:\Program Files (x86)\ASCOM\Platform 6 Developer Components\Installer Generator\Resources\WizardImage.bmp"
LicenseFile="C:\Program Files (x86)\ASCOM\Platform 6 Developer Components\Installer Generator\Resources\CreativeCommons.txt"
; {cf}\ASCOM\Uninstall\Telescope folder created by Platform, always
UninstallFilesDir="{cf}\ASCOM\Uninstall\Telescope\SS2K"

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Dirs]
Name: "{cf}\ASCOM\Uninstall\Telescope\SS2K"
; TODO: Add subfolders below {app} as needed (e.g. Name: "{app}\MyFolder")

[Files]
; regserver flag only if native COM, not .NET
Source: "G:\T_Astro\Astro_Software\Ascom\V6-SS2K\6.1.8d_Move\SS2K Driver.dll"; DestDir: "{app}" ;AfterInstall: RegASCOM(); Flags: regserver
; Require a read-me HTML to appear after installation, maybe driver's Help doc
Source: "G:\T_Astro\Astro_Software\Ascom\V6-SS2K\6.1.8d_Move\SS2K Driver.htm"; DestDir: "{app}"; Flags: isreadme
; TODO: Add other files needed by your driver here (add subfolders above)
Source: "G:\T_Astro\Astro_Software\Ascom\V6-SS2K\6.1.8d_Move\Tabctl32.ocx"; DestDir: "{sys}"; Flags: regserver









[CODE]
//
// Before the installer UI appears, verify that the (prerequisite)
// ASCOM Platform 5.5 or greater is installed, including both Helper
// components. Utility is required for all types (COM and .NET)!
//
function InitializeSetup(): Boolean;
var
   U : Variant;
   H : Variant;
begin
   Result := FALSE;  // Assume failure
   // check that the DriverHelper and Utilities objects exist, report errors if they don't
   try
      H := CreateOLEObject('DriverHelper.Util');
   except
      MsgBox('The ASCOM DriverHelper object has failed to load, this indicates a serious problem with the ASCOM installation', mbInformation, MB_OK);
   end;
   try
      U := CreateOLEObject('ASCOM.Utilities.Util');
   except
      MsgBox('The ASCOM Utilities object has failed to load, this indicates that the ASCOM Platform has not been installed correctly', mbInformation, MB_OK);
   end;
   try
      if (U.IsMinimumRequiredVersion(5,5)) then	// this will work in all locales
         Result := TRUE;
   except
   end;
   if(not Result) then
      MsgBox('The ASCOM Platform 5.5 or greater is required for this driver.', mbInformation, MB_OK);
end;

//
// Register and unregister the driver with the Chooser
// We already know that the Helper is available
//
procedure RegASCOM();
var
   P: Variant;
begin
   P := CreateOleObject('ASCOM.Utilities.Profile');
   P.DeviceType := 'Telescope';
   P.Register('SS2K.Telescope', 'SkySensor2000-PC-V6');
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
   P: Variant;
begin
   if CurUninstallStep = usUninstall then
   begin
     P := CreateOleObject('ASCOM.Utilities.Profile');
     P.DeviceType := 'Telescope';
     P.Unregister('SS2K.Telescope');
  end;
end;
