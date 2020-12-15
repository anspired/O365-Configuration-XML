enum OfficeClientEdition
{
  x86 = 32
  x64 = 64
}
enum Channel
{
  Monthly
  Broad
  Targeted
}
enum channelvlk
{
  PerpetualVL2019
}
enum OfficeProduct
{
  O365ProPlusRetail
  O365BusinessRetail
  O365SmallBusPremRetail
}
enum OfficeProductVolume
{
  ProPlus2019Volume
  Standard2019Volume
}
enum OfficeApps
{
  AccessRetail
  Access2019Retail
  Access2019Volume
  ExcelRetail
  Excel2019Retail
  Excel2019Volume
  HomeBusinessRetail
  HomeBusiness2019Retail
  HomeStudentRetail
  HomeStudent2019Retail
  O365HomePremRetail
  OneNoteRetail
  OutlookRetail
  Outlook2019Retail
  Outlook2019Volume
  Personal2019Retail
  PowerPointRetail
  PowerPoint2019Retail
  PowerPoint2019Volume
  ProfessionalRetail
  Professional2019Retail
  ProjectProXVolume
  ProjectPro2019Retail
  ProjectPro2019Volume
  ProjectStdRetail
  ProjectStdXVolume
  ProjectStd2019Retail
  ProjectStd2019Volume
  ProPlus2019Volume
  PublisherRetail
  Publisher2019Retail
  Publisher2019Volume
  Standard2019Volume
  VisioProXVolume
  VisioPro2019Retail
  VisioPro2019Volume
  VisioStdRetail
  VisioStdXVolume
  VisioStd2019Retail
  VisioStd2019Volume
  WordRetail
  Word2019Retail
  Word2019Volume
  LyncEntryRetail
  LyncRetail
  SkypeforBusinessEntryRetail
  SkypeforBusinessRetail
  SkypeforBusiness2019Volume
  SkypeforBusiness2019Retail
}
enum ExcludeApp
{
  Access
  Excel
  Groove
  Lync
  OneDrive
  OneNote
  Outlook
  PowerPoint
  Publisher
  Teams
  Word
  Bing
}
enum Lang
{
  MatchInstalled
  MatchOS
  MatchPreviousMSI
  af_za
  sq_al
  ar_sa
  hy_am
  as_in
  az_Latn_az
  bn_bd
  bn_in
  eu_es
  bs_latn_ba
  ca_es_valencia
  zh_cn
  zh_tw
  bg_bg
  ca_es
  hr_hr
  cs_cz
  da_dk
  nl_nl
  en_us
  et_ee
  fi_fi
  fr_fr
  de_de
  el_gr
  gl_es
  ka_ge
  gu_in
  ha_Latn_ng
  he_il
  hi_in
  hu_hu
  id_id
  is_is
  ig_ng
  ga_ie
  xh_za
  zu_za
  it_it
  ja_jp
  kk_kz
  kn_in
  rw_rw
  sw_ke
  kok_in
  ko_kr
  ky_kg
  lv_lv
  lt_lt
  lb_lu
  mk_mk
 	ms_my
  ml_in
  mt_mt
  mi_nz
  mr_in
  ne_np
  Bokmål
  Nynorsk
  or_in
  ps_af
  fa_ir
  pl_pl
  pt_pt
  pt_br
  pa_in
  ro_ro
  rm_ch
  ru_ru
  gd_gb
  sr_cyrl_rs
  sr_latn_rs
  sr_cyrl_ba
  nso_za
  tn_za
  si_lk
  sk_sk
  sl_si
  es_es
  sv_se
  ta_in
  tt_ru
  te_in
  th_th
  tr_tr
  uk_ua
  ur_pk
  uz_Latn_uz
  vi_vn
  cy_gb
  wo_sn
  yo_ng
}

enum DisplayLevel
{
  Full
  None
}
enum Logging
{
  Off
  Standard
}
enum RemoveMSI
{
  InfoPath
  InfoPathR
  PrjPro
  PrjStd
  SharePointDesigner
  VisPro
  VisStd
}
class AppSettings
{
  [string]$Key
  [string]$Name
  [string]$Value
  [string]$Type
  [string]$App
  [string]$Id
  AppSettings() { }
}
$Appsettings = @(
  @{
    Key   = $Key
    Name  = $Name
    Value = $Value
    Type  = $Type
    App   = $App
    ID    = $ID
  }
)

function New-ODTXML
{
  [CmdletBinding(DefaultParametersetName = '0')]
  param (
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 0)][OfficeClientEdition]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 0)][OfficeClientEdition]
    [Parameter(ParameterSetName = '2', Mandatory = $false, ValueFromPipeline = $true, Position = 1)][OfficeClientEdition]
    [Parameter(ParameterSetName = '3', Mandatory = $false, ValueFromPipeline = $true, Position = 1)][OfficeClientEdition]
    $OfficeClientEdition = [OfficeClientEdition]::x64,
    [Parameter(ParameterSetName = '2', Mandatory = $false, ValueFromPipeline = $true, Position = 0)][switch]
    $DownloadOnly,
    [Parameter(ParameterSetName = '3', Mandatory = $false, ValueFromPipeline = $true, Position = 0)][switch]
    $DownloadOnlyVLK,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 1)][OfficeProduct]
    [Parameter(ParameterSetName = '2', Mandatory = $false, ValueFromPipeline = $true, Position = 2)][OfficeProduct]
    $Product = [OfficeProduct]::O365ProPlusRetail,
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 1)][OfficeProductVolume]
    [Parameter(ParameterSetName = '3', Mandatory = $false, ValueFromPipeline = $true, Position = 2)][OfficeProductVolume]
    $ProductVLK = [OfficeProductVolume]::ProPlus2019Volume,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 2)][Channel]
    [Parameter(ParameterSetName = '2', Mandatory = $false, ValueFromPipeline = $true, Position = 3)][Channel]
    $Channels = [channel]::Monthly,
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 2)][channelvlk]
    [Parameter(ParameterSetName = '3', Mandatory = $false, ValueFromPipeline = $true, Position = 3)][channelvlk]
    $ChannelsVlk = [channelvlk]::PerpetualVL2019,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 3)][string]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 3)][string]
    [Parameter(ParameterSetName = '2', Mandatory = $false, ValueFromPipeline = $true, Position = 4)][string]
    [Parameter(ParameterSetName = '3', Mandatory = $false, ValueFromPipeline = $true, Position = 4)][string]
    $SourcePath,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 4)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 4)]
    [OfficeApps[]]$Apps,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 5)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 5)]
    [ExcludeApp[]]$ExcludeApp = ([ExcludeApp]::Groove, [ExcludeApp]::OneNote, [ExcludeApp]::Bing),
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 6)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 6)]
    [string]$Version,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 7)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 7)]
    [switch]$OfficeMgmtCOM,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 8)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 8)]
    [switch]$AllowCdnFallback,
    [Lang]$Language = "en_us",
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 9)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 9)]
    [Lang]$FallbackLanguage,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 10)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 10)]
    [switch]$PinIconsToTaskbar,
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 11)]
    [switch]$SharedLic,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 12)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 12)]
    [switch]$SCLCacheOverride,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 13)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 13)]
    [string]$SCLCacheOverrideDirectory,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 14)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 14)]
    [switch]$AutoActivate,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 15)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 15)]
    [switch]$ForceAppShutdown,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 16)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 16)]
    [switch]$DeviceBasedLicensing,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 17)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 17)]
    [switch]$UpdatesEnabled,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 18)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 18)]
    [Channel]$UpdateChannel,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 19)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 19)]
    [string]$UpdatePath,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 20)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 20)]
    [string]$UpdateTargetVersion,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 21)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 21)]
    [DisplayLevel]$DisplayLevel,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 22)][switch]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 22)][switch]
    $AcceptEULA = [switch]::0,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 23)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 23)]
    [Logging]$Logging,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 24)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 24)]
    [string]$LogPath,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 25)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 25)]
    [appsettings]$AppSettings,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 26)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 26)]
    [switch]$Remove,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 27)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 27)]
    [switch]$RemoveAll,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 28)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 28)]
    [OfficeProduct]$RemoveOfficeProduct,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 29)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 29)]
    [Lang]$removelanguage,
    [Parameter(ParameterSetName = '0', Mandatory = $false, ValueFromPipeline = $true, Position = 30)]
    [Parameter(ParameterSetName = '1', Mandatory = $false, ValueFromPipeline = $true, Position = 30)]
    [RemoveMSI[]]$RemoveIgnoreApps
  )
  switch ($pscmdlet.ParameterSetName)
  {
    0
    {
      $products = $product
      $channel = $channels
    }
    1
    {
      $products = $ProductVLK
      $channel = $ChannelsVlk
    }
    2
    {
      $products = $product
      $channel = $channels
    }
    3
    {
      $channel = $ChannelsVlk
      $products = $ProductVLK
    }
  }
  $lang = $Language -replace "_", "-"
  $fbLang = $FallbackLanguage -replace "_", "-"
  $removeLang = $removelanguage -replace "_", "-"

  $w = [xml.XmlTextWriter]::Create(($s = [IO.MemoryStream]::new()), [xml.XmlWriterSettings]@{Encoding = [Text.Encoding]::UTF8; OmitXmlDeclaration = $true; Indent = $true; IndentChars = "`t"; WriteEndDocumentOnClose = $false })

  $w.WriteStartDocument()

  $w.WriteStartElement("Configuration")

  $w.WriteStartElement("Add")
  $w.WriteAttributeString("OfficeClientEdition", "$($OfficeClientEdition.value__)")
  $w.WriteAttributeString("Channel", "$($Channel)")
  if ($Omgtcom) { $w.WriteAttributeString("OfficeMgmtCOM", "$($Omgtcom)") }
  if ($SourcePath) { $w.WriteAttributeString("SourcePath", "$($SourcePath)") }
  if ($AllowCdnFallback) { $w.WriteAttributeString("AllowCdnFallback", "$($AllowCdnFallback)") }
  if ($Version) { $w.WriteAttributeString("Version", "$($Version)") }
  #if ($SourcePath) { $w.WriteAttributeString("DownloadPath", "$($SourcePath)") }

  foreach ($p in $products)
  {
    $w.WriteStartElement("Product")
    $w.WriteAttributeString("ID", "$($p)")
    $w.WriteStartElement("Language")
    $w.WriteAttributeString("ID", "$($lang)")
    if ($fbLang) { $w.WriteAttributeString("Fallback", "$($fbLang)") }
    $w.WriteEndElement()
    if (-not ($DownloadOnly -or $DownloadOnlyVLK))
    {
      if ($ExcludeApp)
      {
        foreach ($e in $excludeapp)
        {
          $w.WriteStartElement("ExcludeApp")
          $w.WriteAttributeString("ID", "$($e)")
          $w.WriteEndElement()
        }
      }
    }
    $w.WriteFullEndElement()
  }

  foreach ($p in $apps)
  {
    $w.WriteStartElement("Product")
    $w.WriteAttributeString("ID", "$($p)")
    $w.WriteStartElement("Language")
    $w.WriteAttributeString("ID", "$($lang)")
    if ($fbLang) { $w.WriteAttributeString("Fallback", "$($fbLang)") }
    $w.WriteEndElement()
    $w.WriteFullEndElement()
  }
  $w.WriteFullEndElement()

  if ($DisplayLevel)
  {
    $w.WriteStartElement("Display")
    $w.WriteAttributeString("Level", "$($DisplayLevel)")
    $w.WriteAttributeString("AcceptEULA", "$($AcceptEULA)")
    $w.WriteEndElement()
  }

  if ($logging)
  {
    $w.WriteStartElement("Logging")
    $w.WriteAttributeString("Level", "$($logging)")
    if ($Logging -eq "standard") { $w.WriteAttributeString("Path", "$($LogPath)") }
    $w.WriteEndElement()
  }
  if ($SharedLic.IsPresent)
  {
    $ShrLic = [int]([bool]::Parse($SharedLicOn.IsPresent))
    $w.WriteAttributeString("SharedComputerLicensing", "$($ShrLic)")
    $w.WriteEndElement()
  }
  if ($PinIconsToTaskbar.IsPresent)
  {
    $w.WriteAttributeString("PinIconsToTaskbar", "$($PinIconsToTaskbar.IsPresent.ToString())")
    $w.WriteEndElement()
  }
  if ($SCLCacheOverride.IsPresent)
  {
    $SCL = [int]([bool]::Parse($SCLCacheOverrideOn.IsPresent))
    $w.WriteAttributeString("SCLCacheOverride", "$($SCL)")
    $w.WriteEndElement()
  }
  if ($SCLCacheOverrideDirectory)
  {
    $w.WriteAttributeString("SCLCacheOverrideDirectory", "$($SCLCacheOverrideDirectory)")
    $w.WriteEndElement()
  }
  if ($AutoActivate.IsPresent)
  {
    $AAOff = [int]([bool]::Parse($AutoActivateOff.IsPresent))
    $w.WriteAttributeString("AutoActivate", "$($AAOff)")
    $w.WriteEndElement()
  }
  if ($ForceAppShutdown.IsPresent)
  {
    $w.WriteAttributeString("FORCEAPPSHUTDOWN", "$($ForceAppShutdown.IsPresent.ToString())")
    $w.WriteEndElement()
  }
  if ($OfficeMgmtCOM.IsPresent)
  {
    $Omgtcom = [int]([bool]::Parse($OfficeMgmtCOM.IsPresent))
    $w.WriteAttributeString("OfficeMgmtCOM", "$($Omgtcom)")
    $w.WriteEndElement()
  }


  if ($pscmdlet.ParameterSetName -eq "1")
  {
    if ($DeviceBasedLicensingOn.IsPresent)
    {
      $DBLon = [int]([bool]::Parse($DeviceBasedLicensingOn.IsPresent))
      $w.WriteStartElement("Property")
      $w.WriteAttributeString("Name", "DeviceBasedLicensing")
      $w.WriteAttributeString("Value", "$($DBLon)")
      $w.WriteEndElement()
    }
  }

  if ($remove)
  {
    $w.WriteStartElement("RemoveMSI")
    if ($RemoveIgnoreApps)
    {
      foreach ($e in $RemoveIgnoreApps)
      {
        $w.WriteStartElement("IgnoreProduct")
        $w.WriteAttributeString("ID", "$($e)")
        $w.WriteEndElement()
      }
    }
    $w.WriteEndElement()
  }

  if ($RemoveOfficeProduct)
  {
    $w.WriteStartElement("Remove")
    $w.WriteAttributeString("All", "$($RemoveAll)")
    $w.WriteStartElement("Product")
    $w.WriteAttributeString("ID", "$($RemoveOfficeProduct)")
    $w.WriteStartElement("Language")
    $w.WriteAttributeString("ID", "$($RemoveLang)")
    $w.WriteEndElement()
  }

  if ($UpdatesEnabled)
  {
    $w.WriteStartElement("Updates")
    $w.WriteAttributeString("Enabled", "$($UpdatesEnabled)")
    if ($UpdatePath)
    {
      $w.WriteAttributeString("UpdatePath", "$($UpdatePath)")
      $w.WriteAttributeString("Channel", "$($UpdateChannel)")
    }
    if ($UpdateTargetVersion)
    {
      $w.WriteAttributeString("TargetVersion", "$($UpdateTargetVersion)")
    }
    $w.WriteEndElement()
  }

  if ($AppSettings)
  {
    $w.WriteStartElement("AppSettings")
    $w.WriteStartElement("User")
    $w.WriteAttributeString("Key", "$($AppSettings.Key)")
    $w.WriteAttributeString("Name", "$($AppSettings.Name)")
    $w.WriteAttributeString("Value", "$($AppSettings.Value)")
    $w.WriteAttributeString("Type", "$($AppSettings.Type)")
    $w.WriteAttributeString("App", "$($AppSettings.App)")
    $w.WriteAttributeString("Id", "$($AppSettings.Id)")
    $w.WriteFullEndElement()
  }

  $w.WriteFullEndElement()

  $w.WriteEndDocument()
  $w.Flush();

  $r = [System.IO.StreamReader]::new([IO.MemoryStream]::new($s.ToArray())).readtoend()

  $w.Close(); $w.Dispose()

  return $r
}
#new-odtxml