# Berry IT Network Utility - BERRY IT Network Diagnostics
# Requires Windows PowerShell 5.1, WPF and .NET Framework 4.7.2+.
# Copyright (c) 2026 Berry IT. Distributed without warranty.
# This is a Windows-only, read-only diagnostic utility.
if ($PSVersionTable.PSEdition -ne 'Desktop') { Write-Error 'Start Berry IT Network Utility in Windows PowerShell 5.1 (powershell.exe).'; return }
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Error 'Berry IT Network Utility requires STA. Launch with powershell.exe -STA -File .\NetworkUtility.ps1'; return
}
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Berry IT Network Utility | BERRY IT" Width="1190" Height="790" MinWidth="980" MinHeight="660" WindowStartupLocation="CenterScreen" Background="#0B0B0E" Foreground="#F4F4F4" FontFamily="Segoe UI" FontSize="13">
<Window.Resources>
 <SolidColorBrush x:Key="Panel" Color="#1B1B20"/><SolidColorBrush x:Key="Line" Color="#303036"/><SolidColorBrush x:Key="Muted" Color="#ACAAB2"/>
 <Style TargetType="Button"><Setter Property="Background" Value="#29292F"/><Setter Property="Foreground" Value="White"/><Setter Property="BorderBrush" Value="#44444B"/><Setter Property="Padding" Value="13,8"/><Setter Property="Margin" Value="0,0,8,0"/><Setter Property="Cursor" Value="Hand"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border></ControlTemplate></Setter.Value></Setter></Style>
 <Style TargetType="TextBox"><Setter Property="Background" Value="#111116"/><Setter Property="Foreground" Value="White"/><Setter Property="CaretBrush" Value="#FF3C3C"/><Setter Property="BorderBrush" Value="#474750"/><Setter Property="Padding" Value="9,7"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style>
 <Style TargetType="ComboBox"><Setter Property="Background" Value="#28282E"/><Setter Property="Foreground" Value="#111116"/><Setter Property="Padding" Value="6"/><Setter Property="Margin" Value="0,2,8,8"/></Style>
 <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/></Style>
 <Style TargetType="DataGrid"><Setter Property="Background" Value="#1B1B20"/><Setter Property="Foreground" Value="#F4F4F4"/><Setter Property="RowBackground" Value="#1B1B20"/><Setter Property="AlternatingRowBackground" Value="#25252B"/><Setter Property="BorderBrush" Value="#303036"/><Setter Property="HorizontalGridLinesBrush" Value="#303036"/><Setter Property="VerticalGridLinesBrush" Value="#303036"/><Setter Property="HeadersVisibility" Value="Column"/><Setter Property="CanUserAddRows" Value="False"/><Setter Property="IsReadOnly" Value="True"/><Setter Property="AutoGenerateColumns" Value="True"/></Style>
 <Style x:Key="Section" TargetType="TextBlock"><Setter Property="FontSize" Value="23"/><Setter Property="FontFamily" Value="Orbitron, Segoe UI"/><Setter Property="FontWeight" Value="Bold"/><Setter Property="Margin" Value="0,0,0,12"/></Style>
 <Style x:Key="Card" TargetType="Border"><Setter Property="Background" Value="#1B1B20"/><Setter Property="BorderBrush" Value="#303036"/><Setter Property="BorderThickness" Value="1"/><Setter Property="CornerRadius" Value="9"/><Setter Property="Padding" Value="18"/><Setter Property="Margin" Value="0,0,12,12"/></Style>
 <Style x:Key="Nav" TargetType="Button"><Setter Property="HorizontalContentAlignment" Value="Left"/><Setter Property="Margin" Value="0,0,0,7"/><Setter Property="Padding" Value="14,13"/><Setter Property="Background" Value="#1B1B20"/><Setter Property="BorderBrush" Value="#303036"/></Style>
</Window.Resources>
<DockPanel>
 <Border DockPanel.Dock="Top" Background="#141417" BorderBrush="#303036" BorderThickness="0,0,0,1" Padding="22,14"><DockPanel><StackPanel Orientation="Horizontal" DockPanel.Dock="Left"><Image x:Name="BrandIcon" Width="48" Height="48" Stretch="Uniform" Margin="0,0,18,0"/><StackPanel Orientation="Horizontal" VerticalAlignment="Center"><Image x:Name="BrandLogo" Height="66" Width="225" Stretch="Uniform" Margin="0,0,24,0"/><Border Width="1" Background="#49434A" Margin="0,8,24,8"/><StackPanel VerticalAlignment="Center"><TextBlock Text="BERRY IT NETWORK UTILITY" FontFamily="Orbitron, Segoe UI" FontSize="23" FontWeight="Bold" Foreground="#F4F4F4"/><TextBlock Text="NETWORK DIAGNOSTICS" Foreground="#ACAAB2" FontSize="11" Margin="0,3,0,0"/></StackPanel></StackPanel></StackPanel><TextBlock x:Name="HeaderStatus" Text="Ready · No tests running" HorizontalAlignment="Right" Foreground="#ACAAB2" VerticalAlignment="Center"/></DockPanel></Border>
 <Border DockPanel.Dock="Bottom" Background="#141417" Padding="13,7" BorderBrush="#303036" BorderThickness="0,1,0,0"><DockPanel><TextBlock Text="Berry IT Network Utility v0.4.0 · Read-only network diagnostics" Foreground="#ACAAB2"/><TextBlock x:Name="Footer" Text="No administrator privileges required for normal tests" Foreground="#ACAAB2" HorizontalAlignment="Right"/></DockPanel></Border>
 <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="214"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Border Grid.Column="0" Background="#141417" BorderThickness="0,0,1,0" BorderBrush="#303036" Padding="13,20"><StackPanel>
   <Button x:Name="NavDashboard" Style="{StaticResource Nav}" Content="⌂  Dashboard"/><Button x:Name="NavPing" Style="{StaticResource Nav}" Content="◉  Ping Monitor"/><Button x:Name="NavSpeed" Style="{StaticResource Nav}" Content="⇅  Speed Test"/><Button x:Name="NavRoute" Style="{StaticResource Nav}" Content="↳  Traceroute"/><Button x:Name="NavDNS" Style="{StaticResource Nav}" Content="◎  DNS Lookup"/><Button x:Name="NavPort" Style="{StaticResource Nav}" Content="▣  Port Tester"/><Button x:Name="NavAdapters" Style="{StaticResource Nav}" Content="▤  Network Adapters"/><Button x:Name="NavReports" Style="{StaticResource Nav}" Content="▧  Reports"/>
   <TextBlock Text="BERRY IT" FontFamily="Orbitron, Segoe UI" Foreground="#FF3C3C" Margin="13,32,0,2" FontWeight="Bold"/><TextBlock Text="Consult · Secure · Manage · Grow" FontSize="10" Foreground="#ACAAB2" Margin="13,0,0,0"/>
  </StackPanel></Border>
  <Grid Grid.Column="1" Margin="22,20,12,10">
   <ScrollViewer x:Name="PageDashboard" VerticalScrollBarVisibility="Auto"><StackPanel>
     <TextBlock Text="Network Overview" Style="{StaticResource Section}"/><TextBlock Text="Your active connection and quick diagnostics" Foreground="#ACAAB2" Margin="0,0,0,18"/>
     <UniformGrid Columns="2">
      <Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="ACTIVE ADAPTER" Foreground="#ACAAB2"/><TextBlock x:Name="DashAdapter" Text="Loading..." FontSize="20" FontWeight="SemiBold"/><TextBlock x:Name="DashAddress" Text="—" Foreground="#ACAAB2"/></StackPanel></Border>
      <Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="DEFAULT GATEWAY" Foreground="#ACAAB2"/><TextBlock x:Name="DashGateway" Text="—" FontSize="20" FontWeight="SemiBold"/><TextBlock x:Name="DashDNS" Text="DNS: —" Foreground="#ACAAB2"/></StackPanel></Border>
     </UniformGrid>
     <Border Style="{StaticResource Card}"><StackPanel>
       <DockPanel Margin="0,0,0,12"><TextBlock Text="THIS DEVICE &amp; NETWORK" FontSize="17" FontWeight="Bold"/><Button x:Name="DashDeviceRefresh" Content="Refresh details" HorizontalAlignment="Right"/></DockPanel>
       <UniformGrid Columns="2">
        <StackPanel Margin="0,0,12,14"><TextBlock Text="DEVICE HOSTNAME" Foreground="#ACAAB2"/><TextBlock x:Name="DashHostname" Text="Loading..." FontSize="19" FontWeight="SemiBold"/></StackPanel>
        <StackPanel Margin="0,0,12,14"><TextBlock Text="LOCAL IPv4" Foreground="#ACAAB2"/><TextBlock x:Name="DashLocalIP" Text="Loading..." FontSize="19" FontWeight="SemiBold"/></StackPanel>
        <StackPanel Margin="0,0,12,14"><TextBlock Text="PUBLIC IPv4" Foreground="#ACAAB2"/><TextBlock x:Name="DashPublicIP" Text="Not checked" FontSize="19" FontWeight="SemiBold"/></StackPanel>
        <StackPanel Margin="0,0,12,14"><TextBlock Text="WINDOWS VERSION" Foreground="#ACAAB2"/><TextBlock x:Name="DashOS" Text="Loading..." FontSize="16"/></StackPanel>
        <StackPanel Margin="0,0,12,8"><TextBlock Text="NETWORK ADAPTER" Foreground="#ACAAB2"/><TextBlock x:Name="DashDeviceAdapter" Text="Loading..." FontSize="16"/></StackPanel>
        <StackPanel Margin="0,0,12,8"><TextBlock Text="LINK SPEED" Foreground="#ACAAB2"/><TextBlock x:Name="DashLinkSpeed" Text="Unknown" FontSize="16"/></StackPanel>
       </UniformGrid>
       <TextBlock Text="Public IP requires an HTTPS lookup and may show your VPN or internet gateway address." Foreground="#ACAAB2" FontSize="11"/>
     </StackPanel></Border>
     <Border Style="{StaticResource Card}"><StackPanel><DockPanel><TextBlock Text="INTERNET SPEED TEST" FontSize="17" FontWeight="Bold"/><Button x:Name="DashSpeedButton" Content="Run Speed Test" Background="#D00000" BorderBrush="#D00000" HorizontalAlignment="Right"/></DockPanel>
       <UniformGrid Columns="4" Margin="0,15,0,8"><StackPanel><TextBlock Text="DOWNLOAD" Foreground="#ACAAB2"/><TextBlock x:Name="DashDownload" Text="— Mbps" FontSize="24" FontWeight="Bold"/></StackPanel><StackPanel><TextBlock Text="UPLOAD" Foreground="#ACAAB2"/><TextBlock x:Name="DashUpload" Text="— Mbps" FontSize="24" FontWeight="Bold"/></StackPanel><StackPanel><TextBlock Text="PING" Foreground="#ACAAB2"/><TextBlock x:Name="DashSpeedPing" Text="— ms" FontSize="24" FontWeight="Bold"/></StackPanel><StackPanel><TextBlock Text="JITTER" Foreground="#ACAAB2"/><TextBlock x:Name="DashSpeedJitter" Text="— ms" FontSize="24" FontWeight="Bold"/></StackPanel></UniformGrid>
       <Button x:Name="DashSpeedDetails" Content="View detailed speed test" HorizontalAlignment="Left" Margin="0,10,0,10"/>
       <TextBlock x:Name="DashSpeedStatus" Text="Not tested · Built-in Cloudflare speed test" Foreground="#ACAAB2"/>
     </StackPanel></Border>
     <Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="QUICK DIAGNOSTICS" FontWeight="Bold" FontSize="16" Margin="0,0,0,12"/><WrapPanel><Button x:Name="DashPingButton" Content="Ping Monitor"/><Button x:Name="DashDNSButton" Content="DNS Lookup"/><Button x:Name="DashPortButton" Content="Port Tester"/><Button x:Name="DashRefreshButton" Content="Refresh Network"/></WrapPanel></StackPanel></Border>
   </StackPanel></ScrollViewer>
   <ScrollViewer x:Name="PagePing" Visibility="Collapsed" VerticalScrollBarVisibility="Auto"><StackPanel>
    <TextBlock Text="Ping Monitor" Style="{StaticResource Section}"/><TextBlock Text="Continuous ICMP latency measurements" Foreground="#ACAAB2" Margin="0,0,0,18"/>
    <Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="TARGET HOST / IP" Foreground="#ACAAB2"/><TextBox x:Name="PingTarget" Text="1.1.1.1" Margin="0,5,0,12"/><WrapPanel><StackPanel Width="130"><TextBlock Text="Interval (ms)"/><TextBox x:Name="PingInterval" Text="1000"/></StackPanel><StackPanel Width="130" Margin="12,0"><TextBlock Text="Timeout (ms)"/><TextBox x:Name="PingTimeout" Text="2000"/></StackPanel><StackPanel Width="130"><TextBlock Text="Payload (bytes)"/><TextBox x:Name="PingPayload" Text="32"/></StackPanel></WrapPanel><WrapPanel Margin="0,15,0,0"><Button x:Name="PingStart" Content="Start monitoring" Background="#D00000" BorderBrush="#D00000"/><Button x:Name="PingStop" Content="Stop"/><Button x:Name="PingClear" Content="Reset results"/><Button x:Name="PingExport" Content="Export CSV"/></WrapPanel></StackPanel></Border>
    <UniformGrid Columns="4"><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="LATEST" Foreground="#ACAAB2"/><TextBlock x:Name="PingLatest" Text="— ms" FontSize="21"/></StackPanel></Border><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="AVERAGE" Foreground="#ACAAB2"/><TextBlock x:Name="PingAverage" Text="— ms" FontSize="21"/></StackPanel></Border><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="PACKET LOSS" Foreground="#ACAAB2"/><TextBlock x:Name="PingLoss" Text="— %" FontSize="21"/></StackPanel></Border><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="SENT / RECEIVED" Foreground="#ACAAB2"/><TextBlock x:Name="PingCount" Text="0 / 0" FontSize="21"/></StackPanel></Border></UniformGrid>
    <Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="LIVE LATENCY (LAST 90 PROBES)" Foreground="#ACAAB2" Margin="0,0,0,8"/><Canvas x:Name="PingCanvas" Height="175" Background="#111116" ClipToBounds="True"/><TextBlock Text="Red line: successful probes · Gaps: timeouts" Foreground="#ACAAB2" Margin="0,8,0,0"/></StackPanel></Border>
    <DataGrid x:Name="PingGrid" Height="240" Margin="0,0,12,10"/>
   </StackPanel></ScrollViewer>
   <ScrollViewer x:Name="PageSpeed" Visibility="Collapsed" VerticalScrollBarVisibility="Auto"><StackPanel>
    <TextBlock Text="Speed Test" Style="{StaticResource Section}"/><TextBlock Text="Real download / upload measurement using built-in PowerShell and .NET" Foreground="#ACAAB2" Margin="0,0,0,18"/>
    <Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="SPEEDTEST ENGINE" Foreground="#ACAAB2"/><TextBlock x:Name="SpeedEngine" Text="Built-in measurement engine" Margin="0,6,0,12"/><WrapPanel><Button x:Name="SpeedStart" Content="Run Speed Test" Background="#D00000" BorderBrush="#D00000"/><Button x:Name="SpeedLocate" Content="About test server"/></WrapPanel><TextBlock x:Name="SpeedStatus" Text="No speed test has been run." Foreground="#ACAAB2" Margin="0,12,0,0"/></StackPanel></Border>
    <UniformGrid Columns="2"><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="DOWNLOAD" Foreground="#ACAAB2"/><TextBlock x:Name="SpeedDownload" Text="— Mbps" FontSize="32" FontWeight="Bold"/></StackPanel></Border><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="UPLOAD" Foreground="#ACAAB2"/><TextBlock x:Name="SpeedUpload" Text="— Mbps" FontSize="32" FontWeight="Bold"/></StackPanel></Border><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="LATENCY" Foreground="#ACAAB2"/><TextBlock x:Name="SpeedLatency" Text="— ms" FontSize="32"/></StackPanel></Border><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="JITTER" Foreground="#ACAAB2"/><TextBlock x:Name="SpeedJitter" Text="— ms" FontSize="32"/></StackPanel></Border></UniformGrid>
    <Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="TEST HISTORY" Foreground="#ACAAB2" Margin="0,0,0,8"/><DataGrid x:Name="SpeedGrid" Height="220"/><Button x:Name="SpeedExport" Content="Export History CSV" Margin="0,12,0,0" HorizontalAlignment="Left"/></StackPanel></Border>
    <TextBlock Text="Speed tests transfer data to Cloudflare speed test endpoints and may use significant bandwidth. Berry IT Network Utility never starts a speed test automatically. Results are approximate and may differ from commercial multi-connection tests." Foreground="#ACAAB2" Margin="3,0,10,12"/>
   </StackPanel></ScrollViewer>
   <ScrollViewer x:Name="PageRoute" Visibility="Collapsed"><StackPanel><TextBlock Text="Traceroute" Style="{StaticResource Section}"/><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="Destination IP or hostname" Foreground="#ACAAB2"/><TextBox x:Name="RouteTarget" Text="1.1.1.1" Margin="0,6,0,12"/><WrapPanel><Button x:Name="RouteStart" Content="Run traceroute" Background="#D00000" BorderBrush="#D00000"/><Button x:Name="RouteExport" Content="Export TXT"/></WrapPanel><TextBlock x:Name="RouteStatus" Text="Ready" Foreground="#ACAAB2" Margin="0,10,0,0"/></StackPanel></Border><TextBox x:Name="RouteOutput" Height="350" FontFamily="Consolas" FontSize="12" IsReadOnly="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/></StackPanel></ScrollViewer>
   <ScrollViewer x:Name="PageDNS" Visibility="Collapsed"><StackPanel><TextBlock Text="DNS Lookup" Style="{StaticResource Section}"/><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="Hostname / record" Foreground="#ACAAB2"/><TextBox x:Name="DNSTarget" Text="example.com" Margin="0,6,0,10"/><WrapPanel><StackPanel Width="135"><TextBlock Text="Record type"/><ComboBox x:Name="DNSType"><ComboBoxItem Content="A" IsSelected="True"/><ComboBoxItem Content="AAAA"/><ComboBoxItem Content="MX"/><ComboBoxItem Content="TXT"/><ComboBoxItem Content="NS"/><ComboBoxItem Content="CNAME"/><ComboBoxItem Content="PTR"/><ComboBoxItem Content="SOA"/></ComboBox></StackPanel><StackPanel Width="245"><TextBlock Text="Resolver (blank = system)"/><TextBox x:Name="DNSServer" Text=""/></StackPanel></WrapPanel><Button x:Name="DNSStart" Content="Query DNS" Background="#D00000" BorderBrush="#D00000" HorizontalAlignment="Left" Margin="0,13,0,0"/><TextBlock x:Name="DNSStatus" Text="Ready" Foreground="#ACAAB2" Margin="0,8,0,0"/></StackPanel></Border><TextBox x:Name="DNSOutput" Height="350" FontFamily="Consolas" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/></StackPanel></ScrollViewer>
   <ScrollViewer x:Name="PagePort" Visibility="Collapsed"><StackPanel><TextBlock Text="TCP Port Tester" Style="{StaticResource Section}"/><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="Target hostname / IP" Foreground="#ACAAB2"/><TextBox x:Name="PortTarget" Text="1.1.1.1" Margin="0,6,0,12"/><WrapPanel><StackPanel Width="140"><TextBlock Text="TCP port"/><TextBox x:Name="PortNumber" Text="443"/></StackPanel><StackPanel Width="160" Margin="12,0"><TextBlock Text="Timeout (ms)"/><TextBox x:Name="PortTimeout" Text="3000"/></StackPanel></WrapPanel><WrapPanel Margin="0,13,0,5"><Button x:Name="PortStart" Content="Test connection" Background="#D00000" BorderBrush="#D00000"/><Button x:Name="Port443" Content="HTTPS · 443"/><Button x:Name="Port80" Content="HTTP · 80"/><Button x:Name="Port3389" Content="RDP · 3389"/></WrapPanel><TextBlock x:Name="PortStatus" Text="Ready" Foreground="#ACAAB2" Margin="0,10,0,0" FontSize="16"/></StackPanel></Border></StackPanel></ScrollViewer>
   <ScrollViewer x:Name="PageAdapters" Visibility="Collapsed"><StackPanel><TextBlock Text="Network Adapters" Style="{StaticResource Section}"/><WrapPanel Margin="0,0,0,12"><Button x:Name="AdaptersRefresh" Content="Refresh adapter information" Background="#D00000" BorderBrush="#D00000"/><Button x:Name="AdaptersExport" Content="Export TXT"/></WrapPanel><TextBox x:Name="AdaptersOutput" IsReadOnly="True" FontFamily="Consolas" TextWrapping="NoWrap" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" Height="510"/></StackPanel></ScrollViewer>
   <ScrollViewer x:Name="PageReports" Visibility="Collapsed"><StackPanel><TextBlock Text="Reports &amp; History" Style="{StaticResource Section}"/><TextBlock Text="Create local diagnostic reports. These files may contain addresses and hostnames; review before sharing." Foreground="#ACAAB2" Margin="0,0,0,18"/><Border Style="{StaticResource Card}"><StackPanel><TextBlock Text="SESSION REPORT" FontWeight="Bold"/><TextBlock Text="Includes adapter details, ping samples and speed test history." Foreground="#ACAAB2" Margin="0,10,0,14"/><Button x:Name="ReportSave" Content="Save Report (TXT)" Background="#D00000" BorderBrush="#D00000" HorizontalAlignment="Left"/></StackPanel></Border></StackPanel></ScrollViewer>
  </Grid>
 </Grid>
</DockPanel>
</Window>
'@
try {
 $reader = [System.Xml.XmlNodeReader]::new($xaml)
 $window = [Windows.Markup.XamlReader]::Load($reader)
} catch { Write-Error "Berry IT Network Utility GUI XAML failed to load: $($_.Exception.Message)"; return }
$names = @('BrandIcon','BrandLogo','HeaderStatus','Footer','DashHostname','DashLocalIP','DashPublicIP','DashOS','DashDeviceAdapter','DashLinkSpeed','DashAdapter','DashAddress','DashGateway','DashDNS','DashDownload','DashUpload','DashSpeedPing','DashSpeedJitter','DashSpeedStatus','PingTarget','PingInterval','PingTimeout','PingPayload','PingLatest','PingAverage','PingLoss','PingCount','PingCanvas','PingGrid','SpeedEngine','SpeedStatus','SpeedDownload','SpeedUpload','SpeedLatency','SpeedJitter','SpeedGrid','RouteTarget','RouteStatus','RouteOutput','DNSTarget','DNSType','DNSServer','DNSStatus','DNSOutput','PortTarget','PortNumber','PortTimeout','PortStatus','AdaptersOutput')
foreach($n in $names) { Set-Variable -Name $n -Value $window.FindName($n) -Scope Script }
$script:pingRows = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$script:speedRows = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
# Load the original, approved BERRY IT assets from the website repository.
# Cache in LocalAppData so subsequent launches work without a network connection.
function Initialize-BerryBranding {
 $assetFolder=Join-Path $env:LOCALAPPDATA 'BerryIT\\Berry IT Network Utility\\Branding'
 try { [IO.Directory]::CreateDirectory($assetFolder) | Out-Null } catch { return }
 $origin='https://raw.githubusercontent.com/BERRY-2508/BERRY.IT.WEB/main/'
 $assets=@(@('berry-it-icon.png','Berry IT Network Utility-Icon.png'),@('berry-it-full-logo.png','BerryIT-Full-Logo.png'))
 foreach($asset in $assets){
  $file=Join-Path $assetFolder $asset[1]
  if(-not (Test-Path -LiteralPath $file)){
   try {
    $wc=[System.Net.WebClient]::new()
    try { $wc.DownloadFile($origin+$asset[0],$file) } finally { $wc.Dispose() }
   } catch { if(Test-Path -LiteralPath $file){Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue} }
  }
 }
 function Get-BrandBitmap($file){
  if(-not (Test-Path -LiteralPath $file)){return $null}
  try {
   $bitmap=[Windows.Media.Imaging.BitmapImage]::new()
   $bitmap.BeginInit();$bitmap.CacheOption=[Windows.Media.Imaging.BitmapCacheOption]::OnLoad
   $bitmap.UriSource=[Uri]::new($file,[UriKind]::Absolute)
   $bitmap.EndInit();$bitmap.Freeze()
   return $bitmap
  }catch{return $null}
 }
 $iconBitmap=Get-BrandBitmap (Join-Path $assetFolder 'Berry IT Network Utility-Icon.png')
 if($iconBitmap){$BrandIcon.Source=$iconBitmap;$window.Icon=$iconBitmap}
 $logoBitmap=Get-BrandBitmap (Join-Path $assetFolder 'BerryIT-Full-Logo.png')
 if($logoBitmap){$BrandLogo.Source=$logoBitmap}else{$BrandLogo.Visibility='Collapsed'}
 if(-not $iconBitmap){$BrandIcon.Visibility='Collapsed'}
}
Initialize-BerryBranding

$PingGrid.ItemsSource = $script:pingRows; $SpeedGrid.ItemsSource = $script:speedRows
$script:pingSamples = [System.Collections.Generic.List[double]]::new()
$script:pingSent=0; $script:pingReceived=0; $script:pingActive=$false; $script:pingTask=$null
$script:speedExe=$null; $script:speedJob=$null; $script:workJob=$null
function Show-Page([string]$page) {
 foreach($p in @('Dashboard','Ping','Speed','Route','DNS','Port','Adapters','Reports')) {
  $window.FindName('Page'+$p).Visibility = if($p -eq $page){'Visible'}else{'Collapsed'}
  $window.FindName('Nav'+$p).Background = if($p -eq $page){ [Windows.Media.BrushConverter]::new().ConvertFromString('#68252B') }else{ [Windows.Media.BrushConverter]::new().ConvertFromString('#1B1B20') }
 }
}
foreach($p in @('Dashboard','Ping','Speed','Route','DNS','Port','Adapters','Reports')) { $b=$window.FindName('Nav'+$p); $pageName=$p; $b.Add_Click(([scriptblock]::Create("Show-Page '$pageName'"))) }
$script:publicIPJob=$null
function Refresh-PublicIP {
 if($script:publicIPJob -and $script:publicIPJob.State -in @('Running','NotStarted')) { return }
 $DashPublicIP.Text='Checking...'
 try {
  $script:publicIPJob=Start-Job -ScriptBlock {
   try {
    [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
    $ip=(Invoke-RestMethod -Uri 'https://api4.ipify.org' -TimeoutSec 8 -ErrorAction Stop).ToString().Trim()
    $out=[Net.IPAddress]::None
    if(-not [Net.IPAddress]::TryParse($ip,[ref]$out) -or $out.AddressFamily -ne [Net.Sockets.AddressFamily]::InterNetwork){throw 'Unexpected response'}
    $ip
   }catch{'Unavailable'}
  }
 }catch{$DashPublicIP.Text='Unavailable'}
}
function Refresh-Network {
 try {
  $DashHostname.Text=$env:COMPUTERNAME
  try {
   $os=Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
   $DashOS.Text=($os.Caption + ' (' + $os.OSArchitecture + ')')
  }catch{$DashOS.Text=[Environment]::OSVersion.VersionString}
  $nics = @(Get-NetIPConfiguration -ErrorAction Stop | Where-Object { $_.IPv4DefaultGateway } | Select-Object -First 1)
  if($nics.Count) {
   $nic=$nics[0];$DashAdapter.Text=$nic.InterfaceAlias
   $ipList=(($nic.IPv4Address | Select-Object -ExpandProperty IPv4Address) -join ', ')
   $DashAddress.Text='IPv4: '+$ipList;$DashLocalIP.Text=$ipList
   $DashDeviceAdapter.Text=$nic.InterfaceAlias
   try{$adapter=Get-NetAdapter -InterfaceIndex $nic.InterfaceIndex -ErrorAction Stop;$DashLinkSpeed.Text=[string]$adapter.LinkSpeed}catch{$DashLinkSpeed.Text='Unknown'}
   $DashGateway.Text=(($nic.IPv4DefaultGateway | Select-Object -ExpandProperty NextHop) -join ', ')
   $dns=@(Get-DnsClientServerAddress -InterfaceIndex $nic.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ServerAddresses)
   $DashDNS.Text='DNS: '+($dns -join ', ')
  }else{$DashAdapter.Text='No default IPv4 adapter';$DashAddress.Text='No active default route';$DashGateway.Text='—';$DashDNS.Text='DNS: —';$DashLocalIP.Text='Unavailable';$DashDeviceAdapter.Text='No active default adapter';$DashLinkSpeed.Text='Unknown'}
 } catch {$DashAdapter.Text='Network unavailable';$DashAddress.Text=$_.Exception.Message;$DashLocalIP.Text='Unavailable';$DashDeviceAdapter.Text='Unavailable';$DashLinkSpeed.Text='Unknown'}
}
function Save-Text([string]$text,[string]$filename,[string]$filter='Text files (*.txt)|*.txt') {
 $dlg=[Microsoft.Win32.SaveFileDialog]::new();$dlg.FileName=$filename;$dlg.Filter=$filter
 if($dlg.ShowDialog($window)) { [IO.File]::WriteAllText($dlg.FileName,$text,[Text.UTF8Encoding]::new($true)) }
}
function Start-Work([scriptblock]$code,[object[]]$arguments,[scriptblock]$onDone) {
 if($script:workJob -and $script:workJob.State -in @('Running','NotStarted')) { [Windows.MessageBox]::Show('Another diagnostic is running. Try again after it completes.','Berry IT Network Utility') | Out-Null; return }
 $script:workJob=Start-Job -ScriptBlock $code -ArgumentList $arguments
 $script:workComplete=$onDone
 $HeaderStatus.Text='Diagnostic in progress...'
}
$script:pingTimer=[Windows.Threading.DispatcherTimer]::new();$script:pingTimer.Interval=[TimeSpan]::FromMilliseconds(150)
$script:nextPing=[DateTime]::MinValue
$script:pingTimer.Add_Tick({
 if(-not $script:pingActive){return}
 if($script:pingTask -and $script:pingTask.IsCompleted) {
  $r=$null;try{$r=$script:pingTask.GetAwaiter().GetResult()}catch{}
  $script:pingSent++
  $ms=[double]::NaN; $state='Timeout/Error'
  if($r -and $r.Status -eq [Net.NetworkInformation.IPStatus]::Success){$ms=[double]$r.RoundtripTime;$state='Success';$script:pingReceived++}
  $script:pingSamples.Add($ms);if($script:pingSamples.Count -gt 90){$script:pingSamples.RemoveAt(0)}
  $script:pingRows.Insert(0,[pscustomobject]@{Time=(Get-Date).ToString('HH:mm:ss');Target=$PingTarget.Text;Latency=if([double]::IsNaN($ms)){'—'}else{"$ms ms"};Status=$state})
  if($script:pingRows.Count -gt 500){$script:pingRows.RemoveAt($script:pingRows.Count-1)}
  $good=@($script:pingSamples | Where-Object {-not [double]::IsNaN($_)})
  $PingLatest.Text=if([double]::IsNaN($ms)){'Timeout'}else{"$ms ms"}
  if($good.Count){$PingAverage.Text=('{0:N1} ms' -f (($good|Measure-Object -Average).Average))}
  $PingLoss.Text=('{0:N1} %' -f (100*($script:pingSent-$script:pingReceived)/[Math]::Max(1,$script:pingSent)))
  $PingCount.Text="$($script:pingSent) / $($script:pingReceived)"
  $PingCanvas.Children.Clear()
  $w=[Math]::Max(100,$PingCanvas.ActualWidth);$h=[Math]::Max(100,$PingCanvas.ActualHeight)
  $valid=@($script:pingSamples | Where-Object {-not [double]::IsNaN($_)})
  $max=30; if($valid.Count){$max=[Math]::Max(30,($valid|Measure-Object -Maximum).Maximum * 1.2)}
  $lastPoint=$null
  for($i=0;$i -lt $script:pingSamples.Count;$i++){
   $v=$script:pingSamples[$i];if([double]::IsNaN($v)){$lastPoint=$null;continue}
   $point=[Windows.Point]::new(($i/89.0)*($w-12)+6,($h-12)-($v/$max)*($h-20))
   if($lastPoint){$line=[Windows.Shapes.Line]::new();$line.X1=$lastPoint.X;$line.Y1=$lastPoint.Y;$line.X2=$point.X;$line.Y2=$point.Y;$line.Stroke=[Windows.Media.Brushes]::Tomato;$line.StrokeThickness=2.3;$PingCanvas.Children.Add($line)|Out-Null}
   $lastPoint=$point
  }
  $script:pingTask=$null
 }
 if(-not $script:pingTask -and [DateTime]::UtcNow -ge $script:nextPing){
  try{$script:pingTask=$script:pinger.SendPingAsync($PingTarget.Text.Trim(),$script:pingTimeout,[byte[]]::new($script:pingPayload))}
  catch{$HeaderStatus.Text='Ping error: '+$_.Exception.Message;$script:pingActive=$false;return}
  $script:nextPing=[DateTime]::UtcNow.AddMilliseconds($script:pingInterval)
 }
})
$window.FindName('PingStart').Add_Click({
 $intervalNumber=0;$timeoutNumber=0;$payloadNumber=0
 if(-not [int]::TryParse($PingInterval.Text,[ref]$intervalNumber) -or $intervalNumber -lt 250 -or $intervalNumber -gt 60000 -or -not [int]::TryParse($PingTimeout.Text,[ref]$timeoutNumber) -or $timeoutNumber -lt 250 -or $timeoutNumber -gt 30000 -or -not [int]::TryParse($PingPayload.Text,[ref]$payloadNumber) -or $payloadNumber -lt 0 -or $payloadNumber -gt 1400 -or [string]::IsNullOrWhiteSpace($PingTarget.Text)){[Windows.MessageBox]::Show('Enter a target and valid interval (250–60000 ms), timeout (250–30000 ms) and payload (0–1400 bytes).','Berry IT Network Utility')|Out-Null;return}
 $script:pingInterval=$intervalNumber;$script:pingTimeout=$timeoutNumber;$script:pingPayload=$payloadNumber
 if(-not $script:pinger){$script:pinger=[Net.NetworkInformation.Ping]::new()}
 $script:pingActive=$true;$script:nextPing=[DateTime]::MinValue;$script:pingTimer.Start();$HeaderStatus.Text='Monitoring '+$PingTarget.Text
})
$window.FindName('PingStop').Add_Click({$script:pingActive=$false;$script:pingTimer.Stop();$HeaderStatus.Text='Ping monitoring stopped'})
$window.FindName('PingClear').Add_Click({$script:pingRows.Clear();$script:pingSamples.Clear();$script:pingSent=0;$script:pingReceived=0;$PingCanvas.Children.Clear();$PingLatest.Text='— ms';$PingAverage.Text='— ms';$PingLoss.Text='— %';$PingCount.Text='0 / 0'})
$window.FindName('PingExport').Add_Click({$rows=@($script:pingRows);if($rows.Count){Save-Text (($rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n") 'berry-it-network-utility-ping.csv' 'CSV files (*.csv)|*.csv'}})
function Display-SpeedEngine {
 $SpeedEngine.Text='Built-in .NET engine  |  Cloudflare edge  |  No additional install required'
}
$window.FindName('SpeedLocate').Add_Click({
 [Windows.MessageBox]::Show('Berry IT Network Utility uses HTTPS requests to https://speed.cloudflare.com/__down and /__up. Download, upload and latency measurements use real data transfers, not demo figures. The test runs only when you press Run Speed Test. Results can vary with connection conditions.','Built-in speed test') | Out-Null
})
$window.FindName('SpeedStart').Add_Click({
 if($script:speedJob -and $script:speedJob.State -in @('Running','NotStarted')){return}
 $SpeedStatus.Text='Measuring latency, download and upload... (may consume ~20 MB of data)'
 $DashSpeedStatus.Text='Speed test running...'
 $script:speedJob=Start-Job -ScriptBlock {
  $ErrorActionPreference='Stop'
  [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
  Add-Type -AssemblyName System.Net.Http
  $http=[Net.Http.HttpClient]::new()
  $http.Timeout=[TimeSpan]::FromSeconds(25)
  $http.DefaultRequestHeaders.UserAgent.ParseAdd('Berry IT Network Utility/0.3')
  try {
   $latencies=New-Object 'System.Collections.Generic.List[double]'
   for($i=0;$i -lt 5;$i++){
    $watch=[Diagnostics.Stopwatch]::StartNew()
    $response=$http.GetAsync('https://speed.cloudflare.com/__down?bytes=0',[Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
    try {$response.EnsureSuccessStatusCode() | Out-Null;$watch.Stop();$latencies.Add($watch.Elapsed.TotalMilliseconds)}finally{$response.Dispose()}
   }
   $samplesDown=New-Object 'System.Collections.Generic.List[double]'
   foreach($size in @(1000000,4000000,8000000)) {
    $watch=[Diagnostics.Stopwatch]::StartNew()
    $response=$http.GetAsync("https://speed.cloudflare.com/__down?bytes=$size",[Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
    try {
     $response.EnsureSuccessStatusCode() | Out-Null
     $stream=$response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
     $buffer=New-Object byte[] 65536
     $count=0L
     try {while(($n=$stream.Read($buffer,0,$buffer.Length)) -gt 0){$count+=$n}}finally{$stream.Dispose()}
     $watch.Stop()
     if($count -lt ($size * 0.95)){throw "Unexpected download size: received $count bytes, requested $size"}
     if($watch.Elapsed.TotalSeconds -gt 0){$samplesDown.Add(($count * 8 / $watch.Elapsed.TotalSeconds / 1000000))}
    }finally{$response.Dispose()}
   }
   $samplesUp=New-Object 'System.Collections.Generic.List[double]'
   foreach($size in @(500000,1000000,2000000)) {
    $bytes=New-Object byte[] $size
    $content=[Net.Http.ByteArrayContent]::new($bytes)
    $watch=[Diagnostics.Stopwatch]::StartNew()
    try {
     $response=$http.PostAsync('https://speed.cloudflare.com/__up',$content).GetAwaiter().GetResult()
     try {$response.EnsureSuccessStatusCode() | Out-Null;$watch.Stop()}finally{$response.Dispose()}
     if($watch.Elapsed.TotalSeconds -gt 0){$samplesUp.Add(($size * 8 / $watch.Elapsed.TotalSeconds / 1000000))}
    }finally{$content.Dispose()}
   }
   if($samplesDown.Count -eq 0 -or $samplesUp.Count -eq 0){throw 'Insufficient successful transfers to measure speed.'}
   # Use the highest successful sample to reduce short-transfer startup overhead.
   $down=($samplesDown | Measure-Object -Maximum).Maximum
   $up=($samplesUp | Measure-Object -Maximum).Maximum
   $lat=($latencies | Measure-Object -Average).Average
   $diffs=@(for($i=1;$i -lt $latencies.Count;$i++){[Math]::Abs($latencies[$i]-$latencies[$i-1])})
   $jitter=if($diffs.Count){($diffs | Measure-Object -Average).Average}else{0}
   [pscustomobject]@{ok=$true;downloadMbps=[Math]::Round($down,2);uploadMbps=[Math]::Round($up,2);latencyMs=[Math]::Round($lat,1);jitterMs=[Math]::Round($jitter,1);server='Cloudflare edge'} | ConvertTo-Json -Compress
  }catch {
   [pscustomobject]@{ok=$false;error=$_.Exception.Message} | ConvertTo-Json -Compress
  }finally {$http.Dispose()}
 }
})
$window.FindName('SpeedExport').Add_Click({$rows=@($script:speedRows);if($rows.Count){Save-Text (($rows | ConvertTo-Csv -NoTypeInformation)-join "`r`n") 'berry-it-network-utility-speedtest.csv' 'CSV files (*.csv)|*.csv'}})
$window.FindName('RouteStart').Add_Click({$target=$RouteTarget.Text.Trim();if(-not $target){return};$RouteStatus.Text='Tracing...';Start-Work {param($hostname) & tracert.exe -d -w 1200 -h 20 $hostname | Out-String} @($target) {param($result)$RouteOutput.Text=$result;$RouteStatus.Text='Traceroute completed'}})
$window.FindName('RouteExport').Add_Click({Save-Text $RouteOutput.Text 'berry-it-network-utility-traceroute.txt'})
$window.FindName('DNSStart').Add_Click({$target=$DNSTarget.Text.Trim();if(-not $target){return};$type=([Windows.Controls.ComboBoxItem]$DNSType.SelectedItem).Content.ToString();$server=$DNSServer.Text.Trim();$DNSStatus.Text='Querying...';Start-Work {param($hostname,$recordType,$resolver) if($resolver){Resolve-DnsName -Name $hostname -Type $recordType -Server $resolver -ErrorAction Stop | Format-List | Out-String}else{Resolve-DnsName -Name $hostname -Type $recordType -ErrorAction Stop | Format-List | Out-String}} @($target,$type,$server) {param($result)$DNSOutput.Text=$result;$DNSStatus.Text='DNS query completed'}})
$window.FindName('PortStart').Add_Click({
 $num=0;$time=0;if(-not [int]::TryParse($PortNumber.Text,[ref]$num) -or $num -lt 1 -or $num -gt 65535 -or -not [int]::TryParse($PortTimeout.Text,[ref]$time) -or $time -lt 100 -or $time -gt 30000 -or -not $PortTarget.Text.Trim()){[Windows.MessageBox]::Show('Enter a valid hostname, TCP port 1–65535, and timeout 100–30000 ms.','Berry IT Network Utility')|Out-Null;return}
 $PortStatus.Text='Connecting...';Start-Work {param($hostname,$port,$timeout) $client=[Net.Sockets.TcpClient]::new();try{$sw=[Diagnostics.Stopwatch]::StartNew();$task=$client.ConnectAsync($hostname,$port);if(-not $task.Wait($timeout)){return "TCP $hostname`:$port — connection timed out after $timeout ms (cause unknown)"};if($task.IsFaulted){return "TCP connection failed: $($task.Exception.GetBaseException().Message)"};$sw.Stop();return "TCP connection succeeded: $hostname`:$port ($($sw.ElapsedMilliseconds) ms)"}catch{return "TCP connection failed: $($_.Exception.Message)"}finally{$client.Dispose()}} @($PortTarget.Text.Trim(),$num,$time) {param($result)$PortStatus.Text=$result}
})
foreach($portPreset in @(80,443,3389)){$btn=$window.FindName('Port'+$portPreset);$p=$portPreset;$btn.Add_Click(([scriptblock]::Create("`$PortNumber.Text='$p'")))}
function Refresh-Adapters {try{$AdaptersOutput.Text=(Get-NetIPConfiguration -Detailed | Format-List | Out-String -Width 180)+"`r`n--- Network adapter link details ---`r`n"+(Get-NetAdapter | Format-Table -AutoSize | Out-String -Width 180)}catch{$AdaptersOutput.Text=$_.Exception.Message}}
$window.FindName('AdaptersRefresh').Add_Click({Refresh-Adapters});$window.FindName('AdaptersExport').Add_Click({Save-Text $AdaptersOutput.Text 'berry-it-network-utility-adapters.txt'})
$window.FindName('ReportSave').Add_Click({$report="Berry IT Network Utility diagnostic session — $(Get-Date -Format o)`r`n`r`nADAPTERS`r`n$($AdaptersOutput.Text)`r`nPING`r`n$(($script:pingRows | Format-Table -AutoSize | Out-String))`r`nSPEED TEST HISTORY`r`n$(($script:speedRows | Format-Table -AutoSize | Out-String))";Save-Text $report 'berry-it-network-utility-report.txt'})
$window.FindName('DashSpeedButton').Add_Click({
 $window.FindName('SpeedStart').RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent))
});$window.FindName('DashSpeedDetails').Add_Click({Show-Page 'Speed'});
$window.FindName('DashDeviceRefresh').Add_Click({Refresh-Network;Refresh-PublicIP});$window.FindName('DashPingButton').Add_Click({Show-Page 'Ping'});$window.FindName('DashDNSButton').Add_Click({Show-Page 'DNS'});$window.FindName('DashPortButton').Add_Click({Show-Page 'Port'});$window.FindName('DashRefreshButton').Add_Click({Refresh-Network;Refresh-PublicIP})
$script:jobTimer=[Windows.Threading.DispatcherTimer]::new();$script:jobTimer.Interval=[TimeSpan]::FromMilliseconds(350)
$script:jobTimer.Add_Tick({
 if($script:publicIPJob -and $script:publicIPJob.State -notin @('Running','NotStarted')) {
  $ipJob=$script:publicIPJob;$script:publicIPJob=$null
  $val=(Receive-Job -Job $ipJob -ErrorAction SilentlyContinue | Select-Object -Last 1)
  $DashPublicIP.Text=if($val){[string]$val}else{'Unavailable'}
  Remove-Job -Job $ipJob -Force -ErrorAction SilentlyContinue
 }

 if($script:workJob -and $script:workJob.State -notin @('Running','NotStarted')) {
  $j=$script:workJob;$script:workJob=$null
  $result=(Receive-Job -Job $j -ErrorAction SilentlyContinue | Out-String);$errors=@($j.ChildJobs[0].Error)
  if($errors.Count){$result += "`r`nERROR: " + (($errors | Out-String).Trim())}
  Remove-Job $j -Force -ErrorAction SilentlyContinue
  try{& $script:workComplete $result}catch{$HeaderStatus.Text='Diagnostic error: '+$_.Exception.Message;return}
  $HeaderStatus.Text=if($errors.Count){'Diagnostic failed — see results'}else{'Ready · Diagnostic completed'}
 }
 if($script:speedJob -and $script:speedJob.State -notin @('Running','NotStarted')) {
  $j=$script:speedJob;$script:speedJob=$null;$output=(Receive-Job $j -ErrorAction SilentlyContinue | Out-String);Remove-Job $j -Force -ErrorAction SilentlyContinue
  try {
   $json=($output -split '\r?\n' | Where-Object {$_.Trim().StartsWith('{')} | Select-Object -Last 1)
   $v=ConvertFrom-Json -InputObject $json -ErrorAction Stop
   if(-not $v.ok){throw [string]$v.error}
   $down=[double]$v.downloadMbps;$up=[double]$v.uploadMbps
   $lat=[double]$v.latencyMs;$jit=[double]$v.jitterMs
   $SpeedDownload.Text="$down Mbps";$SpeedUpload.Text="$up Mbps";$SpeedLatency.Text="$lat ms";$SpeedJitter.Text="$jit ms"
   $DashDownload.Text=$SpeedDownload.Text;$DashUpload.Text=$SpeedUpload.Text;$DashSpeedPing.Text=$SpeedLatency.Text;$DashSpeedJitter.Text=$SpeedJitter.Text
   $script:speedRows.Insert(0,[pscustomobject]@{Time=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss');DownloadMbps=$down;UploadMbps=$up;LatencyMs=$lat;JitterMs=$jit;Server='Cloudflare edge'})
   $SpeedStatus.Text='Completed · '+(Get-Date).ToString('HH:mm:ss');$DashSpeedStatus.Text='Last tested '+(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
  }catch{$SpeedStatus.Text='Test failed: '+$_.Exception.Message+' '+$output.Substring(0,[Math]::Min(220,$output.Length));$DashSpeedStatus.Text='Speed test failed (open Speed Test for details)'}
 }
})
$window.Add_Closed({$script:pingTimer.Stop();$script:jobTimer.Stop();$script:pingActive=$false;if($script:pinger){$script:pinger.Dispose()};foreach($j in @($script:workJob,$script:speedJob,$script:publicIPJob)){if($j){Stop-Job $j -ErrorAction SilentlyContinue;Remove-Job $j -Force -ErrorAction SilentlyContinue}}})
Refresh-Network;Refresh-PublicIP;Refresh-Adapters;Display-SpeedEngine;Show-Page 'Dashboard';$script:jobTimer.Start()
[void]$window.ShowDialog()