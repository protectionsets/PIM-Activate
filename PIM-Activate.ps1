<#
PIM Bulk Activator (WPF GUI) for Entra ID roles
- WPF template style: STA, robust XAML loader, status bar, optional Mica/DarkTitle interop
- Shows eligible roles as: ROLE - (Tenant wide \ AU Name) - MemberType
- Active roles highlighted in green; rows disabled for already-active items
- Synchronous active check (same runspace)
- Bulk-activate selected roles; bulk-deactivate all active roles
- Quick Admin portal buttons open in already-open default browser window/tab
Author: Yoni + Copilot
#>
$ErrorActionPreference = 'Stop'

#region --- Graph helpers ---
function Connect-GraphForRead {
    Connect-MgGraph -Scopes @(
        'User.Read',
        'RoleEligibilitySchedule.Read.Directory',
        'RoleAssignmentSchedule.Read.Directory',
        'AdministrativeUnit.Read.All'
    ) -NoWelcome
}
function Connect-GraphForWrite {
    Connect-MgGraph -Scopes @(
        'User.Read',
        'RoleManagement.ReadWrite.Directory',
        'RoleAssignmentSchedule.ReadWrite.Directory',
        'AdministrativeUnit.Read.All'
    ) -NoWelcome
}
function Invoke-GraphGet { param([Parameter(Mandatory)][string]$Uri) Invoke-MgGraphRequest -Method GET -Uri $Uri }
function Invoke-GraphPost {
    param([Parameter(Mandatory)][string]$Uri,[Parameter(Mandatory)]$BodyObject)
    $json = $BodyObject | ConvertTo-Json -Depth 10
    Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $json -ContentType 'application/json'
}
function Get-CurrentUser { Invoke-GraphGet -Uri 'https://graph.microsoft.com/v1.0/me' }

# --- Activation / Deactivation request functions ---
function Request-PIMActivation {
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$RoleDefinitionId,
        [Parameter(Mandatory)][string]$DirectoryScopeId,
        [Parameter(Mandatory)][string]$Justification,
        [Parameter(Mandatory)][TimeSpan]$Duration,
        [string]$TicketNumber,
        [string]$TicketSystem
    )
    $dur = [System.Xml.XmlConvert]::ToString($Duration)
    $body = @{
        action           = "selfActivate"
        principalId      = $UserId
        roleDefinitionId = $RoleDefinitionId
        directoryScopeId = $DirectoryScopeId
        justification    = $Justification
        scheduleInfo     = @{
            startDateTime = (Get-Date).ToUniversalTime().ToString("o")
            expiration    = @{ type = "afterDuration"; duration = $dur }
        }
    }
    if ($TicketNumber -or $TicketSystem) {
        $body.ticketInfo = @{ ticketNumber = ($TicketNumber ?? ""); ticketSystem = ($TicketSystem ?? "") }
    }
    try {
        $resp = Invoke-GraphPost -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests" -BodyObject $body
        [pscustomobject]@{ Status = $resp.status; RequestId = $resp.id; Message = "Submitted" }
    } catch {
        [pscustomobject]@{ Status = "Error"; RequestId = $null; Message = ($_.ErrorDetails.Message ?? $_.Exception.Message) }
    }
}
function Request-PIMDeactivation {
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][string]$RoleDefinitionId,
        [Parameter(Mandatory)][string]$DirectoryScopeId,
        [string]$Justification = "Deactivating all active roles"
    )
    $body = @{
        action           = "selfDeactivate"
        principalId      = $UserId
        roleDefinitionId = $RoleDefinitionId
        directoryScopeId = $DirectoryScopeId
        justification    = $Justification
    }
    try {
        $resp = Invoke-GraphPost -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests" -BodyObject $body
        [pscustomobject]@{ Status = $resp.status; RequestId = $resp.id; Message = "Submitted" }
    } catch {
        [pscustomobject]@{ Status = "Error"; RequestId = $null; Message = ($_.ErrorDetails.Message ?? $_.Exception.Message) }
    }
}
#endregion

# AU display cache
$script:AuDisplayCache = @{}
function Resolve-ScopeDisplay {
    param([string]$ScopeId)
    # Tenant-wide or empty scope -> exact wording requested
    if (-not $ScopeId -or $ScopeId -eq "/") { return "(Tenant wide)" }
    # Normalize resource path to GUID
    $normalizedId = $ScopeId
    if ($normalizedId -match '^/administrativeUnits/([0-9a-fA-F\-]+)$' -or
        $normalizedId -match '^/AdministrativeUnits/([0-9a-fA-F\-]+)$') { $normalizedId = $Matches[1] }
    # Cache?
    if ($script:AuDisplayCache.ContainsKey($normalizedId)) { return $script:AuDisplayCache[$normalizedId] }
    # Resolve via Graph
    try {
        $au = Invoke-GraphGet -Uri ("https://graph.microsoft.com/v1.0/directory/administrativeUnits/{0}" -f $normalizedId)
        $name = if ($au.displayName) { [string]$au.displayName } else { $normalizedId }
        $script:AuDisplayCache[$normalizedId] = $name
        return $name
    } catch {
        $script:AuDisplayCache[$normalizedId] = $normalizedId
        return $normalizedId
    }
}
function Get-PIMEligibleDirectoryRoles {
    param([Parameter(Mandatory)][string]$UserId)
    $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules" +
           "/filterByCurrentUser(on='principal')?`$expand=roleDefinition&`$top=999"
    $res = Invoke-GraphGet -Uri $uri
    $items = @()
    while ($true) {
        if ($res.value) { $items += $res.value }
        if (-not $res.'@odata.nextLink') { break }
        $res = Invoke-GraphGet -Uri $res.'@odata.nextLink'
    }
    $items |
    ForEach-Object {
        $scopeId = $_.directoryScopeId
        [pscustomobject]@{
            RoleDefinitionId = $_.roleDefinitionId
            RoleName         = $_.roleDefinition.displayName
            PrincipalId      = $_.principalId
            DirectoryScopeId = if ([string]::IsNullOrWhiteSpace($scopeId)) { "/" } else { $scopeId }
            ScopeDisplay     = Resolve-ScopeDisplay -ScopeId $scopeId
            EligibilityId    = $_.id
            MemberType       = $_.memberType # "Direct" | "Group"
            IsActive         = $false
            IsChecked        = $false
        }
    } | Sort-Object ScopeDisplay, RoleName
}
function Get-ActivePIMAssignmentsForUser {
    param([Parameter(Mandatory)][string]$UserId)
    $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentSchedules" +
           "?`$filter=principalId eq '$UserId' and assignmentType eq 'Activated'&`$expand=roleDefinition&`$top=999"
    $res = Invoke-GraphGet -Uri $uri
    $items = @()
    while ($true) {
        if ($res.value) { $items += $res.value }
        if (-not $res.'@odata.nextLink') { break }
        $res = Invoke-GraphGet -Uri $res.'@odata.nextLink'
    }
    $items |
    ForEach-Object {
        [pscustomobject]@{
            RoleDefinitionId = $_.roleDefinitionId
            RoleName         = $_.roleDefinition.displayName
            DirectoryScopeId = if ([string]::IsNullOrWhiteSpace($_.directoryScopeId)) { "/" } else { $_.directoryScopeId }
            ScopeDisplay     = Resolve-ScopeDisplay -ScopeId $_.directoryScopeId
        }
    }
}

#region --- Ensure STA & load WPF assemblies
try {
    if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        $thisScript = $PSCommandPath; if (-not $thisScript) { $thisScript = $MyInvocation.MyCommand.Path }
        if ($thisScript) {
            $hostPath = (Get-Process -Id $PID).Path
            $args = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File', $thisScript)
            Start-Process -FilePath $hostPath -ArgumentList $args -WorkingDirectory (Get-Location) | Out-Null
            return
        }
    }
} catch {}
#Requires -Version 5.1
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,System.Xaml
#endregion

#region --- Optional Windows 11 title bar & Mica
$dwTypes = @"
using System;
using System.Runtime.InteropServices;
public static class Dwm {
    [DllImport("dwmapi.dll", PreserveSig=true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@
try { Add-Type -TypeDefinition $dwTypes -ErrorAction Stop } catch {}
function Enable-DarkTitleBar { param([IntPtr]$Handle) foreach ($attr in 20,19) { try { $val = 1; [void][Dwm]::DwmSetWindowAttribute($Handle, $attr, [ref]$val, 4); break } catch {} } }
function Enable-Mica { param([IntPtr]$Handle) try { $DWMWA_SYSTEMBACKDROP_TYPE=38; $DWMSBT_MAINWINDOW=2; [void][Dwm]::DwmSetWindowAttribute($Handle,$DWMWA_SYSTEMBACKDROP_TYPE,[ref]$DWMSBT_MAINWINDOW,4) } catch {} }
#endregion

#region --- XAML (row selection replaces checkbox; disabled when active)
$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PIM Bulk Activator" Width="1000" Height="1200"
        WindowStartupLocation="Manual"
        Background="{DynamicResource WindowBackgroundBrush}"
        FontFamily="Segoe UI"
        SnapsToDevicePixels="True">
  <Window.Resources>
    <SolidColorBrush x:Key="WindowBackgroundBrush" Color="#F9F9FB"/>
    <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFFFF"/>
    <SolidColorBrush x:Key="BorderBrush" Color="#DDDDDD"/>
    <SolidColorBrush x:Key="TextBrush" Color="#111111"/>
    <SolidColorBrush x:Key="MutedTextBrush" Color="#666666"/>
    <SolidColorBrush x:Key="AccentStrokeBrush" Color="#4F6BED"/>
    <SolidColorBrush x:Key="ActiveGreenBrush" Color="#107C10"/>
    <CornerRadius x:Key="Radius-M">10</CornerRadius>

    <Style TargetType="Button">
      <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
      <Setter Property="Background" Value="{DynamicResource PanelBrush}"/>
      <Setter Property="Padding" Value="14,10"/>
      <Setter Property="Margin" Value="0,8,0,0"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="MinHeight" Value="36"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="HorizontalAlignment" Value="Stretch"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" CornerRadius="{DynamicResource Radius-M}"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}">
              <ContentPresenter Margin="4,0" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource AccentStrokeBrush}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource AccentStrokeBrush}"/>
                <Setter Property="Opacity" Value="0.95"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="SectionHeader" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource MutedTextBrush}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,16,0,4"/>
    </Style>

    <!-- Grouping source -->
    <CollectionViewSource x:Key="RolesView" Source="{Binding RoleItems}">
      <CollectionViewSource.GroupDescriptions>
        <PropertyGroupDescription PropertyName="ScopeDisplay"/>
      </CollectionViewSource.GroupDescriptions>
    </CollectionViewSource>

    <Style TargetType="StatusBar">
      <Setter Property="Background" Value="{DynamicResource PanelBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
      <Setter Property="BorderThickness" Value="1,1,1,0"/>
      <Setter Property="Padding" Value="8,2"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="280"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Left command rail -->
    <Border Grid.Column="0" Grid.Row="0" Background="{DynamicResource PanelBrush}"
            BorderBrush="{DynamicResource BorderBrush}" BorderThickness="0,0,1,0">
      <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="16" x:Name="LeftStack">
          <TextBlock Text="Account" Style="{StaticResource SectionHeader}"/>
          <Button x:Name="BtnConnect" Content="Sign in / Load roles"/>
          <Button x:Name="BtnDisconnect" Content="Disconnect"/>

          <TextBlock Text="Activation" Style="{StaticResource SectionHeader}"/>
          <Button x:Name="BtnActivate" Content="Activate selected" IsEnabled="False"/>
          <Button x:Name="BtnDeactivateAll" Content="Deactivate ALL active" IsEnabled="False"/>

          <TextBlock Text="Quick Admin Portals" Style="{StaticResource SectionHeader}"/>
          <Button x:Name="BtnM365" Content="M365 Admin"/>
          <Button x:Name="BtnDefender" Content="Defender Admin"/>
          <Button x:Name="BtnIntune" Content="Intune Admin"/>
          <Button x:Name="BtnPurview" Content="Purview Admin"/>
          <Button x:Name="BtnEntra" Content="EntraID Admin"/>
          <Button x:Name="BtnEXO" Content="EXO Admin"/>
          <Button x:Name="BtnTeams" Content="Teams Admin"/>
          <TextBlock Text="Exit" Style="{StaticResource SectionHeader}"/>
          <Button x:Name="BtnExit" Content="EXIT"/>

        </StackPanel>
      </ScrollViewer>
    </Border>

    <!-- Right work area -->
    <Grid Grid.Column="1" Grid.Row="0">
      <Grid.RowDefinitions>
        <RowDefinition Height="3*"/>
        <RowDefinition Height="6"/>
        <RowDefinition Height="2*"/>
      </Grid.RowDefinitions>

      <!-- Roles list (grouped) -->
      <Border Grid.Row="0" Margin="16,0,16,8" BorderBrush="{DynamicResource BorderBrush}"
              BorderThickness="1" Background="{DynamicResource PanelBrush}">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <TextBlock Text="Eligible roles" Margin="8" Foreground="{DynamicResource MutedTextBrush}"/>

          <ListView x:Name="RolesList" Grid.Row="1" Margin="8" BorderThickness="0"
                    ItemsSource="{Binding Source={StaticResource RolesView}}"
                    SelectionMode="Extended"
                    HorizontalContentAlignment="Stretch">

            <!-- Row container: click toggles; we sync highlight manually -->
            <ListView.ItemContainerStyle>
              <Style TargetType="ListViewItem">
                <Setter Property="Padding" Value="6"/>
                <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
                <Setter Property="BorderThickness" Value="0"/>
                <Style.Triggers>
                  <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FFEFEFF5"/>
                  </Trigger>
                  <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#DCE2F9"/>
                    <Setter Property="BorderThickness" Value="1"/>
                  </Trigger>
                  <!-- Prevent selecting roles that are already active -->
                  <DataTrigger Binding="{Binding IsActive}" Value="True">
                    <Setter Property="IsEnabled" Value="False"/>
                  </DataTrigger>
                </Style.Triggers>
              </Style>
            </ListView.ItemContainerStyle>

            <ListView.Resources>
              <!-- Single-line cell (no checkbox), keeps your green "active" visual -->
              <DataTemplate x:Key="RoleCell">
                <Grid Margin="0,3,0,3">
                  <TextBlock TextTrimming="CharacterEllipsis"
                             VerticalAlignment="Center"
                             Margin="6,0,0,0">
                    <TextBlock.Style>
                      <Style TargetType="TextBlock">
                        <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
                        <Style.Triggers>
                          <DataTrigger Binding="{Binding IsActive}" Value="True">
                            <Setter Property="Foreground" Value="{DynamicResource ActiveGreenBrush}"/>
                          </DataTrigger>
                        </Style.Triggers>
                      </Style>
                    </TextBlock.Style>
                    <TextBlock.Text>
                      <MultiBinding StringFormat="{}{0} - {1} - {2}">
                        <Binding Path="RoleName"/>
                        <Binding Path="ScopeDisplay"/>
                        <Binding Path="MemberType"/>
                      </MultiBinding>
                    </TextBlock.Text>
                  </TextBlock>
                </Grid>
              </DataTemplate>

              <!-- Group header (AU) -->
              <Style TargetType="{x:Type GroupItem}">
                <Setter Property="Template">
                  <Setter.Value>
                    <ControlTemplate TargetType="{x:Type GroupItem}">
                      <StackPanel>
                        <Border Background="#F5F6F8" BorderBrush="{DynamicResource BorderBrush}"
                                BorderThickness="0,1,0,0" Padding="6,4">
                          <TextBlock Text="{Binding Name, StringFormat=AU: {0}}"
                                     Foreground="{DynamicResource MutedTextBrush}" FontWeight="Bold"/>
                        </Border>
                        <ItemsPresenter/>
                      </StackPanel>
                    </ControlTemplate>
                  </Setter.Value>
                </Setter>
              </Style>
            </ListView.Resources>

            <ListView.ItemTemplate>
              <StaticResource ResourceKey="RoleCell"/>
            </ListView.ItemTemplate>
          </ListView>
        </Grid>
      </Border>

      <GridSplitter Grid.Row="1" Height="6" HorizontalAlignment="Stretch" Background="#FFEFEFEF" ShowsPreview="True"/>

      <!-- Activation options + log -->
      <Border Grid.Row="2" Margin="16,8,16,16" BorderBrush="{DynamicResource BorderBrush}"
              BorderThickness="1" Background="{DynamicResource PanelBrush}">
        <Grid Margin="8">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="2*"/>
            <ColumnDefinition Width="3*"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <StackPanel Grid.Column="0" Grid.Row="0">
            <TextBlock Text="Justification" Style="{StaticResource SectionHeader}"/>
            <TextBox x:Name="TxtJust" Text="Operational need" Height="32"/>
            <TextBlock Text="Duration (hours)" Style="{StaticResource SectionHeader}"/>
            <StackPanel Orientation="Horizontal">
              <TextBox x:Name="TxtDur" Width="80" Height="32" Text="2"/>
              <TextBlock Text="(1-8)" Margin="8,8,0,0" Foreground="{DynamicResource MutedTextBrush}"/>
            </StackPanel>
          </StackPanel>

          <StackPanel Grid.Column="1" Grid.Row="0">
            <TextBlock Text="Ticket number" Style="{StaticResource SectionHeader}"/>
            <TextBox x:Name="TxtTicketNum" Height="32"/>
            <TextBlock Text="Ticket system" Style="{StaticResource SectionHeader}"/>
            <TextBox x:Name="TxtTicketSys" Height="32"/>
          </StackPanel>

          <TextBox x:Name="TxtLog" Grid.ColumnSpan="2" Grid.Row="2" Margin="0,8,0,0"
                   TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"
                   IsReadOnly="True" Background="#FAFAFB" BorderThickness="0" FontFamily="Consolas" FontSize="12"/>
        </Grid>
      </Border>
    </Grid>

    <StatusBar Grid.ColumnSpan="2" Grid.Row="1" x:Name="StatusBar" Margin="0">
      <StatusBarItem HorizontalAlignment="Stretch">
        <TextBlock x:Name="StatusLeft" Text="Ready."/>
      </StatusBarItem>
      <StatusBarItem HorizontalAlignment="Right">
        <TextBlock x:Name="StatusRight" Text="--:--:--"/>
      </StatusBarItem>
    </StatusBar>
  </Grid>
</Window>
'@
#endregion

#region --- Parse XAML
try { $window = [Windows.Markup.XamlReader]::Parse($Xaml) }
catch {
    $sr = New-Object System.IO.StringReader($Xaml)
    try { $xr = [System.Xml.XmlReader]::Create($sr); $window = [Windows.Markup.XamlReader]::Load($xr) }
    finally { if ($xr) { $xr.Dispose() } ; if ($sr) { $sr.Dispose() } }
}
if (-not $window) { throw "Failed to parse window XAML." }

# Position top-left and ensure it's not maximized
$window.WindowStartupLocation = 'Manual'
$window.WindowState = 'Normal'
$window.Left = 0
$window.Top  = 0

try {
    $interop = [System.Windows.Interop.WindowInteropHelper]::new($window)
    $null = $interop.EnsureHandle()
    Enable-DarkTitleBar -Handle $interop.Handle
    Enable-Mica       -Handle $interop.Handle
} catch {}
#endregion

#region --- Grab controls
$BtnConnect      = $window.FindName('BtnConnect')
$BtnDisconnect   = $window.FindName('BtnDisconnect')
$BtnActivate     = $window.FindName('BtnActivate')
$BtnDeactivateAll= $window.FindName('BtnDeactivateAll')
$RolesList       = $window.FindName('RolesList')
$BtnM365         = $window.FindName('BtnM365')
$BtnDefender     = $window.FindName('BtnDefender')
$BtnIntune       = $window.FindName('BtnIntune')
$BtnPurview      = $window.FindName('BtnPurview')
$BtnEntra        = $window.FindName('BtnEntra')
$BtnEXO          = $window.FindName('BtnEXO')
$BtnTeams        = $window.FindName('BtnTeams')
$BtnExit         = $window.FindName('BtnExit')
$TxtJust         = $window.FindName('TxtJust')
$TxtDur          = $window.FindName('TxtDur')
$TxtTicketNum    = $window.FindName('TxtTicketNum')
$TxtTicketSys    = $window.FindName('TxtTicketSys')
$TxtLog          = $window.FindName('TxtLog')
$StatusLeft      = $window.FindName('StatusLeft')
$StatusRight     = $window.FindName('StatusRight')
#endregion

#region --- Data context & collection
$script:RoleItems = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$dataCtx = New-Object psobject -Property @{ RoleItems = $script:RoleItems }
$window.DataContext = $dataCtx
#endregion

#region --- Helpers (robust active matching + marking + view refresh)
function Update-Status([string]$msg) {
    try {
        if ($StatusLeft)  { $StatusLeft.Text  = $msg }
        $ts = (Get-Date).ToString('HH:mm:ss')
        if ($StatusRight) { $StatusRight.Text = $ts }
        if ($TxtLog)      { $TxtLog.AppendText("[$ts] $msg`r`n"); $TxtLog.ScrollToEnd() }
    } catch {}
}
# Open URL in default browser (reuse existing window/tab)
function Open-Url([string]$url) {
    try {
        $clean = $url.Trim()
        if ([string]::IsNullOrWhiteSpace($clean)) { return }
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $clean
        $psi.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($psi) | Out-Null
        Update-Status "Opened: $clean"
    } catch {
        [System.Windows.MessageBox]::Show("Could not open:`r`n$url`r`n$($_.Exception.Message)","Open link failed",'OK','Error') | Out-Null
    }
}
# Force UI to refresh bindings (PSCustomObject doesn't raise PropertyChanged)
function Refresh-RolesView {
    try {
        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($RolesList.ItemsSource)
        if ($view) { $view.Refresh(); return }
    } catch {}
    try { $RolesList.Items.Refresh() } catch {}
}
# Normalize scopeId for matching ("/" | "" | GUID | "/administrativeUnits/GUID")
function Get-RoleKey {
    param(
        [Parameter(Mandatory)][string]$RoleDefinitionId,
        [Parameter()][string]$DirectoryScopeId
    )
    $scope = $DirectoryScopeId
    if ([string]::IsNullOrWhiteSpace($scope)) { $scope = "/" }
    if ($scope -match '^/administrativeUnits/([0-9a-fA-F\-]+)$' -or
        $scope -match '^/AdministrativeUnits/([0-9a-fA-F\-]+)$') {
        $scope = $Matches[1]
    }
    return ($RoleDefinitionId + "`n" + $scope).ToLowerInvariant()
}
# Apply IsActive to each item in RoleItems and clear IsChecked on active items; refresh view
function Mark-ActiveRoles {
    param([Parameter()][object[]]$ActiveAssignments = @())
    try {
        $activeLookup = @{}
        foreach ($a in $ActiveAssignments) {
            $key = Get-RoleKey -RoleDefinitionId $a.RoleDefinitionId -DirectoryScopeId $a.DirectoryScopeId
            $activeLookup[$key] = $true
        }
        foreach ($item in $script:RoleItems) {
            $k = Get-RoleKey -RoleDefinitionId $item.RoleDefinitionId -DirectoryScopeId $item.DirectoryScopeId
            $wasActive = [bool]$item.IsActive
            $nowActive = [bool]$activeLookup[$k]
            if ($nowActive -ne $wasActive) { $item.IsActive = $nowActive }
            if ($nowActive) { $item.IsChecked = $false }
        }
    } catch {
        Update-Status ("WARN: Mark-ActiveRoles failed: {0}" -f $_.Exception.Message)
    } finally {
        Refresh-RolesView
    }
}
#endregion

#region --- Row click toggling & Spacebar selection
# Toggle selection by clicking anywhere on the line; respect Ctrl/Shift default behaviors
$RolesList.Add_PreviewMouseLeftButtonDown({
    param($sender,$e)
    try {
        # If user holds Ctrl/Shift, let WPF handle the selection (range/multi-select)
        if ([System.Windows.Input.Keyboard]::Modifiers -ne [System.Windows.Input.ModifierKeys]::None) { return }
        $dep = [System.Windows.DependencyObject]$e.OriginalSource
        $container = [System.Windows.Controls.ItemsControl]::ContainerFromElement($RolesList, $dep)
        if ($container -and ($container -is [System.Windows.Controls.ListViewItem])) {
            $item = $container.DataContext
            if ($item -and -not [bool]$item.IsActive) {
                # Toggle check
                $item.IsChecked = -not [bool]$item.IsChecked
                # Mirror selection highlight to the new state
                $container.IsSelected = [bool]$item.IsChecked

                Refresh-RolesView
                $e.Handled = $true
                try {
                    $state = if ($item.IsChecked) { "Selected" } else { "Deselected" }
                    Update-Status ("{0}: {1}" -f $state, $item.RoleName)
                } catch {}
            }
        }
    } catch {}
})

# Allow Spacebar to toggle selection on the current/focused row
$RolesList.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq 'Space') {
        try {
            $view    = [System.Windows.Data.CollectionViewSource]::GetDefaultView($RolesList.ItemsSource)
            $current = $view.CurrentItem
            if (-not $current) {
                if ($RolesList.Items.Count -gt 0) { $current = $RolesList.Items[0] } else { return }
            }
            if ([bool]$current.IsActive) { return } # skip active items

            # Find container to update IsSelected explicitly
            $container = $RolesList.ItemContainerGenerator.ContainerFromItem($current)
            # Toggle check
            $current.IsChecked = -not [bool]$current.IsChecked
            # Mirror selection highlight
            if ($container -is [System.Windows.Controls.ListViewItem]) {
                $container.IsSelected = [bool]$current.IsChecked
            }

            Refresh-RolesView
            $e.Handled = $true
            try {
                $state = if ($current.IsChecked) { "Selected" } else { "Deselected" }
                Update-Status ("{0}: {1}" -f $state, $current.RoleName)
            } catch {}
        } catch {}
    }
})

# Keep IsChecked in sync when selection changes via keyboard/mouse (e.g., Shift/Ctrl)
$RolesList.Add_SelectionChanged({
    param($sender,$e)
    try {
        foreach ($added in $e.AddedItems) {
            if ($added -and -not [bool]$added.IsActive) { $added.IsChecked = $true }
        }
        foreach ($removed in $e.RemovedItems) {
            if ($removed) { $removed.IsChecked = $false }
        }
        Refresh-RolesView
    } catch {}
})
#endregion

#region --- Sign in / load roles (synchronous active check)
$script:SignedInUser = $null
$BtnConnect.Add_Click({
    try {
        Update-Status "Connecting (read scopes)…"
        Connect-GraphForRead
        $me = Get-CurrentUser
        $script:SignedInUser = $me
        $mail = if ($me.mail) { $me.mail } else { $me.userPrincipalName }
        Update-Status ("Signed in as: {0} <{1}>" -f $me.displayName, $mail)

        Update-Status "Fetching eligible roles…"
        $script:RoleItems.Clear()
        $roles = Get-PIMEligibleDirectoryRoles -UserId $me.id
        foreach ($r in $roles) { $script:RoleItems.Add($r) | Out-Null }

        if ($script:RoleItems.Count -eq 0) {
            Update-Status "No eligible roles found."
            $BtnActivate.IsEnabled = $false
            $BtnDeactivateAll.IsEnabled = $true
        }
        else {
            Update-Status ("Loaded {0} eligible roles." -f $script:RoleItems.Count)
            $BtnActivate.IsEnabled = $true
            $BtnDeactivateAll.IsEnabled = $true
        }

        # --- Synchronous active check in the same runspace ---
        Update-Status "Checking active assignments…"
        $active = @()
        try {
            $active = Get-ActivePIMAssignmentsForUser -UserId $me.id
        } catch {
            Update-Status ("WARN: Active check failed: {0}" -f $_.Exception.Message)
            $active = @()
        }

        # Mark already-active roles (green & disabled row)
        try {
            Mark-ActiveRoles -ActiveAssignments $active
            Refresh-RolesView
            Update-Status "Active state refreshed."
        } catch {
            Update-Status ("WARN: marking active roles failed: {0}" -f $_.Exception.Message)
        }
    }
    catch {
        Update-Status "ERROR: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show(
            "Failed to connect or load roles.`r`n$($_.Exception.Message)",
            "Error",
            'OK',
            'Error'
        ) | Out-Null
    }
})
#endregion

#region --- Disconnect
$BtnDisconnect.Add_Click({
    try {
        Update-Status "Disconnecting from Microsoft Graph…"
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $script:SignedInUser = $null
        $script:RoleItems.Clear()
        $script:AuDisplayCache = @{}
        Refresh-RolesView
        Update-Status "Disconnected. State cleared."
        $BtnActivate.IsEnabled = $false
        $BtnDeactivateAll.IsEnabled = $false
    } catch {
        Update-Status "WARN: Disconnect issue: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Disconnect completed with warnings:`r`n$($_.Exception.Message)","Warning",'OK','Warning') | Out-Null
    }
})
#endregion

#region --- Activate selected (refresh active flags after run; optimistic green)
$BtnActivate.Add_Click({
    if (-not $script:SignedInUser) {
        [System.Windows.MessageBox]::Show("Please sign in first.","Info",'OK','Information') | Out-Null
        return
    }
    try {
        Update-Status "Requesting write scope…"
        Connect-GraphForWrite
    } catch {
        Update-Status "Admin consent required for write scope."
        [System.Windows.MessageBox]::Show(
            "Admin consent required for RoleManagement.ReadWrite.Directory / RoleAssignmentSchedule.ReadWrite.Directory.",
            "Admin consent needed",'OK','Warning') | Out-Null
        return
    }

    $toActivate = @()
    foreach ($item in $script:RoleItems) { if ($item.IsChecked -and -not $item.IsActive) { $toActivate += $item } }
    if ($toActivate.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Select at least one non-active role by clicking its line or using Space.","Info",'OK','Information') | Out-Null
        return
    }

    $hours = 2
    try {
        $hStr = $TxtDur.Text.Trim()
        if ($hStr -match '^\d+$') { $hours = [int]$hStr }
        if ($hours -lt 1) { $hours = 1 } elseif ($hours -gt 8) { $hours = 8 }
    } catch {}
    $duration = [TimeSpan]::FromHours($hours)
    $just = if ([string]::IsNullOrWhiteSpace($TxtJust.Text)) { "Operational need" } else { $TxtJust.Text }
    $ticketNumber = $TxtTicketNum.Text
    $ticketSystem = $TxtTicketSys.Text

    try {
        foreach ($role in $toActivate) {
            $label = if ($role.ScopeDisplay -and $role.ScopeDisplay -ne "(Tenant wide)") { "$( $role.RoleName ) — $( $role.ScopeDisplay )" } else { $role.RoleName }
            Update-Status "Activating: $label …"

            $res = Request-PIMActivation -UserId $script:SignedInUser.id `
                                         -RoleDefinitionId $role.RoleDefinitionId `
                                         -DirectoryScopeId $role.DirectoryScopeId `
                                         -Justification $just `
                                         -Duration $duration `
                                         -TicketNumber $ticketNumber `
                                         -TicketSystem $ticketSystem
            switch ($res.Status) {
                "Granted"        {
                    Update-Status "✓ Activated: $label (RequestId: $($res.RequestId))"
                    # Optimistically mark green now
                    try { $role.IsActive = $true; $role.IsChecked = $false; Refresh-RolesView } catch {}
                }
                "PendingApproval" { Update-Status "⏳ Pending approval: $label (RequestId: $($res.RequestId))" }
                "NotStarted"      { Update-Status "⌛ Submitted: $label (RequestId: $($res.RequestId))" }
                default           { Update-Status "⚠️ $($res.Status): $label — $($res.Message)" }
            }
        }

        # Confirm with fresh active list
        try {
            $activeNow = Get-ActivePIMAssignmentsForUser -UserId $script:SignedInUser.id
            Mark-ActiveRoles -ActiveAssignments $activeNow
            Refresh-RolesView
            Update-Status "Active state refreshed."
        } catch {
            Update-Status ("WARN: Could not refresh active state: {0}" -f $_.Exception.Message)
        }

        [System.Windows.MessageBox]::Show("Activation requests submitted.","Completed",'OK','Information') | Out-Null
    } catch {
        Update-Status "ERROR during activation: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Activation errors occurred.`r`n$($_.Exception.Message)","Error",'OK','Error') | Out-Null
    }
})
#endregion

#region --- Deactivate ALL active (refresh active flags after run)
$BtnDeactivateAll.Add_Click({
    if (-not $script:SignedInUser) {
        [System.Windows.MessageBox]::Show("Please sign in first.","Info",'OK','Information') | Out-Null
        return
    }
    try {
        Update-Status "Requesting write scope…"
        Connect-GraphForWrite
    } catch {
        Update-Status "Admin consent required for write scope."
        [System.Windows.MessageBox]::Show(
            "Admin consent required for RoleManagement.ReadWrite.Directory / RoleAssignmentSchedule.ReadWrite.Directory.",
            "Admin consent needed",'OK','Warning') | Out-Null
        return
    }

    try {
        Update-Status "Loading active PIM assignments…"
        $active = Get-ActivePIMAssignmentsForUser -UserId $script:SignedInUser.id
        if ($active.Count -eq 0) {
            Update-Status "No active PIM assignments found."
            [System.Windows.MessageBox]::Show("No active roles to deactivate.","Info",'OK','Information') | Out-Null
            return
        }

        foreach ($a in $active) {
            $label = if ($a.ScopeDisplay -and $a.ScopeDisplay -ne "(Tenant wide)") { "$( $a.RoleName ) — $( $a.ScopeDisplay )" } else { $a.RoleName }
            Update-Status "Deactivating: $label …"
            $res = Request-PIMDeactivation -UserId $script:SignedInUser.id `
                                           -RoleDefinitionId $a.RoleDefinitionId `
                                           -DirectoryScopeId $a.DirectoryScopeId `
                                           -Justification "Bulk deactivation from GUI"
            switch ($res.Status) {
                "Granted"        { Update-Status "✓ Deactivated: $label (RequestId: $($res.RequestId))" }
                "PendingApproval"{ Update-Status "⏳ Pending approval for deactivation: $label (RequestId: $($res.RequestId))" }
                "NotStarted"     { Update-Status "⌛ Deactivation submitted: $label (RequestId: $($res.RequestId))" }
                default          { Update-Status "⚠️ $($res.Status): $label — $($res.Message)" }
            }
        }

        # Refresh active flags right after deactivation
        try {
            $activeNow = Get-ActivePIMAssignmentsForUser -UserId $script:SignedInUser.id
            Mark-ActiveRoles -ActiveAssignments $activeNow
            foreach ($item in $script:RoleItems) { $item.IsChecked = $false }
            Refresh-RolesView
            Update-Status "Active state refreshed."
        } catch {
            Update-Status ("WARN: Could not refresh active state: {0}" -f $_.Exception.Message)
        }

        [System.Windows.MessageBox]::Show("Deactivation requests submitted.","Completed",'OK','Information') | Out-Null
    } catch {
        Update-Status "ERROR during deactivation: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Deactivation errors occurred.`r`n$($_.Exception.Message)","Error",'OK','Error') | Out-Null
    }
})

#region --- Exit button
$BtnExit.Add_Click({ $window.Close() })
#endregion

#endregion

#region --- Quick Admin links
$BtnM365.Add_Click({    Open-Url "https://admin.cloud.microsoft/" })
$BtnDefender.Add_Click({Open-Url "https://security.microsoft.com/homepage" })
$BtnIntune.Add_Click({  Open-Url "https://intune.microsoft.com/?ref=AdminCenter#home" })
$BtnPurview.Add_Click({ Open-Url "https://purview.microsoft.com/?rfr=AdminCenter" })
$BtnEntra.Add_Click({   Open-Url "https://entra.microsoft.com/" })
$BtnEXO.Add_Click({     Open-Url "https://admin.exchange.microsoft.com/" })
$BtnTeams.Add_Click({   Open-Url "https://admin.teams.microsoft.com/" })
#endregion

#region --- Clock & show
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(800)
$timer.add_Tick({ if ($StatusRight) { $StatusRight.Text = (Get-Date).ToString('HH:mm:ss') } })
$timer.Start()
Update-Status "Ready."
$null = $window.ShowDialog()
#endregion
