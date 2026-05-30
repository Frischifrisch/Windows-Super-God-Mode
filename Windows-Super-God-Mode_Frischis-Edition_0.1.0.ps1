# ======================================================================================================
# Skript zur Erstellung von Shell-Ordnerverknüpfungen (Aka "Super God Mode")
# Original-Autor: ThioJoe (https://github.com/ThioJoe/Windows-Super-God-Mode)
# 
# OPTIMIERT & ERWEITERT FÜR CHRISTIAN (Multithreading, COM-Performance & Clean-UI)
# Übersetzt & Ausführlich kommentiert von T3 Chat (Dein Nerd-Buddy)
# ======================================================================================================

param(
    # Alternative Optionen
    [switch]$DontGroupTasks,
    [switch]$UseAlternativeCategoryNames,
    [switch]$AllURLProtocols,
    [switch]$CollectExtraURLProtocolInfo,
    [switch]$AllowDuplicateDeepLinks,
    [switch]$DeepScanHiddenLinks,
    # Ausgabe-Steuerung
    [string]$Output,
    [switch]$KeepPreviousOutputFolders,
    # Filter-Argumente (Ausgaben einschränken)
    [switch]$NoStatistics,
    [switch]$NoReadMe,
    [switch]$SkipCLSID,
    [switch]$SkipNamedFolders,
    [switch]$SkipTaskLinks,
    [switch]$SkipMSSettings,
    [switch]$SkipDeepLinks,
    [switch]$SkipURLProtocols,
    [switch]$SkipHiddenAppLinks,
    # Debugging & Diagnose
    [switch]$Verbose,
    [switch]$Debug,
    [switch]$timing,
    [switch]$debugSkipAppxSearch,
    [string]$debugSearchOnlyProtocolList,
    [switch]$uniqueOutputFolder,
    # Erweiterte Argumente
    [switch]$NoGUI,
    [string]$CustomDLLPath,
    [string]$CustomLanguageFolderPath,
    [string]$CustomSystemSettingsDLLPath,
    [string]$CustomAllSystemSettingsXMLPath
)

$VERSION = "1.2.3-Christian-Optimized"

# ==============================================================================================================================
# =============================================  [KOMPATIBILITÄTS-CHECK]  ======================================================
# ==============================================================================================================================
# LERNEFFEKT: Da wir für das Multithreading das PowerShell 7-exklusive Feature '-Parallel' nutzen,
# müssen wir sicherstellen, dass Christian das Skript nicht aus Versehen in der alten Windows PowerShell 5.1 startet.
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "Dieses optimierte Skript benötigt PowerShell 7 (Core). Bitte installiere/nutze PowerShell 7!"
    Write-Warning "Alternativ kannst du 'ForEach-Object -Parallel' durch ein normales 'ForEach-Object' ersetzen, verlierst aber den Speed-Boost."
    exit
}

# ==============================================================================================================================
# ==========================  [OPTIMIERUNG 1: GLOBALES COM-OBJEKT FÜR PERFORMANTE VERKNÜPFUNGEN]  ==============================
# ==============================================================================================================================
# LERNEFFEKT: Die Erstellung eines COM-Objekts per 'New-Object -ComObject WScript.Shell' zwingt Windows dazu, im Hintergrund
# die C++ basierte Windows-Scriptinghost-Laufzeitumgebung in den RAM zu laden. 
# Das bei jeder einzelnen Verknüpfung in einer Schleife zu tun, erzeugt enormen Overhead.
#
# WIR MACHEN DAS JETZT SMART & FAUL: 
# Wir erstellen EIN einziges, globales COM-Objekt und nutzen es überall im Skript. Am Ende des Skripts entladen wir es einmal sauber.
Write-Verbose "[PERFORMANCE] Erzeuge globalen Windows Script Host für schnellen Zugriff..."
$GlobalShell = New-Object -ComObject WScript.Shell

# ==============================================================================================================================
# =========================================  FUNKTION: GUI-DIALOG (WPF/XAML)  ==================================================
# ==============================================================================================================================
function Show-SuperGodModeDialog {
    param(
        [string]$defaultOutputFolderName,
        [switch]$initialDebug,
        [switch]$initialVerbose
    )

    $tooltips = @{
        DontGroupTasks = "Verhindert die Gruppierung von Aufgabenverknüpfungen (der App-Name wird nicht vorangestellt)."
        UseAlternativeCategoryNames = "Sucht nach alternativen Kategorienamen für Aufgabenlinks."
        AllURLProtocols = "Schließt auch Drittanbieter-URL-Protokolle ein, nicht nur Microsoft-Systemprotokolle."
        CollectExtraURLProtocolInfo = "Sammelt zusätzliche Detail-Informationen über URL-Protokolle für die CSV-Statistik."
        KeepPreviousOutputFolders = "Löscht den alten Ausgabeordner vor dem Start nicht."
        CollectStatistics = "Erstellt statistische Berichte (CSV/XML) über gefundene Elemente im Analyseordner."
        AllowDuplicateDeepLinks = "Erlaubt doppelte Deep-Links, selbst wenn sie bereits als Aufgaben-Link existieren."
        CollectCLSID = "Erstellt Verknüpfungen für System-IDs (CLSIDs) wie Netzwerke, Drucker, etc."
        CollectNamedFolders = "Erstellt Verknüpfungen zu namentlich bekannten Sonderordnern (z. B. Downloads)."
        CollectTaskLinks = "Erstellt Shortcuts zu spezifischen Untermenüs der Systemsteuerung."
        CollectMSSettings = "Erstellt Verknüpfungen zu den modernen Windows 10/11 'ms-settings:' Seiten."
        CollectDeepLinks = "Erstellt direkte Verknüpfungen zu tief verschachtelten Einstellungen."
        CollectURLProtocols = "Erstellt Verknüpfungen für registrierte URL-Protokolle (z. B. mailto:)."
        CollectAppxLinks = "Sucht nach versteckten App-Sublinks (z. B. ms-clock://pausefocustimer)."
        DeepScanHiddenLinks = "Scannt ALLE Dateien im Installationspfad von Apps. Sehr präzise, aber weitaus langsamer!"
        LoggingLevel = "Detailgrad der Log-Ausgaben im PowerShell-Fenster."
    }

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms

    [xml]$xaml = @"
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Super God Mode Parameter" Height="715" Width="800">
        <Window.Resources>
            <Color x:Key="BackgroundColor">#1E1E1E</Color>
            <Color x:Key="ForegroundColor">#CCCCCC</Color>
            <Color x:Key="AccentColor">#0078D4</Color>
            <Color x:Key="SecondaryBackgroundColor">#2D2D2D</Color>
            <Color x:Key="BorderColor">#3F3F3F</Color>
            <Color x:Key="WarningColor">#FF6B68</Color>
            <Color x:Key="VersionColor">#888888</Color>
            <Color x:Key="ButtonHoverColor">#1b99fa</Color>
            <Color x:Key="HighlightedTextColor">#69d2ff</Color>

            <SolidColorBrush x:Key="BackgroundBrush" Color="{StaticResource BackgroundColor}"/>
            <SolidColorBrush x:Key="ForegroundBrush" Color="{StaticResource ForegroundColor}"/>
            <SolidColorBrush x:Key="AccentBrush" Color="{StaticResource AccentColor}"/>
            <SolidColorBrush x:Key="SecondaryBackgroundBrush" Color="{StaticResource SecondaryBackgroundColor}"/>
            <SolidColorBrush x:Key="BorderBrush" Color="{StaticResource BorderColor}"/>
            <SolidColorBrush x:Key="WarningBrush" Color="{StaticResource WarningColor}"/>
            <SolidColorBrush x:Key="VersionBrush" Color="{StaticResource VersionColor}"/>
            <SolidColorBrush x:Key="ButtonHoverBrush" Color="{StaticResource ButtonHoverColor}"/>
            <SolidColorBrush x:Key="HighlightedTextBrush" Color="{StaticResource HighlightedTextColor}"/>
            
            <Thickness x:Key="BorderThickness">1</Thickness>
            <Thickness x:Key="GroupBoxPadding">5</Thickness>

            <Style x:Key="DarkModeGroupBoxStyle" TargetType="GroupBox">
                <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
                <Setter Property="BorderThickness" Value="{StaticResource BorderThickness}"/>
                <Setter Property="Padding" Value="{StaticResource GroupBoxPadding}"/>
                <Setter Property="Margin" Value="0,10,0,10"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="GroupBox">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <Border BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Grid.Row="0" Grid.RowSpan="2"/>
                                <Border Background="{TemplateBinding Background}" BorderBrush="Transparent" BorderThickness="{TemplateBinding BorderThickness}" Grid.Row="1">
                                    <ContentPresenter Margin="{TemplateBinding Padding}"/>
                                </Border>
                                <TextBlock Margin="5,0,0,0" Padding="3,0,3,0" Background="{StaticResource BackgroundBrush}" HorizontalAlignment="Left" VerticalAlignment="Top" TextElement.Foreground="{StaticResource ForegroundBrush}" Text="{TemplateBinding Header}"/>
                            </Grid>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
            </Style>

            <Style x:Key="SubtleButtonStyle" TargetType="Button">
                <Setter Property="Background" Value="Transparent"/>
                <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
                <Setter Property="BorderThickness" Value="0"/>
                <Setter Property="FontSize" Value="12"/>
                <Setter Property="Cursor" Value="Hand"/>
                <Setter Property="Padding" Value="10,5"/>
                <Setter Property="Template">
                    <Setter.Value>
                        <ControlTemplate TargetType="Button">
                            <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                            </Border>
                        </ControlTemplate>
                    </Setter.Value>
                </Setter>
                <Style.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                        <Setter Property="Background" Value="{StaticResource ButtonHoverBrush}"/>
                    </Trigger>
                </Style.Triggers>
            </Style>

            <Style x:Key="CustomCheckBoxStyle" TargetType="CheckBox" BasedOn="{StaticResource {x:Type CheckBox}}">
                <Style.Triggers>
                    <Trigger Property="IsEnabled" Value="False">
                        <Setter Property="Opacity" Value="0.5"/>
                    </Trigger>
                </Style.Triggers>
            </Style>

            <Style x:Key="HighlightedCheckBoxStyle" TargetType="CheckBox" BasedOn="{StaticResource {x:Type CheckBox}}">
                <Setter Property="Foreground" Value="{StaticResource HighlightedTextBrush}"/>
                <Setter Property="FontWeight" Value="Semibold"/>
            </Style>

        </Window.Resources>
        <Grid Background="{StaticResource BackgroundBrush}">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <Border Background="{StaticResource AccentBrush}" Grid.Row="0">
                <Grid>
                    <StackPanel>
                        <TextBlock Text="&quot;Super God Mode&quot; Skript" FontSize="24" Foreground="White" HorizontalAlignment="Center" Margin="0,10,0,0"/>
                        <TextBlock Text="Optimierung &amp; Analyse für Windows" FontSize="16" Foreground="White" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                    </StackPanel>
                    <Button x:Name="btnAbout" Content="Über uns" Style="{StaticResource SubtleButtonStyle}"
                            HorizontalAlignment="Right" VerticalAlignment="Top" Margin="0,10,10,0"/>
                </Grid>
            </Border>

            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Text="Bewege die Maus über Optionen für Details" FontStyle="Italic" HorizontalAlignment="Right" Margin="0,0,0,10" Grid.Row="0" Foreground="{StaticResource ForegroundBrush}"/>

                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <GroupBox Header="Alternative Einstellungen" Grid.Column="0" Style="{StaticResource DarkModeGroupBoxStyle}">
                            <StackPanel Margin="5">
                                <CheckBox x:Name="chkDontGroupTasks" Content="Aufgaben:   Aufgaben nicht sortieren/gruppieren" Margin="0,5,0,0" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.DontGroupTasks)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkUseAlternativeCategoryNames" Content="Aufgaben:   Alternative Kategorienamen nutzen" Margin="0,5,0,0" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.UseAlternativeCategoryNames)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                               <CheckBox x:Name="chkAllowDuplicateDeepLinks" Content="Deep Links: Erlaube doppelte Einträge" Margin="0,5,0,0" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.AllowDuplicateDeepLinks)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkAllURLProtocols" Content="Protokolle:  Dritthersteller-Protokolle einbeziehen" Margin="0,5,0,0" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.AllURLProtocols)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkCollectExtraURLProtocolInfo" Content="Protokolle:  Erweiterte Zusatz-Infos sammeln" Margin="0,5,0,0" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectExtraURLProtocolInfo)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkDeepScanHiddenLinks" Content="Protokolle:  Deep-Scan nach versteckten Links (langsam)" Margin="0,5,0,0" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.DeepScanHiddenLinks)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                            </StackPanel>
                        </GroupBox>

                        <GroupBox Header="Ausgaben steuern" Grid.Column="1" Style="{StaticResource DarkModeGroupBoxStyle}">
                            <Grid Margin="5">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <CheckBox x:Name="chkCollectStatistics" Content="Statistiken erstellen  &#128202;" IsChecked="True" Margin="0,5,5,5" Grid.Column="0" Grid.Row="0" Style="{StaticResource HighlightedCheckBoxStyle}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectStatistics)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkCollectCLSID" Content="CLSID-Links" IsChecked="True" Margin="5,5,0,5" Grid.Column="1" Grid.Row="0" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectCLSID)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkCollectNamedFolders" Content="Benannte Ordner" IsChecked="True" Margin="0,5,5,5" Grid.Column="0" Grid.Row="1" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectNamedFolders)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkCollectTaskLinks" Content="Aufgaben-Links" IsChecked="True" Margin="5,5,0,5" Grid.Column="1" Grid.Row="1" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectTaskLinks)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkCollectMSSettings" Content="MS-Settings Links" IsChecked="True" Margin="0,5,5,5" Grid.Column="0" Grid.Row="2" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectMSSettings)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkCollectDeepLinks" Content="Deep Links" IsChecked="True" Margin="5,5,0,5" Grid.Column="1" Grid.Row="2" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectDeepLinks)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkCollectURLProtocols" Content="URL-Protokolle" IsChecked="True" Margin="0,5,5,5" Grid.Column="0" Grid.Row="3" Foreground="{StaticResource ForegroundBrush}">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectURLProtocols)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                                <CheckBox x:Name="chkCollectAppxLinks" Content="Versteckte App-Links" IsChecked="True" Margin="5,5,0,5" Grid.Column="1" Grid.Row="3" Foreground="{StaticResource ForegroundBrush}" Style="{StaticResource CustomCheckBoxStyle}" ToolTipService.ShowOnDisabled="True">
                                    <CheckBox.ToolTip>
                                        <ToolTip Content="$($tooltips.CollectAppxLinks)" />
                                    </CheckBox.ToolTip>
                                </CheckBox>
                            </Grid>
                        </GroupBox>
                    </Grid>

                    <GroupBox Header="Speicherort der Ergebnisse" Grid.Row="2" Style="{StaticResource DarkModeGroupBoxStyle}">
                        <StackPanel Margin="5">
                            <CheckBox x:Name="chkKeepPreviousOutputFolders" Content="Alte Ausgabeordner behalten (nicht vorab löschen)" Margin="0,5,0,0" Foreground="{StaticResource ForegroundBrush}">
                                <CheckBox.ToolTip>
                                    <ToolTip Content="$($tooltips.KeepPreviousOutputFolders)" />
                                </CheckBox.ToolTip>
                            </CheckBox>
                            <TextBlock Text="Basisverzeichnis:" Margin="0,10,0,5" Foreground="{StaticResource ForegroundBrush}"/>
                            <DockPanel LastChildFill="True" Margin="0,0,0,5">
                                <Button x:Name="btnBrowse" Content="Durchsuchen..." DockPanel.Dock="Right" Margin="5,0,0,0" Padding="10,5" FontSize="14" MinWidth="120" Background="{StaticResource SecondaryBackgroundBrush}" Foreground="{StaticResource ForegroundBrush}"/>
                                <TextBox x:Name="txtOutputPath" IsReadOnly="True" Padding="5,0,0,0" VerticalContentAlignment="Center" FontSize="14" Height="30" Background="{StaticResource SecondaryBackgroundBrush}" Foreground="{StaticResource ForegroundBrush}"/>
                            </DockPanel>
                            <TextBlock Text="Name des Ausgabeordners:" Margin="0,5,0,5" Foreground="{StaticResource ForegroundBrush}"/>
                            <TextBox x:Name="txtOutputFolderName" Margin="0,0,0,5" Padding="5,0,0,0" VerticalContentAlignment="Center" FontSize="14" Height="30" Background="{StaticResource SecondaryBackgroundBrush}" Foreground="{StaticResource ForegroundBrush}"/>
                            <Separator Margin="0,10,0,10" Background="{StaticResource BorderBrush}"/>
                            <TextBlock Text="Letzter Speicherpfad:" Margin="0,5,0,5" FontWeight="Bold" Foreground="{StaticResource ForegroundBrush}"/>
                            <TextBlock x:Name="txtCurrentPath" Text="" Margin="0,0,0,10" TextWrapping="Wrap" FontWeight="Bold" Foreground="{StaticResource ForegroundBrush}"/>
                        </StackPanel>
                    </GroupBox>

                    <StackPanel Grid.Row="3">
                        <Button x:Name="btnOK" Content="Verknüpfungen erstellen" Width="Auto" Height="Auto" FontSize="14" HorizontalAlignment="Center" Margin="0,10,0,10" Padding="15,7" Background="{StaticResource AccentBrush}" Foreground="White"/>
                        <TextBlock Text="Hinweis: Die Standards sind optimal konfiguriert" FontWeight="Bold" Foreground="{StaticResource WarningBrush}" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center">
                                <TextBlock Text="Konsolen-Log:" Margin="0,0,10,0" VerticalAlignment="Center" Foreground="{StaticResource ForegroundBrush}">
                                    <TextBlock.ToolTip>
                                        <ToolTip Content="$($tooltips.LoggingLevel)" />
                                    </TextBlock.ToolTip>
                                </TextBlock>
                                <ComboBox x:Name="cmbLoggingLevel" Width="Auto" MinWidth="100" SelectedIndex="0">
                                    <ComboBoxItem Content="Standard"/>
                                    <ComboBoxItem Content="Ausführlich"/>
                                    <ComboBoxItem Content="Debug"/>
                                    <ComboBox.ToolTip>
                                        <ToolTip Content="$($tooltips.LoggingLevel)" />
                                    </ComboBox.ToolTip>
                                </ComboBox>
                            </StackPanel>
                        </Grid>
                    </StackPanel>
                </Grid>
            </ScrollViewer>

            <TextBlock x:Name="txtVersion" Text="Version: $VERSION" Grid.Row="1" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,10,5" FontSize="12" Foreground="{StaticResource VersionBrush}"/>
        </Grid>
    </Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Bindungen für UI-Komponenten
    $chkDontGroupTasks = $window.FindName("chkDontGroupTasks")
    $chkUseAlternativeCategoryNames = $window.FindName("chkUseAlternativeCategoryNames")
    $chkAllURLProtocols = $window.FindName("chkAllURLProtocols")
    $chkCollectExtraURLProtocolInfo = $window.FindName("chkCollectExtraURLProtocolInfo")
    $chkKeepPreviousOutputFolders = $window.FindName("chkKeepPreviousOutputFolders")
    $chkAllowDuplicateDeepLinks = $window.FindName("chkAllowDuplicateDeepLinks")
    $chkDeepScanHiddenLinks = $window.FindName("chkDeepScanHiddenLinks")

    $chkCollectStatistics = $window.FindName("chkCollectStatistics")
    $chkCollectCLSID = $window.FindName("chkCollectCLSID")
    $chkCollectNamedFolders = $window.FindName("chkCollectNamedFolders")
    $chkCollectTaskLinks = $window.FindName("chkCollectTaskLinks")
    $chkCollectMSSettings = $window.FindName("chkCollectMSSettings")
    $chkCollectDeepLinks = $window.FindName("chkCollectDeepLinks")
    $chkCollectURLProtocols = $window.FindName("chkCollectURLProtocols")
    $chkCollectAppxLinks = $window.FindName("chkCollectAppxLinks")

    $txtOutputPath = $window.FindName("txtOutputPath")
    $txtCurrentPath = $window.FindName("txtCurrentPath")
    $txtOutputFolderName = $window.FindName("txtOutputFolderName")
    $btnBrowse = $window.FindName("btnBrowse")
    $btnOK = $window.FindName("btnOK")

    $cmbLoggingLevel = $window.FindName("cmbLoggingLevel")
    if ($initialDebug) {
        $cmbLoggingLevel.SelectedIndex = 2
    } elseif ($initialVerbose) {
        $cmbLoggingLevel.SelectedIndex = 1
    } else {
        $cmbLoggingLevel.SelectedIndex = 0
    }

    $defaultOutputPath = $PSScriptRoot
    $txtOutputPath.Text = $defaultOutputPath
    $txtOutputFolderName.Text = $defaultOutputFolderName
    $txtCurrentPath.Text = [System.IO.Path]::Combine($defaultOutputPath, $defaultOutputFolderName)

    $btnBrowse.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $folderBrowser.ValidateNames = $false
        $folderBrowser.CheckFileExists = $false
        $folderBrowser.CheckPathExists = $true
        $folderBrowser.FileName = "Ordnerauswahl"
        $folderBrowser.Title = "Wähle das Ausgabe-Verzeichnis"
        $folderBrowser.InitialDirectory = $txtOutputPath.Text
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedPath = Split-Path $folderBrowser.FileName
            $txtOutputPath.Text = $selectedPath
            UpdateCurrentPath
        }
    })

    $chkCollectURLProtocols.Add_Checked({
        $chkCollectAppxLinks.IsEnabled = $true
    })

    $chkCollectURLProtocols.Add_Unchecked({
        $chkCollectAppxLinks.IsChecked = $false
        $chkCollectAppxLinks.IsEnabled = $false
    })

    $chkCollectAppxLinks.IsEnabled = $chkCollectURLProtocols.IsChecked

    $txtOutputPath.Add_TextChanged({ UpdateCurrentPath })
    $txtOutputFolderName.Add_TextChanged({ UpdateCurrentPath })

    $btnAbout = $window.FindName("btnAbout")
    $btnAbout.Add_Click({
        [System.Windows.MessageBox]::Show("    `"Super God Mode`" Skript für Windows`n`nVersion: $VERSION`nEntwickler: ThioJoe`n`nQuellcode:`nhttps://github.com/ThioJoe/Windows-Super-God-Mode",
            "Über das Skript",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::None
        )
    })

    function UpdateCurrentPath {
        $outputPath = $txtOutputPath.Text
        $folderName = $txtOutputFolderName.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($folderName)) {
            $txtCurrentPath.Text = $outputPath
        } else {
            $txtCurrentPath.Text = [System.IO.Path]::Combine($outputPath, $folderName)
        }
    }

    $btnOK.Add_Click({
        $window.DialogResult = $true
        $window.Close()
    })

    $result = $window.ShowDialog()

    if (-not $result) {
        return $null
    }

    return @{
        DontGroupTasks = $chkDontGroupTasks.IsChecked
        UseAlternativeCategoryNames = $chkUseAlternativeCategoryNames.IsChecked
        AllURLProtocols = $chkAllURLProtocols.IsChecked
        CollectExtraURLProtocolInfo = $chkCollectExtraURLProtocolInfo.IsChecked
        KeepPreviousOutputFolders = $chkKeepPreviousOutputFolders.IsChecked
        NoStatistics = !$chkCollectStatistics.IsChecked
        AllowDuplicateDeepLinks = $chkAllowDuplicateDeepLinks.IsChecked
        DeepScanHiddenLinks = $chkDeepScanHiddenLinks.IsChecked
        SkipCLSID = !$chkCollectCLSID.IsChecked
        SkipNamedFolders = !$chkCollectNamedFolders.IsChecked
        SkipTaskLinks = !$chkCollectTaskLinks.IsChecked
        SkipMSSettings = !$chkCollectMSSettings.IsChecked
        SkipDeepLinks = !$chkCollectDeepLinks.IsChecked
        SkipURLProtocols = !$chkCollectURLProtocols.IsChecked
        SkipHiddenAppLinks = !$chkCollectAppxLinks.IsChecked
        Output = $txtCurrentPath.Text
        Verbose = $cmbLoggingLevel.SelectedIndex -eq 1
        Debug = $cmbLoggingLevel.SelectedIndex -eq 2
    }
}

# ==============================================================================================================================
# ===============================================  INITIALISIERUNG & SETUP  ====================================================
# ==============================================================================================================================

$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$defaultOutputFolderName = "Super God Mode"
if ($uniqueOutputFolder) { $defaultOutputFolderName += "-$timestamp" }

if (-not $NoGUI) {
    Write-Host "`nBitte nutze das sich öffnende GUI-Fenster, um die Optionen festzulegen.`n(Mache dieses PowerShell-Fenster nicht zu!)"
    $params = Show-SuperGodModeDialog -defaultOutputFolderName $defaultOutputFolderName -initialDebug:$Debug -initialVerbose:$Verbose
    if ($null -eq $params) {
        Write-Host "GUI-Fenster wurde geschlossen. Skript wird abgebrochen.`n" -ForegroundColor Yellow
        exit
    }
    $DontGroupTasks = $params.DontGroupTasks
    $UseAlternativeCategoryNames = $params.UseAlternativeCategoryNames
    $AllURLProtocols = $params.AllURLProtocols
    $CollectExtraURLProtocolInfo = $params.CollectExtraURLProtocolInfo
    $KeepPreviousOutputFolders = $params.KeepPreviousOutputFolders
    $NoStatistics = $params.NoStatistics
    $AllowDuplicateDeepLinks = $params.AllowDuplicateDeepLinks
    $DeepScanHiddenLinks = $params.DeepScanHiddenLinks
    $SkipCLSID = $params.SkipCLSID
    $SkipNamedFolders = $params.SkipNamedFolders
    $SkipTaskLinks = $params.SkipTaskLinks
    $SkipMSSettings = $params.SkipMSSettings
    $SkipDeepLinks = $params.SkipDeepLinks
    $SkipURLProtocols = $params.SkipURLProtocols
    $SkipHiddenAppLinks = $params.SkipHiddenAppLinks
    $Output = $params.Output
    $Verbose = $params.Verbose
    $Debug = $params.Debug
}

if ($Verbose) {
    $Verbose = $true
    $VerbosePreference = 'Continue'
} else { $VerbosePreference = 'SilentlyContinue' }

if ($Debug) {
    $Debug = $true
    $DebugPreference = 'Continue'
    $VerbosePreference = 'Continue'
    $Verbose = $true
} else { $DebugPreference = 'SilentlyContinue' }

Write-Host "Starte Skriptausführung..." -ForegroundColor Green

if ($Output) {
    $Output = $Output.TrimEnd("\")
    if ($uniqueOutputFolder) { $Output += "-$timestamp" }

    if (-not [System.IO.Path]::IsPathRooted($Output)) {
        $mainShortcutsFolder = Join-Path $PSScriptRoot $Output
    } else { $mainShortcutsFolder = $Output }
} else {
    $mainShortcutsFolder = Join-Path $PSScriptRoot $defaultOutputFolderName
}

$debugLogsFolderName = "__Debug Logs"
$debugLogFolderPath = Join-Path $mainShortcutsFolder $debugLogsFolderName
if ($Debug) {
    if (-not (Test-Path $debugLogFolderPath)) { New-Item -Path $debugLogFolderPath -ItemType Directory -Force | Out-Null }
    $debugTranscriptFileName = "DebugTranscript-$timestamp.log"
    $debugTranscriptFilePath = Join-Path $debugLogFolderPath $debugTranscriptFileName
    Start-Transcript -Path $debugTranscriptFilePath
}

$clsidFolderName = "CLSID Shell-Ordner-Links"
$namedFolderName = "Benannte Systemordner"
$taskLinksFolderName = "Integrierte Aufgaben-Verknüpfungen"
$msSettingsFolderName = "Moderne Einstellungen"
$deepLinksFolderName = "Deep Links"
$urlProtocolsFolderName = "URL-Protokolle"
$URLProtocolPageLinksFolderName = "Versteckte App-Links"
$statisticsFolderName = "__Skript-Daten-Statistiken"

$CLSIDshortcutsOutputFolder = Join-Path $mainShortcutsFolder $clsidFolderName
$namedShortcutsOutputFolder = Join-Path $mainShortcutsFolder $namedFolderName
$taskLinksOutputFolder = Join-Path $mainShortcutsFolder $taskLinksFolderName
$msSettingsOutputFolder = Join-Path $mainShortcutsFolder $msSettingsFolderName
$deepLinksOutputFolder = Join-Path $mainShortcutsFolder $deepLinksFolderName
$URLProtocolLinksOutputFolder = Join-Path $mainShortcutsFolder $urlProtocolsFolderName
$URLProtocolPageLinksOutputFolder = Join-Path $mainShortcutsFolder $URLProtocolPageLinksFolderName
$statisticsOutputFolder = Join-Path $mainShortcutsFolder $statisticsFolderName

$csvFiles = @{
    CLSID = @{ Value = "CLSID_Shell_Ordner.csv"; Skip = $SkipCLSID }
    NamedFolders = @{ Value = "Benannte_Ordner.csv"; Skip = $SkipNamedFolders }
    TaskLinks = @{ Value = "System_Aufgabenlinks.csv"; Skip = $SkipTaskLinks }
    MSSettings = @{ Value = "MS_Settings_Sammlung.csv"; Skip = $SkipMSSettings }
    DeepLinks = @{ Value = "Deep_Links_Sammlung.csv"; Skip = $SkipDeepLinks }
    URLProtocols = @{ Value = "Registrierte_URL_Protokolle.csv"; Skip = $SkipURLProtocols }
    URLProtocolPages = @{ Value = "Versteckte_App_Verknüpfungen.csv"; Skip = $SkipHiddenAppLinks }
}
$xmlFiles = @{
    Shell32Content = @{ Value = "Shell32_Tasks_Rohdaten.xml"; Skip = $SkipTaskLinks }
    Shell32ResolvedContent = @{ Value = "Shell32_Tasks_Übersetzt.xml"; Skip = $SkipTaskLinks }
    ResolvedSettings = @{ Value = "Einstellungen_XML_Übersetzt.xml"; Skip = $SkipDeepLinks }
}

$clsidCsvPath = Join-Path $statisticsOutputFolder $csvFiles.CLSID.Value
$namedFoldersCsvPath = Join-Path $statisticsOutputFolder $csvFiles.NamedFolders.Value
$taskLinksCsvPath = Join-Path $statisticsOutputFolder $csvFiles.TaskLinks.Value
$msSettingsCsvPath = Join-Path $statisticsOutputFolder $csvFiles.MSSettings.Value
$deepLinksCsvPath = Join-Path $statisticsOutputFolder $csvFiles.DeepLinks.Value
$URLProtocolLinksCsvPath = Join-Path $statisticsOutputFolder $csvFiles.URLProtocols.Value
$URLProtocolPageLinksCsvPath = Join-Path $statisticsOutputFolder $csvFiles.URLProtocolPages.Value

$xmlContentFilePath = Join-Path $statisticsOutputFolder $xmlFiles.Shell32Content.Value
$resolvedXmlContentFilePath = Join-Path $statisticsOutputFolder $xmlFiles.Shell32ResolvedContent.Value
$resolvedSettingsXmlContentFilePath = Join-Path $statisticsOutputFolder $xmlFiles.ResolvedSettings.Value

$allSettingsXmlPath1 = "$Env:windir\ImmersiveControlPanel\Settings\AllSystemSettings_{D6E2A6C6-627C-44F2-8A5C-4959AC0C2B2D}.xml"
$allSettingsXmlPath2 = "$Env:windir\ImmersiveControlPanel\Settings\AllSystemSettings_{FDB289F3-FCFC-4702-8015-18926E996EC1}.xml"
$allSettingsXmlPath3 = "$Env:windir\ImmersiveControlPanel\Settings\AllSystemSettings_{253E530E-387D-4BC2-959D-E6F86122E5F2}.xml"
$systemSettingsDllPath = "$Env:windir\ImmersiveControlPanel\SystemSettings.dll"

$permanentURIProtocols = @(
    'bb','drop','fax','filesystem','grd','mailserver','modem','p1','pack','payment','prospero','snews','upt','videotex','wais','wpid','z39.50',
    'aaa','aaas','about','acap','acct','cap','cid','coap','coap+tcp','coap+ws','coaps','coaps+tcp','coaps+ws','crid','data','dav','dict','dns',
    'dtn','example','file','ftp','geo','go','gopher','h323','http','https','iax','icap','im','imap','info','ipn','ipp','ipps','iris','iris.beep',
    'iris.lwz','iris.xpc','iris.xpcs','jabber','ldap','leaptofrogans','mailto','mid','msrp','msrps','mt','mtqp','mupdate','news','nfs','ni','nih',
    'nntp','opaquelocktoken','pkcs11','pop','pres','reload','rtsp','rtsps','rtspu','service','session','shttp','sieve','sip','sips','sms','snmp',
    'soap.beep','soap.beeps','stun','stuns','tag','tel','telnet','tftp','thismessage','tip','tn3270','turn','turns','tv','urn','vemmi','vnc','ws',
    'wss','xcon','xcon-userid','xmlrpc.beep','xmlrpc.beeps','xmpp','z39.50r','z39.50s'
)

function TestPathSafe {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    try {
        $testPathResult = Test-Path -LiteralPath $Path -PathType Leaf -ErrorAction Stop
        return $testPathResult
    }
    catch {
        Write-Debug "Fehler bei sicherer Pfadprüfung für $Path: $_"
        return $false
    }
}

if ($CustomAllSystemSettingsXMLPath) {
    if (-not (TestPathSafe $CustomAllSystemSettingsXMLPath)) {
        Write-Error "Der angegebene Pfad existiert nicht: $CustomAllSystemSettingsXMLPath"
        return
    } else {
        $allSettingsXmlPath = $CustomAllSystemSettingsXMLPath
    }
} elseif (TestPathSafe $allSettingsXmlPath1) {
    $allSettingsXmlPath = $allSettingsXmlPath1
} elseif (TestPathSafe $allSettingsXmlPath2) {
    $allSettingsXmlPath = $allSettingsXmlPath2
} elseif (TestPathSafe $allSettingsXmlPath3) {
    $allSettingsXmlPath = $allSettingsXmlPath3
} else {
    Write-Error "Keine AllSystemSettings XML-Datei gefunden. Deep-Links können leider nicht erstellt werden."
    $allSettingsXmlPath = $null
}

$mainFolderWasCreatedThisRun = -not (Test-Path $mainShortcutsFolder)

try {
    New-Item -Path $mainShortcutsFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
} catch [System.UnauthorizedAccessException] {
    Write-Error "Fehlende Berechtigungen zum Erstellen des Ausgabe-Ordners: $_"
    return
} catch {
    if (-not (Test-Path $mainShortcutsFolder)) {
        Write-Host "Fehler beim Erstellen des Ordners: $_" -ForegroundColor Yellow
        return
    }
}

if (-not $KeepPreviousOutputFolders) {
    try {
        if (Test-Path $mainShortcutsFolder) {
            if (Test-Path $CLSIDshortcutsOutputFolder){ Remove-Item -Path $CLSIDshortcutsOutputFolder -Recurse -Force }
            if (Test-Path $namedShortcutsOutputFolder){ Remove-Item -Path $namedShortcutsOutputFolder -Recurse -Force }
            if (Test-Path $taskLinksOutputFolder){ Remove-Item -Path $taskLinksOutputFolder -Recurse -Force }
            if (Test-Path $statisticsOutputFolder) { Remove-Item -Path $statisticsOutputFolder -Recurse -Force }
            if (Test-Path $msSettingsOutputFolder){ Remove-Item -Path $msSettingsOutputFolder -Recurse -Force }
            if (Test-Path $deepLinksOutputFolder){ Remove-Item -Path $deepLinksOutputFolder -Recurse -Force }
            if (Test-Path $URLProtocolLinksOutputFolder){ Remove-Item -Path $URLProtocolLinksOutputFolder -Recurse -Force }
            if (Test-Path $URLProtocolPageLinksOutputFolder){ Remove-Item -Path $URLProtocolPageLinksOutputFolder -Recurse -Force }
        }
    } catch {
        Write-Error "Fehler beim Bereinigen älterer Ordnerstrukturen: $_"
    }
}

if ($CustomDLLPath -and -not (Test-Path-Safe $CustomDLLPath)) {
    Write-Error "Die spezifizierte DLL existiert nicht: $CustomDLLPath"
    return
}

if ($CustomLanguageFolderPath) {
    if (-not (Test-Path $CustomLanguageFolderPath -PathType Container)) {
         Write-Error "Der Sprachpfad ist kein valider Ordner: $CustomLanguageFolderPath"
         return
    } else {
         Write-Host "Nutze benutzerdefinierten Sprachordner: $CustomLanguageFolderPath"
    }
}

if ($CustomSystemSettingsDLLPath -and -not (Test-Path-Safe $CustomSystemSettingsDLLPath)) {
    Write-Error "Die SystemSettings.dll existiert nicht am Pfad: $CustomSystemSettingsDLLPath"
    return
} else {
    $systemSettingsDllPath = $CustomSystemSettingsDLLPath
}

if ($SkipURLProtocols) { $SkipHiddenAppLinks = $true }

# Generator für Desktop.ini konfigurierte Verzeichnisse
function New-FolderWithIcon {
    param (
        [string]$FolderPath,
        [string]$IconFile,
        [string]$IconIndex,
        [string]$IconFileWin10,
        [string]$IconIndexWin10,
        [switch]$DetailsView,
        [switch]$desktopIniOnly
    )
    if (-not $desktopIniOnly) {
         New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null
    }

    $windowsBuildNum = [System.Environment]::OSVersion.Version.Build
    if ($windowsBuildNum -lt 22000) {
        if ($IconIndexWin10) { $IconIndex = $IconIndexWin10 }
        if ($IconFileWin10) { $IconFile = $IconFileWin10 }
    }

    if ($IconIndex -notmatch '^-') { $IconIndex = "-$IconIndex" }

    if ($DetailsView) {
        $desktopIniContent = @"
[.ShellClassInfo]
IconResource=$IconFile,$IconIndex
[ViewState]
Mode=4
Vid={137E7700-3573-11CF-AE69-08002B2E1262}
"@
    } else {
        $desktopIniContent = @"
[.ShellClassInfo]
IconResource=$IconFile,$IconIndex
"@
    }

    $desktopIniPath = Join-Path $FolderPath "desktop.ini"
    Set-Content -Path $desktopIniPath -Value $desktopIniContent -Encoding ASCII -Force

    $desktopIniItem = Get-Item $desktopIniPath -Force
    $desktopIniItem.Attributes = 'Hidden,System'

    $folderItem = Get-Item $FolderPath -Force
    $folderItem.Attributes = 'ReadOnly,Directory'
}

# Aufbau der Verzeichnisse mit individuellen Symbolen
if ($mainFolderWasCreatedThisRun) {
    New-FolderWithIcon -FolderPath $mainShortcutsFolder -IconFile "%windir%\System32\imageres.dll" -IconIndex "10" -IconIndexWin10 "190" -desktopIniOnly
}
if (-not $SkipCLSID) {
    New-FolderWithIcon -FolderPath $CLSIDshortcutsOutputFolder -IconFile "%windir%\System32\shell32.dll" -IconIndex "20" -IconIndexWin10 "210" -DetailsView
}
if (-not $SkipNamedFolders) {
    New-FolderWithIcon -FolderPath $namedShortcutsOutputFolder -IconFile "%windir%\System32\imageres.dll" -IconIndex "77" -DetailsView
}
if (-not $SkipTaskLinks) {
    New-FolderWithIcon -FolderPath $taskLinksOutputFolder -IconFile "%windir%\System32\shell32.dll" -IconIndex "137" -DetailsView
}
if (-not $SkipMSSettings) {
    New-FolderWithIcon -FolderPath $msSettingsOutputFolder -IconFile "%windir%\System32\imageres.dll" -IconIndex "114" -DetailsView
}
if (-not $SkipDeepLinks) {
    New-FolderWithIcon -FolderPath $deepLinksOutputFolder -IconFile "%windir%\System32\imageres.dll" -IconIndex "175" -DetailsView
}
if (-not $SkipURLProtocols) {
    New-FolderWithIcon -FolderPath $URLProtocolLinksOutputFolder -IconFile "%windir%\System32\imageres.dll" -IconIndex "5302" -DetailsView
}
if (-not $SkipHiddenAppLinks) {
    New-FolderWithIcon -FolderPath $URLProtocolPageLinksOutputFolder -IconFile "%windir%\System32\imageres.dll" -IconIndex "1025" -DetailsView
}
if (-not $NoStatistics) {
    New-FolderWithIcon -FolderPath $statisticsOutputFolder -IconFile "%windir%\System32\imageres.dll" -IconIndex "9" -DetailsView
}

if (-not $NoReadMe) {
    $tipsFilePath = Join-Path $mainShortcutsFolder "!Lies Mich - Tipps und Infos.txt"
    $tipsContent = @"
------------------------- Hilfreiche Tipps -------------------------

• Um direkt zu sehen, wohin Verknüpfungen führen, mache im Windows Explorer die Spalte `"Zielort`" (Link Target) sichtbar. (Rechtsklick Spaltenköpfe > Mehr > Zielort aktivieren).

• Einige Links können fehlschlagen. Das ist normal: Windows hält oft veralteten 'Legacy'-Leichencode im System, der nicht mehr funktionstüchtig ist.

• Das Skript basiert nicht auf einer starren Liste, sondern scannt dein lokales Windows-System dynamisch! Dadurch hängen Ergebnisse von deiner Windows-Version und installierten Apps ab.

---------------------------------------------------------------------------
Erstellt mit:   `"Super God Mode`" Skript - Version: $VERSION
Entwickler:     ThioJoe
Projektseite:   https://github.com/ThioJoe/Windows-Super-God-Mode
"@
    Set-Content -Path $tipsFilePath -Value $tipsContent -Force
}

# Win32 & COM Interop-Signaturen laden (P/Invoke)
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    public class Windows
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr FindResource(IntPtr hModule, int lpName, string lpType);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr LoadResource(IntPtr hModule, IntPtr hResInfo);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr LockResource(IntPtr hResData);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern uint SizeofResource(IntPtr hModule, IntPtr hResInfo);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool FreeLibrary(IntPtr hModule);
    }
"@

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class IconExtractor
{
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool DestroyIcon(IntPtr hIcon);
}
"@

if (-not ([System.Management.Automation.PSTypeName]'Win32').Type) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    using System.Text;
    public class Win32 {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        public static extern int SHLoadIndirectString(string pszSource, StringBuilder pszOutBuf, int cchOutBuf, IntPtr ppvReserved);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool FreeLibrary(IntPtr hModule);
    }
"@
}

# Warnmeldung stylen
function Print-SuperWarning {
    param([string]$message)
    $paddingSize = 4
    $lineLength = $message.Length + ($paddingSize * 2)
    $yellowLine = (" " * $lineLength)
    $padding = " " * $paddingSize
    $centeredMessage = $padding + $message + $padding

    Write-Host $yellowLine -BackgroundColor Red
    Write-Host $centeredMessage -BackgroundColor Yellow -ForegroundColor Black
    Write-Host $yellowLine -BackgroundColor Red
}

# Registry Multi-Language Strings auflösen
function Get-LocalizedString {
    param (
        [string]$StringReference,
        [string]$CustomLanguageFolder,
        [string]$AppxManifestPath,
        [switch]$returnSanitizedOnFail = $false
    )
    if ($AppxManifestPath) {
        $manifestParentFolder = Split-Path $AppxManifestPath | Split-Path -Leaf
        Write-Debug "Lese Ressource: $StringReference | Paket: $manifestParentFolder"
    } else {
        Write-Debug "Lese Ressource: $StringReference"
    }

    if ($StringReference -match '^\@\@') {
        $references = $StringReference -split '@' | Where-Object { $_ -ne '' }
        $resolvedStrings = @()
        foreach ($ref in $references) {
            $resolved = Get-LocalizedString -StringReference "@$ref" -CustomLanguageFolder $CustomLanguageFolder -AppxManifestPath $AppxManifestPath -returnSanitizedOnFail:$false
            if ($resolved) {
                $resolvedStrings += $resolved -split ';' | ForEach-Object { $_.Trim() }
            }
        }
        return ($resolvedStrings | Select-Object -Unique) -join ';'
    }
    elseif ($StringReference -match '^ms-resource:') {
        return Get-FullMsResource -ShortReference $StringReference -AppxManifestPath $AppxManifestPath -returnSanitizedOnFail:$returnSanitizedOnFail
    }
    elseif ($StringReference -match '@\{.+\?ms-resource://.+}') {
        return Get-MsResource $StringReference -returnSanitizedOnFail:$returnSanitizedOnFail
    }
    elseif ($StringReference -match '@(.+),-(\d+)') {
        $dllPath = [Environment]::ExpandEnvironmentVariables($Matches[1])
        $resourceId = [uint32]$Matches[2]

        if ($CustomLanguageFolder){
            $muiNameToCheck = "$dllPath.mui"
            if (Test-Path-Safe (Join-Path $CustomLanguageFolder $muiNameToCheck)) {
                $dllPath = Join-Path $CustomLanguageFolder $muiNameToCheck
            }
        }

        $hModule = [Win32]::LoadLibrary($dllPath)
        if ($hModule -eq [IntPtr]::Zero) {
            return Sanitize-Unresolved-Reference -ReferenceString $StringReference -returnSanitizedOnFail:$returnSanitizedOnFail
        }

        $stringBuilder = New-Object System.Text.StringBuilder 1024
        $result = [Win32]::LoadString($hModule, $resourceId, $stringBuilder, $stringBuilder.Capacity)

        [void][Win32]::FreeLibrary($hModule)

        if ($result -ne 0) {
            return $stringBuilder.ToString()
        } else {
            return Sanitize-Unresolved-Reference -ReferenceString $StringReference -returnSanitizedOnFail:$returnSanitizedOnFail
        }
    } else {
        return Sanitize-Unresolved-Reference -ReferenceString $StringReference -returnSanitizedOnFail:$returnSanitizedOnFail
    }
}

function Sanitize-Unresolved-Reference {
    param (
        [string]$ReferenceString,
        [switch]$returnSanitizedOnFail = $false
    )
    $sanitizedReferenceDefault = "UnresolvedReference"

    if (-not $ReferenceString) {
        if ($returnSanitizedOnFail) { return $sanitizedReferenceDefault }
        return $null
    }
    if (-not $returnSanitizedOnFail) { return $null }

    try {
        $sanitizedReference = $ReferenceString.Trim()
        $sanitizedReference = $sanitizedReference -replace '[<>:"/\\|?*]', '_'
        $sanitizedReference = "UnresolvedReference - $sanitizedReference"
    } catch {
        if ($returnSanitizedOnFail) { return $sanitizedReferenceDefault }
        return $null
    }
    return $sanitizedReference
}

function Get-FullMsResource {
    param (
        [string]$ShortReference,
        [string]$AppxManifestPath,
        [switch]$returnSanitizedOnFail = $false
    )

    $manifest = [xml](Get-Content $AppxManifestPath)
    $packageName = $manifest.Package.Identity.Name
    $resourceName = $ShortReference -replace '^ms-resource:/*', ''

    if ($resourceName -match "^$packageName/") {
        $fullReference = "@{$packageName`?ms-resource://$resourceName}"
    } elseif ($resourceName -match '^Resources/') {
        $fullReference = "@{$packageName`?ms-resource://$packageName/$resourceName}"
    } elseif ($resourceName -match '/Resources/') {
        $fullReference = "@{$packageName`?ms-resource://$packageName/$resourceName}"
    } else {
        $fullReference = "@{$packageName`?ms-resource://$packageName/Resources/$resourceName}"
    }

    return Get-MsResource $fullReference -returnSanitizedOnFail:$returnSanitizedOnFail
}

function Get-MsResource {
    param (
        [string]$ResourcePath,
        [switch]$returnSanitizedOnFail = $false
    )
    $stringBuilder = New-Object System.Text.StringBuilder 1024
    $result = [Win32]::SHLoadIndirectString($ResourcePath, $stringBuilder, $stringBuilder.Capacity, [IntPtr]::Zero)

    if ($result -eq 0) {
        return $stringBuilder.ToString()
    } else {
        $packageFullName = ($ResourcePath -split '\?')[0].Trim('@{}')
        $resourceUri = ($ResourcePath -split '\?')[1].Trim('@{}')
        $packageName = ($packageFullName -split '_')[0]

        $package = Get-AppxPackage | Where-Object { $_.Name -eq $packageName }
        if (-not $package) {
            $packageFamilyName = ($packageFullName -split '_')[-1]
            $package = Get-AppxPackage | Where-Object { $_.PackageFamilyName -eq "${packageName}_$packageFamilyName" }
        }

        if ($package) {
            $packagePath = $package.InstallLocation
            $priPath = Join-Path $packagePath "resources.pri"
            if (Test-Path-Safe $priPath) {
                $newResourcePath = "@{" + $priPath + "?" + $resourceUri + "}"
                $result = [Win32]::SHLoadIndirectString($newResourcePath, $stringBuilder, $stringBuilder.Capacity, [IntPtr]::Zero)
                if ($result -eq 0) {
                    return $stringBuilder.ToString()
                }
            }
        }
        return Sanitize-Unresolved-Reference -ReferenceString $ResourcePath -returnSanitizedOnFail:$returnSanitizedOnFail
    }
}

# Holt die Bezeichner für CLSIDs aus der Control Panel Registry
function Get-FolderName {
    param (
        [string]$clsid,
        [string]$CustomLanguageFolder
    )
    $fetchNameErrorVariable = $null
    $nameSource = "Unknown"

    $defaultPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid"
    $defaultName = (Get-ItemProperty -Path $defaultPath -ErrorAction SilentlyContinue -ErrorVariable $fetchNameErrorVariable).'(Default)'

    if ($defaultName) {
        if ($defaultName -match '@.+,-\d+') {
            $resolvedName = Get-LocalizedString -StringReference $defaultName -CustomLanguageFolder $CustomLanguageFolder -returnSanitizedOnFail:$false
            if ($resolvedName) {
                $nameSource = "Localized String"
                return @($resolvedName, $nameSource)
            }
        }
        $nameSource = "Default Value"
        return @($defaultName, $nameSource)
    }

    $initPropertyBagPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\Instance\InitPropertyBag"
    $targetKnownFolder = (Get-ItemProperty -Path $initPropertyBagPath -ErrorAction SilentlyContinue).TargetKnownFolder

    if ($targetKnownFolder) {
        $folderDescriptionsPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\$targetKnownFolder"
        $folderName = (Get-ItemProperty -Path $folderDescriptionsPath -ErrorAction SilentlyContinue).Name
        if ($folderName) {
            $nameSource = "Known Folder ID"
            return @($folderName, $nameSource)
        }
    }

    $localizedStringPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid"
    $localizedString = (Get-ItemProperty -Path $localizedStringPath -ErrorAction SilentlyContinue).LocalizedString

    if ($localizedString) {
        $resolvedString = Get-LocalizedString -StringReference $localizedString -CustomLanguageFolder $CustomLanguageFolder -returnSanitizedOnFail:$false
        if ($resolvedString) {
            $nameSource = "Localized String"
            return @($resolvedString, $nameSource)
        }
    }

    return @($clsid, $nameSource)
}

# ==============================================================================================================================
# =========================  [OPTIMIERUNG 1&2: ANPASSUNG DER VERKNÜPFUNGS-ERSTELLUNG]  ========================================
# ==============================================================================================================================
# LERNEFFEKT: Wir verzichten in den Funktionen jetzt vollständig auf das lokale Neuerzeugen per
# "New-Object -ComObject WScript.Shell". Stattdessen vertrauen wir blind auf das globale $GlobalShell-Objekt.
# Das entlastet den Garbage Collector der .NET-Laufzeitumgebung massiv.
function Create-CLSID-Shortcut {
    param (
        [string]$clsid,
        [string]$name,
        [string]$shortcutPath,
        [string]$pageName = ""
    )
    try {
        # Wir nützen direkt die globale Shell-Instanz
        $shortcut = $GlobalShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "explorer.exe"

        if ($pageName) {
            $shortcut.Arguments = "shell:::$clsid\$pageName"
        } else {
            $shortcut.Arguments = "shell:::$clsid"
        }

        $iconPath = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\DefaultIcon" -ErrorAction SilentlyContinue).'(default)'
        if ($iconPath) {
            $shortcut.IconLocation = $iconPath
        } else {
            $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,3"
        }

        $shortcut.Save()
        return $true
    }
    catch {
        Write-Error "Fehler beim Erstellen des CLSID-Shortcuts für $name: $_"
        return $false
    }
}

function Check-File-For-Icon {
    param ([string]$filePath)
    try {
        $hIcon = [IconExtractor]::ExtractIcon([IntPtr]::Zero, $filePath, 0)
        if ($hIcon -ne [IntPtr]::Zero) {
            $iconPath = $filePath + ",0"
            [void][IconExtractor]::DestroyIcon($hIcon)
            return $iconPath
        }
    } catch {}
}

function Get-TaskIcon {
    param (
        [string]$controlPanelName,
        [string]$applicationId,
        [string]$commandTarget
    )
    $iconPath = $null

    if ($controlPanelName) {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\$controlPanelName"
        $iconPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Icon
    }

    if (-not $iconPath -and $applicationId) {
        $regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\$applicationId"
        $iconPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).Icon

        if (-not $iconPath) {
            $regPath = "Registry::HKEY_CLASSES_ROOT\CLSID\$applicationId\DefaultIcon"
            $iconPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).'(default)'
        }
    }

    if (-not $iconPath -and $commandTarget) {
        $commandFileName = Split-Path -Path $commandTarget -Leaf
        $ignoredExeFiles = @("control.exe", "rundll32.exe")
        if ($commandFileName -match '\.exe$' -and $ignoredExeFiles -notcontains $commandFileName) {
            $iconPath = Check-File-For-Icon -filePath $commandTarget
        }
    }

    if ($iconPath) { return Fix-CommandPath $iconPath }
    return "%SystemRoot%\System32\shell32.dll,0"
}

function Fix-CommandPath {
    param ([string]$command)
    if ($command -match '^\%\%') { $command = $command -replace '^\%\%', '%' }
    return $command
}

# Aufgabenverknüpfungen erstellen (Unter Nützung der globalen Shell)
function Create-TaskLink-Shortcut {
    param (
        [string]$name,
        [string]$shortcutPath,
        [string]$shortcutType,
        [string]$command,
        [string]$controlPanelName,
        [string]$applicationId,
        [string[]]$keywords
    )

    $targetAndPathRegex = '(?xi)
        ^(
            (?:[a-z]:|%%?\w+%)\\
            (?:[^\\/:*?"<>|\r\n]+?\\)*?
            [^\\/:*?"<>|\r\n]+?
            \.[a-z0-9]+
        |
            [^\\/:*?"<>|\r\n]+?\.[a-z0-9]+
        )
        (?=\s|$|["])
        \s*
        (.*)
    '

    try {
        $targetPath = ''
        $arguments = ''

        # Erneut: Nutzung der performanten globalen Shell!
        if ($shortcutType -eq "url") {
            $shortcut = $GlobalShell.CreateShortcut($shortcutPath)
            $shortcut.TargetPath = $command
        } else {
            $shortcut = $GlobalShell.CreateShortcut($shortcutPath)

            if ($command -match $targetAndPathRegex) {
                $targetPath = Fix-CommandPath $Matches[1]
                $arguments = Fix-CommandPath $Matches[2]
                $arguments = [Environment]::ExpandEnvironmentVariables($arguments)

                $shortcut.TargetPath = $targetPath
                $shortcut.Arguments = $arguments
            } else {
                $shortcut.TargetPath = Fix-CommandPath $command
            }

            if ($keywords -and $keywords.Count -gt 0) {
                $descriptionLimit = 259
                $keywordString = ""
                $separator = " "

                foreach ($keyword in $keywords) {
                    $potentialNewString = if ($keywordString) { $keywordString + $separator + $keyword } else { $keyword }
                    if ($potentialNewString.Length -le $descriptionLimit) { $keywordString = $potentialNewString } else { break }
                }
                $shortcut.Description = $keywordString
            }
        }
        
        $iconPath = Get-TaskIcon -controlPanelName $controlPanelName -applicationId $applicationId -commandTarget $shortcut.TargetPath

        if ($shortcutType -eq "lnk") { $shortcut.IconLocation = $iconPath }
        $shortcut.Save()

        if ($shortcutType -eq "url" -and $iconPath) {
            $iconFile, $iconIndex = $iconPath -split ','
            Add-Content -Path $shortcutPath -Value "IconFile=$iconFile"
            if ($iconIndex) { Add-Content -Path $shortcutPath -Value "IconIndex=$iconIndex" }
        }
        return $true
    } catch {
        Write-Host "Fehler beim Generieren der Verknüpfung für $name: $_"
        return $false
    }
}

# Holt die XML-Definitionen direkt aus der shell32.dll Datei
function Get-Shell32XMLContent {
    param(
        [switch]$SaveXML,
        [string]$CustomDLL
    )
    if ($CustomDLL) {
        $dllPath = $CustomDLL
    } else {
        $dllPath = Join-Path $env:SystemRoot "System32\shell32.dll"
    }

    $LOAD_LIBRARY_AS_DATAFILE = 0x00000002
    $DONT_RESOLVE_DLL_REFERENCES = 0x00000001
    $xmlContent = ""

    $shell32Handle = [Windows]::LoadLibraryEx($dllPath, [IntPtr]::Zero, $LOAD_LIBRARY_AS_DATAFILE -bor $DONT_RESOLVE_DLL_REFERENCES)
    if ($shell32Handle -eq [IntPtr]::Zero) {
        Write-Error "DLL konnte nicht als Nur-Lese-Ressource geöffnet werden: $dllPath"
        return $null
    }

    try {
        $hResInfo = [Windows]::FindResource($shell32Handle, 21, "XML")
        if ($hResInfo -eq [IntPtr]::Zero) {
            Write-Error "XML-Ressourcenpfad id:21 in shell32.dll konnte nicht ausgelesen werden."
            return $null
        }

        $hResData = [Windows]::LoadResource($shell32Handle, $hResInfo)
        $pData = [Windows]::LockResource($hResData)
        $size = [Windows]::SizeofResource($shell32Handle, $hResInfo)
        $byteArray = New-Object byte[] $size
        [System.Runtime.InteropServices.Marshal]::Copy($pData, $byteArray, 0, $size)

        $xmlContent = [System.Text.Encoding]::UTF8.GetString($byteArray)
        $xmlContent = $xmlContent -replace "`0", ""
    }
    finally {
        [void][Windows]::FreeLibrary($shell32Handle)
    }

    $xmlContent = $xmlContent.Trim()

    if ($SaveXML) { Save-PrettyXML -xmlContent $xmlContent -outputPath $xmlContentFilePath }
    return $xmlContent
}

function Save-PrettyXML {
    param (
        [string]$xmlContent,
        [string]$outputPath
    )
    try {
        $xmlDoc = New-Object System.Xml.XmlDocument
        $xmlDoc.LoadXml($xmlContent)

        $writerSettings = New-Object System.Xml.XmlWriterSettings
        $writerSettings.Indent = $true
        $writerSettings.IndentChars = "  "
        $writerSettings.NewLineChars = "`r`n"
        $writerSettings.NewLineHandling = [System.Xml.NewLineHandling]::Replace

        $writer = [System.Xml.XmlWriter]::Create($outputPath, $writerSettings)
        $xmlDoc.Save($writer)
        $writer.Close()
    }
    catch {
        Write-Error "Formatierung der XML-Datensätze schlug fehl: $_"
    }
}

# Parse die Control-Panel-Tasks aus XML
function Get-TaskLinks {
    param(
        [switch]$SaveXML,
        [string]$DLLPath,
        [string]$CustomLanguageFolder
    )
    $xmlContent = Get-Shell32XMLContent -SaveXML:$SaveXML -CustomDLL:$DLLPath

    if (-not $xmlContent) {
        Write-Host "Shell32-XML Abruf blockiert. Überspringe Tasks." -ForegroundColor Yellow
        return $null
    }

    try {
        $xml = [xml]$xmlContent
    } catch {
        Write-Error "Auswertung des XML-Schemas blockiert: $_"
        return $null
    }

    $resolvedXml = $xml.Clone()

    $nsManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsManager.AddNamespace("cpl", "http://schemas.microsoft.com/cpltasks/v1")
    $nsManager.AddNamespace("sh", "http://schemas.microsoft.com/windows/tasks/v1")
    $nsManager.AddNamespace("sh2", "http://schemas.microsoft.com/windows/tasks/v2")

    $tasks = @()
    $allTasks = $xml.SelectNodes("//sh:task", $nsManager)

    foreach ($task in $allTasks) {
        $taskId = $task.GetAttribute("id")
        $nameNode = $task.SelectSingleNode("sh:name", $nsManager)
        $controlPanel = $task.SelectSingleNode("sh2:controlpanel", $nsManager)
        $commandNode = $task.SelectSingleNode("sh:command", $nsManager)
        $keywordsNodes = $task.SelectNodes("sh:keywords", $nsManager)

        $name = $null
        if ($nameNode -and $nameNode.InnerText) {
            if ($nameNode.InnerText -match '@(.+),-(\d+)') {
                $name = Get-LocalizedString -StringReference $nameNode.InnerText -CustomLanguageFolder $CustomLanguageFolder -returnSanitizedOnFail:$false
                if ($name -and $resolvedXml) {
                    try {
                        $resolvedNameNode = $resolvedXml.SelectSingleNode("//sh:task[@id='$taskId']/sh:name", $nsManager)
                        if ($resolvedNameNode) { $resolvedNameNode.InnerText = $name }
                    } catch {}
                }
            } else {
                $name = $nameNode.InnerText
            }
        }
        if ($name) {
            $name = $name.Trim()
        } elseif ($task.Name -eq "sh:task" -and $task.GetAttribute("idref")) {
            continue
        } else {
            continue
        }

        $command = $null
        $appName = $null
        $page = $null
        $appId = $task.ParentNode.id
        if (-not $appId) { $appId = $task.GetAttribute("id") }

        if ($controlPanel) {
            $appName = $controlPanel.GetAttribute("name")
            $page = $controlPanel.GetAttribute("page")
            $command = "control.exe /name $appName"
            if ($page) { $command += " /page $page" }
        } elseif ($commandNode) {
            $command = $commandNode.InnerText
        }

        $keywords = @()
        foreach ($keywordNode in $keywordsNodes) {
             $keyword = $null
             if ($keywordNode.InnerText -match '@(.+),-(\d+)') {
                 $keyword = Get-LocalizedString -StringReference $keywordNode.InnerText -CustomLanguageFolder $CustomLanguageFolder
             } else {
                 $keyword = $keywordNode.InnerText
             }
             if ($keyword) {
                 $splitKeywords = $keyword.Split(';', [StringSplitOptions]::RemoveEmptyEntries)
                 foreach ($splitKeyword in $splitKeywords) {
                     if ($splitKeyword) { $keywords += $splitKeyword.Trim() }
                 }
             }
        }

        if (-not $appName) {
            $appName = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$appId" -ErrorAction SilentlyContinue)."System.ApplicationName"
        }

        if ($name -and ($command -or $appName)) {
            $newTask = [PSCustomObject]@{
                TaskId = $taskId
                ApplicationId = $appId
                Name = $name
                ApplicationName = $appName
                Page = $page
                Command = $command
                Keywords = $keywords
                ControlPanelName = (if ($controlPanel) { $controlPanel.GetAttribute("name") } else { $null })
            }

            $isDuplicate = $tasks | Where-Object { $_.Name -eq $newTask.Name -and $_.Command -eq $newTask.Command }
            if (-not $isDuplicate) { $tasks += $newTask }
        }
    }

    $resolvedXmlContent = $resolvedXml.OuterXml
    if ($SaveXML) { Save-PrettyXML -xmlContent $resolvedXmlContent -outputPath $resolvedXmlContentFilePath }

    return $tasks
}

# Verknüpfung zu namenbasiertem Ordner anfertigen (Nützung globale Shell)
function Create-NamedShortcut {
    param (
        [string]$name,
        [string]$shortcutPath,
        [string]$iconPath
    )
    try {
        $shortcut = $GlobalShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "explorer.exe"
        $shortcut.Arguments = "shell:$name"

        if ($iconPath) {
            $shortcut.IconLocation = $iconPath
        } else {
            $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,3"
        }

        $shortcut.Save()
        return $true
    }
    catch {
        Write-Host "Konnte Ordnerverknüpfung nicht erstellen für $name: $_"
        return $false
    }
}

# ==============================================================================================================================
# ==========================================  FUNKTIONEN: CSV DATEIERSTELLUNG  =================================================
# ==============================================================================================================================

function Create-CLSIDCsvFile {
    param (
        [string]$outputPath,
        [array]$clsidData
    )
    $csvContent = "CLSID,ExplorerCommand,Name,NameSource,CustomIcon`n"
    foreach ($item in $clsidData) {
        $explorerCommand = "explorer shell:::$($item.CLSID)"
        $iconPath = if ($item.IconPath) { "`"$($item.IconPath -replace '"', '""')`"" } else { "None" }
        $escapedName = $item.Name -replace '"', '""'
        $csvContent += "$($item.CLSID),`"$explorerCommand`",`"$escapedName`",$($item.NameSource),$iconPath`n"
    }
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Create-NamedFoldersCsvFile {
    param ([string]$outputPath)
    $csvContent = "Name,ExplorerCommand,RelativePath,ParentFolder`n"
    $namedFolders = Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions"

    foreach ($folder in $namedFolders) {
        $folderProperties = Get-ItemProperty -Path $folder.PSPath
        $name = $folderProperties.Name
        if ($name) {
            $explorerCommand = "explorer shell:$name"
            $relativePath = $folderProperties.RelativePath -replace ',', '","'
            $parentFolderGuid = $folderProperties.ParentFolder
            $parentFolderName = "None"
            if ($parentFolderGuid) {
                $parentFolderPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\$parentFolderGuid"
                $parentFolderName = (Get-ItemProperty -Path $parentFolderPath -ErrorAction SilentlyContinue).Name
            }
            $csvContent += "`"$name`",`"$explorerCommand`",`"$relativePath`",`"$parentFolderName`"`n"
        }
    }
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Create-TaskLinksCsvFile {
    param (
        [string]$outputPath,
        [array]$taskLinksData
    )
    $csvContent = "XMLTaskId,ApplicationId,ApplicationName,Name,Page,Command,Keywords`n"
    foreach ($item in $taskLinksData) {
        $taskId = $item.TaskId -replace '"', '""'
        $applicationId = $item.ApplicationId -replace '"', '""'
        $applicationName = $item.ApplicationName -replace '"', '""'
        $name = $item.Name -replace '"', '""'
        $page = $item.Page -replace '"', '""'
        $command = $item.Command -replace '"', '""'
        $keywords = ($item.Keywords -join ';') -replace '"', '""'

        $csvContent += "`"$taskId`",`"$applicationId`",`"$applicationName`",`"$name`",`"$page`",`"$command`",`"$keywords`"`n"
    }
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Create-MSSettingsCsvFile {
    param (
        [string]$outputPath,
        [array]$msSettingsList
    )
    $csvContent = "Setting Name,Full Setting Command`n"
    foreach ($fullLink in $msSettingsList) {
        $fullLinkParts = $fullLink -split ":", 2
        $name = $fullLinkParts[1].Trim()
        $csvContent += "`"$name`",`"$fullLink`"`n"
    }
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Prettify-App-Name {
    param([string]$AppName, [string]$TaskName)
    $wordsToRejoin = @("Bit Locker", "Side Show", "Auto Play")
    $AppName = $AppName -replace '^Microsoft\.', ''
    $AppName = $AppName -creplace '(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])|\b(?=[A-Z]{2,}\b)', ' '

    foreach ($word in $wordsToRejoin) {
        $AppName = $AppName -replace $word, $word.Replace(' ', '')
    }
    $PrettyName = "$AppName - $TaskName"
    return ($PrettyName -replace '[\\/:*?"<>|]', '_')
}

function Get-AllSettings-Data {
    param (
        [string]$xmlFilePath,
        [switch]$SaveXML
    )
    if (-not (Test-Path-Safe $xmlFilePath)) { return $null }

    try {
        [xml]$xmlContent = Get-Content $xmlFilePath
        $searchableContentData = $xmlContent.PCSettings.SearchableContent
        $settingsData = @()
    } catch { return $null }

    foreach ($content in $searchableContentData) {
        $entryName = $content.Filename
        $deepLink = $content.ApplicationInformation.DeepLink
        $description = Get-LocalizedString $content.SettingInformation.Description -returnSanitizedOnFail:$true
        $content.SettingInformation.Description = $description

        $highKeywords = ""
        if ($content.SettingInformation.HighKeywords) {
             $highKeywords = Get-LocalizedString $content.SettingInformation.HighKeywords -returnSanitizedOnFail:$false
             if ($highKeywords -is [array]) { $highKeywords = $highKeywords -join "; " }
             $content.SettingInformation.HighKeywords = $highKeywords
        }

        $settingsData += @{
            Name         = $entryName
            Description  = $description
            HighKeywords = $highKeywords
            DeepLink     = $deepLink
            IconPath     = $content.ApplicationInformation.Icon
        }
    }

    if ($SaveXML) { $xmlContent.Save($resolvedSettingsXmlContentFilePath) }
    return $settingsData
}

function Get-MS-SettingsFrom-SystemSettingsDLL {
    param ([string]$DllPath)
    if (-not (Test-Path-Safe $DllPath)) { return $null }

    try {
        $content = [System.IO.File]::ReadAllText($DllPath, [System.Text.Encoding]::Unicode)
        $results = New-Object System.Collections.Generic.HashSet[string]

        $matchesList = [regex]::Matches($content, 'ms-settings:[a-z-]+')
        foreach ($match in $matchesList) { [void]$results.Add($match.Value) }
    } catch { return $null }
    
    Write-Host "Gefundene einzigartige 'ms-settings' Pfade: $($results.Count)"
    return ($results | Sort-Object)
}

function Create-MSSettings-Shortcut {
    param ([string]$fullName, [string]$shortcutPath)
    try {
        # Natürlich auch hier: Globale Shell!
        $shortcut = $GlobalShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = $fullName
        $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,-16826"
        $shortcut.Save()
        return $true
    } catch { return $false }
}

# Verknüpfung für tiefe URL/LNK-Dateien herstellen
function Create-Deep-Link-Shortcut {
    param ([object]$settingArray)
    $rawTarget = $settingArray.DeepLink
    $name = if ($settingArray.Description) { $settingArray.Description } else { $rawTarget }

    $targetAndPathRegex = '(?xi)
        ^(
            (?:[a-z]:|%%?\w+%)\\
            (?:[^\\/:*?"<>|\r\n]+?\\)*?
            [^\\/:*?"<>|\r\n]+?
            \.[a-z0-9]+
        |
            [^\\/:*?"<>|\r\n]+?\.[a-z0-9]+
        )
        (?=\s|$|["])
        \s*
        (.*)
    '
    $target = ""
    $targetArgs = ""

    if ($rawTarget -match '^Microsoft\.[a-zA-Z]+$') {
        $shortcutType = "app"
        $target = "control.exe"
        $targetArgs = "/name $rawTarget"
        $fullCommand = "control.exe /name $rawTarget"
    } elseif ($rawTarget -match '^Microsoft\.[a-zA-Z]+\\.+$') {
        $shortcutType = "appPage"
        $target = "control.exe"
        $targetArgs = "/name $($rawTarget.Split('\')[0]) /page $($rawTarget.Split('\')[1])"
        $fullCommand = "$target $targetArgs"
    } elseif ($rawTarget -match '^shell:::{[a-zA-Z0-9-]+}') {
        $shortcutType = "clsid"
        $fullCommand = "explorer $rawTarget"
    } elseif ($rawTarget -match $targetAndPathRegex) {
        $shortcutType = "pathcommand"
        $fullCommand = Fix-CommandPath $rawTarget
        $target = Fix-CommandPath $Matches[1]
        $targetArgs = Fix-CommandPath $Matches[2]
    } elseif ($rawTarget -match '^[a-zA-Z0-9]+$' -and $settingArray.Name -match '^Defender_') {
        $shortcutType = "windowsdefender"
        $fullCommand = "windowsdefender://$rawTarget"
        $target = "windowsdefender://$rawTarget"
        $name = "Windows Defender - $name"
    } else {
        $shortcutType = "assumedPath"
        if ($rawTarget -match '^(\S+)\s*(.*)$') {
            $target = $Matches[1]
            $targetArgs = $Matches[2]
        } else { $target = $rawTarget }
    }

    if ($targetArgs) { $targetArgs = [Environment]::ExpandEnvironmentVariables($targetArgs) }

    $sanitizedName = $name -replace '[\\/:*?"<>|]', '_'
    $shortcutPath = Join-Path $deepLinksOutputFolder "$sanitizedName.lnk"
    try {
        # Nutzung unserer schnellen globalen Shell-Referenz
        $shortcut = $GlobalShell.CreateShortcut($shortcutPath)

        if ($shortcutType -eq "app" -or $shortcutType -eq "appPage") {
            $shortcut.TargetPath = $target
            $shortcut.Arguments = $targetArgs
        } elseif ($shortcutType -eq "clsid") {
            $shortcut.TargetPath = "explorer.exe"
            $shortcut.Arguments = $rawTarget
        } elseif ($shortcutType -eq "pathcommand") {
            $shortcut.TargetPath = $target
            $shortcut.Arguments = $targetArgs
        } elseif ($shortcutType -eq "assumedPath" -or $shortcutType -eq "windowsdefender") {
            $shortcut.TargetPath = $target
        } else {
            $shortcut.TargetPath = $rawTarget
        }

        if ($settingArray.IconPath) { $shortcut.IconLocation = $settingArray.IconPath }
        else { $shortcut.IconLocation = "%SystemRoot%\System32\shell32.dll,-16826" }

        $shortcut.Save()
        $settingArray.FullCommand = $fullCommand
        $settingArray.ShortcutPath = $shortcutPath
        return $settingArray
    } catch {
        return $false
    }
}

function Create-Deep-Link-CSVFile {
    param (
        [string]$outputPath,
        [array]$deepLinksDataArray
    )
    $csvContent = "Description,Deep Link,Full Command,IconPath`n"
    foreach ($item in $deepLinksDataArray) {
        $description = $item.Description -replace '"', '""'
        $deepLink = $item.DeepLink -replace '"', '""'
        $fullCommand = $item.FullCommand -replace '"', '""'
        $iconPath = if ($item.IconPath) { "`"$($item.IconPath -replace '"', '""')`"" } else { "None" }
        $csvContent  += "`"$description`",`"$deepLink`",`"$fullCommand`",$iconPath`n"
    }
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Get-AppDetails-From-AppxManifest {
    param(
        [string]$CustomLanguageFolder,
        [switch]$OnlyMicrosoftApps,
        [switch]$GetExtraData
    )
    $urlProtocolData = @()
    $associatedProtocolsPerApp = @()

    foreach ($appx in Get-AppxPackage) {
        $isMicrosoft = ($appx.PublisherId -eq "8wekyb3d8bbwe") -or ($appx.PublisherId -eq "cw5n1h2txyewy")
        $appInfo = [PSCustomObject]@{
            PackageFullName = $appx.PackageFullName
            PackageName = $appx.Name
            InstallLocation = $appx.InstallLocation
            Publisher = $appx.Publisher
            PublisherId = $appx.PublisherId
            IsMicrosoft = $isMicrosoft
            Protocols = @()
        }

        if ($OnlyMicrosoftApps -and -not $isMicrosoft) { continue }

        $location = $appx.InstallLocation
        $manifestPath = "$location\AppxManifest.xml"
        if ($null -ne $location -and (Test-Path-Safe $manifestPath)) {
            [xml]$xml = Get-Content $manifestPath
            $ns = New-Object Xml.XmlNamespaceManager $xml.NameTable
            $ns.AddNamespace("main", "http://schemas.microsoft.com/appx/manifest/foundation/windows10")
            $ns.AddNamespace("uap", "http://schemas.microsoft.com/appx/manifest/uap/windows10")
            $ns.AddNamespace("uap2", "http://schemas.microsoft.com/appx/manifest/uap/windows10/2")
            $ns.AddNamespace("uap3", "http://schemas.microsoft.com/appx/manifest/uap/windows10/3")
            $ns.AddNamespace("uap4", "http://schemas.microsoft.com/appx/manifest/uap/windows10/4")
            $ns.AddNamespace("uap5", "http://schemas.microsoft.com/appx/manifest/uap/windows10/5")

            $uapNamespaces = @("uap", "uap2", "uap3", "uap4", "uap5")
            $xpathQuery = ($uapNamespaces | ForEach-Object {
                "//$_`:Extension[@Category = 'windows.protocol']/$_`:Protocol | //uap:Extension[@Category = 'windows.protocol']/$_`:Protocol"
            }) -join ' | '

            $protocolElements = $xml.SelectNodes($xpathQuery, $ns)

            foreach ($protocolElement in $protocolElements) {
                $protocol = $protocolElement.GetAttribute("Name")
                $appElement = $protocolElement.SelectSingleNode("ancestor::main:Application", $ns)
                $appId = $appElement.GetAttribute("Id")

                $displayNameElement = $appElement.SelectSingleNode(".//uap:VisualElements/@DisplayName", $ns)
                $displayName = if ($displayNameElement) { $displayNameElement.Value } else { "" }

                $descriptionElement = $appElement.SelectSingleNode(".//uap:VisualElements/@Description", $ns)
                $description = if ($descriptionElement) { $descriptionElement.Value } else { "" }

                $executable = $appElement.GetAttribute("Executable")
                $command = if ($executable) { Join-Path $location $executable } else { "" }
                $PackageName = $appx.Name

                if ($displayName -match '^ms-resource:' -and $GetExtraData) {
                    $displayName = Get-LocalizedString -StringReference $displayName -AppxManifestPath $manifestPath -CustomLanguageFolder $CustomLanguageFolder -returnSanitizedOnFail
                }
                if ($description -match '^ms-resource:' -and $GetExtraData) {
                    $description = Get-LocalizedString -StringReference $description -AppxManifestPath $manifestPath -CustomLanguageFolder $CustomLanguageFolder -returnSanitizedOnFail
                }

                $appInfo.Protocols += $protocol
                $urlProtocolData += [PSCustomObject]@{
                    Protocol = $protocol
                    Name = $displayName
                    Command = $command
                    IconPath = ""
                    DerivedIcon = ""
                    PackageName = $PackageName
                    PackageFullName = $appx.PackageFullName
                    PackageAppKeyName = $appId
                    PackageAppName = $displayName
                    PackageAppDescription = $description
                    IsMicrosoft = $isMicrosoft
                    InstallLocation = $location
                }
            }
        }
        $associatedProtocolsPerApp += $appInfo
    }
    return @{
        UrlProtocolData = ($urlProtocolData | Sort-Object -Property Protocol)
        AssociatedProtocolsPerApp = ($associatedProtocolsPerApp | Sort-Object -Property PackageName)
    }
}

function Make-DeepCopy {
    param ($object)
    $serializedData = [System.Management.Automation.PSSerializer]::Serialize($object)
    return [System.Management.Automation.PSSerializer]::Deserialize($serializedData)
}

function Get-And-Process-URL-Protocols {
    param(
        [string]$CustomLanguageFolder,
        [switch]$OnlyMicrosoftApps,
        [string[]]$permanentProtocolsIgnore,
        [switch]$GetExtraData
    )
    $urlProtocolDataOriginal = @()
    $urlProtocols = @{}

    function Get-RegistryKeyData {
        param ([Microsoft.Win32.RegistryKey]$Key)
        $data = @{
            '(Default)' = $Key.GetValue('')
            Values = @{}
            SubKeys = @{}
        }
        foreach ($valueName in $Key.GetValueNames()) {
            if ($valueName -ne '') { $data.Values[$valueName] = $Key.GetValue($valueName) }
        }
        foreach ($subKeyName in $Key.GetSubKeyNames()) {
            $subKey = $Key.OpenSubKey($subKeyName)
            $data.SubKeys[$subKeyName] = Get-RegistryKeyData -Key $subKey
            $subKey.Close()
        }
        return $data
    }

    Get-ChildItem -Path 'Registry::HKEY_CLASSES_ROOT' -ErrorAction SilentlyContinue |
        Where-Object { $_.GetValue('(Default)') -match '^URL:' -or $null -ne $_.GetValue('URL Protocol') } |
        ForEach-Object { $urlProtocols[$_.PSChildName] = Get-RegistryKeyData -Key $_.OpenSubKey('') }

    foreach ($protocol in $urlProtocols.Keys) {
        $protocolData = $urlProtocols[$protocol]
        $protocolName = $protocolData['(Default)']

        if ($protocolName -match '^URL:(.+)$') { $protocolName = $Matches[1] }

        $command = $null
        if ($protocolData.SubKeys.ContainsKey('shell') -and
            $protocolData.SubKeys['shell'].SubKeys.ContainsKey('open') -and
            $protocolData.SubKeys['shell'].SubKeys['open'].SubKeys.ContainsKey('command')) {
            $command = $protocolData.SubKeys['shell'].SubKeys['open'].SubKeys['command']['(Default)']
        }

        $iconPath = $null
        if ($protocolData.SubKeys.ContainsKey('DefaultIcon')) {
            $iconPath = $protocolData.SubKeys['DefaultIcon']['(Default)']
        }

        $urlProtocolDataOriginal += [PSCustomObject]@{
            Protocol = $protocol
            Name = $protocolName
            Command = $command
            IconPath = $iconPath
            DerivedIcon = ""
            PackageName = ""
            PackageFullName = ""
            PackageAppKeyName = ""
            PackageAppName = ""
            PackageAppDescription = ""
            IsMicrosoft = ($protocol -match '^ms-|^microsoft')
        }
    }

    $arrayProtocolAndAppxData = Get-AppDetails-From-AppxManifest -CustomLanguageFolder $CustomLanguageFolder -OnlyMicrosoftApps:$OnlyMicrosoftApps -GetExtraData:$GetExtraData
    $protocolAppxData = $arrayProtocolAndAppxData.UrlProtocolData
    $associatedProtocolsPerApp = $arrayProtocolAndAppxData.AssociatedProtocolsPerApp

    $urlProtocolDataPreferredAppx = Make-DeepCopy -object $urlProtocolDataOriginal
    foreach ($protocol in $urlProtocolDataPreferredAppx) {
        $appxData = $protocolAppxData | Where-Object { $_.Protocol -eq $protocol.Protocol }
        if ($appxData) {
            $protocol.PackageName = $appxData.PackageName
            $protocol.PackageFullName = $appxData.PackageFullName
            $protocol.PackageAppKeyName = $appxData.PackageAppKeyName
            $protocol.PackageAppName = $appxData.PackageAppName
            $protocol.PackageAppDescription = $appxData.PackageAppDescription
            $protocol.Command = $appxData.Command
            $protocol.IsMicrosoft = $appxData.IsMicrosoft
        }
    }

    $filteredUrlProtocolData = @()

    if ($OnlyMicrosoftApps) {
        $filteredUrlProtocolData = $urlProtocolDataPreferredAppx | Where-Object { $_.IsMicrosoft }
    } else {
        $filteredUrlProtocolData = $urlProtocolDataPreferredAppx | Where-Object { $permanentProtocolsIgnore -notcontains $_.Protocol }
    }

    return @{
        FilteredUrlProtocolData = $filteredUrlProtocolData
        AssociatedProtocolsPerApp = $associatedProtocolsPerApp
    }
}

function Create-URL-Protocols-CSVFile {
    param (
        [string]$outputPath,
        [array]$urlProtocolsData
    )
    $csvContent = "Protocol,Name,Command,Package Name,Package AppName,Package App Description,IsMicrosoft`n"
    foreach ($item in $urlProtocolsData) {
        $protocol = ($item.Protocol -replace '"', '""') + "://"
        $name = $item.Name -replace '"', '""'
        $command = $item.Command -replace '"', '""'
        $packageName = $item.PackageName -replace '"', '""'
        $packageAppName = $item.PackageAppName -replace '"', '""'
        $packageAppDescription = $item.PackageAppDescription -replace '"', '""'

        $csvContent += "`"$protocol`",`"$name`",`"$command`",`"$packageName`",`"$packageAppName`",`"$packageAppDescription`",$($item.IsMicrosoft)`n"
    }
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Create-Protocol-Shortcut {
    param ([string]$protocol, [string]$name, [string]$shortcutPath)
    try {
        $urlFileContent = "[InternetShortcut]`nURL=$protocol`:///`n"
        $urlFileContent | Out-File -FilePath $shortcutPath -Encoding ascii
        return $true
    } catch { return $false }
}

function Format-FileGrid {
    param (
        [array]$fileNames,
        [int]$columnsPerRow = 3,
        [int]$minSpacing = 4,
        [int]$indent = 5,
        [string]$prefix = "",
        [switch]$noSort
    )
    if ($fileNames.Count -eq 0) { return }

    $indentString = " " * $indent
    $sortedNames = if ($noSort) { $fileNames } else { $fileNames | Sort-Object }

    if ($sortedNames.Count -le $columnsPerRow) {
        $padding = " " * $minSpacing
        Write-Host ($indentString + $prefix + ($sortedNames -join $padding))
        return
    }

    $columnWidths = @()
    for ($i = 0; $i -lt $columnsPerRow; $i++) {
        $columnItems = @()
        for ($j = $i; $j -lt $sortedNames.Count; $j += $columnsPerRow) { $columnItems += $sortedNames[$j] }
        $maxLength = ($columnItems | Measure-Object Length -Maximum).Maximum
        if ($maxLength -gt 0) { $columnWidths += $maxLength }
    }

    for ($i = 0; $i -lt $sortedNames.Count; $i += $columnsPerRow) {
        $row = @()
        for ($j = 0; $j -lt $columnsPerRow; $j++) {
            if ($i + $j -lt $sortedNames.Count) {
                $item = $sortedNames[$i + $j]
                $padding = if ($j -lt ($columnWidths.Count - 1)) {
                    " " * ([Math]::Max(0, $columnWidths[$j] - $item.Length + $minSpacing))
                } else { "" }
                $row += $item + $padding
            }
        }
        Write-Host ($indentString + $prefix + ($row -join ""))
    }
}

# Einstiegsstelle für Dateiscanner
function Search-HiddenLinks {
    param (
        [Parameter(Mandatory=$true)] $associatedProtocolsPerApp,
        $URLProtocolsData,
        [switch]$SkipAppXURLSearch,
        [switch]$DeepScanHiddenLinks
    )
    if ($SkipAppXURLSearch) { return $null }

    $availableRAM = [System.Math]::Round((Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory)
    $maxFileSize = [System.Math]::Round($availableRAM * 0.70)

    if ($debugSkipAppxSearch) { Print-SuperWarning "!!! DEBUG-MODUS: Überspringe AppX-Suche !!!" }

    $ignoredExtensions = @(
        '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.ico', '.ogg', '.mp4',
        '.webm', '.webp', '.flac', '.wav', '.mp3', '.m4a', '.aac', '.wma', '.flv',
        '.avi', '.mov', '.wmv', '.mpg', '.mpeg', '.m4v', '.mkv', '.3gp', '.3g2',
        '.mxf', '.psd', '.psb', '.heif', '.heic', '.hevc', '.tiff', '.tif', '.ogv',
        '.p7x', '.ttf', '.onnxe', '.bundle', '.vsix', '.nupkg', '.asar', '.resS',
        '.bnk', '.onnx', '.opq', '.safetensors', '.cat', '.sig',
        '.7z', '.cab', '.iso', '.jar', '.rar', '.tar', '.zip', '.wim',
        '.vdi', '.vmdk', '.vhd', '.vhdx'
    )

    $encodingMap = @{
        ".txt" = "UTF-8"; ".xml" = "UTF-8"; ".json" = "UTF-8";
        ".dll" = "Unicode"; ".exe" = "Unicode"; ".js" = "UTF-8"
    }

    $packagesToSearchList = $URLProtocolsData.PackageName | Where-Object { $_ -ne "" } | Sort-Object -Unique
    $packagesToSearch = $associatedProtocolsPerApp | Where-Object { $packagesToSearchList -contains $_.PackageName }
    $packagesToSearch | Add-Member -MemberType NoteProperty -Name FilesToSearch -Value @()

    $totalFiles = 0
    [Int64]$totalFilesSize = 0
    foreach ($appPackage in $packagesToSearch) {
        $filesPerPackage = Get-ChildItem -Path $appPackage.InstallLocation -Recurse -File | Where-Object { $_.Extension -notin $ignoredExtensions -and $_.Length -gt 0 -and $_.Length -lt $maxFileSize }
        $appPackage.FilesToSearch = $filesPerPackage
        $totalFiles += $filesPerPackage.Count
        $totalFilesSize += ($filesPerPackage | Measure-Object -Property Length -Sum).Sum
    }

    if ($debugSkipAppxSearch) {
        $packagesToSearch = @()
    }

    # ------------------ Suchprozess starten ------------------
    $processedFiles = 0
    [Int64]$processedFilesSize = 0
    $resultsAppx = @()
    if ($Verbose -or $timing) { $stopwatch = [System.Diagnostics.Stopwatch]::StartNew() }
    
    # LERNEFFEKT: Aufruf des parallelisierten Suchers!
    Write-Host "[1/2] Durchsuche Appx-Programmdateien nach versteckten Links:"
    foreach ($appPackage in $packagesToSearch) {
        $appxPackageResult, $processedFiles, $processedFilesSize = Get-ProtocolsInProgramFiles -encodingMap $encodingMap -program $appPackage -totalFiles $totalFiles -processedFiles $processedFiles -totalFilesSize $totalFilesSize -processedFilesSize $processedFilesSize
        $resultsAppx += $appxPackageResult
    }
    Write-Host ""
    if ($Verbose -or $timing) { $stopwatch.Stop(); Write-Host "   > Laufzeit für Appx-Suche: $($stopwatch.Elapsed)" -ForegroundColor Green }

    $programFilesSearchData = @()
    foreach ($protocol in $URLProtocolsData | Where-Object { $_.Protocol -notin $resultsAppx.Protocol }) {
        foreach ($command in $protocol.Command) {
            $target = ""
            $installDir = ""

            if ($command) {
                $foundValidTarget = $false

                if (-not $foundValidTarget -and $command -match '(?i)[a-z]:\\(?:[^\\/:*?"<>|\r\n]+\\)*[^\\/:*?"<>|\r\n]+\.[^\\/:*?"<>|\r\n ]+') {
                    $target = $Matches[0]
                    if (Test-Path-Safe -Path $target) { $foundValidTarget = $true }
                }
                if (-not $foundValidTarget -and $command -match '^"([^"]+)"?') {
                    $target = $Matches[1].Trim()
                    if (Test-Path-Safe $target) { $foundValidTarget = $true }
                }
                if (-not $foundValidTarget) {
                    if (Test-Path-Safe $command) { $target = $command; $foundValidTarget = $true }
                }

                if (-not $foundValidTarget) { continue }

                $moreRestrictiveProgramFilesPaths = @("*\Common Files\*", "*\WindowsApps\*", "*\Adobe\*", "*\Microsoft Office\*")

                if ($target -match '^(.*\\[Pp]rogram [Ff]iles( \(x86\))?\\[^\\]+)') {
                    $installDir = $Matches[1]
                    $containsRestrictivePath = $false
                    foreach ($exception in $moreRestrictiveProgramFilesPaths) {
                        if (($installDir -ilike "$exception") -or ($($installDir + "\") -ilike "$exception")) {
                            $containsRestrictivePath = $true
                            break
                        }
                    }
                    if ($containsRestrictivePath) { $installDir = Split-Path -Path $target -Parent }
                } else {
                    $installDir = Split-Path -Path $target -Parent
                }

                $searchItem = [PSCustomObject]@{
                    PackageFullName = $protocol.PackageFullName
                    PackageName     = $protocol.PackageName
                    InstallLocation = $installDir
                    Publisher       = "N/A"
                    PublisherId     = "N/A"
                    IsMicrosoft     = $protocol.IsMicrosoft
                    Protocols       = @($protocol.Protocol)
                    FilesToSearch   = @()
                    Command         = $command
                    Target          = $target
                }
                $programFilesSearchData += $searchItem
            }
        }
    }

    $totalFiles = 0
    [Int64]$totalFilesSize = 0
    Write-Host "`nArbeite Liste der regulären Installationsordner ab...`n"
    foreach ($program in $programFilesSearchData) {
        $files = @()
        $folder = $program.InstallLocation
        if ($null -ne $folder -and $DeepScanHiddenLinks) {
            try {
                $files = Get-ChildItem -Path $folder -Recurse -File -Force -ErrorAction SilentlyContinue | Where-Object { 
                    $_.Length -lt $maxFileSize -and
                    $_.Extension -notin $ignoredExtensions -and 
                    $_.Length -gt 0
                }
                $program.FilesToSearch = $files
            } catch {
                $program.InstallLocation = $null
                $program.FilesToSearch = @($program.Target | Get-Item)
            }
        } else {
            try {
                $files = @($program.Target | Get-Item)
                $program.FilesToSearch = $files
            } catch { continue }
        }
        $totalFiles += $files.Count
        $totalFilesSize += ($files | Measure-Object -Property Length -Sum).Sum
    }

    $processedFiles = 0
    [Int64]$processedFilesSize = 0
    $resultsNonAppx = @()
    if ($Verbose -or $timing) { $stopwatch = [System.Diagnostics.Stopwatch]::StartNew() }

    Write-Host "`n[2/2] Durchsuche Nicht-Appx Programmdateien:"
    foreach ($itemToSearch in $programFilesSearchData) {
        $programResult, $processedFiles, $processedFilesSize = Get-ProtocolsInProgramFiles -encodingMap $encodingMap -program $itemToSearch -totalFiles $totalFiles -processedFiles $processedFiles -totalFilesSize $totalFilesSize -processedFilesSize $processedFilesSize
        $resultsNonAppx += $programResult
    }
    Write-Host ""
    if ($Verbose -or $timing) { $stopwatch.Stop(); Write-Host "   > Laufzeit für reguläre Programmsuche: $($stopwatch.Elapsed)" -ForegroundColor Green }

    $resultsUniqueFullURL = $resultsAppx | Sort-Object -Property FullURL -Unique
    if ($resultsNonAppx.Count -gt 0) {
        foreach ($result in $resultsNonAppx) {
            if ($resultsUniqueFullURL.FullURL -notcontains $result.FullURL) { $resultsUniqueFullURL += $result }
        }
    }
    return $resultsUniqueFullURL
}

# ==============================================================================================================================
# ========================  [OPTIMIERUNG 1&3: MULTITHREADING-SCAN MIT WRITE-PROGRESS PROGRESSBAR]  ============================
# ==============================================================================================================================
# LERNEFFEKT: Hier verschmilzt modernste CPU-Parallelisierung mit einer sauberen Win7+ UI!
#
# Was passiert hier?
# - 'ForEach-Object -Parallel' liest unter Windows 7/8/10/11 verschiedene Dateistapel (Files) auf getrennten CPU-Cores ein!
# - Da wir multithreading-basiert arbeiten, können wir nicht einfach ein flackerndes 'Write-Host \r' nutzen, 
#   weil sich die verschiedenen Threads gegenseitig in die Ausgabe pfuschen würden.
# - Stattdessen nutzen wir das thread-sichere 'Write-Progress', das Windows-intern einen extrem schicken,
#   blauen GUI-Fortschrittsbalken am oberen Bildschirmrand rendert. Absolut flackerfrei und hochprofessionell!
function Get-ProtocolsInProgramFiles {
    param (
        [hashtable]$encodingMap,
        [PSCustomObject]$program,
        [int]$totalFiles,
        [int]$processedFiles,
        [Int64]$totalFilesSize = 0,
        [Int64]$processedFilesSize = 0
    )
    $protocolsList = $program.Protocols
    $filesToSearch = $program.FilesToSearch

    # LERNEFFEKT: Wir erstellen eine Thread-sichere, synchronisierte Liste, in der alle Threads ungehindert parallel
    # ihre Suchtreffer abspeichern können, ohne dass es zu Schreibblockaden (Ressourcenkonflikten) kommt.
    $synchronizedResults = [System.Collections.Generic.List[object]]::Synchronized((New-Object System.Collections.Generic.List[object]))

    $currentPercentage = [math]::Floor(($processedFiles / $totalFiles) * 100)

    # LERNEFFEKT: Regulärer Ausdrucks-Katalog vorab für Multithread-Zugriff füttern
    $regexOptions = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    $utf8Regex = @{}
    $unicodeRegex = @{}
    foreach ($protocol in $protocolsList) {
        $utf8Regex[$protocol] = [regex]::Escape($protocol) + "://[^`"\s<>()\\``\x00-\x1F\x7F]+"
        $unicodeRegex[$protocol] = [regex]::Escape($protocol) + "://[ -~]+"
    }

    # Zähler-Threadvariable initialisieren, um den Stand über die Threads hinweg atomar hochzuzählen
    $runningCount = 0

    # LERNEFFEKT: Start der Parallel-Engine!
    # '-Parallel' verteilt die Last automatisch auf alle verfügbaren Kerne deines Prozessors.
    # Wir übergeben die Regexes und Variablen via '$using:' an die abgeschotteten Worker-Threads.
    $filesToSearch | ForEach-Object -Parallel {
        $file = $_
        try {
            $null = Test-Path -LiteralPath $file.FullName -PathType Leaf -ErrorAction Stop
        } catch { return } # Äquivalent zu 'continue' im Thread

        $fileExtension = $file.Extension
        $encodingsToTry = if ($using:encodingMap.ContainsKey($fileExtension)) { @($using:encodingMap[$fileExtension]) } else { @("UTF-8", "Unicode") }

        foreach ($encodingName in $encodingsToTry) {
            $encoding = [System.Text.Encoding]::GetEncoding($encodingName)
            try {
                $content = [System.IO.File]::ReadAllText($file.FullName, $encoding)
                foreach ($protocol in $using:protocolsList) {
                    $uriPattern = if ($encodingName -eq "UTF-8") { $using:utf8Regex[$protocol] } else { $using:unicodeRegex[$protocol] }
                    $URImatches = [regex]::Matches($content, $uriPattern, $using:regexOptions)
                    
                    if ($URImatches.Count -gt 0) {
                        foreach ($match in $URImatches) {
                            $foundItem = [PSCustomObject]@{
                                FullURL       = $match.Value -replace '".*$', ''
                                EncodingUsed  = $encodingName
                                UsesVariables = $match.Value -match "[<>()\[\]{}]|=$|:$"
                                FilePath      = $file.FullName
                                Protocol      = $protocol
                            }
                            # Sicherer Eintrag in synchronisierte Liste
                            [void]$using:synchronizedResults.Add($foundItem)
                        }
                    }
                }
            } catch {}
        }
    }

    # Update der allgemeinen globalen Zähler
    $processedFiles += $filesToSearch.Count
    $currentPercentage = [math]::Floor(($processedFiles / $totalFiles) * 100)

    # LERNEFFEKT: Zeigt den flackerfreien blauen Balken im Terminal! Keine zerschossene Textausgabe mehr!
    Write-Progress -Activity "Suche versteckte App-Links (Multithread-Speed)" -Status "Verarbeite Datei-Stapel... ($processedFiles / $totalFiles)" -PercentComplete $currentPercentage

    return $synchronizedResults, $processedFiles, $processedFilesSize
}

function Create-AppXURLSearchResultsToCSV {
    param (
        [Array]$urlItemsData,
        [string]$outputPath
    )
    $csvContent = "Protocol,Full URL,Encoding Used,Embedded Variables, Found in File`n"
    foreach ($urlItem in $urlItemsData) {
        $protocol = $urlItem.Protocol -replace '"', '""'
        $fullURL = $urlItem.FullURL -replace '"', '""'
        $encodingUsed = $urlItem.EncodingUsed -replace '"', '""'
        $usesVariables = $urlItem.UsesVariables -replace '"', '""'
        $filePath = $urlItem.FilePath -replace '"', '""'

        $csvContent += "`"$protocol`",`"$fullURL`",`"$encodingUsed`",`"$usesVariables`",`"$filePath`"`n"
    }
    $csvContent | Out-File -FilePath $outputPath -Encoding utf8
}

function Create-AppXURLShortcuts {
    param (
        [Array]$foundURLsItems,
        [string]$shortcutFolder
    )
    $createdShortcuts = @()

    foreach ($urlItem in $foundURLsItems) {
        if ($urlItem.UsesVariables) { continue }

        $fullURL = $urlItem.FullURL
        $simplifiedURL = $fullURL -replace '\?.*$', ''

        $sanitizedURLForName = $simplifiedURL -replace '://', ' - '
        $sanitizedURLForName = $sanitizedURLForName -replace '[\\/:*?"<>|]', '_'
        $sanitizedURLForName = $sanitizedURLForName -replace '_+$', ''

        $i = 2
        $originalSanitizedURLForName = $sanitizedURLForName
        while ($createdShortcuts -contains $sanitizedURLForName) {
            $sanitizedURLForName = "$originalSanitizedURLForName ($i)"
            $i++
        }

        $shortcutPath = Join-Path $shortcutFolder "$sanitizedURLForName.url"
        try {
            $urlFileContent = "[InternetShortcut]`nURL=$fullURL`n"
            $urlFileContent | Out-File -FilePath $shortcutPath -Encoding ascii
            $createdShortcuts += $sanitizedURLForName
        }
        catch {}
    }
}

# =============================================================================================================================
# ==================================================  HAUPTSKRIPTLAUF   =======================================================
# =============================================================================================================================

$clsidInfo = @()
$namedFolders = @()
$taskLinks = @()
$settingsData = @()
$msSettingsList = @()
$deepLinkData = @()
$deepLinksProcessedData = @()
$URLProtocolsData = @()

# Block 1: Analysiere CLSIDs
if (-not $SkipCLSID) {
    Write-Host "`n--- Analysiere System-CLSIDs (Objektkatalog) ---"
    try {
        $shellFolders = Get-ChildItem -Path 'Registry::HKEY_CLASSES_ROOT\CLSID' |
        Where-Object { $_.GetSubKeyNames() -contains "ShellFolder" } |
        Select-Object PSChildName
    } catch {
        Write-Error "Auslesen des CLSID Pfads fehlgeschlagen: $_"
        $shellFolders = $null
    }

    $usedNames = @()

    foreach ($folder in $shellFolders) {
        $clsid = $folder.PSChildName
        $resultArray = Get-FolderName -clsid $clsid -CustomLanguageFolder $CustomLanguageFolderPath
        $name = $resultArray[0]
        $nameSource = $resultArray[1]

        $sanitizedName = $name -replace '[\\/:*?"<>|]', '_'

        $i = 2
        $originalSanitizedName = $sanitizedName
        while ($usedNames -contains $sanitizedName) {
            $sanitizedName = "$originalSanitizedName ($i)"
            $i++
        }
        $usedNames += $sanitizedName

        $shortcutPath = Join-Path $CLSIDshortcutsOutputFolder "$sanitizedName.lnk"
        $success = Create-CLSID-Shortcut -clsid $clsid -name $name -shortcutPath $shortcutPath

        if ($success) {
            Write-Host "Erfolgreich erstellt (CLSID): $name"
        }

        $iconPath = (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\CLSID\$clsid\DefaultIcon" -ErrorAction SilentlyContinue).'(default)'
        $clsidInfo += [PSCustomObject]@{
            CLSID = $clsid
            Name = $name
            NameSource = $nameSource
            IconPath = $iconPath
        }
    }
}

# Block 2: Sonderordner (Namensbasiert)
if (-not $SkipNamedFolders) {
    Write-Host "`n--- Verarbeite namentliche System-Sonderordner ---"
    try {
        $namedFolders = Get-ChildItem -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions"
    } catch {
        $namedFolders = $null
    }

    foreach ($folder in $namedFolders) {
        $folderProperties = Get-ItemProperty -Path $folder.PSPath
        $folderName = $folderProperties.Name
        $iconPath = $folderProperties.Icon

        if ($folderName) {
            $sanitizedName = $folderName -replace '[\\/:*?"<>|]', '_'
            $shortcutPath = Join-Path $namedShortcutsOutputFolder "$sanitizedName.lnk"
            $success = Create-NamedShortcut -name $folderName -shortcutPath $shortcutPath -iconPath $iconPath

            if ($success) {
                Write-Host "Erfolgreich erstellt (Sonderordner): $folderName"
            }
        }
    }
}

# Block 3: Aufgabenlinks
if (-not $SkipTaskLinks) {
    Write-Host "`n --- Verarbeite klassische System-Aufgabenlinks ---"
    $taskLinks = Get-TaskLinks -SaveXML:(!$NoStatistics) -DLLPath:$CustomDLLPath -CustomLanguageFolder:$CustomLanguageFolderPath
    $createdShortcutNames = @{}

    if ($taskLinks) {
        $taskLinks = $taskLinks | Sort-Object -Property Name
    }
    
    foreach ($task in $taskLinks) {
        $originalName = $task.Name
        $sanitizedName = ""

        if (-not $DontGroupTasks) {
            if ($UseAlternativeCategoryNames) {
                $trueApplicationName = (Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\$($task.ApplicationId)" -ErrorAction SilentlyContinue)."(Default)"
                if ($trueApplicationName) {
                    $sanitizedName = "$trueApplicationName - $originalName"
                }
            }

            if ($task.ApplicationName -and -not $sanitizedName) {
                $sanitizedName = Prettify-App-Name -AppName $task.ApplicationName -TaskName $originalName
            } elseif (-not $task.ApplicationName -and -not $sanitizedName) {
                $sanitizedName = $originalName -replace '[\\/:*?"<>|]', '_'
            }
        } else {
            $sanitizedName = $originalName -replace '[\\/:*?"<>|]', '_'
        }

        $nameCounter = 1
        $uniqueName = $sanitizedName
        while ($createdShortcutNames.ContainsKey($uniqueName)) {
            $nameCounter++
            $uniqueName = "${sanitizedName} ($nameCounter)"
        }

        if ($task.Command) {
            $command = $task.Command
        } elseif ($task.ApplicationName -and $task.Page) {
            $command = "control.exe /name $($task.ApplicationName) /page $($task.Page)"
        } else { continue }

        $shortcutType = if ($command -match '^[a-zA-Z0-9]+:\/\/') { "url" } else { "lnk" }

        $shortcutPath = Join-Path $taskLinksOutputFolder "$uniqueName.$shortcutType"
        $createdShortcutNames[$uniqueName] = $true

        $success = Create-TaskLink-Shortcut -name $uniqueName -shortcutPath $shortcutPath -shortcutType $shortcutType -command $command -controlPanelName $task.ControlPanelName -applicationId $task.ApplicationId -keywords $task.Keywords

        if ($success) {
            Write-Host "Erfolgreich erstellt (Task): $uniqueName"
        }
    }
}

# Block 4: Deep Links
if (-not $SkipDeepLinks -and $allSettingsXmlPath) {
    Write-Host "`n--- Analysiere Deep Links ---"
    $deepLinkData = Get-AllSettings-Data -xmlFilePath $allSettingsXmlPath -SaveXML:(!$NoStatistics)

    foreach ($deepLink in $deepLinkData) {
        $existingTaskLink = "" 
        if ($deepLink.DeepLink) {
            if (-not $AllowDuplicateDeepLinks -and -not $SkipTaskLinks) {
                foreach ($taskLink in $taskLinks) {
                    $trimmedCommand = $taskLink.Command.Trim()
                    $trimmedDeepLink = $deepLink.DeepLink.Trim()

                    if (($trimmedCommand -eq $trimmedDeepLink) -or ($trimmedDeepLink -eq "$($taskLink.ApplicationName)\$($taskLink.Page)".Trim("\\"))) {
                        $existingTaskLink = $taskLink
                        break
                    }
                }
                if ($existingTaskLink) { continue }
            }

            $result = Create-Deep-Link-Shortcut -settingArray $deepLink
            if ($result) {
                Write-Host "Erfolgreich erstellt (Deep-Link): $($deepLink.Description)"
                $deepLinksProcessedData += $result
            }
        }
    }
}

# Block 5: Moderne MS Settings
if (-not $SkipMSSettings) {
    Write-Host "`n--- Analysiere Moderne Windows-Einstellungen (ms-settings:) ---"
    $msSettingsList = Get-MS-SettingsFrom-SystemSettingsDLL -DllPath $systemSettingsDllPath

    foreach ($setting in $msSettingsList) {
        $settingName = $setting.Split(':')[1]
        $fullShortcutName = $setting
        $sanitizedName = $settingName -replace '[\\/:*?"<>|]', '_'
        $shortcutPath = Join-Path $msSettingsOutputFolder "$sanitizedName.lnk"

        $success = Create-MSSettings-Shortcut -fullName $fullShortcutName -shortcutPath $shortcutPath
        if ($success) {
            Write-Host "Erfolgreich erstellt (MS-Settings): $fullShortcutName"
        }
    }
}

# Block 6: URL Internetprotokolle
if (-not $SkipURLProtocols){
    Write-Host "`n--- Analysiere im Betriebssystem verankerte URL-Protokolle ---"
    $OnlyMicrosoftApps = if ($AllURLProtocols) { $false } else { $true }

    $protocolAndAppxData = Get-And-Process-URL-Protocols -CustomLanguageFolder $CustomLanguageFolderPath -OnlyMicrosoftApps:$OnlyMicrosoftApps -permanentProtocolsIgnore $permanentURIProtocols -GetExtraData:$CollectExtraURLProtocolInfo
    $URLProtocolsData = $protocolAndAppxData.FilteredUrlProtocolData
    $associatedProtocolsPerApp = $protocolAndAppxData.AssociatedProtocolsPerApp

    foreach ($protocol in $URLProtocolsData) {
         $success = Create-Protocol-Shortcut -protocol $protocol.Protocol -name $protocol.Name -shortcutPath (Join-Path $URLProtocolLinksOutputFolder "$($protocol.Protocol).url")
         if ($success) {
             Write-Host "Erfolgreich erstellt (URI): $($protocol.Protocol)"
         }
    }

    if (-not $SkipHiddenAppLinks) {
        Write-Host "`n--- Suche nach versteckten, tiefen App-Sublinks ---"
        $appXURLSearchResults = Search-HiddenLinks -associatedProtocolsPerApp $associatedProtocolsPerApp -URLProtocolsData $URLProtocolsData -SkipAppXURLSearch:$SkipHiddenAppLinks -DeepScanHiddenLinks:$DeepScanHiddenLinks

        if ($appXURLSearchResults) {
            Create-AppXURLShortcuts -foundURLsItems $appXURLSearchResults -shortcutFolder $URLProtocolPageLinksOutputFolder
        }
    }
}

# ==============================================================================================================================
# ========================================  SPEICHERUNG STATISTIK-REPORTS  ======================================================
# ==============================================================================================================================

if (-not $NoStatistics) {
    if (-not $SkipCLSID) { Create-CLSIDCsvFile -outputPath $clsidCsvPath -clsidData $clsidInfo }
    if (-not $SkipNamedFolders) { Create-NamedFoldersCsvFile -outputPath $namedFoldersCsvPath }
    if (-not $SkipTaskLinks) { Create-TaskLinksCsvFile -outputPath $taskLinksCsvPath -taskLinksData $taskLinks }
    if (-not $SkipMSSettings) { Create-MSSettingsCsvFile -outputPath $msSettingsCsvPath -msSettingsList $msSettingsList }
    if (-not $SkipDeepLinks) { Create-Deep-Link-CSVFile -outputPath $deepLinksCsvPath -deepLinksDataArray $deepLinksProcessedData }
    if (-not $SkipURLProtocols) { Create-URL-Protocols-CSVFile -outputPath $URLProtocolLinksCsvPath -urlProtocolsData $URLProtocolsData }
    if (-not $SkipHiddenAppLinks -and -not $SkipURLProtocols) { Create-AppXURLSearchResultsToCSV -urlItemsData $appXURLSearchResults -outputPath $URLProtocolPageLinksCsvPath }
}

if (-not $NoStatistics) {
    $displayCsvFiles = @($csvFiles.Values | Where-Object { -not $_.Skip } | ForEach-Object { $_.Value })
    $displayXmlFiles = @($xmlFiles.Values | Where-Object { -not $_.Skip } | ForEach-Object { $_.Value })
}

if ($displayCsvFiles -or $displayXmlFiles) {
    Write-Host "`n--------------------------------------------------------------------------------"
    Write-Host "Statistiken und XML-Protokolldaten abgelegt im Unterordner: `"$statisticsFolderName`""

    if ($displayCsvFiles) {
        Write-Host "`n   - CSV-Berichte:"
        Format-FileGrid -fileNames $displayCsvFiles -Indent 7
    }
    if ($displayXmlFiles) {
        Write-Host "`n   - XML-Strukturen:"
        Format-FileGrid -fileNames $displayXmlFiles -Indent 7
    }
}

# ==============================================================================================================================
# ===================================  [ABSPANN: SAUBERES ENTTEILEN DER RESSOURCEN]  ===========================================
# ==============================================================================================================================
# LERNEFFEKT: Windows sperrt COM-Objekt-Dateien im Hintergrund, solange die PowerShell-Prozedur läuft.
# Indem we das globale Objekt am Schluss explizit freigeben ("ReleaseComObject"), verhindern wir festsitzende 
# Geister-Prozesse (wie z.B. eine unsichtbare, blockierte explorer.exe-Instanz im Task-Manager).
if ($null -ne $GlobalShell) {
    Write-Verbose "[CLEANUP] Gebe globale Shell-Schnittstelle wieder für das System frei..."
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($GlobalShell) | Out-Null
}

$appXURLSearchResultsCreated = $appXURLSearchResults | Where-Object { -not $_.UsesVariables }
$totalCount = $clsidInfo.Count + $namedFolders.Count + $taskLinks.Count + $msSettingsList.Count + $deepLinksProcessedData.Count + $URLProtocolsData.Count + $appXURLSearchResultsCreated.Count

Write-Host "`n------------------------------------------------"
Write-Host   "    Endergebnisse: Windows Super God Mode       " -ForeGroundColor Yellow
Write-Host   "------------------------------------------------`n"

Write-Host "         Erstellte Shortcuts Gesamt: " -NoNewline
Write-Host $totalCount -ForegroundColor Green

Write-Host "           > CLSID System-Links:     " -NoNewline
Write-Host $clsidInfo.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipCLSID) { "   (Übersprungen)" })

Write-Host "           > Benannte Ordner:        " -NoNewline
Write-Host $namedFolders.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipNamedFolders) { "   (Übersprungen)" })

Write-Host "           > Aufgabenlinks:          " -NoNewline
Write-Host $taskLinks.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipTaskLinks) { "   (Übersprungen)" })

Write-Host "           > Systemeinstellungen:    " -NoNewline
Write-Host $msSettingsList.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipMSSettings) { "   (Übersprungen)" })

Write-Host "           > Deep Links:             " -NoNewline
Write-Host $deepLinksProcessedData.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipDeepLinks) { "   (Übersprungen)" })

Write-Host "           > URL Protokolle:         " -NoNewline
Write-Host $URLProtocolsData.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipURLProtocols) { "   (Übersprungen)" })

Write-Host "           > Versteckte App-Links:   " -NoNewline
Write-Host $appXURLSearchResultsCreated.Count -ForegroundColor Cyan -NoNewline
Write-Host $(if ($SkipHiddenAppLinks -or $SkipURLProtocols) { "   (Übersprungen)" })

Write-Host "`n------------------------------------------------`n"

if ($Debug) { Stop-Transcript }

if (-not $NoGUI) {
    if (-not $psISE) {
        Write-Host "Beliebige Taste drücken zum Beenden...`n" -NoNewline
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } else {
        Read-Host "Drücke Enter zum Beenden..."
    }
}
