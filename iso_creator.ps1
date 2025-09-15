Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Include the original New-IsoFile function


function New-IsoFile 
{  
  [CmdletBinding(DefaultParameterSetName='Source')]Param( 
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$true, ParameterSetName='Source')]$Source,  
    [parameter(Position=2)][string]$Path = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss.ffff')).iso",  
    [ValidateScript({Test-Path -LiteralPath $_ -PathType Leaf})][string]$BootFile = $null, 
    [ValidateSet('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')][string] $Media = 'DVDPLUSRW_DUALLAYER', 
    [string]$Title = (Get-Date).ToString("yyyyMMdd-HHmmss.ffff"),  
    [switch]$Force, 
    [parameter(ParameterSetName='Clipboard')][switch]$FromClipboard 
  ) 
  
  Begin {  
    ($cp = new-object System.CodeDom.Compiler.CompilerParameters).CompilerOptions = '/unsafe' 
    if (!('ISOFile' -as [type])) {  
      Add-Type -CompilerParameters $cp -TypeDefinition @'
public class ISOFile  
{ 
  public unsafe static void Create(string Path, object Stream, int BlockSize, int TotalBlocks)  
  {  
    int bytes = 0;  
    byte[] buf = new byte[BlockSize];  
    var ptr = (System.IntPtr)(&bytes);  
    var o = System.IO.File.OpenWrite(Path);  
    var i = Stream as System.Runtime.InteropServices.ComTypes.IStream;  
   
    if (o != null) { 
      while (TotalBlocks-- > 0) {  
        i.Read(buf, BlockSize, ptr); o.Write(buf, 0, bytes);  
      }  
      o.Flush(); o.Close();  
    } 
  } 
}  
'@  
    } 
   
    if ($BootFile) { 
      if('BDR','BDRE' -contains $Media) { Write-Warning "Bootable image doesn't seem to work with media type $Media" } 
      ($Stream = New-Object -ComObject ADODB.Stream -Property @{Type=1}).Open()  # adFileTypeBinary 
      $Stream.LoadFromFile((Get-Item -LiteralPath $BootFile).Fullname) 
      ($Boot = New-Object -ComObject IMAPI2FS.BootOptions).AssignBootImage($Stream) 
    } 
  
    $MediaType = @('UNKNOWN','CDROM','CDR','CDRW','DVDROM','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','HDDVDROM','HDDVDR','HDDVDRAM','BDROM','BDR','BDRE') 
  
    Write-Verbose -Message "Selected media type is $Media with value $($MediaType.IndexOf($Media))"
    ($Image = New-Object -com IMAPI2FS.MsftFileSystemImage -Property @{VolumeName=$Title}).ChooseImageDefaultsForMediaType($MediaType.IndexOf($Media)) 
   
    if (!($Target = New-Item -Path $Path -ItemType File -Force:$Force -ErrorAction SilentlyContinue)) { Write-Error -Message "Cannot create file $Path. Use -Force parameter to overwrite if the target file already exists."; break } 
  }  
  
  Process { 
    if($FromClipboard) { 
      if($PSVersionTable.PSVersion.Major -lt 5) { Write-Error -Message 'The -FromClipboard parameter is only supported on PowerShell v5 or higher'; break } 
      $Source = Get-Clipboard -Format FileDropList 
    } 
  
    foreach($item in $Source) { 
      if($item -isnot [System.IO.FileInfo] -and $item -isnot [System.IO.DirectoryInfo]) { 
        $item = Get-Item -LiteralPath $item
      } 
  
      if($item) { 
        Write-Verbose -Message "Adding item to the target image: $($item.FullName)"
        try { $Image.Root.AddTree($item.FullName, $true) } catch { Write-Error -Message ($_.Exception.Message.Trim() + ' Try a different media type.') } 
      } 
    } 
  } 
  
  End {  
    if ($Boot) { $Image.BootImageOptions=$Boot }  
    $Result = $Image.CreateResultImage()  
    [ISOFile]::Create($Target.FullName,$Result.ImageStream,$Result.BlockSize,$Result.TotalBlocks) 
    Write-Verbose -Message "Target image ($($Target.FullName)) has been created"
    $Target
  } 
}

# XAML for the GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="ISO File Creator" 
        Height="600" Width="800"
        ResizeMode="CanResize"
        WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="100"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Title -->
        <TextBlock Grid.Row="0" Text="ISO File Creator" FontSize="24" FontWeight="Bold" 
                   HorizontalAlignment="Center" Margin="0,0,0,20"/>

        <!-- Source Files/Folders -->
        <GroupBox Grid.Row="1" Header="Source Files and Folders" Margin="0,0,0,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <ListBox Name="lstSources" Grid.Row="0" Margin="5"
                         ScrollViewer.HorizontalScrollBarVisibility="Auto"/>
                
                <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="5">
                    <Button Name="btnAddFiles" Content="Add Files" Width="80" Margin="0,0,5,0"/>
                    <Button Name="btnAddFolders" Content="Add Folders" Width="80" Margin="0,0,5,0"/>
                    <Button Name="btnFromClipboard" Content="From Clipboard" Width="100" Margin="0,0,5,0"/>
                    <Button Name="btnClearSources" Content="Clear All" Width="80" Margin="0,0,5,0"/>
                    <Button Name="btnRemoveSelected" Content="Remove Selected" Width="120"/>
                </StackPanel>
            </Grid>
        </GroupBox>

        <!-- Output Path -->
        <GroupBox Grid.Row="2" Header="Output ISO File" Margin="0,0,0,10">
            <Grid Margin="5">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Name="txtOutputPath" Grid.Column="0" Margin="0,0,5,0"/>
                <Button Name="btnBrowseSave" Grid.Column="1" Content="Browse..." Width="80"/>
            </Grid>
        </GroupBox>

        <!-- ISO Title -->
        <GroupBox Grid.Row="3" Header="ISO Volume Title" Margin="0,0,0,10">
            <TextBox Name="txtTitle" Margin="5"/>
        </GroupBox>

        <!-- Media Type -->
        <GroupBox Grid.Row="4" Header="Media Type" Margin="0,0,0,10">
            <ComboBox Name="cmbMediaType" Margin="5"/>
        </GroupBox>

        <!-- Boot File (Optional) -->
        <GroupBox Grid.Row="5" Header="Boot File (Optional)" Margin="0,0,0,10">
            <Grid Margin="5">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox Name="txtBootFile" Grid.Column="0" Margin="0,0,5,0"/>
                <Button Name="btnBrowseBoot" Grid.Column="1" Content="Browse..." Width="80"/>
            </Grid>
        </GroupBox>

        <!-- Progress and Output -->
        <GroupBox Grid.Row="6" Header="Progress and Output" Margin="0,0,0,10">
            <Grid Margin="5">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <ProgressBar Name="progressBar" Grid.Row="0" Height="20" Margin="0,0,0,5"/>
                <TextBox Name="txtOutput" Grid.Row="1" IsReadOnly="True" 
                         VerticalScrollBarVisibility="Auto" 
                         TextWrapping="Wrap" FontFamily="Consolas"/>
            </Grid>
        </GroupBox>

        <!-- Control Buttons -->
        <StackPanel Grid.Row="7" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,0">
            <CheckBox Name="chkForce" Content="Force Overwrite" Margin="0,0,20,0" VerticalAlignment="Center"/>
            <Button Name="btnCreateISO" Content="Create ISO" Width="100" Height="30" 
                    FontWeight="Bold" Background="LightGreen" Margin="0,0,10,0"/>
            <Button Name="btnExit" Content="Exit" Width="80" Height="30"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Parse XAML and create window
$reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$lstSources = $window.FindName("lstSources")
$btnAddFiles = $window.FindName("btnAddFiles")
$btnAddFolders = $window.FindName("btnAddFolders")
$btnFromClipboard = $window.FindName("btnFromClipboard")
$btnClearSources = $window.FindName("btnClearSources")
$btnRemoveSelected = $window.FindName("btnRemoveSelected")
$txtOutputPath = $window.FindName("txtOutputPath")
$btnBrowseSave = $window.FindName("btnBrowseSave")
$txtTitle = $window.FindName("txtTitle")
$cmbMediaType = $window.FindName("cmbMediaType")
$txtBootFile = $window.FindName("txtBootFile")
$btnBrowseBoot = $window.FindName("btnBrowseBoot")
$progressBar = $window.FindName("progressBar")
$txtOutput = $window.FindName("txtOutput")
$chkForce = $window.FindName("chkForce")
$btnCreateISO = $window.FindName("btnCreateISO")
$btnExit = $window.FindName("btnExit")

# Initialize controls
$mediaTypes = @('CDR','CDRW','DVDRAM','DVDPLUSR','DVDPLUSRW','DVDPLUSR_DUALLAYER','DVDDASHR','DVDDASHRW','DVDDASHR_DUALLAYER','DISK','DVDPLUSRW_DUALLAYER','BDR','BDRE')
foreach ($media in $mediaTypes) {
    $cmbMediaType.Items.Add($media) | Out-Null
}
$cmbMediaType.SelectedItem = 'DVDPLUSRW_DUALLAYER'

# Set default values
$txtOutputPath.Text = "$env:temp\$((Get-Date).ToString('yyyyMMdd-HHmmss')).iso"
$txtTitle.Text = (Get-Date).ToString("yyyyMMdd-HHmmss")

# Helper function to add output text
function Add-OutputText {
    param([string]$Text)
    $txtOutput.AppendText("$Text`r`n")
    $txtOutput.ScrollToEnd()
    [System.Windows.Forms.Application]::DoEvents()
}

# Event handlers
$btnAddFiles.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Multiselect = $true
    $openFileDialog.Title = "Select Files to Add to ISO"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($file in $openFileDialog.FileNames) {
            if ($lstSources.Items -notcontains $file) {
                $lstSources.Items.Add($file) | Out-Null
                Add-OutputText "Added file: $file"
            }
        }
    }
})

$btnAddFolders.Add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Select Folder to Add to ISO"
    
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $folder = $folderBrowserDialog.SelectedPath
        if ($lstSources.Items -notcontains $folder) {
            $lstSources.Items.Add($folder) | Out-Null
            Add-OutputText "Added folder: $folder"
        }
    }
})

$btnFromClipboard.Add_Click({
    try {
        if ($PSVersionTable.PSVersion.Major -lt 5) {
            Add-OutputText "ERROR: The From Clipboard feature requires PowerShell v5 or higher"
            return
        }
        
        $clipboardItems = Get-Clipboard -Format FileDropList -ErrorAction Stop
        if ($clipboardItems) {
            foreach ($item in $clipboardItems) {
                if ($lstSources.Items -notcontains $item) {
                    $lstSources.Items.Add($item) | Out-Null
                    Add-OutputText "Added from clipboard: $item"
                }
            }
        } else {
            Add-OutputText "No files or folders found in clipboard"
        }
    }
    catch {
        Add-OutputText "ERROR: Failed to get items from clipboard. Copy files/folders in Explorer first."
    }
})

$btnClearSources.Add_Click({
    $lstSources.Items.Clear()
    Add-OutputText "Cleared all source items"
})

$btnRemoveSelected.Add_Click({
    if ($lstSources.SelectedItem) {
        $selected = $lstSources.SelectedItem
        $lstSources.Items.Remove($selected)
        Add-OutputText "Removed: $selected"
    }
})

$btnBrowseSave.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "ISO Files (*.iso)|*.iso|All Files (*.*)|*.*"
    $saveFileDialog.DefaultExt = "iso"
    $saveFileDialog.Title = "Save ISO File As"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutputPath.Text = $saveFileDialog.FileName
    }
})

$btnBrowseBoot.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Boot Files (*.bin;*.com)|*.bin;*.com|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Boot File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtBootFile.Text = $openFileDialog.FileName
    }
})

$btnCreateISO.Add_Click({
    if ($lstSources.Items.Count -eq 0) {
        Add-OutputText "ERROR: Please add at least one file or folder to create an ISO"
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($txtOutputPath.Text)) {
        Add-OutputText "ERROR: Please specify an output path for the ISO file"
        return
    }
    
    try {
        $btnCreateISO.IsEnabled = $false
        $progressBar.IsIndeterminate = $true
        
        Add-OutputText "Starting ISO creation..."
        Add-OutputText "Output file: $($txtOutputPath.Text)"
        Add-OutputText "Media type: $($cmbMediaType.SelectedItem)"
        Add-OutputText "Volume title: $($txtTitle.Text)"
        
        # Prepare parameters
        $params = @{
            Source = @($lstSources.Items)
            Path = $txtOutputPath.Text
            Media = $cmbMediaType.SelectedItem
            Title = $txtTitle.Text
            Force = $chkForce.IsChecked
            Verbose = $true
        }
        
        if (![string]::IsNullOrWhiteSpace($txtBootFile.Text)) {
            if (Test-Path $txtBootFile.Text) {
                $params.BootFile = $txtBootFile.Text
                Add-OutputText "Using boot file: $($txtBootFile.Text)"
            } else {
                Add-OutputText "WARNING: Boot file not found, creating non-bootable ISO"
            }
        }
        
        # Create the ISO
        $result = New-IsoFile @params 4>&1 5>&1
        
        # Handle verbose output
        foreach ($message in $result) {
            if ($message -is [System.Management.Automation.VerboseRecord]) {
                Add-OutputText "VERBOSE: $($message.Message)"
            } elseif ($message -is [System.IO.FileInfo]) {
                Add-OutputText "SUCCESS: ISO file created successfully!"
                Add-OutputText "File: $($message.FullName)"
                Add-OutputText "Size: $([math]::Round($message.Length / 1MB, 2)) MB"
            }
        }
        
        Add-OutputText "ISO creation completed!"
        
    }
    catch {
        Add-OutputText "ERROR: $($_.Exception.Message)"
    }
    finally {
        $progressBar.IsIndeterminate = $false
        $btnCreateISO.IsEnabled = $true
    }
})

$btnExit.Add_Click({
    $window.Close()
})

# Show the window
Add-OutputText "ISO File Creator - Ready"
Add-OutputText "Add files and folders, configure settings, then click 'Create ISO'"
$window.ShowDialog() | Out-Null
