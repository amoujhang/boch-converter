Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Windows.Forms 

$Form = New-Object system.Windows.Forms.Form
$Label = New-Object System.Windows.Forms.ListBox
$Form.Controls.Add($Label)
$Form.Text="點雲模型轉檔工具"
$Label.Items.Add("開始轉檔")
$Label.Width = $Form.Width
$Label.Height = $Form.Height
#$Label.AutoSize = $True


$modelexefolder = 'C:\boch-converter\ATP_720\'
$cloudpointexefolder = 'C:\boch-converter\ATP_Cloud\'
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter           = '壓縮檔、點雲檔 (*.zip, *.e57)|*.zip;*.e57'
}
$dialogresult = $FileBrowser.ShowDialog()
if ($dialogresult -eq [System.Windows.Forms.DialogResult]::OK) {    
    $tmpfolder = $env:TEMP + '\converter\'
    Remove-Item ($tmpfolder + '*') -Recurse -Force
    if ([System.IO.Path]::GetExtension($FileBrowser.FileName) -eq '.zip') {
        $Form.Visible = $True        
        $Label.Items.Add("解壓縮...")
        $Form.Update()
        [System.IO.Compression.ZipFile]::ExtractToDirectory($FileBrowser.FileName, $tmpfolder)
        $texturepath = Get-ChildItem -Path $tmpfolder -Filter 'texture' -Recurse -ErrorAction SilentlyContinue -Force
        $fbxpath = Get-ChildItem -Path $tmpfolder -Filter '*.fbx' -Recurse -ErrorAction SilentlyContinue -Force
        $xmlpath = Get-ChildItem -Path $modelexefolder -Filter 'localSetting.xml' -Recurse -ErrorAction SilentlyContinue -Force
        $exepath = Get-ChildItem -Path $modelexefolder -Filter 'AutoTakePhoto720.exe' -Recurse -ErrorAction SilentlyContinue -Force
        $ffmpegpath = Get-ChildItem -Path $modelexefolder -Filter 'ffmpegPNG.bat' -Recurse -ErrorAction SilentlyContinue -Force
        $inputpath = Get-ChildItem -Path $modelexefolder -Filter 'Input' -Recurse -ErrorAction SilentlyContinue -Force
        $outputpath = Get-ChildItem -Path $modelexefolder -Filter 'Output' -Recurse -ErrorAction SilentlyContinue -Force
        Remove-Item ($inputpath.FullName + '\*') -Recurse -Force
        Remove-Item ($outputpath.FullName + '\*') -Recurse -Force
        [xml]$xmlDoc = Get-Content $xmlpath.FullName
        $xmlDoc.LocalSetting.FileFBX = $fbxpath.FullName
        $xmlDoc.LocalSetting.TextureFolder = $texturepath.FullName
        $xmlDoc.LocalSetting.CameraAngle = '1'
        $xmlDoc.Save($xmlpath.FullName)
        Set-Location $modelexefolder
        $Label.Items.Add("處理模型中...")
        $Form.Update()
        Start-Process $exepath.FullName -WindowStyle Hidden -Wait
        $Label.Items.Add("錄製...")
        $Form.Update()
        Start-Process $ffmpegpath.FullName -WindowStyle Hidden -Wait
        Invoke-Item $outputpath.FullName
    } else {
        $Form.Visible = $True
        $Form.Update()
        $xmlpath = Get-ChildItem -Path $cloudpointexefolder -Filter 'localSetting.xml' -Recurse -ErrorAction SilentlyContinue -Force
        $exepath = Get-ChildItem -Path $cloudpointexefolder -Filter 'AutoTakePhoto720.exe' -Recurse -ErrorAction SilentlyContinue -Force 
        $outputpath = Get-ChildItem -Path $cloudpointexefolder -Filter 'Output' -ErrorAction SilentlyContinue -Force       
        $inputpath = Get-ChildItem -Path $cloudpointexefolder -Filter 'Input' -ErrorAction SilentlyContinue -Force    
        $ffmpegpath = Get-ChildItem -Path $cloudpointexefolder -Filter 'ffmpegPNG.bat' -Recurse -ErrorAction SilentlyContinue -Force                     
        Remove-Item ($inputpath.FullName + '\*') -Recurse -Force
        Remove-Item ($outputpath.FullName + '\*') -Recurse -Force
        New-Item -ErrorAction Ignore -ItemType directory -Path ($tmpfolder+'output') | Out-Null          
        $Label.Items.Add("e57點雲格式轉換中...")
        $Form.Update()
        Set-Location 'C:\boch-converter\ATP_Cloud\PointCloudConverter\'
        Start-Process 'run.bat' -Args """$($FileBrowser.FileName)"" ""$($tmpfolder+'output')""" -WindowStyle Hidden -Wait        
        [xml]$xmlDoc = Get-Content $xmlpath.FullName        
        $xmlDoc.LocalSetting.CloudDataFolder = ($tmpfolder+'output')
        $xmlDoc.LocalSetting.CameraAngle = '1'
        $xmlDoc.Save($xmlpath.FullName)
        Set-Location $cloudpointexefolder
        $Label.Items.Add("點雲模型處理中(約需1小時)...")
        $Form.Update()
        Start-Process $exepath.FullName -WindowStyle Hidden -Wait  
        $Label.Items.Add("錄製中...")
        $Form.Update()
        Start-Process $ffmpegpath.FullName -WindowStyle Hidden -Wait      
        Invoke-Item $outputpath.FullName
    }        
    $Label.Text = "完成"
    Start-Sleep 3
    $Form.Close()
}
else {
    [System.Windows.Forms.MessageBox]::Show('請選擇檔案, 3D模型請用.zip, 點雲模型用.e57.')
}