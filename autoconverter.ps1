Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem
$modelexefolder = 'Z:\BOCH\boch-converter\ATP_720\'
$cloudpointexefolder = 'Z:\BOCH\boch-converter\ATP_Cloud\'
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter           = '壓縮檔 (*.zip)|*.zip|點雲檔(*.e57)|*.e57'
}
$dialogresult = $FileBrowser.ShowDialog()
if ($dialogresult -eq [System.Windows.Forms.DialogResult]::OK) {    
    $tmpfolder = $env:TEMP + '\converter\'
    Remove-Item ($tmpfolder + '*') -Recurse -Force
    if ([System.IO.Path]::GetExtension($FileBrowser.FileName) -eq '.zip') {
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
        Start-Process $exepath.FullName -WindowStyle Minimized -Wait
        Start-Process $ffmpegpath.FullName -WindowStyle Minimized -Wait
        Invoke-Item $outputpath.FullName
    } else {
        $xmlpath = Get-ChildItem -Path $cloudpointexefolder -Filter 'localSetting.xml' -Recurse -ErrorAction SilentlyContinue -Force
        $exepath = Get-ChildItem -Path $cloudpointexefolder -Filter 'AutoTakePhoto720.exe' -Recurse -ErrorAction SilentlyContinue -Force
        $inputpath = Get-ChildItem -Path $cloudpointexefolder -Filter 'Input' -Recurse -ErrorAction SilentlyContinue -Force
        $outputpath = Get-ChildItem -Path $cloudpointexefolder -Filter 'Output' -Recurse -ErrorAction SilentlyContinue -Force
        $e572las =  Get-ChildItem -Path $cloudpointexefolder -Filter 'e572las.exe' -Recurse -ErrorAction SilentlyContinue -Force        
        $PotreeConverter =  Get-ChildItem -Path $cloudpointexefolder -Filter 'PotreeConverter.exe' -Recurse -ErrorAction SilentlyContinue -Force

        Start-Process $e572las -ArgumentList "-v -i $($FileBrowser.FileName) -o $($tmpfolder+'out.laz')" -WindowStyle Minimized -Wait
        New-Item -ErrorAction Ignore -ItemType directory -Path ($tmpfolder+'output')
        Start-Process $PotreeConverter -ArgumentList "$($tmpfolder+'out.laz') --overwrite -o $($tmpfolder+'output')" -WindowStyle Minimized -Wait

        [xml]$xmlDoc = Get-Content $xmlpath.FullName        
        $xmlDoc.LocalSetting.CloudDataFolder = ($tmpfolder+'output')
        $xmlDoc.LocalSetting.CameraAngle = '1'
        $xmlDoc.Save($xmlpath.FullName)
        Set-Location $cloudpointexefolder
        Start-Process $exepath.FullName -WindowStyle Minimized -Wait
        Invoke-Item $outputpath.FullName
    }        
}
else {
    [System.Windows.Forms.MessageBox]::Show('請選擇檔案, 3D模型請用.zip, 點雲模型用.e57.')
}