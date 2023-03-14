param ([switch]$usecustomlist=$false,[string]$listpath)

If($usecustomlist -eq $false){
    Write-Host "All users will be evaulated" -ForegroundColor Green
}

else{
    try{$userpath = import-csv $listpath}
    catch{
        Add-Type -AssemblyName System.Windows.Forms
        
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
            InitialDirectory = [Environment]::GetFolderPath('Desktop') 
            Filter = 'CSV (*.csv)|*.csv|SpreadSheet (*.xlsx)|*.xlsx'
        }
        
        $null = $FileBrowser.ShowDialog()
        $userpath = import-csv $FileBrowser.FileName
        Write-Host "we are using the file selector" -ForegroundColor Red
    }
    
}

Write-Host "value of the path is $pathtolist"
$userpath | ft
