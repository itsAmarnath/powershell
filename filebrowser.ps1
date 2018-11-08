$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog #-Property @{ 
    #InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    #Filter = 'Documents (*.docx)|*.docx|SpreadSheet (*.xlsx)|*.xlsx'
#}
$null = $FileBrowser.ShowDialog()