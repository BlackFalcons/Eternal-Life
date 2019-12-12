using namespace System.Management.Automation
function addTypeFile {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$objectName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String]$filePath
    )
    
    if (-not ([PSTypeName]'EternalFrame').Type)
    {
        Add-type -Path .\EternalFrame.cs
    }
}
