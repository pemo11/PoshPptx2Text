<#
 .SYNOPSIS
 Several helper functions
#>

$DllPath = Join-Path -Path $PSScriptRoot -ChildPath "PoshPptx2Text.dll"
Import-Module -Name $DllPath
