<#
.SYNOPSIS
Cmdlet help is awesome.  Autogenerate via template so I never forget.

.DESCRIPTION
.PARAMETER Files
One or more CSV files to be convert to XLS format.

.PARAMETER Password
Password

.INPUTS
.OUTPUTS
.EXAMPLE
Encrypt-File  @("C:\Users\xxx\Desktop\foo.csv","C:\Users\xxx\Desktop\bar.csv") "Happy1"

.LINK
#>

Function Encrypt-File {
	
    [CmdletBinding()] 
    param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True,Position=0)]
        [String[]] $Files,
        [Parameter(Mandatory=$True,Position=1)]
        [String] $Password
    ) 

    begin { 
        Write-Verbose "$($MyInvocation.MyCommand.Name)::Begin" 

        # open Excel
        Write-Verbose "Opening Excel..."
        $Excel = new-object -comobject Excel.application
        $Excel.Visible = $True 
        $Excel.DisplayAlerts = $False 

        # define constants
        New-Variable -Option Constant -Name xlDelimited -Value 1
        New-Variable -Option Constant -Name xlTextQualifierNone -Value -4142
        New-Variable -Option Constant -Name xlWorkbookDefault -Value 51   

        # eliminate race conditions (http://stackoverflow.com/a/461327/134367)
        Start-Sleep -sec 2

    }

    process {
        Write-Verbose "$($MyInvocation.MyCommand.Name)::Process" 

        try {

            Foreach ($File in $Files) {

                Write-Verbose "Opening file $($File)..."
                $Workbook = $Excel.Workbooks.open($File)
                $Destination = $File.replace(".csv", ".xls")

                Write-Verbose "Saving file $($Destination)..."
                $Workbook.Password = $Password
                $Workbook.SaveAs($Destination, $xlWorkbookDefault) | Out-Null
                $Workbook.Close()

            }

        }
        catch [Exception] {
            Write-Host $_.Exception.ToString()
        }
        finally {}

    }

    end {
        Write-Verbose "$($MyInvocation.MyCommand.Name)::End" 

        Write-Verbose "Closing Excel..."
        $Excel.Quit() 

    }
}

Export-ModuleMember Encrypt-File
Set-Alias ef Encrypt-File
Export-ModuleMember -Alias ef