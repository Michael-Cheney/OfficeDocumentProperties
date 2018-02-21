<#
.Synopsis
   Get the extended file details from Microsoft Office document
.DESCRIPTION
   This module will allow you to read and write to all of the extended file details in Microsoft Office documents. This includes fields such as "Title, Subject, Author, Keywords, Comments... and so on
.EXAMPLE
   Get-OfficeDocumentProperties -FilePath C:\temp\excel.xls
.EXAMPLE
   Get-OfficeDocumentProperties -FilePath C:\temp\excel.xls | Select N
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Get-OfficeDocumentProperties
{
    [CmdletBinding(SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Low')]
    [OutputType([String])]
    Param
    (
        # FilePath 
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({Test-Path $_})]
        [string]$FilePath
    )

    Begin
    {
        $document_properties = @()
        # Create Application Object
        $excel = New-Object -ComObject Excel.Application
        # Set Application to not be visible
        $excel.Visible = $false
        # Read the file with the application
        $workbook = $excel.Workbooks.Open($FilePath)
        $binding = "System.Reflection.BindingFlags" -as [type]
    }
    Process
    {
        foreach($book in $workbook.BuiltInDocumentProperties) {
            $pn = [System.__ComObject].invokemember("name",$binding::GetProperty,$null,$book,$null)
            trap [system.exception] {
                continue
            }
            $property = new-object System.Object
            $property | Add-Member -MemberType NoteProperty `
                                   -Name "Name" `
                                   -Value $pn
            $property | Add-Member -MemberType NoteProperty `
                                   -Name "Value" `
                                   -Value ([System.__ComObject].invokemember("value",$binding::GetProperty,$null,$book,$null))
            $document_properties += $property
        }
    }
    End
    {
        $excel.quit()
        $document_properties
    }
}
