# Put me in C:\Users\<user>\Documents\WindowsPowerShell

Function Invoke-NamedParameter{

    [CmdletBinding(DefaultParameterSetName = "Named")]
    param(
        [Parameter(ParameterSetName = "Named", Position = 0, Mandatory = $true)]
        [Parameter(ParameterSetName = "Positional", Position = 0, Mandatory = $true)]
        [ValidateNotNull()]
        [System.Object]$Object
        , 
        [Parameter(ParameterSetName = "Named", Position = 1, Mandatory = $true)]
        [Parameter(ParameterSetName = "Positional", Position = 1, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$Method
        ,
        [Parameter(ParameterSetName = "Named", Position = 2, Mandatory = $true)]
        [ValidateNotNull()]
        [Hashtable]$Parameter
        ,
        [Parameter(ParameterSetName = "Positional")]
        [Object[]]$Argument
        
    )

    end{
        ## Just being explicit that this does not support pipelines
        if($PSCmdlet.ParameterSetName -eq "Named") { 
            ## Invoke method with parameter names
            ## Note: It is ok to use a hashtable here because the keys (parameter names) and values (args)
            ## will be output in the same order.  We don't need to worry about the order so long as
            ## all parameters have names
            $Object.GetType().InvokeMember($Method, 
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null, ## Binder
            $Object, ## Target
            ([Object[]]($Parameter.Values)), # Args
            $null, # Modifiers
            $null, # Culture
            ([String[]]($Parameter.Keys)) ## NamedParameters
            )
        } else {
            ## Invoke method witout parameter names
            $Object.GetType().InvokeMember($Method,
            [System.Reflection.BindingFlags]::InvokeMethod,
            $null, ## Binder
            $Object, ## Target
            $Argument, # Args
            $null, # Modifiers
            $null, # Culture
            $null ## NamedParameters
            )

        }

    }

}

function PPT_Template_Dynamic {

    param($title)
    if(!$title)
        {$title = @()}
    
    $ppt = New-Object -ComObject Powerpoint.Application
    $fname = "YYYY-MM-DD Powerpoint Template.pptm"
    # Need to be in the directory containing file, or add path info
    
    $ppt.Presentations.Open($fname)
    $authordate = $fname+"!Module1.Add_Author_Date_Boxes"
    Invoke-NamedParameter $ppt "Run" -argument @($authordate, $title)
    $save = $fname+"!Module1.Save_File"
    Invoke-NamedParameter $ppt "Run" -argument @($save, $title)

}

function Test-Macro {
    param ()

    $ppt = New-Object -ComObject Powerpoint.Application
    $fname = "test.pptm"
    $ppt.Presentations.Open($fname)
    $authordate = $fname + "!Module1.Add_Author_Date_Boxes"
    $title = "mytitle"
    Invoke-NamedParameter $ppt "Run" -Argument @($authordate, $title)

}

# Need to set install location for utilities.ps1
# Set-Alias twoway 
# Set-Alias hardlink


