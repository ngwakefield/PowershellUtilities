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
            $Object.GetType().InvokeMember(($Method, 
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
}