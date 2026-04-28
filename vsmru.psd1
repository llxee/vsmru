@{
    RootModule        = 'vsmru.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a8b3c4d5-e6f7-4a5b-9c8d-7e6f5a4b3c2d'
    Author            = 'vsmru'
    CompanyName       = 'vsmru'
    Copyright         = '(c) vsmru. All rights reserved.'
    Description       = 'Visual Studio Most Recently Used (MRU) Projects CLI. Lists, searches, and deletes recent VS solutions and projects.'

    PowerShellVersion = '7.0'

    FunctionsToExport = @('Get-VSMRU')
    AliasesToExport   = @('vsmru')
    CmdletsToExport   = @()
    VariablesToExport = @()
}
