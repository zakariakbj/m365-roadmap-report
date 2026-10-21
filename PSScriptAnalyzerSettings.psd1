@{
    Severity     = @("Error", "Warning")
    # Pester blocks share variables across scriptblocks, which this rule cannot follow
    ExcludeRules = @("PSUseDeclaredVarsMoreThanAssignments")
}
