$ScriptConfig = @{
    # This needs to be a single, valid CW company.
    #   Use the * wildcard for fuzzy search, but if multiple values are returned, it will error.
    CWCompany = '*roost*'

    # OrganisationalUnit contains an array of any OUs you wish to search.
    OrganisationalUnit = @(
        'MY OU GOES HERE'
    )
}