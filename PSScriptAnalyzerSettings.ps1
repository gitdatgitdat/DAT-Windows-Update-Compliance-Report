@{
  Severity            = @('Error','Warning')
  IncludeDefaultRules = $true
  ExcludeRules        = @(
    'PSAvoidUsingWriteHost',        # allowed for CLI UX
    'PSAvoidTrailingWhitespace',    # relaxed early
    'PSUseConsistentWhitespace'     # relaxed early
  )
}
