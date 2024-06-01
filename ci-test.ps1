param(
  [Parameter(Mandatory, Position = 0)]
  [string]$targetFramework
)


dotnet test --no-build -c Release -f $targetFramework
