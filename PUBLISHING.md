# Publishing RSR v.1

These steps update the existing repository history at `https://github.com/Esra11/RSR`.
Use the existing clean clone at `C:\work\test github rsrs\RSR` as the staging repository.

## 1. Refresh and verify the staging clone

```powershell
$source = 'C:\work\RSR-Github\RSR-main'
$repo = 'C:\work\test github rsrs\RSR'

git -C $repo switch main
git -C $repo pull --ff-only origin main
git -C $repo status --short
```

Stop if the final command prints local changes.

## 2. Copy publishable files

Copy source and tests while excluding generated output:

```powershell
robocopy "$source\DesktopSteps" "$repo\DesktopSteps" /E /XD bin obj
if ($LASTEXITCODE -ge 8) { throw "DesktopSteps copy failed: $LASTEXITCODE" }

robocopy "$source\DesktopSteps.Tests" "$repo\DesktopSteps.Tests" /E /XD bin obj
if ($LASTEXITCODE -ge 8) { throw "DesktopSteps.Tests copy failed: $LASTEXITCODE" }

Copy-Item -LiteralPath "$source\README.md" -Destination "$repo\README.md" -Force
Copy-Item -LiteralPath "$source\CHANGELOG.md" -Destination "$repo\CHANGELOG.md" -Force
Copy-Item -LiteralPath "$source\PUBLISHING.md" -Destination "$repo\PUBLISHING.md" -Force
Copy-Item -LiteralPath "$source\.gitignore" -Destination "$repo\.gitignore" -Force
Copy-Item -LiteralPath "$source\.env.example" -Destination "$repo\.env.example" -Force
```

Remove the old tracked build folder. Release binaries will be attached to the GitHub release instead:

```powershell
git -C $repo rm -r --ignore-unmatch bin
```

Do not copy `.env`, `bin`, `obj`, or any `Recordings` folder.

## 3. Review and test the staged repository

```powershell
git -C $repo status --short
git -C $repo diff --stat
git -C $repo diff --check

dotnet build "$repo\DesktopSteps\DesktopSteps.csproj" -c Release
dotnet build "$repo\DesktopSteps.Tests\DesktopSteps.Tests.csproj" -c Release
dotnet run --project "$repo\DesktopSteps.Tests\DesktopSteps.Tests.csproj" -c Release --no-build -- --intent-input-tests
```

Confirm ignored private/runtime content is not staged:

```powershell
git -C $repo status --short --ignored | Select-String -Pattern '\.env$|Recordings|\\bin\\|\\obj\\'
git -C $repo diff --cached --name-status
```

The cached file list must contain no added or modified `.env`, recordings, screenshots,
generated `bin`/`obj` files, or diagnostics. Existing tracked `bin` files may appear only
with status `D`, because this release removes them from source control.

## 4. Commit and publish the source

```powershell
git -C $repo add DesktopSteps DesktopSteps.Tests README.md CHANGELOG.md PUBLISHING.md .gitignore .env.example
git -C $repo status --short
git -C $repo diff --cached --stat
git -C $repo diff --cached --check

git -C $repo commit -m "Release RSR v.1"
git -C $repo tag -a v.1 -m "RSR v.1"
git -C $repo push origin main
git -C $repo push origin v.1
```

Review `git status` and the staged diff before the commit. Pushing changes GitHub and should happen only after that review.

## 5. Build a recording-free release package

Build outside both source trees so runtime recordings cannot enter the archive:

```powershell
$release = 'C:\work\RSR-Github\release-v1'
if (Test-Path -LiteralPath $release) { Remove-Item -LiteralPath $release -Recurse -Force }
New-Item -ItemType Directory -Path $release | Out-Null

dotnet publish "$repo\DesktopSteps\DesktopSteps.csproj" -c Release -o "$release\RSR"
Copy-Item -LiteralPath "$repo\.env.example" -Destination "$release\RSR\.env.example"
Copy-Item -LiteralPath "$repo\README.md" -Destination "$release\RSR\README.md"

Get-ChildItem -LiteralPath $release -Recurse -Force |
    Where-Object { $_.Name -eq '.env' -or $_.FullName -match '[\\/]Recordings([\\/]|$)' } |
    ForEach-Object { throw "Private runtime content found in release: $($_.FullName)" }

Compress-Archive -Path "$release\RSR" -DestinationPath "$release\RSR-v.1.zip" -Force
```

After installing GitHub CLI and running `gh auth login`, create the GitHub release:

```powershell
gh release create v.1 "$release\RSR-v.1.zip" `
    --repo Esra11/RSR `
    --title "RSR v.1" `
    --notes-file "$repo\CHANGELOG.md"
```
