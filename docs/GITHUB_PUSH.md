# GitHub push checklist

## 1. Add the installer and portable exe

Run in the repo folder (`BundleScript`):

```powershell
git add .gitignore
git add dist/Create_Bundle.exe Output/Create_Bundle_Setup.exe
```

If Git still says the paths are ignored, force-add:

```powershell
git add -f dist/Create_Bundle.exe Output/Create_Bundle_Setup.exe
```

Then commit:

```powershell
git commit -m "Add installer and portable exe for end-user downloads"
```

## 2. Point remote at your real repo

Create the repository on GitHub first (github.com → New repository, name e.g. `Create_Bundle`).

Then set the remote (use your GitHub username or org instead of `cmcejas`):

```powershell
git remote set-url origin https://github.com/cmcejas/Create_Bundle.git
```

Check it:

```powershell
git remote -v
```

## 3. Push

```powershell
git push -u origin main
```

If you see **"Failed to connect to github.com port 443"**:

- Try from another network (e.g. mobile hotspot) if you’re on a corporate VPN/firewall.
- Or use SSH instead of HTTPS:  
  `git remote set-url origin git@github.com:cmcejas/Create_Bundle.git`  
  (requires SSH keys set up on GitHub.)
