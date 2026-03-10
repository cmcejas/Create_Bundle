# Create the GitHub repo and push

## Step 1: Create the repo on GitHub (in the browser)

1. Go to **https://github.com/new**
2. **Repository name:** `Create_Bundle` (or any name you like)
3. **Description (optional):** e.g. "PDF/Word/Outlook bundler – merge documents into one PDF"
4. Choose **Public**
5. **Do not** check "Add a README", "Add .gitignore", or "Choose a license" – you already have those locally. Leave the repo **empty**.
6. Click **Create repository**

GitHub will show you a page with setup commands. You can ignore that and do Step 2 below.

## Step 2: Point your local repo at the new GitHub repo

Replace `YOUR_GITHUB_USERNAME` with your actual GitHub username (e.g. `cmcejas`):

```powershell
git remote set-url origin https://github.com/YOUR_GITHUB_USERNAME/Create_Bundle.git
```

Example:

```powershell
git remote set-url origin https://github.com/cmcejas/Create_Bundle.git
```

Check it:

```powershell
git remote -v
```

## Step 3: Push and set upstream

```powershell
git push --set-upstream origin main
```

Enter your GitHub credentials if prompted. After this, the repo and the installer/exe will be on GitHub and you can share the link.

---

**Direct download link for colleagues (after push):**  
`https://github.com/YOUR_GITHUB_USERNAME/Create_Bundle/raw/main/Output/Create_Bundle_Setup.exe`
