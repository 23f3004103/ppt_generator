# Quick Vercel Deployment Commands

## Copy and paste these commands in PowerShell (one by one):

### Step 1: Initialize Git (if not already done)
```powershell
cd "C:\Users\goura\Downloads\BONUS 1"
git init
git add .
git commit -m "PowerPoint Generator with NVIDIA LLM support"
```

### Step 2: Create GitHub Repository
1. Go to: https://github.com/new
2. Repository name: `powerpoint-generator`
3. Make it **Public**
4. Click "Create repository"

### Step 3: Push to GitHub (replace YOUR_USERNAME)
```powershell
git remote add origin https://github.com/YOUR_USERNAME/powerpoint-generator.git
git branch -M main
git push -u origin main
```

### Step 4: Deploy on Vercel
1. Go to: https://vercel.com
2. Sign up/login with GitHub
3. Click "New Project"
4. Import your `powerpoint-generator` repository
5. Click "Deploy"

## Your live links will be:
- **Main App**: `https://powerpoint-generator.vercel.app/app.html`
- **Demo**: `https://powerpoint-generator.vercel.app/`

## Alternative: Vercel CLI Method
```powershell
npm install -g vercel
vercel login
vercel
```

Follow the prompts and your app will be deployed!
