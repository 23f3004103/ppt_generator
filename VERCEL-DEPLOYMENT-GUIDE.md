# Vercel Deployment Guide for PowerPoint Generator

## Step 1: Prepare Your Project for Vercel

### Create vercel.json Configuration
```json
{
  "functions": {
    "api/*.py": {
      "runtime": "python3.9"
    }
  },
  "routes": [
    { "src": "/", "dest": "/app.html" },
    { "src": "/app", "dest": "/app.html" },
    { "src": "/demo", "dest": "/index.html" },
    { "src": "/(.*)", "dest": "/$1" }
  ],
  "headers": [
    {
      "source": "/app.html",
      "headers": [
        {
          "key": "Cache-Control",
          "value": "public, max-age=0, must-revalidate"
        }
      ]
    }
  ]
}
```

## Step 2: Initialize Git Repository (if not already done)

```bash
# Navigate to your project directory
cd "C:\Users\goura\Downloads\BONUS 1"

# Initialize git repository
git init

# Add all files
git add .

# Create initial commit
git commit -m "Initial commit: PowerPoint Generator with NVIDIA LLM support"
```

## Step 3: Create GitHub Repository

1. Go to https://github.com/new
2. Repository name: `powerpoint-generator` or `your-text-your-style`
3. Description: "AI-powered PowerPoint generator with template style preservation"
4. Set to **Public** (required for Vercel free tier)
5. Don't initialize with README (you already have files)
6. Click "Create repository"

## Step 4: Push to GitHub

```bash
# Add GitHub remote (replace YOUR_USERNAME with your GitHub username)
git remote add origin https://github.com/YOUR_USERNAME/powerpoint-generator.git

# Push to GitHub
git branch -M main
git push -u origin main
```

## Step 5: Deploy on Vercel

### Option A: Automatic Deployment (Recommended)
1. Go to https://vercel.com
2. Click "Continue with GitHub"
3. Click "Import Project"
4. Select your `powerpoint-generator` repository
5. Configure:
   - **Framework Preset**: Other
   - **Root Directory**: `./` (default)
   - **Build Command**: Leave empty
   - **Output Directory**: Leave empty
   - **Install Command**: Leave empty
6. Click "Deploy"

### Option B: Vercel CLI (Alternative)
```bash
# Install Vercel CLI globally
npm install -g vercel

# Login to Vercel
vercel login

# Deploy from project directory
cd "C:\Users\goura\Downloads\BONUS 1"
vercel

# Follow the prompts:
# - Set up and deploy? Y
# - Which scope? (select your account)
# - Link to existing project? N
# - Project name: powerpoint-generator
# - Directory: ./
# - Override settings? N
```

## Step 6: Configure Custom Domain (Optional)

1. In Vercel dashboard, go to your project
2. Click "Domains" tab
3. Add custom domain or use provided Vercel URL

## Step 7: Environment Variables (If Needed)

If you add any server-side functionality:
1. Go to project settings in Vercel
2. Click "Environment Variables"
3. Add any required variables

## Expected Hosted URLs

After deployment, you'll get URLs like:
- **Main App**: `https://powerpoint-generator.vercel.app/app.html`
- **Demo**: `https://powerpoint-generator.vercel.app/`
- **LLM Agent**: `https://powerpoint-generator.vercel.app/index.html`

## Step 8: Update Documentation

Update your README.md with the live demo links:

```markdown
## 🌐 Live Demo

- **PowerPoint Generator**: [https://your-project.vercel.app/app.html](https://your-project.vercel.app/app.html)
- **LLM Agent Demo**: [https://your-project.vercel.app/index.html](https://your-project.vercel.app/index.html)
```

## Troubleshooting

### Common Issues:
1. **Build Fails**: Ensure all file paths are relative
2. **404 Errors**: Check vercel.json routing configuration
3. **CORS Issues**: Add proper headers in vercel.json
4. **Large Files**: Vercel has file size limits (check documentation)

### File Size Considerations:
- Vercel free tier has deployment size limits
- Consider optimizing large assets
- External libraries loaded via CDN work well

## Vercel Benefits for This Project

✅ **Free Hosting**: Perfect for demos and portfolios
✅ **Automatic HTTPS**: Secure by default
✅ **Global CDN**: Fast loading worldwide
✅ **GitHub Integration**: Auto-deploy on push
✅ **Custom Domains**: Professional URLs
✅ **Analytics**: Built-in usage statistics

## Next Steps After Deployment

1. Test all functionality on the live site
2. Update your project documentation with live links
3. Share the hosted demo link
4. Consider setting up automatic deployments for updates
