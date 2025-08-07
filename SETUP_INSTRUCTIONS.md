# Setup Instructions for GitHub and NPM Publishing

## 📋 Prerequisites

1. **GitHub Account**: Make sure you're logged in at [github.com](https://github.com)
2. **NPM Account**: Create one at [npmjs.com](https://www.npmjs.com/signup) if you don't have one
3. **Node.js**: Ensure Node.js 16+ is installed (`node --version`)

## 🚀 Step 1: Push to GitHub

### 1.1 Create a New Repository on GitHub

1. Go to [https://github.com/new](https://github.com/new)
2. Repository name: `powerpoint-creator`
3. Description: "Professional PowerPoint presentation generator with template support"
4. Keep it **Public** (required for free NPM publishing)
5. **DON'T** initialize with README, .gitignore, or license (we already have them)
6. Click "Create repository"

### 1.2 Push Your Local Repository

After creating the empty repository on GitHub, run these commands:

```bash
# Add GitHub as remote origin
git remote add origin https://github.com/wapdat/powerpoint-creator.git

# Push to GitHub
git push -u origin main
```

If you get an authentication error, you may need to:
- Use a Personal Access Token instead of password
- Create one at: https://github.com/settings/tokens/new
- Select scopes: `repo` (full control)

Alternative using SSH:
```bash
git remote set-url origin git@github.com:wapdat/powerpoint-creator.git
git push -u origin main
```

## 📦 Step 2: Publish to NPM

### 2.1 Check Package Name Availability

First, check if the name is available:
```bash
npm view powerpoint-creator
```

If you get a 404 error, the name is available! If not, update the name in `package.json`.

### 2.2 Login to NPM

```bash
npm login
```

Enter your NPM username, password, and email when prompted.

### 2.3 Build the Project

```bash
# Install dependencies
npm install

# Build the TypeScript code
npm run build
```

### 2.4 Test Locally (Optional but Recommended)

```bash
# Test the package locally
npm link

# In another directory, test installing
npm link powerpoint-creator

# Test the CLI
powerpoint-creator --help

# Unlink when done
npm unlink powerpoint-creator
```

### 2.5 Publish to NPM

```bash
# For first-time publishing
npm publish

# If the package name is taken, you can scope it to your username
# Update package.json name to: "@wapdat/powerpoint-creator"
# Then publish with public access
npm publish --access public
```

## 🔄 Step 3: Set Up Automated Publishing (Optional)

### 3.1 Get NPM Token

1. Go to [https://www.npmjs.com/settings/~/tokens](https://www.npmjs.com/settings/~/tokens)
2. Click "Generate New Token"
3. Choose "Automation" type
4. Copy the token (starts with `npm_`)

### 3.2 Add Token to GitHub Secrets

1. Go to your GitHub repository
2. Settings → Secrets and variables → Actions
3. Click "New repository secret"
4. Name: `NPM_TOKEN`
5. Value: Paste your NPM token
6. Click "Add secret"

### 3.3 Create a Release to Trigger Publishing

1. Go to your repository on GitHub
2. Click "Releases" → "Create a new release"
3. Choose a tag (e.g., `v1.0.0`)
4. Release title: "v1.0.0 - Initial Release"
5. Describe the features
6. Click "Publish release"

This will trigger the GitHub Action to automatically publish to NPM!

## 📝 Step 4: Update Package Version

For future updates:

```bash
# Update patch version (1.0.0 → 1.0.1)
npm version patch

# Update minor version (1.0.0 → 1.1.0)
npm version minor

# Update major version (1.0.0 → 2.0.0)
npm version major

# Push the version tag to GitHub
git push --follow-tags

# Publish to NPM
npm publish
```

## ✅ Verification

After publishing, verify your package:

1. **View on NPM**: https://www.npmjs.com/package/powerpoint-creator
2. **Test installation**:
   ```bash
   # In a new directory
   npm install -g powerpoint-creator
   powerpoint-creator --help
   ```

## 🎉 Success Checklist

- [ ] Repository visible at https://github.com/wapdat/powerpoint-creator
- [ ] Package visible at https://www.npmjs.com/package/powerpoint-creator
- [ ] CLI works: `npx powerpoint-creator --help`
- [ ] GitHub Actions badge shows passing

## 🆘 Troubleshooting

### NPM Publish Errors

1. **E403 Forbidden**: You need to login: `npm login`
2. **E402 Payment Required**: Name might be too similar to existing package
3. **E403 Package name not allowed**: Try scoping: `@wapdat/powerpoint-creator`

### GitHub Push Errors

1. **Authentication failed**: Use Personal Access Token or SSH
2. **Repository not found**: Check the remote URL: `git remote -v`
3. **Permission denied**: Make sure you own the repository

### Build Errors

```bash
# Clean install
rm -rf node_modules package-lock.json
npm install
npm run build
```

## 📚 Useful Commands Summary

```bash
# GitHub
git add .
git commit -m "Update message"
git push

# NPM
npm login
npm run build
npm publish
npm version patch
npm view powerpoint-creator

# Testing
npm link
npm run example
npx powerpoint-creator --help
```

## 🔗 Important Links

- **Your GitHub Repo**: https://github.com/wapdat/powerpoint-creator
- **Your NPM Package**: https://www.npmjs.com/package/powerpoint-creator
- **NPM Tokens**: https://www.npmjs.com/settings/~/tokens
- **GitHub Tokens**: https://github.com/settings/tokens

Good luck with your package! 🚀