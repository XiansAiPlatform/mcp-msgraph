# NPM Publishing Guide

This document provides step-by-step instructions for patching and publishing the MCP Microsoft Graph server package to npm.

## Prerequisites

Before you begin, ensure you have:

1. **npm Account**: An npm account with publishing permissions for the `@99xio` organization
2. **GitHub Access**: Push access to this repository
3. **NPM Token**: A valid npm access token configured in GitHub Secrets

## Setup NPM Token (One-time setup)

1. Generate an npm access token:

   ```bash
   npm login
   npm token create --access=public
   ```

2. Add the token to GitHub Secrets:
   - Go to your GitHub repository
   - Navigate to **Settings** > **Secrets and variables** > **Actions**
   - Click **New repository secret**
   - Name: `NPM_TOKEN`
   - Value: Your npm token

## Publishing Process

### 1. Version Bumping

The package follows semantic versioning (semver). Choose the appropriate version bump:

- **Patch** (0.3.0 → 0.3.1): Bug fixes, minor changes
- **Minor** (0.3.0 → 0.4.0): New features, backward compatible
- **Major** (0.3.0 → 1.0.0): Breaking changes

#### Manual Version Update

Navigate to the package directory and update the version:

```bash
cd src/mcp

# For patch version
npm version patch

# For minor version  
npm version minor

# For major version
npm version major

# Or set specific version
npm version 1.2.3
```

#### Using npm Scripts

You can also use the built-in npm version commands:

```bash
cd src/mcp

# Patch version (0.3.0 → 0.3.1)
npm version patch --no-git-tag-version

# Minor version (0.3.0 → 0.4.0) 
npm version minor --no-git-tag-version

# Major version (0.3.0 → 1.0.0)
npm version major --no-git-tag-version
```

### 2. Create and Push Git Tag

After updating the version, create a git tag and push it:

```bash
# Get the current version from package.json
VERSION=$(node -p "require('./src/mcp/package.json').version")

# Create and push the tag
git add src/mcp/package.json
git commit -m "Bump version to v$VERSION"
git tag "v$VERSION"
git push origin main
git push origin "v$VERSION"
```

### 3. Automated Publishing

Once you push the tag, GitHub Actions will automatically:

1. ✅ Extract version from the tag
2. ✅ Validate version format
3. ✅ Install dependencies
4. ✅ Build the TypeScript code
5. ✅ Run tests
6. ✅ Verify build output
7. ✅ Publish to npm
8. ✅ Generate build summary

### 4. Manual Publishing (Alternative)

If you need to publish manually:

```bash
cd src/mcp

# Install dependencies
npm ci

# Build the package
npm run build

# Test the build
npm test

# Publish to npm
npm publish --access public
```

## Version Management Examples

### Patch Release Example
```bash
cd src/mcp
npm version patch --no-git-tag-version  # 0.3.0 → 0.3.1
git add package.json
git commit -m "Bump version to v0.3.1"
git tag "v0.3.1"
git push origin main && git push origin "v0.3.1"
```

### Minor Release Example
```bash
cd src/mcp
npm version minor --no-git-tag-version  # 0.3.0 → 0.4.0
git add package.json
git commit -m "Bump version to v0.4.0"
git tag "v0.4.0"
git push origin main && git push origin "v0.4.0"
```

### Major Release Example
```bash
cd src/mcp
npm version major --no-git-tag-version  # 0.3.0 → 1.0.0
git add package.json
git commit -m "Bump version to v1.0.0"
git tag "v1.0.0"
git push origin main && git push origin "v1.0.0"
```

## Quick Release Script

You can use this one-liner for quick patch releases:

```bash
cd src/mcp && npm version patch --no-git-tag-version && VERSION=$(node -p "require('./package.json').version") && cd ../.. && git add src/mcp/package.json && git commit -m "Bump version to v$VERSION" && git tag "v$VERSION" && git push origin main && git push origin "v$VERSION"
```

## Package Usage After Publishing

Once published, users can install and use the package:

### Global Installation
```bash
npm install -g @99xio/mcp-msgraph
```

### Using with npx
```bash
npx @99xio/mcp-msgraph
```

### MCP Configuration
```json
{
  "mcpServers": {
    "msgraph": {
      "command": "npx",
      "args": ["@99xio/mcp-msgraph"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret",
        "AZURE_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

## Troubleshooting

### Common Issues

1. **NPM Token Expired**: Regenerate token and update GitHub Secrets
2. **Version Already Exists**: Bump to a new version number
3. **Build Failures**: Check TypeScript compilation errors
4. **Permission Denied**: Ensure you have publish rights to `@99xio` organization

### Checking Published Versions

```bash
npm view @99xio/mcp-msgraph versions --json
```

### Unpublishing (Use with caution)

```bash
# Unpublish specific version (only within 72 hours)
npm unpublish @99xio/mcp-msgraph@1.0.0

# Deprecate version instead (recommended)
npm deprecate @99xio/mcp-msgraph@1.0.0 "This version has been deprecated"
```

## Best Practices

1. **Test Locally**: Always test the build locally before publishing
2. **Semantic Versioning**: Follow semver principles strictly
3. **Changelog**: Update changelog for each release
4. **Tag Messages**: Use descriptive tag messages
5. **Branch Protection**: Consider requiring PR reviews for version bumps

## Monitoring

After publishing, monitor:

- **npm Downloads**: Check package statistics on npmjs.com
- **GitHub Actions**: Monitor workflow runs for any failures  
- **Issues**: Watch for user-reported issues with new versions

---

For questions or issues with the publishing process, please create an issue in this repository.
