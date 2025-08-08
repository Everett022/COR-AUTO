# COR-AUTO Excel Add-in

This is an Excel add-in for COR-AUTO functionality.

## Deployment to GitHub Pages

This project is configured to deploy to GitHub Pages automatically. The deployment process:

1. **Build Process**: The project uses webpack to build the add-in files into the `dist/` folder
2. **GitHub Actions**: A workflow automatically builds and deploys the `dist/` folder to GitHub Pages
3. **URL Structure**: The add-in is accessible at `https://everett022.github.io/COR-AUTO/`

## Files Structure for GitHub Pages

The following files are deployed to GitHub Pages:
- `index.html` - Main entry point (redirects to taskpane.html)
- `taskpane.html` - Main taskpane interface
- `commands.html` - Commands interface
- `display.html` - Display dialog
- `orderingSet.html` - Ordering settings dialog
- `inventorySet.html` - Inventory settings dialog
- `assets/` - All image and asset files
- `manifest.xml` - Office add-in manifest

## Local Development

To run locally:
```bash
npm install
npm start
```

## Building for Production

To build for production:
```bash
npm run build
```

The built files will be in the `dist/` folder and ready for deployment. 