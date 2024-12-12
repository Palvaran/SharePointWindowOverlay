# SharePoint Window Overlay

A SharePoint Framework (SPFx) extension that adds a customizable iframe overlay to your SharePoint site, allowing for dynamic integrations such as AI chatbots or other embedded content.

---

## Features

- **Floating Iframe:** Easily embed external content within your SharePoint site.
- **Minimize and Restore:** Control the visibility of the iframe with intuitive buttons.
- **Responsive Design:** Adapts to various screen sizes for optimal viewing on desktops, tablets, and mobile devices.
- **Accessibility Enhancements:** Keyboard navigable and screen reader friendly.
- **LazyLoading** For Performance Enhancements

---

## Configuration

### **1. Update `config.json`**

The iframe URL is configured in the `config.json` file located at:
`/src/extensions/tobyOverlayCustomizer/config.json`

Example configuration:
{
  "iframeUrl": "https://copilotstudio.microsoft.com/environments/Default-XXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX/bots/cr0fd_copilotPocTemporary/webchat?__version__=2"
}

2. Check and Update package-solution.json

Before building and deploying, verify the package-solution.json file located in /config/.

Key Properties to Check:

skipFeatureDeployment: Set this to false to ensure the feature is scoped properly:

"skipFeatureDeployment": false
version: Increment the version number with each new build to ensure SharePoint detects the update. Example:
"version": "1.0.0.0"
Steps to Update:

Open package-solution.json in a text editor.
Ensure skipFeatureDeployment is false.
Update the version field. Follow semantic versioning (1.0.0.0, 1.0.0.1, etc.).
Save the file before proceeding.

---

## Deployment Steps ##

1. Clone this repository:
  git clone https://github.com/Palvaran/SharePointWindowOverlay.git
2. Install dependencies:
  npm install
3. Build the project:
  gulp build
4. Bundle and package the solution:
  gulp bundle --ship
  gulp package-solution --ship
5. Deploy the package to your SharePoint App Catalog.

---

## Notes for New Users ##

Things to Check Before Deploying:
config.json: Ensure the iframeUrl is set to the desired value.
package-solution.json:
skipFeatureDeployment must be false.
The version number should be incremented for new builds.
Dependencies: Run npm install to ensure all dependencies are installed.
Test Before Production: Always test the solution in a development environment.



