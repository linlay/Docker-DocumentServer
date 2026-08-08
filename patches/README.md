# ONLYOFFICE source UI patch

`onlyoffice-web-apps-9.4.0-source-ui.patch` applies to the unminified web-apps
sources shipped by `onlyoffice-documentserver` package version `9.4.0-129`.
`Dockerfile.source-ui` refuses to build against a different package version,
applies the patch, and regenerates each changed `.js.gz` resource.

The patch adds two explicit editor customization capabilities:

- `compactHeaderQuickAccess`: when `compactHeader` is enabled, render native
  Save, Print, Undo, and Redo controls in the toolbar header and hide their
  ribbon duplicates.
- `compactHeaderHideLogo`: do not mount the ONLYOFFICE logo into the compact
  toolbar row.
- `forceCompactToolbar`: use the signed `compactToolbar` setting instead of a
  previously stored browser preference.
- `sourceUiLayout`: apply `customization.layout` visibility settings without
  enabling extended branding, logo, or white-label customization.

The existing ONLYOFFICE source remains responsible for `compactHeader`,
`layout.leftMenu`, toolbar tab visibility, and double-click ribbon folding.
The AI Bridge contains no editor DOM observer, CSS injection, or element-moving
fallback.

Build the image from the repository root:

```sh
docker build --pull=false \
  --build-arg DOCUMENTSERVER_BASE_IMAGE=onlyoffice/documentserver:latest \
  --build-arg ONLYOFFICE_PACKAGE_VERSION=9.4.0-129 \
  -f Dockerfile.source-ui \
  -t onlyoffice/documentserver:9.4.0-129-source-ui .
```
