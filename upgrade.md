# Upgrade project leave-mgmt-client-side-solution to v1.18.2

Date: 2/15/2024

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.18.2. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i -SE @microsoft/sp-core-library@1.18.2
```

File: [./package.json:16:5](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm i -SE @microsoft/sp-lodash-subset@1.18.2
```

File: [./package.json:17:5](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm i -SE @microsoft/sp-office-ui-fabric-core@1.18.2
```

File: [./package.json:18:5](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-webpart-base@1.18.2
```

File: [./package.json:20:5](./package.json)

### FN001021 @microsoft/sp-property-pane | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-property-pane

Execute the following command:

```sh
npm i -SE @microsoft/sp-property-pane@1.18.2
```

File: [./package.json:19:5](./package.json)

### FN001034 @microsoft/sp-adaptive-card-extension-base | Optional

Install SharePoint Framework dependency package @microsoft/sp-adaptive-card-extension-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-adaptive-card-extension-base@1.18.2
```

File: [./package.json:14:3](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i -DE @microsoft/sp-build-web@1.18.2
```

File: [./package.json:48:5](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i -DE @microsoft/sp-module-interfaces@1.18.2
```

File: [./package.json:49:5](./package.json)

### FN002022 @microsoft/eslint-plugin-spfx | Required

Install SharePoint Framework dev dependency package @microsoft/eslint-plugin-spfx

Execute the following command:

```sh
npm i -DE @microsoft/eslint-plugin-spfx@1.18.2
```

File: [./package.json:45:3](./package.json)

### FN002023 @microsoft/eslint-config-spfx | Required

Install SharePoint Framework dev dependency package @microsoft/eslint-config-spfx

Execute the following command:

```sh
npm i -DE @microsoft/eslint-config-spfx@1.18.2
```

File: [./package.json:45:3](./package.json)

### FN010001 .yo-rc.json version | Recommended

Update version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.18.2"
  }
}
```

File: [./.yo-rc.json:5:5](./.yo-rc.json)

### FN001022 office-ui-fabric-react | Required

Remove SharePoint Framework dependency package office-ui-fabric-react

Execute the following command:

```sh
npm un -S office-ui-fabric-react
```

File: [./package.json:35:5](./package.json)

### FN001035 @fluentui/react | Required

Install SharePoint Framework dependency package @fluentui/react

Execute the following command:

```sh
npm i -SE @fluentui/react@8.106.4
```

File: [./package.json:14:3](./package.json)

### FN002026 typescript | Required

Install SharePoint Framework dev dependency package typescript

Execute the following command:

```sh
npm i -DE typescript@4.7.4
```

File: [./package.json:45:3](./package.json)

### FN002028 @microsoft/rush-stack-compiler-4.7 | Required

Install SharePoint Framework dev dependency package @microsoft/rush-stack-compiler-4.7

Execute the following command:

```sh
npm i -DE @microsoft/rush-stack-compiler-4.7@0.1.0
```

File: [./package.json:45:3](./package.json)

### FN010010 .yo-rc.json @microsoft/teams-js SDK version | Recommended

Update @microsoft/teams-js SDK version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/teams-js": "2.12.0"
    }
  }
}
```

File: [./.yo-rc.json:2:3](./.yo-rc.json)

### FN012017 tsconfig.json extends property | Required

Update tsconfig.json extends property

```json
{
  "extends": "./node_modules/@microsoft/rush-stack-compiler-4.7/includes/tsconfig-web.json"
}
```

File: [./tsconfig.json:2:3](./tsconfig.json)

### FN021003 package.json engines.node | Required

Update package.json engines.node property

```json
{
  "engines": {
    "node": ">=16.13.0 <17.0.0 || >=18.17.1 <19.0.0"
  }
}
```

File: [./package.json:7:5](./package.json)

### FN002024 eslint | Required

Install SharePoint Framework dev dependency package eslint

Execute the following command:

```sh
npm i -DE eslint@8.7.0
```

File: [./package.json:45:3](./package.json)

### FN007002 serve.json initialPage | Required

Update serve.json initialPage URL

```json
{
  "initialPage": "https://{tenantDomain}/_layouts/workbench.aspx"
}
```

File: [./config/serve.json:5:3](./config/serve.json)

### FN015009 config\sass.json | Required

Add file config\sass.json

Execute the following command:

```sh
cat > "config\sass.json" << EOF 
{
  "$schema": "https://developer.microsoft.com/json-schemas/core-build/sass.schema.json"
}
EOF
```

File: [config\sass.json](config\sass.json)

### FN001008 react | Required

Upgrade SharePoint Framework dependency package react

Execute the following command:

```sh
npm i -SE react@17.0.1
```

File: [./package.json:36:5](./package.json)

### FN001009 react-dom | Required

Upgrade SharePoint Framework dependency package react-dom

Execute the following command:

```sh
npm i -SE react-dom@17.0.1
```

File: [./package.json:39:5](./package.json)

### FN002015 @types/react | Required

Upgrade SharePoint Framework dev dependency package @types/react

Execute the following command:

```sh
npm i -DE @types/react@17.0.45
```

File: [./package.json:56:5](./package.json)

### FN002016 @types/react-dom | Required

Upgrade SharePoint Framework dev dependency package @types/react-dom

Execute the following command:

```sh
npm i -DE @types/react-dom@17.0.17
```

File: [./package.json:57:5](./package.json)

### FN010008 .yo-rc.json nodeVersion | Recommended

Update nodeVersion in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "nodeVersion": "18.19.0"
  }
}
```

File: [./.yo-rc.json:2:38](./.yo-rc.json)

### FN010009 .yo-rc.json @microsoft/microsoft-graph-client SDK version | Recommended

Update @microsoft/microsoft-graph-client SDK version in .yo-rc.json

```json
{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/microsoft-graph-client": "3.0.2"
    }
  }
}
```

File: [./.yo-rc.json:2:3](./.yo-rc.json)

### FN022001 Scss file import | Required

Remove scss file import

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

<!-- File: [src\webparts\aboutus\components\Aboutus.module.scss](src\webparts\aboutus\components\Aboutus.module.scss) -->

### FN022001 Scss file import | Required

Remove scss file import

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

<!-- File: [src\webparts\holiday\components\Holiday.module.scss](src\webparts\holiday\components\Holiday.module.scss) -->

### FN022001 Scss file import | Required

Remove scss file import

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

<!-- File: [src\webparts\leaveMgmt\components\LeaveMgmt.module.scss](src\webparts\leaveMgmt\components\LeaveMgmt.module.scss) -->

### FN022001 Scss file import | Required

Remove scss file import

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

File: [src\webparts\leaveMgmtDashboard\components\LeaveMgmtDashboard.module.scss](src\webparts\leaveMgmtDashboard\components\LeaveMgmtDashboard.module.scss)

### FN022001 Scss file import | Required

Remove scss file import

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

<!-- File: [src\webparts\permissionDashboard\components\PermissionDashboard.module.scss](src\webparts\permissionDashboard\components\PermissionDashboard.module.scss) -->

### FN022001 Scss file import | Required

Remove scss file import

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

<!-- File: [src\webparts\permissionRequest\components\PermissionRequest.module.scss](src\webparts\permissionRequest\components\PermissionRequest.module.scss) -->

### FN022002 Scss file import | Optional

Add scss file import

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- File: [src\webparts\aboutus\components\Aboutus.module.scss](src\webparts\aboutus\components\Aboutus.module.scss) -->

### FN022002 Scss file import | Optional

Add scss file import

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- File: [src\webparts\holiday\components\Holiday.module.scss](src\webparts\holiday\components\Holiday.module.scss) -->

### FN022002 Scss file import | Optional

Add scss file import

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- File: [src\webparts\leaveMgmt\components\LeaveMgmt.module.scss](src\webparts\leaveMgmt\components\LeaveMgmt.module.scss) -->

### FN022002 Scss file import | Optional

Add scss file import

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

File: [src\webparts\leaveMgmtDashboard\components\LeaveMgmtDashboard.module.scss](src\webparts\leaveMgmtDashboard\components\LeaveMgmtDashboard.module.scss)

### FN022002 Scss file import | Optional

Add scss file import

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- File: [src\webparts\permissionDashboard\components\PermissionDashboard.module.scss](src\webparts\permissionDashboard\components\PermissionDashboard.module.scss) -->

### FN022002 Scss file import | Optional

Add scss file import

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- File: [src\webparts\permissionRequest\components\PermissionRequest.module.scss](src\webparts\permissionRequest\components\PermissionRequest.module.scss) -->

### FN012020 tsconfig.json noImplicitAny | Required

Add noImplicitAny in tsconfig.json

```json
{
  "compilerOptions": {
    "noImplicitAny": true
  }
}
```

File: [./tsconfig.json:3:22](./tsconfig.json)

### FN007001 serve.json schema | Required

Update serve.json schema URL

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json"
}
```

File: [./config/serve.json:2:3](./config/serve.json)

### FN001033 tslib | Required

Install SharePoint Framework dependency package tslib

Execute the following command:

```sh
npm i -SE tslib@2.3.1
```

File: [./package.json:14:3](./package.json)

### FN002007 ajv | Required

Upgrade SharePoint Framework dev dependency package ajv

Execute the following command:

```sh
npm i -DE ajv@6.12.5
```

File: [./package.json:59:5](./package.json)

### FN002009 @microsoft/sp-tslint-rules | Required

Remove SharePoint Framework dev dependency package @microsoft/sp-tslint-rules

Execute the following command:

```sh
npm un -D @microsoft/sp-tslint-rules
```

File: [./package.json:50:5](./package.json)

### FN002013 @types/webpack-env | Required

Upgrade SharePoint Framework dev dependency package @types/webpack-env

Execute the following command:

```sh
npm i -DE @types/webpack-env@1.15.2
```

File: [./package.json:58:5](./package.json)

### FN002021 @rushstack/eslint-config | Required

Install SharePoint Framework dev dependency package @rushstack/eslint-config

Execute the following command:

```sh
npm i -DE @rushstack/eslint-config@2.5.1
```

File: [./package.json:45:3](./package.json)

### FN002025 eslint-plugin-react-hooks | Required

Install SharePoint Framework dev dependency package eslint-plugin-react-hooks

Execute the following command:

```sh
npm i -DE eslint-plugin-react-hooks@4.3.0
```

File: [./package.json:45:3](./package.json)

### FN015003 tslint.json | Required

Remove file tslint.json

Execute the following command:

```sh
rm "tslint.json"
```

File: [tslint.json](tslint.json)

### FN015008 .eslintrc.js | Required

Add file .eslintrc.js

Execute the following command:

```sh
cat > ".eslintrc.js" << EOF 
require('@rushstack/eslint-config/patch/modern-module-resolution');
export default {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],
  parserOptions: { tsconfigRootDir: __dirname }
};
EOF
```

File: [.eslintrc.js](.eslintrc.js)

### FN023002 .gitignore '.heft' folder | Required

To .gitignore add the '.heft' folder


File: [./.gitignore](./.gitignore)

### FN006005 package-solution.json metadata | Required

In package-solution.json add metadata section

```json
{
  "solution": {
    "metadata": {
      "shortDescription": {
        "default": "leave-mgmt description"
      },
      "longDescription": {
        "default": "leave-mgmt description"
      },
      "screenshotPaths": [],
      "videoUrl": "",
      "categories": []
    }
  }
}
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN006006 package-solution.json features | Required

In package-solution.json add features for components

```json
// {
//   "solution": {
//     "features": [
//       {
//         "title": "leave-mgmt AboutusWebPart Feature",
//         "description": "The feature that activates AboutusWebPart from the leave-mgmt solution.",
//         "id": "7c3add12-7562-435d-b23a-7670af609df2",
//         "version": "3.0.0.3",
//         "componentIds": [
//           "7c3add12-7562-435d-b23a-7670af609df2"
//         ]
//       }
//     ]
//   }
// }
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN006006 package-solution.json features | Required

In package-solution.json add features for components

```json
// {
//   "solution": {
//     "features": [
//       {
//         "title": "leave-mgmt HolidayWebPart Feature",
//         "description": "The feature that activates HolidayWebPart from the leave-mgmt solution.",
//         "id": "a3e66a66-0ff7-4ba3-aec2-144056ebad6c",
//         "version": "3.0.0.3",
//         "componentIds": [
//           "a3e66a66-0ff7-4ba3-aec2-144056ebad6c"
//         ]
//       }
//     ]
//   }
// }
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN006006 package-solution.json features | Required

In package-solution.json add features for components

```json
// {
//   "solution": {
//     "features": [
//       {
//         "title": "leave-mgmt LeaveMgmtWebPart Feature",
//         "description": "The feature that activates LeaveMgmtWebPart from the leave-mgmt solution.",
//         "id": "7c2f80bb-b1eb-4872-a58e-aa44c468db7f",
//         "version": "3.0.0.3",
//         "componentIds": [
//           "7c2f80bb-b1eb-4872-a58e-aa44c468db7f"
//         ]
//       }
//     ]
//   }
// }
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN006006 package-solution.json features | Required

In package-solution.json add features for components

```json
{
  "solution": {
    "features": [
      {
        "title": "leave-mgmt LeaveMgmtDashboardWebPart Feature",
        "description": "The feature that activates LeaveMgmtDashboardWebPart from the leave-mgmt solution.",
        "id": "849fc01b-e6a9-4bee-96e0-a78db56e187a",
        "version": "3.0.0.3",
        "componentIds": [
          "849fc01b-e6a9-4bee-96e0-a78db56e187a"
        ]
      }
    ]
  }
}
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN006006 package-solution.json features | Required

In package-solution.json add features for components

```json
// {
//   "solution": {
//     "features": [
//       {
//         "title": "leave-mgmt PermissionDashboardWebPart Feature",
//         "description": "The feature that activates PermissionDashboardWebPart from the leave-mgmt solution.",
//         "id": "1176bbad-d357-4b2d-96ec-9524a6f012a1",
//         "version": "3.0.0.3",
//         "componentIds": [
//           "1176bbad-d357-4b2d-96ec-9524a6f012a1"
//         ]
//       }
//     ]
//   }
// }
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN006006 package-solution.json features | Required

In package-solution.json add features for components

```json
// {
//   "solution": {
//     "features": [
//       {
//         "title": "leave-mgmt PermissionRequestWebPart Feature",
//         "description": "The feature that activates PermissionRequestWebPart from the leave-mgmt solution.",
//         "id": "338332b6-9a90-4618-9fba-53e9ce815b6e",
//         "version": "3.0.0.3",
//         "componentIds": [
//           "338332b6-9a90-4618-9fba-53e9ce815b6e"
//         ]
//       }
//     ]
//   }
// }
```

File: [./config/package-solution.json:3:3](./config/package-solution.json)

### FN002003 @microsoft/sp-webpart-workbench | Required

Remove SharePoint Framework dev dependency package @microsoft/sp-webpart-workbench

Execute the following command:

```sh
npm un -D @microsoft/sp-webpart-workbench
```

File: [./package.json:51:5](./package.json)

### FN007003 serve.json api | Required

From serve.json remove the api property

```json

```

File: [./config/serve.json:6:3](./config/serve.json)

### FN015007 config\copy-assets.json | Required

Remove file config\copy-assets.json

Execute the following command:

```sh
rm "config\copy-assets.json"
```

File: [config\copy-assets.json](config\copy-assets.json)

### FN024001 Create .npmignore | Required

Create the .npmignore file


File: [./.npmignore](./.npmignore)

### FN005002 deploy-azure-storage.json workingDir | Required

Update deploy-azure-storage.json workingDir

```json
{
  "workingDir": "./release/assets/"
}
```

File: [./config/deploy-azure-storage.json:3:3](./config/deploy-azure-storage.json)

### FN023001 .gitignore 'release' folder | Required

To .gitignore add the 'release' folder


File: [./.gitignore](./.gitignore)

### FN002004 gulp | Required

Upgrade SharePoint Framework dev dependency package gulp

Execute the following command:

```sh
npm i -DE gulp@4.0.2
```

File: [./package.json:60:5](./package.json)

### FN002005 @types/chai | Required

Remove SharePoint Framework dev dependency package @types/chai

Execute the following command:

```sh
npm un -D @types/chai
```

File: [./package.json:52:5](./package.json)

### FN002006 @types/mocha | Required

Remove SharePoint Framework dev dependency package @types/mocha

Execute the following command:

```sh
npm un -D @types/mocha
```

File: [./package.json:55:5](./package.json)

### FN002014 @types/es6-promise | Required

Remove SharePoint Framework dev dependency package @types/es6-promise

Execute the following command:

```sh
npm un -D @types/es6-promise
```

File: [./package.json:54:5](./package.json)

### FN012013 tsconfig.json exclude property | Required

Remove tsconfig.json exclude property

```json
{
  "exclude": []
}
```

File: [./tsconfig.json:35:3](./tsconfig.json)

### FN012018 tsconfig.json es2015.promise lib | Required

Add es2015.promise lib in tsconfig.json

```json
{
  "compilerOptions": {
    "lib": [
      "es2015.promise"
    ]
  }
}
```

File: [./tsconfig.json:25:5](./tsconfig.json)

### FN012019 tsconfig.json es6-promise types | Required

Remove es6-promise type in tsconfig.json

```json
{
  "compilerOptions": {
    "types": [
      "es6-promise"
    ]
  }
}
```

File: [./tsconfig.json:22:7](./tsconfig.json)

### FN013002 gulpfile.js serve task | Required

Before 'build.initialize(require('gulp'));' add the serve task

```js
var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

```

File: [./gulpfile.js](./gulpfile.js)

### FN019002 tslint.json extends | Required

Update tslint.json extends property

```json
{
  "extends": "./node_modules/@microsoft/sp-tslint-rules/base-tslint.json"
}
```

File: [./tslint.json:2:3](./tslint.json)

### FN021002 engines | Required

Remove package.json property

```json
{
  "engines": "undefined"
}
```

File: [./package.json:6:3](./package.json)

### FN017001 Run npm dedupe | Optional

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm un -S office-ui-fabric-react
npm un -D @microsoft/sp-tslint-rules @microsoft/sp-webpart-workbench @types/chai @types/mocha @types/es6-promise
npm i -SE @microsoft/sp-core-library@1.18.2 @microsoft/sp-lodash-subset@1.18.2 @microsoft/sp-office-ui-fabric-core@1.18.2 @microsoft/sp-webpart-base@1.18.2 @microsoft/sp-property-pane@1.18.2 @microsoft/sp-adaptive-card-extension-base@1.18.2 @fluentui/react@8.106.4 react@17.0.1 react-dom@17.0.1 tslib@2.3.1
npm i -DE @microsoft/sp-build-web@1.18.2 @microsoft/sp-module-interfaces@1.18.2 @microsoft/eslint-plugin-spfx@1.18.2 @microsoft/eslint-config-spfx@1.18.2 typescript@4.7.4 @microsoft/rush-stack-compiler-4.7@0.1.0 eslint@8.7.0 @types/react@17.0.45 @types/react-dom@17.0.17 ajv@6.12.5 @types/webpack-env@1.15.2 @rushstack/eslint-config@2.5.1 eslint-plugin-react-hooks@4.3.0 gulp@4.0.2
npm dedupe
cat > "config\sass.json" << EOF 
{
  "$schema": "https://developer.microsoft.com/json-schemas/core-build/sass.schema.json"
}
EOF
rm "tslint.json"
cat > ".eslintrc.js" << EOF 
require('@rushstack/eslint-config/patch/modern-module-resolution');
export default {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],
  parserOptions: { tsconfigRootDir: __dirname }
};
EOF
rm "config\copy-assets.json"
```

### Modify files

#### [./.yo-rc.json](./.yo-rc.json)

Update version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "version": "1.18.2"
  }
}
```

Update @microsoft/teams-js SDK version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/teams-js": "2.12.0"
    }
  }
}
```

Update nodeVersion in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "nodeVersion": "18.19.0"
  }
}
```

Update @microsoft/microsoft-graph-client SDK version in .yo-rc.json:

```json
{
  "@microsoft/generator-sharepoint": {
    "sdkVersions": {
      "@microsoft/microsoft-graph-client": "3.0.2"
    }
  }
}
```

#### [./tsconfig.json](./tsconfig.json)

Update tsconfig.json extends property:

```json
{
  "extends": "./node_modules/@microsoft/rush-stack-compiler-4.7/includes/tsconfig-web.json"
}
```

Add noImplicitAny in tsconfig.json:

```json
{
  "compilerOptions": {
    "noImplicitAny": true
  }
}
```

Remove tsconfig.json exclude property:

```json
{
  "exclude": []
}
```

Add es2015.promise lib in tsconfig.json:

```json
{
  "compilerOptions": {
    "lib": [
      "es2015.promise"
    ]
  }
}
```

Remove es6-promise type in tsconfig.json:

```json
{
  "compilerOptions": {
    "types": [
      "es6-promise"
    ]
  }
}
```

#### [./package.json](./package.json)

Update package.json engines.node property:

```json
{
  "engines": {
    "node": ">=16.13.0 <17.0.0 || >=18.17.1 <19.0.0"
  }
}
```

Remove package.json property:

```json
{
  "engines": "undefined"
}
```

#### [./config/serve.json](./config/serve.json)

Update serve.json initialPage URL:

```json
{
  "initialPage": "https://{tenantDomain}/_layouts/workbench.aspx"
}
```

Update serve.json schema URL:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json"
}
```

From serve.json remove the api property:

```json

```

<!-- #### [src\webparts\aboutus\components\Aboutus.module.scss](src\webparts\aboutus\components\Aboutus.module.scss) -->

Remove scss file import:

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

Add scss file import:

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- #### [src\webparts\holiday\components\Holiday.module.scss](src\webparts\holiday\components\Holiday.module.scss) -->

Remove scss file import:

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

Add scss file import:

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- #### [src\webparts\leaveMgmt\components\LeaveMgmt.module.scss](src\webparts\leaveMgmt\components\LeaveMgmt.module.scss) -->

Remove scss file import:

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

Add scss file import:

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

#### [src\webparts\leaveMgmtDashboard\components\LeaveMgmtDashboard.module.scss](src\webparts\leaveMgmtDashboard\components\LeaveMgmtDashboard.module.scss)

Remove scss file import:

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

Add scss file import:

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- #### [src\webparts\permissionDashboard\components\PermissionDashboard.module.scss](src\webparts\permissionDashboard\components\PermissionDashboard.module.scss) -->

Remove scss file import:

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

Add scss file import:

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

<!-- #### [src\webparts\permissionRequest\components\PermissionRequest.module.scss](src\webparts\permissionRequest\components\PermissionRequest.module.scss) -->

Remove scss file import:

```scss
@import '~office-ui-fabric-react/dist/sass/References.scss'
```

Add scss file import:

```scss
@import '~@fluentui/react/dist/sass/References.scss'
```

#### [./.gitignore](./.gitignore)

To .gitignore add the '.heft' folder:

```text
.heft
```

To .gitignore add the 'release' folder:

```text
release
```

#### [./config/package-solution.json](./config/package-solution.json)

In package-solution.json add metadata section:

```json
{
  "solution": {
    "metadata": {
      "shortDescription": {
        "default": "leave-mgmt description"
      },
      "longDescription": {
        "default": "leave-mgmt description"
      },
      "screenshotPaths": [],
      "videoUrl": "",
      "categories": []
    }
  }
}
```

In package-solution.json add features for components:

```json
{
  "solution": {
    "features": [
      {
        "title": "leave-mgmt AboutusWebPart Feature",
        "description": "The feature that activates AboutusWebPart from the leave-mgmt solution.",
        "id": "7c3add12-7562-435d-b23a-7670af609df2",
        "version": "3.0.0.3",
        "componentIds": [
          "7c3add12-7562-435d-b23a-7670af609df2"
        ]
      }
    ]
  }
}
```

In package-solution.json add features for components:

```json
{
  "solution": {
    "features": [
      {
        "title": "leave-mgmt HolidayWebPart Feature",
        "description": "The feature that activates HolidayWebPart from the leave-mgmt solution.",
        "id": "a3e66a66-0ff7-4ba3-aec2-144056ebad6c",
        "version": "3.0.0.3",
        "componentIds": [
          "a3e66a66-0ff7-4ba3-aec2-144056ebad6c"
        ]
      }
    ]
  }
}
```

In package-solution.json add features for components:

```json
{
  "solution": {
    "features": [
      {
        "title": "leave-mgmt LeaveMgmtWebPart Feature",
        "description": "The feature that activates LeaveMgmtWebPart from the leave-mgmt solution.",
        "id": "7c2f80bb-b1eb-4872-a58e-aa44c468db7f",
        "version": "3.0.0.3",
        "componentIds": [
          "7c2f80bb-b1eb-4872-a58e-aa44c468db7f"
        ]
      }
    ]
  }
}
```

In package-solution.json add features for components:

```json
{
  "solution": {
    "features": [
      {
        "title": "leave-mgmt LeaveMgmtDashboardWebPart Feature",
        "description": "The feature that activates LeaveMgmtDashboardWebPart from the leave-mgmt solution.",
        "id": "849fc01b-e6a9-4bee-96e0-a78db56e187a",
        "version": "3.0.0.3",
        "componentIds": [
          "849fc01b-e6a9-4bee-96e0-a78db56e187a"
        ]
      }
    ]
  }
}
```

In package-solution.json add features for components:

```json
{
  "solution": {
    "features": [
      {
        "title": "leave-mgmt PermissionDashboardWebPart Feature",
        "description": "The feature that activates PermissionDashboardWebPart from the leave-mgmt solution.",
        "id": "1176bbad-d357-4b2d-96ec-9524a6f012a1",
        "version": "3.0.0.3",
        "componentIds": [
          "1176bbad-d357-4b2d-96ec-9524a6f012a1"
        ]
      }
    ]
  }
}
```

In package-solution.json add features for components:

```json
{
  "solution": {
    "features": [
      {
        "title": "leave-mgmt PermissionRequestWebPart Feature",
        "description": "The feature that activates PermissionRequestWebPart from the leave-mgmt solution.",
        "id": "338332b6-9a90-4618-9fba-53e9ce815b6e",
        "version": "3.0.0.3",
        "componentIds": [
          "338332b6-9a90-4618-9fba-53e9ce815b6e"
        ]
      }
    ]
  }
}
```

#### [./.npmignore](./.npmignore)

Create the .npmignore file:

```text
!dist
config

gulpfile.js

release
src
temp

tsconfig.json
tslint.json

*.log

.yo-rc.json
.vscode

```

#### [./config/deploy-azure-storage.json](./config/deploy-azure-storage.json)

Update deploy-azure-storage.json workingDir:

```json
{
  "workingDir": "./release/assets/"
}
```

#### [./gulpfile.js](./gulpfile.js)

Before 'build.initialize(require('gulp'));' add the serve task:

```js
var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

```

#### [./tslint.json](./tslint.json)

Update tslint.json extends property:

```json
{
  "extends": "./node_modules/@microsoft/sp-tslint-rules/base-tslint.json"
}
```
