<!--
Copyright (c) 2025 Snowflake Inc. All rights reserved.
Licensed under the MIT License.
-->

# Snowflake Semantic View Excel Add-in

**Author:** Yogitha Chilukuri
**Version:** 1.0.0
**Date:** December 17, 2025

---

## Overview

This Excel Add-in enables business analysts to securely connect to **Snowflake Semantic Views**, ingest data into Excel, and perform ad-hoc analysis using native Excel features (Pivot Tables, Charts). Unlike traditional ODBC connectors, this solution understands **Semantic Metadata**—ensuring users cannot perform incorrect math (e.g., summing non-additive metrics).

### Key Features

- ✅ **Multi-method Authentication**: OAuth 2.0, Key-Pair, Personal Access Tokens (PAT)
- ✅ **Semantic View Discovery**: Browse databases, schemas, and semantic views with full metadata
- ✅ **Intelligent Pivot Builder**: Drag-and-drop interface with automatic dimension/metric placement
- ✅ **Semantic Validation**: Prevents invalid dimension-metric combinations
- ✅ **AI-Powered Queries**: Natural language queries via Snowflake Cortex Analyst
- ✅ **Pivot Tables & Charts**: Create visualizations directly from semantic queries
- ✅ **Connection Persistence**: Save connection metadata, share workbooks with colleagues
- ✅ **Performance Optimized**: Pagination, chunking, and query optimization for large datasets

### Tech Stack

- **Frontend**: Vanilla JavaScript + HTML + CSS (no frameworks)
- **Backend**: Node.js (Express)
- **Excel Integration**: Office.js (Office Add-ins API)
- **Authentication**: OAuth 2.0, Key-Pair, PAT
- **AI Integration**: Snowflake Cortex Analyst API
- **Build Tool**: Webpack

## How to run this sample

### Prerequisites

- Download and install [Visual Studio Code](https://visualstudio.microsoft.com/downloads/).
- Install the latest version of the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=msoffice.microsoft-office-add-in-debugger) into Visual Studio Code.
- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Microsoft Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).
  
### Run the add-in from the Office Add-ins Development Kit

1. Create a new project with the sample code.

   Open the Office Add-ins Development Kit extension and view samples in the **Sample gallery**. Select the **Create** button in the top-right corner of the sample page. The new project will open in a second Visual Studio Code window. Close the original VSC window.
   
1. Open the Office Add-ins Development Kit.
    
    Select the <img src="./assets/Icon_Office_Add-ins_Development_Kit.png" width="30" alt="The Office Add-ins Development Kit icon in the activity bar of VSCode"/> icon in the **Activity Bar** to open the extension.

1. Preview Your Office Add-in (F5).

    Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code. In the drop down menu, select the option **Desktop (Edge Chromium)**.

    <img src="./assets/devkit_preview.png" width="500" alt="The 'Preview your Office Add-in' option in the Office Add-ins Development Kit's task pane."/>

    The extension checks that the prerequisites are met before debugging starts. The terminal will alert you to any issues with your environment. After this process, the Excel desktop application launches and opens a new workbook with the sample add-in sideloaded. The add-in automatically opens as well.

    If this is the first time that you have sideloaded an Office Add-in on your computer (or the first time in over a month), you may be prompted to delete an old certificate and/or to install a new one. Agree to both prompts. The first run requires installing dependency of this project, which might take 2~3 minutes or longer. During this time, there might be a dialog pop up at the lower right of the VSC screen. You should not interact with this dialog before the Office application launched.

1. Stop Previewing Your Office Add-in.

    Once you are finished testing and debugging the add-in, select the <img src="./assets/Icon_Office_Add-ins_Development_Kit.png" width="30" alt="The Office Add-ins Development Kit icon in the activity bar of VSCode"/> icon and then select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.

## Use the sample add-in

1. Select a cell in the current worksheet.
2. Select `Run` in the task pane.

## Explore sample files

These are the important files in the sample project.

```
| .eslintrc.json
| .gitignore
| .vscode/
|   | extensions.json
|   | launch.json               Launch and debug configurations
|   | settings.json             
|   | tasks.json                
| assets/                       Static assets, such as images
| babel.config.json
| manifest.xml                  Manifest file
| package.json                  
| README.md                     
| RUN_WITH_EXTENSION.md         
| src/                          Add-in source code
|   | taskpane/
|   |   | taskpane.css          Task pane style
|   |   | taskpane.html         Task pane entry HTML
|   |   | taskpane.js           Office JS API calls and add-in logic
| webpack.config.js             Webpack config
```

## Troubleshooting

If you have problems running the sample, take the following steps.

- Close any open instances of Excel.
- Close the previous web server started for the sample with the **Stop Previewing Your Office Add-in** Office Add-ins Development Kit extension option.
- Try running again.

If you still have problems, see [troubleshoot development errors](https://learn.microsoft.com//office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.  

For information on running the sample on Excel on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

For information on debugging on older versions of Office, see [Debug add-ins using developer tools in Microsoft Edge Legacy](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy).

## Make code changes

Once you understand the sample, make it your own! All the information about Office Add-ins is found in our [official documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). You can also explore more samples in the Office Add-ins Development Kit. Select **View Samples** to see more samples of real-world scenarios.

If you edit the manifest as part of your changes, use the **Validate Manifest File** option in the Office Add-ins Development Kit. This shows you any errors in the manifest syntax.

## Engage with the team

Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.

Want to learn more about new features and best practices for the Office platform? [Join the Microsoft Office Add-ins community call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call).

## License

Copyright (c) 2025 Snowflake Inc. All rights reserved.

This project is licensed under the MIT License. See the LICENSE file for details.
