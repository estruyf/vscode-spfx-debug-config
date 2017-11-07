import * as vscode from 'vscode';
const decomment = require('decomment');

/**
 * Configuration constants
 */
const url = "https://localhost:4321/temp/manifests.js";
const location = "https://contoso.sharepoint.com/documents";
const extensionConfig = {
    "name": "",
    "type": "chrome",
    "request": "launch",
    "url": "",
    "webRoot": "${workspaceRoot}",
    "sourceMaps": true,
    "sourceMapPathOverrides": {
        "webpack:///../../../src/*": "${webRoot}/src/*",
        "webpack:///../../../../src/*": "${webRoot}/src/*",
        "webpack:///../../../../../src/*": "${webRoot}/src/*"
    },
    "runtimeArgs": [
        "--remote-debugging-port=9222"
    ]
};

// Activate the debug extension provider
export function activate(context: vscode.ExtensionContext) {
    // register a configuration provider
    context.subscriptions.push(vscode.debug.registerDebugConfigurationProvider('SPFx', new SPFxConfigurationProvider()));
}

export function deactivate() {
    // nothing to do
}

/**
 * SPFx debug configuration provider
 */
class SPFxConfigurationProvider implements vscode.DebugConfigurationProvider {
    provideDebugConfigurations(folder: vscode.WorkspaceFolder | undefined, token?: vscode.CancellationToken): vscode.ProviderResult<vscode.DebugConfiguration[]> {
        return createLaunchConfigFromContext(folder, false);
    }
}

/**
 * Create launch configuration
 * @param folder 
 * @param resolve 
 */
function createLaunchConfigFromContext(folder: vscode.WorkspaceFolder | undefined, resolve: boolean): Promise<vscode.DebugConfiguration[]> {
    const config: vscode.DebugConfiguration[] = [{
        "name": "Local workbench",
        "type": "chrome",
        "request": "launch",
        "url": "https://localhost:4321/temp/workbench.html",
        "webRoot": "${workspaceRoot}",
        "sourceMaps": true,
        "sourceMapPathOverrides": {
            "webpack:///../../../src/*": "${webRoot}/src/*",
            "webpack:///../../../../src/*": "${webRoot}/src/*",
            "webpack:///../../../../../src/*": "${webRoot}/src/*"
        },
        "runtimeArgs": [
            "--remote-debugging-port=9222"
        ]
    },
    {
        "name": "Hosted workbench",
        "type": "chrome",
        "request": "launch",
        "url": "https://contoso.sharepoint.com/_layouts/workbench.aspx",
        "webRoot": "${workspaceRoot}",
        "sourceMaps": true,
        "sourceMapPathOverrides": {
            "webpack:///../../../src/*": "${webRoot}/src/*",
            "webpack:///../../../../src/*": "${webRoot}/src/*",
            "webpack:///../../../../../src/*": "${webRoot}/src/*"
        },
        "runtimeArgs": [
            "--remote-debugging-port=9222"
        ]
    }];

    return getExtensionConfig(config).then(data => {
        return data;
    });
}

/**
 * Retrieve all the SPFx extensions
 */
function getExtensionConfig(currentConfig: vscode.DebugConfiguration[]): Promise<vscode.DebugConfiguration[]> {
    return new Promise<vscode.DebugConfiguration[]>((resolve, reject) => {
        try {
            // Retrieve all manifest files for the extensions
            const manifestFiles = vscode.workspace.findFiles('**/src/**/*.manifest.json').then(async (data: vscode.Uri[]) => {
                // Check if URIs were retrieved
                if (data.length > 0) {
                    // Remove duplicates
                    const files = data.filter((elm, pos, arr) => {
                        return arr.indexOf(elm) === pos;
                    });

                    // Create Promises
                    const proms = files.map(file => readFileContent(file));
                    Promise.all(proms).then((data) => {
                        // Check if data was retrieved
                        if (data.length > 0) {
                            console.log(`SPFx Debugger found ${data.length} extension manifest files.`);
                            // Loop over all the manifest files its  content
                            data.forEach(manifest => {
                                // Check if the manifest file is not null
                                if (manifest !== null) {
                                    // Check the manifest extension type and create a configuration per found type
                                    switch (manifest.extensionType) {
                                        case "ApplicationCustomizer":
                                            const appConfig = Object.assign({}, extensionConfig);
                                            appConfig.name = `Debug ${manifest.alias}`;
                                            appConfig.url = `${location}?loadSPFX=true&debugManifestsFile=${url}&customActions={"${manifest.id}":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"prop1":"val1"}}}`;
                                            currentConfig.push(appConfig);
                                            break;
                                        case "FieldCustomizer":
                                            const fieldConfig = Object.assign({}, extensionConfig);
                                            fieldConfig.name = `Debug ${manifest.alias}`;
                                            fieldConfig.url = `${location}?loadSPFX=true&debugManifestsFile=${url}&fieldCustomizers={"FieldName":{"id":"${manifest.id}","properties":{"prop1":"val1"}}}`;
                                            currentConfig.push(fieldConfig);
                                            break;
                                        case "ListViewCommandSet":
                                            const listviewConfig = Object.assign({}, extensionConfig)
                                            listviewConfig.name = `Debug ${manifest.alias}`;
                                            listviewConfig.url = `${location}?loadSPFX=true&debugManifestsFile=${url}&customActions={"${manifest.id}":{"location":"ClientSideExtension.ListViewCommandSet.CommandBar"}}`;
                                            currentConfig.push(listviewConfig);
                                            break;
                                    }
                                }
                            });
                        } else {
                            console.log(`SPFx Debugger did not find any extension manifest files.`);
                        }
                        resolve(currentConfig);
                    });
                } else {
                    resolve(currentConfig);
                }
            });
        } catch (e) {
            console.error("ERROR:", e);
            resolve(currentConfig);
        }
    });
}

/**
 * Fetch the manifest file to get its content
 * @param fileUri Manifest file URI to parse
 */
const readFileContent = (fileUri: vscode.Uri) => {
    return vscode.workspace.openTextDocument(fileUri).then((file: vscode.TextDocument) => {
        const content = file.getText();
        if (content !== null) {
            // Decomment the content
            let jsonContent = decomment(content);
            jsonContent = JSON.parse(jsonContent);
            if (typeof jsonContent.componentType !== "undefined" &&
                jsonContent["componentType"] === "Extension") {
                return jsonContent;
            }
        }
        return null;
    });
};