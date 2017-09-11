"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const decomment = require('decomment');
function activate(context) {
    console.log('SPFx Debugger is active!');
    // SPFx constants
    const url = "https://localhost:4321/temp/manifests.js";
    const location = "https://contoso.sharepoint.com/documents";
    // Intial debug configuration
    const initialConfigurations = {
        "version": "0.2.0",
        "configurations": [
            {
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
            }
        ]
    };
    // Extension config
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
    // Add a new command
    context.subscriptions.push(vscode.commands.registerCommand('extension.spfx-debug.provideInitialConfigurations', () => {
        return new Promise((resolve, reject) => {
            try {
                // Retrieve all manifest files for the extensions
                const manifestFiles = vscode.workspace.findFiles('**/src/**/*.manifest.json').then((data) => __awaiter(this, void 0, void 0, function* () {
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
                                                const appConfig = JSON.parse(JSON.stringify(extensionConfig));
                                                appConfig.name = `Debug ${manifest.alias}`;
                                                appConfig.url = `${location}?loadSPFX=true&debugManifestsFile=${url}&customActions={"${manifest.id}":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"prop1":"val1"}}}`;
                                                initialConfigurations.configurations.push(appConfig);
                                                break;
                                            case "FieldCustomizer":
                                                const fieldConfig = JSON.parse(JSON.stringify(extensionConfig));
                                                ;
                                                fieldConfig.name = `Debug ${manifest.alias}`;
                                                fieldConfig.url = `${location}?loadSPFX=true&debugManifestsFile=${url}&fieldCustomizers={"FieldName":{"id":"${manifest.id}","properties":{"prop1":"val1"}}}`;
                                                initialConfigurations.configurations.push(fieldConfig);
                                                break;
                                            case "ListViewCommandSet":
                                                const listviewConfig = JSON.parse(JSON.stringify(extensionConfig));
                                                ;
                                                listviewConfig.name = `Debug ${manifest.alias}`;
                                                listviewConfig.url = `${location}?loadSPFX=true&debugManifestsFile=${url}&customActions={"${manifest.id}":{"location":"ClientSideExtension.ListViewCommandSet.CommandBar"}}`;
                                                initialConfigurations.configurations.push(listviewConfig);
                                                break;
                                        }
                                    }
                                });
                            }
                            else {
                                console.log(`SPFx Debugger did not find any extension manifest files.`);
                            }
                            resolve(JSON.stringify(initialConfigurations, null, 2));
                        });
                    }
                    else {
                        resolve(JSON.stringify(initialConfigurations, null, 2));
                    }
                }));
            }
            catch (e) {
                console.error("ERROR:", e);
                resolve(JSON.stringify(initialConfigurations, null, 2));
            }
        });
    }));
}
exports.activate = activate;
/**
 * Fetch the manifest file to get its content
 * @param fileUri Manifest file URI to parse
 */
const readFileContent = (fileUri) => {
    return vscode.workspace.openTextDocument(fileUri).then((file) => {
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
function deactivate() {
    // nothing to do
}
exports.deactivate = deactivate;
//# sourceMappingURL=checkConfig.js.map