# Link IFC and dump props to Excel App bundle for Autodesk APS Design Automation

[![Design Automation](https://img.shields.io/badge/Design%20Automation-v3-green.svg)](http://developer.autodesk.com/)

![Revit](https://img.shields.io/badge/Plugins-Revit-lightgrey.svg)
![.NET](https://img.shields.io/badge/.NET%20Framework-4.8-blue.svg)
[![Revit](https://img.shields.io/badge/Revit-2022-lightgrey.svg)](https://www.autodesk.com/products/revit/overview/)

![Advanced](https://img.shields.io/badge/Level-Advanced-red.svg)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](http://opensource.org/licenses/MIT)

# Description

This sample demonstrates the below on Design Automation:

- How to open IFC files using `Link IFC`  with the `Importer` of `Revit.IFC.Import.dll` from Revit software.
- How to dump IFC object properties by opening the `*.ifc.RVT` file and write properties result into an Excel file using [NPOI](https://www.nuget.org/packages/NPOI/).

# Development Setup

## Prerequisites

1. **APS Account**: Learn how to create a APS Account, activate subscription and create an app at [this tutorial](https://aps.autodesk.com/tutorials).
2. **Visual Studio 2022 and later** (Windows).
3. **Revit 2022 and later**: required to compile changes into the plugin.

## Design Automation Setup

### AppBundle example

```json
{
    "id": "RevitIfcLinkPropDumper",
    "engine": "Autodesk.Revit+2022",
    "description": "Link IFC and dump props to Excel"
}
```

### Activity example

```json
{
    "id": "RevitIfcLinkPropDumperActivity",
    "commandLine": [
        "$(engine.path)\\\\revitcoreconsole.exe /al \"$(appbundles[RevitIfcLinkPropDumper].path)\""
    ],
    "parameters": {
        "inputIFC": {
            "verb": "get",
            "description": "The IFC will be opened by `Link IFC`",
            "required": true,
            "localName": "$(inputIFC)"
        },
        "result": {
            "verb": "put",
            "description": "The result Excel file",
            "localName": "result.xlsx"
        }
    },
    "engine": "Autodesk.Revit+2022",
    "appbundles": [
        "Autodesk.RevitIfcLinkPropDumper+dev"
    ],
    "description": "Activity for linking IFC and dumping props to Excel"
}
```

### Workitem example

```json
{
    "activityId": "Autodesk.RevitIfcLinkPropDumperActivity+dev",
    "arguments": {
        "inputIFC": {
            "verb": "get",
            "url": "https://developer.api.autodesk.com/oss/v2/apptestbucket/97095bbc-1ce3-469f-99ba-0157bbcab73b?region=US"
        },
        "result": {
            "verb": "put",
            "url": "https://developer.api.autodesk.com/oss/v2/apptestbucket/9d3be632-a4fc-457d-bc5d-9e75cefc54b7?region=US"
        }
    }
}
```

## License

This sample is licensed under the terms of the [MIT License](http://opensource.org/licenses/MIT). Please see the [LICENSE](LICENSE) file for full details.

## Written by

Eason Kang [@yiskang](https://twitter.com/yiskang), [Autodesk Developer Advocacy and Support](http://aps.autodesk.com)