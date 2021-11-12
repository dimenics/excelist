<p align="center"><img src="assets/spreadsheet.svg?raw=true" width="350" alt=""></p>

<h1 align="center"> Excelist </h1>

<div align="center">
<img src="https://dev.azure.com/dimesoftware/Utilities/_apis/build/status/dimenics.excelist?branchName=master" />
<img src="https://img.shields.io/azure-devops/coverage/dimesoftware/utilities/193?style=flat-square&color=blue" />
<img src="http://img.shields.io/:license-mit-blue.svg?style=flat-square">
<img src="https://img.shields.io/badge/PRs-welcome-blue.svg?style=flat-square" />
</div>

<br />
Simple library to convert `IEnumerable<T>` to an Excel sheet.

## About the project

The purpose of this project is to detach the implementation of Excel libraries from its contracts. This is particularly useful in an era where it is uncertain that a project will be maintained and will be migrated to .NET Core and .NET 5.

This project then does not introduce any new capabilities. It is merely a generic wrapper that allows you to inject dependencies into your code base and should give you peace of mind that your investments in your code are safe and should not be impacted if you decide to change Excel export libraries, for whatever reason that may be.

## Installation

Use the package manager NuGet to install the base library of Excelist:

`dotnet add package Excelist`

Next it is up to you to decide which _implementations_ you want to use:

| Implementation | Command                               |
| -------------- | ------------------------------------- |
| OpenXml        | `dotnet add package Excelist.OpenXml` |

To speed up the development cycle, there are extension methods at your disposal:

| Extension       | Command                            |
| --------------- | ---------------------------------- |
| System.Net.Http | `dotnet add package Excelist.Http` |

## Usage

The center of this project is the `IEnumerableToExcelConverter<in T>` interface. The implementations are hidden away through this interface, and as such, can be swapped effortlessly for another implementation.

For example, in a good old ASP.NET Web API project, you can use the `Excelist.Http` library to return an Excel file as a response:

```csharp
public class LogsApiController : ApiController
{
    [HttpGet]
    [Route("LogDumpFile")]
    public async t.Task<HttpResponseMessage> GetDump(int limit, string sort, string filter)
    {
        IPage<LogDto> records = await Service.GetAsync(1, limit, 1, sort, filter);
        return Request.ExportToExcel(records.Data, new OpenOfficeEnumerableToExcelConverter<LogDto>());
    }
}
```

## Contributing

![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?style=flat-square)

Pull requests are welcome. Please check out the contribution and code of conduct guidelines.
