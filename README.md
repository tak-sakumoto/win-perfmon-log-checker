# win-perfmon-log-checker

## Introduction

This repository provides a tool to check the result of Windows Performance Monitor (perfmon.exe).
The tool has the following features.

- Convert the output data (blg files) of perfmon.exe into csv files for each counter.
- Extract data by specifying start and end times.
- Get statistics for each counter.
- Plot the measured values for each counter on a line graph using Excel workbook.

## Usage

### Command

```plaintxt
> .\main.ps1 -blgPath \path\to\perfmon-result.blg -startTime "2023/10/09 12:00:00" -endTime "2023/10/10 00:00:00"
```

### Arguments

| Argument | Required | Default | Explanation |
|-|:-:|-|-|
| `-blgPath <path>` | o | - | A path to a perfmon result file (.blg) |
| `-outDirPath <path>` | x | `".\out"` | A path to an output folder |
| `-startTime "yyyy/MM/dd hh:mm:ss"` | x | `""` | A start time by the format "yyyy/MM/dd hh:mm:ss" |
| `-endTime "yyyy/MM/dd hh:mm:ss"` | x | `""` | An end time by the format "yyyy/MM/dd hh:mm:ss" |

## Author

[Takuya Sakumoto (作元 卓也)](https://github.com/tak-sakumoto)
