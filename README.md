## CellStore Excel Add-in
[![Circle CI](https://circleci.com/gh/28msec/cellstore-excel.svg?style=svg)](https://circleci.com/gh/28msec/cellstore-excel)

An Excel Add-in with User Defined Functions to access the Cell Store.

## Development

Prerequisite: [NPM](Prerequisites: NPM)

```bash
$ npm install gulp -g
$ npm install
$ gulp install-dependencies
```

Open src/CellStore.Excel/CellStore.Excel.sln to develop with Visual Studio 2015.

## Install

Close Excel before installing. Then:

```bat
C:\> copy "build\release\CellStore.Excel.xll" "%APPDATA%\Microsoft\AddIns\CellStore.Excel.xll" /y
C:\> start /d "%APPDATA%\Microsoft\AddIns" CellStore.Excel.xll
```
