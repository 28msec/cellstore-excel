#!/bin/bash

if [ -e "$APPDATA" -a -f "build/release/CellStore.Excel.x64.64.xll" ]
then
  echo "build/release/CellStore.Excel.x64.64.xll -> ${APPDATA}\Microsoft\AddIns\CellStore.Excel.x64.64.xll"
  cp -f build/release/CellStore.Excel.x64.64.xll "${APPDATA}\Microsoft\AddIns\CellStore.Excel.x64.64.xll"
else
  echo "Nothing installed"
fi