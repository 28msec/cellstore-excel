#!/bin/bash

if [ -e "$APPDATA" -a -f "build/release/CellStore.Excel.x86.32.xll" ]
then
  echo "build/release/CellStore.Excel.x86.32.xll -> ${APPDATA}\Microsoft\AddIns\CellStore.Excel.x86.32.xll"
  cp -f build/release/CellStore.Excel.x86.32.xll "${APPDATA}\Microsoft\AddIns\CellStore.Excel.x86.32.xll"
else
  echo "Nothing installed"
fi