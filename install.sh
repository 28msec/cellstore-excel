#!/bin/bash

if [ -e "$APPDATA" -a -f "build/release/CellStore.Excel.xll" ]
then
  echo "build/release/CellStore.Excel.xll -> ${APPDATA}\Microsoft\AddIns\CellStore.Excel.xll"
  cp -f build/release/CellStore.Excel.xll "${APPDATA}\Microsoft\AddIns\CellStore.Excel.xll"
fi