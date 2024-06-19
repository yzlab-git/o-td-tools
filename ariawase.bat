@echo off
chcp 65001

echo VBAファイルをdecombineします。
cd %~dp0 
cscript vbac.wsf decombine
echo 完了しました。
